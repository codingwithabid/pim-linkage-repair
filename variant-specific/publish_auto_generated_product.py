#!/usr/bin/env python3
"""
publish_pending_products.py
===========================

Fourth and final script in the variant-repair toolchain. Runs AFTER:
    repair_variants.py            (produces product_states_pre_repair.json)
    push_variants.py              (PUTs variants — may create product drafts)
    repair_archived_variants.py   (the archived-variant flow)

Why this exists
---------------
When the previous scripts PUT variant payloads or do state transitions,
the PIM platform automatically creates a DRAFT revision on each touched
product. If the product was PUBLISHED before the run started, that draft
needs to be promoted to live so the variant changes actually go out.
This script finds those products and publishes them.

Filter
------
From product_states_pre_repair.json, take only products where
    state == "PUBLISHED"
(DRAFT-state products are left alone — they were already drafts before
 the run, so the new draft revision is the right resting state.)

For each PUBLISHED product:

    GET /pim/products/<id>?view=LATEST_INCLUDE_DRAFT&ignoreIfError=1
        → 200 on any site  → draft exists → queue for COMMIT
        → 404 on every site → no draft     → no-op (nothing to publish)

Site fallback
-------------
Same multi-site retry as the variant scripts:
    primary site → fallback site #1 → fallback site #2 → ...
The first site that returns 200 is the "owning site" — that's where
COMMIT gets sent for this product.

Publish flow
------------
Single batched COMMIT per site (READY is removed per operator direction):
    POST /pim/batch/products/change-status  {"ids":[...], "status":"COMMIT"}

A COMMIT failure leaves the product in DRAFT state. No READY gate, no
rollback — manual fix or re-run is the recovery story.

Usage
-----
    # Dry run — see counts without touching the API
    python3 publish_pending_products.py --input ./out --dry-run

    # Just check phase (which products have drafts) — useful preview
    python3 publish_pending_products.py --input ./out --stop-after check

    # Full run
    export API_COOKIE='...'
    python3 publish_pending_products.py --input ./out --workers 8
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import threading
import time
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any

import requests


# ---------------------------------------------------------------------------
# Site config — same as the other scripts
# ---------------------------------------------------------------------------
PRIMARY_SITE_ID = "201cb789-4198-488b-a5eb-4e7df0fb4bee"
FALLBACK_SITE_IDS = [
    "8d3ea3bc-f65b-4227-9fa6-6fae40e4575a",
    "fbfcd92b-d271-4002-9163-4f84986b41be",
]


# ---------------------------------------------------------------------------
# Logging — same three-file pattern
# ---------------------------------------------------------------------------
LOG_FMT = "%(asctime)s | %(levelname)-7s | %(name)-14s | %(message)s"
LOG_DATEFMT = "%Y-%m-%d %H:%M:%S"


def setup_logging(log_dir: Path, verbose: bool = False) -> logging.Logger:
    log_dir.mkdir(parents=True, exist_ok=True)
    formatter = logging.Formatter(LOG_FMT, datefmt=LOG_DATEFMT)

    def _rotating(path: Path, level: int) -> RotatingFileHandler:
        h = RotatingFileHandler(
            path, maxBytes=10 * 1024 * 1024, backupCount=5, encoding="utf-8"
        )
        h.setLevel(level)
        h.setFormatter(formatter)
        return h

    main = logging.getLogger("products")
    main.setLevel(logging.DEBUG if verbose else logging.INFO)
    main.propagate = False
    main.handlers.clear()
    main.addHandler(_rotating(log_dir / "products.run.log", logging.DEBUG if verbose else logging.INFO))
    console = logging.StreamHandler(sys.stderr)
    console.setLevel(logging.INFO)
    console.setFormatter(formatter)
    main.addHandler(console)

    ok = logging.getLogger("products.ok")
    ok.setLevel(logging.INFO)
    ok.propagate = False
    ok.handlers.clear()
    ok.addHandler(_rotating(log_dir / "products.success.log", logging.INFO))

    err = logging.getLogger("products.err")
    err.setLevel(logging.WARNING)
    err.propagate = True
    err.handlers.clear()
    err.addHandler(_rotating(log_dir / "products.errors.log", logging.WARNING))

    main.info("products-publish logging initialised: dir=%s verbose=%s", log_dir, verbose)
    return main


# ---------------------------------------------------------------------------
# HTTP session
# ---------------------------------------------------------------------------

def make_session(
    session_id: str,
    cookie_header: str | None = None,
) -> requests.Session:
    """
    Session shared across workers. X-UpStart-Site goes per-request — same
    rationale as the variant scripts (multi-site retries, thread safety).
    """
    s = requests.Session()
    headers = {
        "X-Upstart-Tenant": "denvermattress",
        "Referer": "https://denvermattress.nochannel-test-1.nochannel-test.upstart.team/",
        "X-Upstart-Session-Id": session_id,
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
    }
    if cookie_header:
        headers["Cookie"] = cookie_header
    s.headers.update(headers)
    return s


# ---------------------------------------------------------------------------
# Step 1 — Draft existence check
# ---------------------------------------------------------------------------

def check_product_has_draft(
    session: requests.Session,
    api_host: str,
    product_id: str,
    site_ids: list[str],
    timeout: float = 30.0,
    delay_ms: int = 0,
) -> tuple[bool, str | None, str | None]:
    """
    GET /pim/products/<id>?view=LATEST_INCLUDE_DRAFT&ignoreIfError=1
    across the site list. Stop at the first site that returns 200.

    Contract (per user direction):
      200 on any site   → draft EXISTS    → return (True, site, None)
      404 on every site → draft does NOT exist → return (False, None, None)
      Other errors      → return (False, None, error_summary)
                          and log; don't publish on uncertain signals.

    Per-attempt failures are at DEBUG level so the run log stays quiet
    when a product simply has no draft (the expected common case). The
    final outcome is recorded by the caller.

    `delay_ms` adds a sleep after EACH HTTP request (including 404 retries
    across sites), so multi-site fallback also gets paced. With workers > 1
    the delay is per-worker.
    """
    main_log = logging.getLogger("products")
    err_log = logging.getLogger("products.err")

    url = f"{api_host.rstrip('/')}/pim/products/{product_id}"
    params = {"view": "LATEST_INCLUDE_DRAFT", "ignoreIfError": "1"}

    last_unexpected_err: str | None = None

    def _sleep():
        if delay_ms > 0:
            time.sleep(delay_ms / 1000.0)

    for site_id in site_ids:
        try:
            r = session.get(
                url,
                params=params,
                headers={"X-UpStart-Site": site_id},
                timeout=timeout,
            )
        except requests.RequestException as e:
            # Network/timeout/DNS — log at DEBUG, try next site.
            main_log.debug(
                "check_network_error productId=%s siteId=%s detail=%r",
                product_id, site_id, str(e),
            )
            last_unexpected_err = f"network: {e}"
            _sleep()
            continue

        if r.status_code == 200:
            # Best case: draft exists on this site. The body itself is fine
            # to ignore for this script's purposes — we don't need contents,
            # only the existence signal.
            main_log.debug("check_draft_found productId=%s siteId=%s", product_id, site_id)
            _sleep()
            return True, site_id, None

        if r.status_code == 404:
            # Expected miss — try next site.
            main_log.debug("check_no_draft productId=%s siteId=%s", product_id, site_id)
            _sleep()
            continue

        if 200 < r.status_code < 300:
            # 2xx but not 200. Treat conservatively — refuse to act.
            err_log.warning(
                "unexpected_2xx productId=%s siteId=%s status=%d body=%r",
                product_id, site_id, r.status_code, r.text[:200],
            )
            last_unexpected_err = f"http {r.status_code}"
            _sleep()
            continue

        if r.status_code in (401, 403):
            # Auth issue — almost certainly the missing-Cookie problem.
            err_log.error(
                "check_auth_error productId=%s siteId=%s status=%d "
                "HINT='Probably missing/expired session cookie. Pass --cookie "
                "or set API_COOKIE env var.'",
                product_id, site_id, r.status_code,
            )
            last_unexpected_err = f"auth http {r.status_code}"
            _sleep()
            # Don't keep trying other sites — auth won't get better.
            return False, None, last_unexpected_err

        # 5xx or anything else unexpected: log, try next site
        last_unexpected_err = f"http {r.status_code}"
        main_log.debug(
            "check_unexpected_status productId=%s siteId=%s status=%d body=%r",
            product_id, site_id, r.status_code, r.text[:200],
        )
        _sleep()

    # Walked every site without a 200.
    if last_unexpected_err is None:
        # All sites returned 404 — definitive "no draft" answer.
        return False, None, None
    # Otherwise we hit something weird (5xx, network); refuse to act and
    # surface the issue so the operator can investigate.
    err_log.warning(
        "check_failed productId=%s tried_sites=%s last_error=%s",
        product_id, site_ids, last_unexpected_err,
    )
    return False, None, last_unexpected_err


# ---------------------------------------------------------------------------
# Step 2 — Batched COMMIT (no READY)
# ---------------------------------------------------------------------------

def chunked(xs: list[str], n: int) -> list[list[str]]:
    if n <= 0:
        return [xs]
    return [xs[i:i + n] for i in range(0, len(xs), n)]


def batch_change_status(
    session: requests.Session,
    api_host: str,
    site_id: str,
    product_ids: list[str],
    new_status: str,  # currently always "COMMIT" (READY removed)
    timeout: float = 60.0,
    dry_run: bool = False,
    delay_ms: int = 0,
) -> tuple[bool, str]:
    """
    POST /pim/batch/products/change-status
    body: {"ids": [...], "status": "COMMIT"}

    Same atomicity assumption as the variant batch endpoint: a 2xx means
    the whole batch applied; anything else means none of it did. If that
    turns out not to be true in practice, split-on-error retry will be
    needed — flag if you see partial behavior.

    `delay_ms` sleeps after the request completes — see check_product_has_draft.
    """
    url = f"{api_host.rstrip('/')}/pim/batch/products/change-status"
    if dry_run:
        return True, (
            f"DRY_RUN POST {url} site={site_id} status={new_status} "
            f"ids_n={len(product_ids)}"
        )

    try:
        r = session.post(
            url,
            headers={"X-UpStart-Site": site_id},
            json={"ids": product_ids, "status": new_status},
            timeout=timeout,
        )
    except requests.RequestException as e:
        if delay_ms > 0:
            time.sleep(delay_ms / 1000.0)
        return False, f"network: {e}"

    if delay_ms > 0:
        time.sleep(delay_ms / 1000.0)

    if 200 <= r.status_code < 300:
        return True, f"http {r.status_code}"
    return False, f"http {r.status_code} body={r.text[:300]!r}"


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument(
        "--input", required=True, type=Path,
        help="Output dir from repair_variants.py (contains "
             "product_states_pre_repair.json)",
    )
    ap.add_argument(
        "--api-host",
        default="https://nochannel-test-1-api.nochannel-test.upstart.team",
        help="API host (no trailing slash)",
    )
    ap.add_argument(
        "--site-id", default=PRIMARY_SITE_ID,
        help="Primary site to try first (default: PRIMARY_SITE_ID constant)",
    )
    ap.add_argument(
        "--fallback-site-ids",
        default=",".join(FALLBACK_SITE_IDS),
        help="Comma-separated fallback sites. Pass empty string to disable.",
    )
    ap.add_argument(
        "--session-id",
        default=os.environ.get("API_SESSION_ID", "example"),
        help="X-Upstart-Session-Id (or API_SESSION_ID env)",
    )
    ap.add_argument(
        "--cookie",
        default=os.environ.get("API_COOKIE", ""),
        help="Raw Cookie header (or API_COOKIE env)",
    )
    ap.add_argument(
        "--workers", type=int, default=1,
        help="Concurrent check + batch workers (default 1)",
    )
    ap.add_argument(
        "--batch-size", type=int, default=1,
        help="Max products per COMMIT batch (default 1)",
    )
    ap.add_argument(
        "--delay-ms", type=int, default=100,
        help="Sleep this many milliseconds after each API request (per "
             "worker), to mimic portal-paced traffic. Default 100. Set to "
             "0 to disable.",
    )
    ap.add_argument(
        "--limit", type=int, default=0,
        help="Process only first N candidates (0=all). For testing.",
    )
    ap.add_argument(
        "--dry-run", action="store_true",
        help="Log what would happen; no HTTP requests made.",
    )
    ap.add_argument(
        "--stop-after",
        choices=["scan", "check", "publish"],
        default="publish",
        help="Stop the pipeline after the named phase. "
             "'check' is useful to preview which products have drafts.",
    )
    ap.add_argument(
        "--skip-publish", action="store_true",
        help="Run the check phase only — skip COMMIT. "
             "Useful for previewing without committing.",
    )
    ap.add_argument(
        "--log-dir", type=Path, default=None,
        help="Where logs go. Default: <input>/logs",
    )
    ap.add_argument(
        "--verbose", "-v", action="store_true",
        help="Enable DEBUG-level logging",
    )
    args = ap.parse_args()

    if not args.input.exists() or not args.input.is_dir():
        print(f"Input dir not found or not a dir: {args.input}", file=sys.stderr)
        return 1

    log_dir = args.log_dir or (args.input / "logs")
    log = setup_logging(log_dir, verbose=args.verbose)
    run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    log.info(
        "run_start run_id=%s input=%s dry_run=%s stop_after=%s "
        "workers=%d batch_size=%d delay_ms=%d",
        run_id, args.input, args.dry_run, args.stop_after,
        args.workers, args.batch_size, args.delay_ms,
    )

    # --- load product_states_pre_repair.json ---
    pstate_path = args.input / "product_states_pre_repair.json"
    if not pstate_path.exists():
        log.error("product_states_pre_repair.json not found at %s", pstate_path)
        return 1
    product_states = json.loads(pstate_path.read_text())
    log.info("loaded product_states_pre_repair.json with %d products", len(product_states))

    # --- filter to state=PUBLISHED only ---
    log.info("[1/3] scan: filtering to PUBLISHED products")
    state_counts = Counter(v.get("state") for v in product_states.values())
    candidates: list[str] = [
        pid for pid, meta in product_states.items()
        if meta.get("state") == "PUBLISHED"
    ]
    log.info(
        "state breakdown: PUBLISHED=%d (candidates) DRAFT=%d (skipped) null=%d (skipped) "
        "other=%d",
        state_counts.get("PUBLISHED", 0),
        state_counts.get("DRAFT", 0),
        state_counts.get(None, 0),
        sum(n for k, n in state_counts.items() if k not in {"PUBLISHED", "DRAFT", None}),
    )

    if args.limit > 0:
        candidates = candidates[: args.limit]
        log.info("--limit %d applied; processing first %d", args.limit, len(candidates))

    if not candidates:
        log.info("no PUBLISHED candidates. run_id=%s", run_id)
        return 0

    if args.stop_after == "scan":
        log.info("stop_after=scan: exiting before any API calls")
        return 0

    # --- build site list ---
    raw_fallbacks = [s.strip() for s in (args.fallback_site_ids or "").split(",")]
    site_ids: list[str] = []
    for s in [args.site_id, *raw_fallbacks]:
        if s and s not in site_ids:
            site_ids.append(s)
    if not site_ids:
        log.error("no site IDs configured")
        return 1
    log.info("siteIds (in retry order): %s", site_ids)

    if not args.cookie and not args.dry_run:
        log.warning(
            "no Cookie provided — if the gateway requires session cookies "
            "you'll get HTML responses. Pass --cookie or API_COOKIE env."
        )

    session = make_session(args.session_id, cookie_header=args.cookie or None)
    t_run = time.time()

    # ----------------------------------------------------------------------
    # Step 2 — Check each product. Group by the site that confirmed the draft.
    # ----------------------------------------------------------------------
    log.info(
        "[2/3] check phase: probing %d products for draft existence across %d workers",
        len(candidates), args.workers,
    )
    ok_log = logging.getLogger("products.ok")
    err_log = logging.getLogger("products.err")

    to_publish_by_site: dict[str, list[str]] = defaultdict(list)
    publish_lock = threading.Lock()
    n_with_draft = 0
    n_no_draft = 0
    n_check_failed = 0

    def _check_one(pid: str) -> tuple[str, bool, str | None, str | None]:
        if args.dry_run:
            # In dry-run, we don't actually hit the API. Instead, pretend
            # every product has a draft on the primary site so the
            # downstream phases get exercised in the dry-run path.
            return (pid, True, site_ids[0], None)
        has_draft, winning_site, err = check_product_has_draft(
            session, args.api_host, pid, site_ids,
            delay_ms=args.delay_ms,
        )
        return (pid, has_draft, winning_site, err)

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        futs = {pool.submit(_check_one, pid): pid for pid in candidates}
        for fut in as_completed(futs):
            pid, has_draft, winning_site, err = fut.result()
            if err is not None:
                n_check_failed += 1
                # check_product_has_draft already logged the error;
                # the count gets surfaced in the summary.
                continue
            if has_draft:
                n_with_draft += 1
                with publish_lock:
                    to_publish_by_site[winning_site].append(pid)
                ok_log.info(
                    "verdict=draft_exists productId=%s siteId=%s",
                    pid, winning_site,
                )
            else:
                n_no_draft += 1
                # No log line for "no draft" — that's the common case and
                # would balloon the log. The count is in the summary.

            done = n_with_draft + n_no_draft + n_check_failed
            if done % 100 == 0:
                log.info(
                    "check progress %d/%d (with_draft=%d no_draft=%d check_failed=%d)",
                    done, len(candidates), n_with_draft, n_no_draft, n_check_failed,
                )

    log.info(
        "[2/3] check done: with_draft=%d no_draft=%d check_failed=%d "
        "to_publish_by_site=%s",
        n_with_draft, n_no_draft, n_check_failed,
        {s: len(v) for s, v in to_publish_by_site.items()},
    )

    if args.stop_after == "check" or args.skip_publish:
        log.info("[3/3] skipping publish phase (stop_after=%s, skip_publish=%s)",
                 args.stop_after, args.skip_publish)
        return _final_summary(log, run_id, t_run,
                              success=(n_with_draft + n_no_draft) > 0)

    if not to_publish_by_site:
        log.info("[3/3] no products to publish (none had drafts)")
        return _final_summary(log, run_id, t_run,
                              success=(n_with_draft + n_no_draft) > 0)

    # ----------------------------------------------------------------------
    # Step 3 — Batched COMMIT (READY removed per operator direction).
    # ----------------------------------------------------------------------
    log.info("[3/3] publish phase: COMMIT only (no READY)")
    _, _, commit_ok, commit_fail = _run_publish_phase(
        session, args.api_host,
        to_publish_by_site,
        batch_size=args.batch_size,
        workers=args.workers,
        dry_run=args.dry_run,
        delay_ms=args.delay_ms,
    )
    log.info(
        "publish phase done: commit_ok=%d commit_fail=%d",
        commit_ok, commit_fail,
    )

    success = commit_ok > 0 and commit_fail == 0 and n_check_failed == 0
    return _final_summary(log, run_id, t_run, success=success)


def _run_publish_phase(
    session: requests.Session,
    api_host: str,
    to_publish_by_site: dict[str, list[str]],
    batch_size: int,
    workers: int,
    dry_run: bool = False,
    delay_ms: int = 0,
) -> tuple[int, int, int, int]:
    """
    Single-phase batched publish: COMMIT only (READY is removed per
    operator direction). A COMMIT failure leaves the product in DRAFT
    state; there's no rollback gate anymore.

    `delay_ms` is forwarded to batch_change_status — see its docstring.

    Returns (ready_ok, ready_failed, commit_ok, commit_failed). ready_*
    are always 0 since no READY calls happen — kept for caller compat.
    """
    log = logging.getLogger("products")
    err_log = logging.getLogger("products.err")
    ok_log = logging.getLogger("products.ok")

    commit_tasks = [
        (site, batch)
        for site, ids in to_publish_by_site.items()
        for batch in chunked(ids, batch_size)
    ]
    log.info(
        "  COMMIT-only: %d batches, %d products, %d sites (READY skipped)",
        len(commit_tasks),
        sum(len(v) for v in to_publish_by_site.values()),
        len(to_publish_by_site),
    )

    commit_ok = commit_fail = 0
    with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futs = {
            pool.submit(
                batch_change_status,
                session, api_host, site, batch, "COMMIT",
                dry_run=dry_run, delay_ms=delay_ms,
            ): (site, batch)
            for site, batch in commit_tasks
        }
        for fut in as_completed(futs):
            site, batch = futs[fut]
            ok, info = fut.result()
            if ok:
                commit_ok += len(batch)
                ok_log.info(
                    "phase=COMMIT verdict=ok siteId=%s n=%d detail=%s",
                    site, len(batch), info,
                )
            else:
                commit_fail += len(batch)
                err_log.error(
                    "phase=COMMIT verdict=failed siteId=%s n=%d ids=%s detail=%s",
                    site, len(batch), batch[:5], info,
                )

    # ready_ok, ready_fail always 0 (no READY phase).
    return 0, 0, commit_ok, commit_fail


def _final_summary(
    log: logging.Logger,
    run_id: str,
    t_start: float,
    success: bool,
) -> int:
    elapsed = time.time() - t_start
    log.info(
        "run_end run_id=%s success=%s elapsed_s=%.1f",
        run_id, success, elapsed,
    )
    return 0 if success else 2


if __name__ == "__main__":
    raise SystemExit(main())