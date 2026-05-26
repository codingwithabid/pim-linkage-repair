#!/usr/bin/env python3
"""
repair_archived_products.py
===========================

Product-side counterpart of repair_archived_variants.py.

This script handles archived products. push_products.py deliberately
skips archived records because an archived product can't be updated
directly — the platform forbids writes to archived. To repair them we
have to:

  1. recover  (state: ARCHIVED → NORMAL)
  2. PUT the new payload
  3. COMMIT          (publish the new version — no READY)
  4. archive again  (state: NORMAL → ARCHIVED)
  5. COMMIT          (commit the archive — no READY)

That's two round-trip state changes plus the actual PUT. push_products.py
deliberately avoids touching archived products so its happy-path stays
simple; this script owns the more involved flow.

Note: READY is intentionally skipped throughout (per operator direction).
A COMMIT failure leaves the product in an intermediate state with no
rollback — re-runs and manual fixes are the recovery story.

Filter
------
Only processes products in product_map.json where:

    state  == "ARCHIVED"
    status in ("PUBLISHED", "DRAFT")

Within ARCHIVED, the status forks the flow:

    status == DRAFT      → PUT only (no state changes — platform allows
                                     PUT directly on archived-DRAFT)
    status == PUBLISHED  → full 6-step pipeline above

Anything else (state != ARCHIVED, status == null, unknown status) is
skipped with a log line — those are someone else's responsibility.

Batched everything-but-PUT
--------------------------
PUT is per-product (URL is product-scoped). RECOVER, COMMIT, and ARCHIVE
all use batch endpoints, grouped by the variant's PRIMARY site
(= site_ids[0]). Batch size defaults to 50.

Failure semantics
-----------------
A batch failure at phase N drops EVERY product in that batch from phase
N+1. This is by design — the API returns success/fail per BATCH not per
product, and the safer default is to stop a half-applied state change
rather than mark some half-broken middle ground.

Failed batches are recorded in errors.log with the offending ids. The
operator can re-run the script (it's idempotent on the platform side for
most operations) or fix the affected products manually.

Usage
-----
    export API_COOKIE='...'
    python3 repair_archived_products.py --input ./out-products --workers 4

    # Dry run — log what would happen, no HTTP
    python3 repair_archived_products.py --input ./out-products --dry-run

    # Stage: only do RECOVER (state change to NORMAL), then stop. Lets
    # the operator inspect the products before they proceed.
    python3 repair_archived_products.py --input ./out-products --stop-after recover
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
# Site configuration — same as repair_products.py / push_products.py
# ---------------------------------------------------------------------------
PRIMARY_SITE_ID = "201cb789-4198-488b-a5eb-4e7df0fb4bee"
FALLBACK_SITE_IDS = [
    "8d3ea3bc-f65b-4227-9fa6-6fae40e4575a",
    "fbfcd92b-d271-4002-9163-4f84986b41be",
]


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
LOG_FMT = "%(asctime)s | %(levelname)-7s | %(name)-15s | %(message)s"
LOG_DATEFMT = "%Y-%m-%d %H:%M:%S"


class ProgressReporter:
    """
    Periodic progress logger for long-running parallel loops. Logs at
    every `step` items OR every `interval_s` seconds, whichever comes
    first. Step auto-computed as ~10% of total (clamped to 10..200).
    Thread-safe.
    """

    def __init__(
        self,
        log: logging.Logger,
        label: str,
        total: int,
        step: int | None = None,
        interval_s: float = 15.0,
    ):
        self.log = log
        self.label = label
        self.total = max(1, total)
        if step is None:
            step = max(10, min(200, self.total // 10 or 1))
        self.step = step
        self.interval_s = interval_s
        self._done = 0
        self._last_log_done = 0
        self._t0 = time.monotonic()
        self._last_log_t = self._t0
        self._lock = threading.Lock()

    def tick(self, extra: str = "") -> None:
        with self._lock:
            self._done += 1
            now = time.monotonic()
            elapsed_since_last = now - self._last_log_t
            steps_since_last = self._done - self._last_log_done
            is_last = self._done >= self.total
            if (
                steps_since_last >= self.step
                or elapsed_since_last >= self.interval_s
                or is_last
            ):
                rate = self._done / max(0.001, now - self._t0)
                pct = 100.0 * self._done / self.total
                suffix = f" {extra}" if extra else ""
                self.log.info(
                    "%s progress %d/%d (%.0f%% — ~%.1f/s)%s",
                    self.label, self._done, self.total, pct, rate, suffix,
                )
                self._last_log_done = self._done
                self._last_log_t = now


def setup_logging(log_dir: Path, verbose: bool = False) -> logging.Logger:
    """
    Three rotating handlers + console mirror at INFO+:

      archived-products.run.log      — everything (DEBUG if --verbose else INFO)
      archived-products.success.log  — per-product verdicts (INFO only)
      archived-products.errors.log   — WARNING+ (propagates to run.log)
    """
    log_dir.mkdir(parents=True, exist_ok=True)
    formatter = logging.Formatter(LOG_FMT, datefmt=LOG_DATEFMT)

    def _rotating(path: Path, level: int) -> RotatingFileHandler:
        h = RotatingFileHandler(
            path, maxBytes=10 * 1024 * 1024, backupCount=5, encoding="utf-8"
        )
        h.setLevel(level)
        h.setFormatter(formatter)
        return h

    main = logging.getLogger("archived")
    main.setLevel(logging.DEBUG if verbose else logging.INFO)
    main.propagate = False
    main.handlers.clear()
    main.addHandler(_rotating(
        log_dir / "archived-products.run.log",
        logging.DEBUG if verbose else logging.INFO,
    ))
    console = logging.StreamHandler(sys.stderr)
    console.setLevel(logging.INFO)
    console.setFormatter(formatter)
    main.addHandler(console)

    ok = logging.getLogger("archived.ok")
    ok.setLevel(logging.INFO)
    ok.propagate = False
    ok.handlers.clear()
    ok.addHandler(_rotating(log_dir / "archived-products.success.log", logging.INFO))

    err = logging.getLogger("archived.err")
    err.setLevel(logging.WARNING)
    err.propagate = True
    err.handlers.clear()
    err.addHandler(_rotating(log_dir / "archived-products.errors.log", logging.WARNING))

    main.info("archived-products logging initialised: dir=%s verbose=%s",
              log_dir, verbose)
    return main


# ---------------------------------------------------------------------------
# HTTP session
# ---------------------------------------------------------------------------

def make_session(
    session_id: str,
    cookie_header: str | None = None,
) -> requests.Session:
    """
    Shared session for all workers. X-UpStart-Site goes per-request
    because workers may target different sites — putting it on the
    session would race.
    """
    s = requests.Session()
    headers = {
        "X-Upstart-Tenant":     "denvermattress",
        "Referer":               "https://denvermattress.nochannel-test-1.nochannel-test.upstart.team/",
        "X-Upstart-Session-Id": session_id,
        "Accept":                "application/json, text/plain, */*",
        "Content-Type":          "application/json",
    }
    if cookie_header:
        headers["Cookie"] = cookie_header
    s.headers.update(headers)
    return s


# ---------------------------------------------------------------------------
# Scan
# ---------------------------------------------------------------------------

def collect_archived_items(
    input_dir: Path,
    product_map: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Find every payload file paired with a product_map entry where:
        state == "ARCHIVED"  AND  status in ("PUBLISHED", "DRAFT")

    Returns one dict per candidate, tagged with its status so the
    pipeline can fork into the two flows:

        {"productId": str,
         "siteIds":   list[str],    # full list from folder name
         "siteId":    str,          # primary (= site_ids[0]), for batch grouping
         "filePath":  Path,
         "payload":   dict,
         "status":    "PUBLISHED" | "DRAFT"}

    Two flows downstream:
      PUBLISHED → full 6-step pipeline
                  (recover → PUT → COMMIT-1 → ARCHIVE → COMMIT-2)
      DRAFT     → PUT only
                  (platform accepts PUT directly on archived-DRAFT;
                   no state change needed, the product stays ARCHIVED)

    Diagnostics tally each skip/candidate category for the post-scan log.
    """
    log = logging.getLogger("archived")
    err_log = logging.getLogger("archived.err")

    items: list[dict[str, Any]] = []
    counts = Counter()
    files_seen = 0

    for tpl_dir in sorted(input_dir.iterdir()):
        if not tpl_dir.is_dir() or tpl_dir.name == "logs":
            continue
        for site_dir in sorted(tpl_dir.iterdir()):
            if not site_dir.is_dir():
                continue
            # Parse comma-joined site list from folder name. Single-site
            # folders have one element; multi-site folders have several.
            site_ids = [s for s in site_dir.name.split(",") if s]
            if not site_ids:
                err_log.warning(
                    "bad_site_folder name=%s reason=empty_after_split",
                    site_dir.name,
                )
                continue
            primary_site = site_ids[0]
            for f in sorted(site_dir.glob("*.json")):
                files_seen += 1
                pid = f.stem
                meta = product_map.get(pid)
                if meta is None:
                    err_log.warning(
                        "orphan_file productId=%s path=%s reason=not_in_product_map",
                        pid, f,
                    )
                    counts["orphan_file"] += 1
                    continue

                state = meta.get("state")
                status = meta.get("status")

                if state != "ARCHIVED":
                    counts["skip_not_archived"] += 1
                    continue

                # Fork on status.
                if status == "PUBLISHED":
                    counts["candidate_published"] += 1
                elif status == "DRAFT":
                    counts["candidate_draft"] += 1
                elif status is None:
                    log.info(
                        "skip_archived_null_status productId=%s siteIds=%s",
                        pid, site_ids,
                    )
                    counts["skip_archived_null"] += 1
                    continue
                else:
                    log.warning(
                        "skip_archived_unknown_status productId=%s status=%r",
                        pid, status,
                    )
                    counts["skip_unknown_status"] += 1
                    continue

                # Load payload now so the API phase stays pure I/O.
                try:
                    payload = json.loads(f.read_text())
                except json.JSONDecodeError as e:
                    err_log.error(
                        "bad_payload_file productId=%s path=%s detail=%r",
                        pid, f, str(e),
                    )
                    counts["bad_payload"] += 1
                    continue

                items.append({
                    "productId": pid,
                    "siteIds":   site_ids,
                    "siteId":    primary_site,
                    "filePath":  f,
                    "payload":   payload,
                    "status":    status,
                })

    log.info(
        "scan complete: files_seen=%d "
        "candidate_published=%d (full pipeline) "
        "candidate_draft=%d (PUT only) "
        "skip_not_archived=%d skip_archived_null=%d "
        "skip_unknown_status=%d orphan_file=%d bad_payload=%d",
        files_seen,
        counts.get("candidate_published", 0),
        counts.get("candidate_draft", 0),
        counts.get("skip_not_archived", 0),
        counts.get("skip_archived_null", 0),
        counts.get("skip_unknown_status", 0),
        counts.get("orphan_file", 0),
        counts.get("bad_payload", 0),
    )

    return items


# ---------------------------------------------------------------------------
# Individual API calls
# ---------------------------------------------------------------------------

def chunked(xs: list[str], n: int) -> list[list[str]]:
    if n <= 0:
        return [xs]
    return [xs[i:i + n] for i in range(0, len(xs), n)]


def put_product(
    session: requests.Session,
    api_host: str,
    product_id: str,
    site_ids: list[str],
    payload: dict[str, Any],
    timeout: float = 30.0,
    dry_run: bool = False,
    delay_ms: int = 0,
) -> tuple[bool, str]:
    """
    PUT /pim/products/<id> with payload body.

    Multi-site contract (same as push_products.put_product):
      X-UpStart-Site: <site_ids[0]>           (header)
      ?sites=<site_ids[1]>,<site_ids[2]>,...  (query, omitted if single-site)
    """
    if not site_ids:
        return False, "no site_ids provided"

    primary_site = site_ids[0]
    other_sites = site_ids[1:]
    url = f"{api_host.rstrip('/')}/pim/products/{product_id}"

    request_kwargs: dict[str, Any] = {
        "headers": {"X-UpStart-Site": primary_site},
        "json":    payload,
        "timeout": timeout,
    }
    if other_sites:
        request_kwargs["params"] = {"sites": ",".join(other_sites)}

    if dry_run:
        params_part = f"?sites={','.join(other_sites)}" if other_sites else ""
        return True, (
            f"DRY_RUN PUT {url}{params_part} "
            f"X-UpStart-Site={primary_site} body_keys={list(payload)}"
        )

    try:
        r = session.put(url, **request_kwargs)
    except requests.RequestException as e:
        if delay_ms > 0:
            time.sleep(delay_ms / 1000.0)
        return False, f"network: {e}"
    if delay_ms > 0:
        time.sleep(delay_ms / 1000.0)
    if 200 <= r.status_code < 300:
        return True, f"http {r.status_code}"
    return False, f"http {r.status_code} body={r.text[:300]!r}"


def batch_recover(
    session: requests.Session,
    api_host: str,
    site_id: str,
    product_ids: list[str],
    timeout: float = 60.0,
    dry_run: bool = False,
    delay_ms: int = 0,
) -> tuple[bool, str]:
    """
    POST /pim/batch/products/recover
    body: {"ids":[...]}

    Moves the listed products from ARCHIVED to NORMAL. The platform
    creates a draft revision representing the change.
    """
    url = f"{api_host.rstrip('/')}/pim/batch/products/recover"
    if dry_run:
        return True, f"DRY_RUN POST {url} site={site_id} ids_n={len(product_ids)}"
    try:
        r = session.post(
            url,
            headers={"X-UpStart-Site": site_id},
            json={"ids": product_ids},
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
    body: {"ids":[...], "status":"COMMIT"}
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


def batch_change_state_archived(
    session: requests.Session,
    api_host: str,
    site_id: str,
    product_ids: list[str],
    timeout: float = 60.0,
    dry_run: bool = False,
    delay_ms: int = 0,
) -> tuple[bool, str]:
    """
    POST /pim/batch/products/change-state/ARCHIVED
    body: {"ids":[...]}

    Moves the listed products from NORMAL to ARCHIVED. This is the
    inverse of batch_recover. Like recover, it creates an uncommitted
    state-change that needs a subsequent COMMIT to make it visible.
    """
    url = f"{api_host.rstrip('/')}/pim/batch/products/change-state/ARCHIVED"
    if dry_run:
        return True, f"DRY_RUN POST {url} site={site_id} ids_n={len(product_ids)}"
    try:
        r = session.post(
            url,
            headers={"X-UpStart-Site": site_id},
            json={"ids": product_ids},
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
# Phase runners — DRY helpers used across all 6 published-path phases.
# ---------------------------------------------------------------------------

def _run_batched_phase(
    work_by_site: dict[str, list[str]],
    runner,  # callable: (site_id, ids_batch) -> (ok: bool, info: str)
    phase_name: str,
    batch_size: int,
    workers: int,
) -> dict[str, list[str]]:
    """
    Run one batched phase across every (site, batch) pair concurrently.
    Returns survivors_by_site — the subset of work_by_site whose batches
    succeeded. Batches that failed are dropped entirely (batch-level
    granularity is by design; see module docstring).

    work_by_site is dict[site → list of product_ids on that site].

    The runner closure captures whatever session/host/dry_run params it
    needs from the caller's scope.
    """
    log = logging.getLogger("archived")
    err_log = logging.getLogger("archived.err")
    ok_log = logging.getLogger("archived.ok")

    tasks = [
        (site, batch)
        for site, ids in work_by_site.items()
        for batch in chunked(ids, batch_size)
    ]
    total_products = sum(len(v) for v in work_by_site.values())
    log.info(
        "  %s: %d batches across %d sites (%d products total, batch_size=%d)",
        phase_name, len(tasks), len(work_by_site), total_products, batch_size,
    )

    survivors_by_site: dict[str, list[str]] = defaultdict(list)
    n_ok = n_fail = 0
    progress = ProgressReporter(log, f"  {phase_name} batches", total=len(tasks))

    with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futs = {pool.submit(runner, site, batch): (site, batch) for site, batch in tasks}
        for fut in as_completed(futs):
            site, batch = futs[fut]
            ok, info = fut.result()
            if ok:
                n_ok += len(batch)
                survivors_by_site[site].extend(batch)
                ok_log.info(
                    "phase=%s verdict=ok siteId=%s n=%d detail=%s",
                    phase_name, site, len(batch), info,
                )
            else:
                n_fail += len(batch)
                err_log.error(
                    "phase=%s verdict=failed siteId=%s n=%d ids=%s detail=%s",
                    phase_name, site, len(batch), batch[:5], info,
                )
            progress.tick(extra=f"ok={n_ok} fail={n_fail}")

    log.info(
        "  %s done: ok=%d fail=%d survivors=%d",
        phase_name, n_ok, n_fail, sum(len(v) for v in survivors_by_site.values()),
    )
    return dict(survivors_by_site)


def _run_put_phase(
    session: requests.Session,
    api_host: str,
    work_by_site: dict[str, list[str]],
    items_by_id: dict[str, dict[str, Any]],
    *,
    workers: int,
    dry_run: bool,
    delay_ms: int = 0,
) -> dict[str, list[str]]:
    """
    PUT phase — per-product, not batched. work_by_site keys are primary
    sites (used only for grouping the survivors back up). The PUT itself
    uses each item's full siteIds list.

    `delay_ms` is forwarded to put_product — see its docstring.

    Returns survivors_by_site: only products whose PUT succeeded.
    """
    log = logging.getLogger("archived")
    err_log = logging.getLogger("archived.err")
    ok_log = logging.getLogger("archived.ok")

    flat_ids = [pid for ids in work_by_site.values() for pid in ids]
    log.info("  PUT: %d products, %d workers", len(flat_ids), workers)

    def _do_one(site_id: str, product_id: str) -> tuple[str, str, bool, str]:
        """
        site_id here is the product's PRIMARY site (used for grouping
        upstream). PUT itself needs the full siteIds list — primary in
        header, rest in ?sites= query. Look it up from items_by_id.
        """
        item = items_by_id[product_id]
        ok, info = put_product(
            session, api_host, product_id,
            item["siteIds"],
            item["payload"],
            dry_run=dry_run, delay_ms=delay_ms,
        )
        return site_id, product_id, ok, info

    survivors_by_site: dict[str, list[str]] = defaultdict(list)
    n_ok = n_fail = 0
    progress = ProgressReporter(log, "  PUT", total=len(flat_ids))

    with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futs = []
        for site, ids in work_by_site.items():
            for pid in ids:
                futs.append(pool.submit(_do_one, site, pid))
        for fut in as_completed(futs):
            site, pid, ok, info = fut.result()
            if ok:
                n_ok += 1
                survivors_by_site[site].append(pid)
                ok_log.info(
                    "phase=PUT verdict=ok productId=%s siteId=%s detail=%s",
                    pid, site, info,
                )
            else:
                n_fail += 1
                err_log.error(
                    "phase=PUT verdict=failed productId=%s siteId=%s detail=%s",
                    pid, site, info,
                )
            progress.tick(extra=f"ok={n_ok} fail={n_fail}")

    log.info("  PUT done: ok=%d fail=%d", n_ok, n_fail)
    return dict(survivors_by_site)


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
        help="Output dir from repair_products.py (contains product_map.json + payloads)",
    )
    ap.add_argument(
        "--api-host",
        default="https://nochannel-test-1-api.nochannel-test.upstart.team",
        help="API host (no trailing slash)",
    )
    ap.add_argument(
        "--site-id", default=PRIMARY_SITE_ID,
        help="Default site (informational; folder names drive actual routing)",
    )
    ap.add_argument(
        "--fallback-site-ids",
        default=",".join(FALLBACK_SITE_IDS),
        help="Comma-separated fallback sites (informational)",
    )
    ap.add_argument(
        "--session-id",
        default=os.environ.get("API_SESSION_ID", "example"),
    )
    ap.add_argument(
        "--cookie",
        default=os.environ.get("API_COOKIE", ""),
    )
    ap.add_argument(
        "--workers", type=int, default=1,
        help="Concurrent batch + PUT workers (default 1)",
    )
    ap.add_argument(
        "--batch-size", type=int, default=1,
        help="Max products per batch call (default 1)",
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
        choices=["scan", "recover", "put", "publish", "archive", "all"],
        default="all",
        help="Stop the pipeline after the named phase. Useful for staged "
             "rollouts: do RECOVER, inspect, continue.",
    )
    ap.add_argument(
        "--log-dir", type=Path, default=None,
        help="Where logs go. Default: <input>/logs",
    )
    ap.add_argument(
        "--verbose", "-v", action="store_true",
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

    # --- load product_map.json ---
    map_path = args.input / "product_map.json"
    if not map_path.exists():
        log.error("product_map.json not found at %s", map_path)
        return 1
    product_map = json.loads(map_path.read_text())
    log.info("loaded product_map.json with %d entries", len(product_map))

    # --- scan + filter ---
    log.info("[1/6] scan: looking for archived-PUBLISHED and archived-DRAFT candidates")
    items = collect_archived_items(args.input, product_map)
    if args.limit > 0:
        items = items[: args.limit]
        log.info("--limit %d applied; processing first %d", args.limit, len(items))

    if not items:
        log.info("nothing to do — no archived candidates. run_id=%s", run_id)
        return 0

    if args.stop_after == "scan":
        log.info("stop_after=scan: exiting before any API calls")
        return 0

    if not args.cookie and not args.dry_run:
        log.warning(
            "no Cookie provided — if the gateway requires session cookies "
            "you'll get HTML responses. Pass --cookie or API_COOKIE env."
        )

    session = make_session(args.session_id, cookie_header=args.cookie or None)

    # Partition by status. The DRAFT path is one step (just PUT); the
    # PUBLISHED path is the full 6-step pipeline. Both keep per-product
    # site bindings so each call goes to the right site.
    published_items = [it for it in items if it["status"] == "PUBLISHED"]
    draft_items     = [it for it in items if it["status"] == "DRAFT"]
    items_by_id: dict[str, dict[str, Any]] = {it["productId"]: it for it in items}
    log.info(
        "candidates partitioned: published=%d (full 6-step pipeline)  draft=%d (PUT only)",
        len(published_items), len(draft_items),
    )

    t_run = time.time()
    any_progress = False

    # ----------------------------------------------------------------------
    # DRAFT path — just PUT, no state changes.
    # An archived product with a DRAFT status means there's an editable
    # draft revision attached to the archived record. The platform allows
    # PUTting directly to that draft; no recover/republish/rearchive needed.
    # ----------------------------------------------------------------------
    if draft_items:
        log.info("[draft] PUT phase: %d archived-DRAFT products (no state changes)",
                 len(draft_items))
        draft_by_site: dict[str, list[str]] = defaultdict(list)
        for it in draft_items:
            draft_by_site[it["siteId"]].append(it["productId"])
        draft_survivors = _run_put_phase(
            session, args.api_host, draft_by_site, items_by_id,
            workers=args.workers, dry_run=args.dry_run, delay_ms=args.delay_ms,
        )
        n = sum(len(v) for v in draft_survivors.values())
        log.info("[draft] done: %d/%d archived-DRAFT products updated",
                 n, len(draft_items))
        if n:
            any_progress = True

    # ----------------------------------------------------------------------
    # PUBLISHED path — the full 6-step pipeline.
    # Each phase consumes the survivors of the previous one. A product
    # that drops out at any step doesn't get touched again by this run.
    # ----------------------------------------------------------------------
    if not published_items:
        log.info("no archived-PUBLISHED candidates; skipping full pipeline")
        return _final_summary(log, run_id, t_run, success=any_progress)

    work_by_site: dict[str, list[str]] = defaultdict(list)
    for it in published_items:
        work_by_site[it["siteId"]].append(it["productId"])
    log.info(
        "[published path] %d products across %d sites",
        len(published_items), len(work_by_site),
    )

    # [2/6] RECOVER  ARCHIVED → NORMAL
    log.info("[2/6] recover phase: ARCHIVED → NORMAL (PUBLISHED candidates only)")
    survivors = _run_batched_phase(
        work_by_site,
        runner=lambda site, batch: batch_recover(
            session, args.api_host, site, batch, dry_run=args.dry_run, delay_ms=args.delay_ms,
        ),
        phase_name="RECOVER",
        batch_size=args.batch_size,
        workers=args.workers,
    )
    if args.stop_after == "recover" or not survivors:
        log.info("stopping after recover (stop_after=%s, survivors=%d)",
                 args.stop_after, sum(len(v) for v in survivors.values()))
        return _final_summary(log, run_id, t_run,
                              success=any_progress or bool(survivors))

    # [3/6] PUT  (per-product)
    log.info("[3/6] PUT phase: pushing repaired payloads")
    survivors = _run_put_phase(
        session, args.api_host, survivors, items_by_id,
        workers=args.workers, dry_run=args.dry_run, delay_ms=args.delay_ms,
    )
    if args.stop_after == "put" or not survivors:
        log.info("stopping after PUT (stop_after=%s, survivors=%d)",
                 args.stop_after, sum(len(v) for v in survivors.values()))
        return _final_summary(log, run_id, t_run,
                              success=any_progress or bool(survivors))

    # [4/6] COMMIT (post-PUT) — publish the recovered+updated revision.
    # READY is removed per operator direction; we go straight from PUT to
    # COMMIT. A COMMIT failure here leaves the product in NORMAL state
    # with an uncommitted draft change.
    log.info("[4/6] COMMIT phase (post-PUT)")
    survivors = _run_batched_phase(
        survivors,
        runner=lambda site, batch: batch_change_status(
            session, args.api_host, site, batch, "COMMIT", dry_run=args.dry_run, delay_ms=args.delay_ms,
        ),
        phase_name="COMMIT-1",
        batch_size=args.batch_size,
        workers=args.workers,
    )
    if args.stop_after == "publish" or not survivors:
        log.info("stopping after publish (stop_after=%s, survivors=%d)",
                 args.stop_after, sum(len(v) for v in survivors.values()))
        return _final_summary(log, run_id, t_run,
                              success=any_progress or bool(survivors))

    # [5/6] ARCHIVE  NORMAL → ARCHIVED  (restoring original state)
    log.info("[5/6] archive phase: NORMAL → ARCHIVED (restoring original state)")
    survivors = _run_batched_phase(
        survivors,
        runner=lambda site, batch: batch_change_state_archived(
            session, args.api_host, site, batch, dry_run=args.dry_run, delay_ms=args.delay_ms,
        ),
        phase_name="ARCHIVE",
        batch_size=args.batch_size,
        workers=args.workers,
    )
    if not survivors:
        log.warning(
            "no survivors after archive; products are stuck in NORMAL state — "
            "needs manual rescue. See errors.log."
        )
        return _final_summary(log, run_id, t_run, success=any_progress)
    if args.stop_after == "archive":
        log.info("stopping after archive (stop_after=archive, survivors=%d)",
                 sum(len(v) for v in survivors.values()))
        return _final_summary(log, run_id, t_run,
                              success=any_progress or bool(survivors))

    # [6/6] COMMIT (post-archive) — finalize the archived state.
    # READY-2 is removed per operator direction; we COMMIT directly. A
    # failure here leaves products in ARCHIVED state with an uncommitted
    # archive change — manual rescue needed.
    log.info("[6/6] COMMIT phase (post-archive)")
    survivors = _run_batched_phase(
        survivors,
        runner=lambda site, batch: batch_change_status(
            session, args.api_host, site, batch, "COMMIT", dry_run=args.dry_run, delay_ms=args.delay_ms,
        ),
        phase_name="COMMIT-2",
        batch_size=args.batch_size,
        workers=args.workers,
    )

    return _final_summary(log, run_id, t_run,
                          success=any_progress or bool(survivors))


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