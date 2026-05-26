#!/usr/bin/env python3
"""
repair_archived_variants.py
===========================

Third script in the variant-repair toolchain. Companion to:
    repair_variants.py   (produces variant_map.json + payload files)
    push_variants.py     (pushes non-archived variants)

This script handles ARCHIVED variants specifically.

Why a separate script
---------------------
An archived variant can't be updated directly — the platform forbids
writes to archived records. To repair an archived variant we have to:

  1. recover it     (state: ARCHIVED → NORMAL)
  2. PUT the new payload
  3. COMMIT          (publish the new version — no READY)
  4. archive again  (state: NORMAL → ARCHIVED)
  5. COMMIT          (commit the archive — no READY)

That's two round-trip state changes plus the actual PUT. push_variants.py
deliberately avoids touching archived variants so its happy-path stays
simple. This script owns the more involved flow.

Note: READY is intentionally skipped throughout (per operator direction).
A COMMIT failure leaves the variant in an intermediate state with no
rollback — re-runs and manual fixes are the recovery story.

Filter
------
Only processes variants in variant_map.json where:

    state == "ARCHIVED"  AND  status == "PUBLISHED"

Archived-DRAFT variants are skipped entirely. Rationale: a DRAFT archived
variant means someone has an editable archived draft — exotic and worth a
human review before automation touches it. Logged at INFO so they're
visible but don't pollute errors.log.

Decision matrix
---------------

    state=ARCHIVED, status=PUBLISHED  →  full 7-step recover/republish/rearchive
    state=ARCHIVED, status=DRAFT       →  skip (logged)
    state=ARCHIVED, status=null        →  skip (no authoritative pre-state)
    state=NORMAL,   any status         →  skip (not this script's job)

Failure handling
----------------
Each batched step is independent: variants whose batch fails at step N do
NOT proceed to step N+1. This means a partial failure may leave a variant
in an intermediate state (e.g. NORMAL after recover but archive failed).
Errors.log records the exact step that failed so a human can decide:
  - re-run the script (recover is idempotent; PUT is idempotent;
    re-archive of an already-NORMAL variant just works)
  - or manually fix in the UI

Sites
-----
The payload file's parent folder name IS the site that owns the variant.
Every API call goes to that site (X-UpStart-Site header). No cross-site
retry — we know exactly where each variant lives.

Usage
-----
    # Dry run first — confirms the candidate count before touching anything
    python3 repair_archived_variants.py --input ./out --dry-run

    # Real run
    export API_COOKIE='...'
    python3 repair_archived_variants.py --input ./out --workers 8
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
# Logging — same three-file pattern as the other scripts
# ---------------------------------------------------------------------------
LOG_FMT = "%(asctime)s | %(levelname)-7s | %(name)-12s | %(message)s"
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

    main = logging.getLogger("archived")
    main.setLevel(logging.DEBUG if verbose else logging.INFO)
    main.propagate = False
    main.handlers.clear()
    main.addHandler(_rotating(log_dir / "archived.run.log", logging.DEBUG if verbose else logging.INFO))
    console = logging.StreamHandler(sys.stderr)
    console.setLevel(logging.INFO)
    console.setFormatter(formatter)
    main.addHandler(console)

    ok = logging.getLogger("archived.ok")
    ok.setLevel(logging.INFO)
    ok.propagate = False
    ok.handlers.clear()
    ok.addHandler(_rotating(log_dir / "archived.success.log", logging.INFO))

    err = logging.getLogger("archived.err")
    err.setLevel(logging.WARNING)
    err.propagate = True
    err.handlers.clear()
    err.addHandler(_rotating(log_dir / "archived.errors.log", logging.WARNING))

    main.info("archived-repair logging initialised: dir=%s verbose=%s", log_dir, verbose)
    return main


# ---------------------------------------------------------------------------
# HTTP session
# ---------------------------------------------------------------------------

def make_session(
    session_id: str,
    cookie_header: str | None = None,
) -> requests.Session:
    """
    Shared across workers. X-UpStart-Site goes per-request so the same
    session works across multiple sites without thread races.
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
# Work-item discovery
# ---------------------------------------------------------------------------

def collect_archived_items(
    input_dir: Path,
    variant_map: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Find every payload file paired with a variant_map entry where:
        state == "ARCHIVED"  AND  status in ("PUBLISHED", "DRAFT")

    Returns one dict per candidate, tagged with its status so the
    pipeline can fork into the two flows:
        {"variantId": ..., "siteId": ..., "filePath": Path,
         "payload": dict, "status": "PUBLISHED" | "DRAFT"}

    Two flows downstream:
      PUBLISHED → full 7-step pipeline
                  (recover → PUT → COMMIT-1 → ARCHIVE → COMMIT-2)
      DRAFT     → PUT only
                  (platform accepts PUT directly on archived-DRAFT;
                   no state change needed, the variant stays ARCHIVED)

    Diagnostics tally each skip category for the run-start log line.
    """
    log = logging.getLogger("archived")
    err_log = logging.getLogger("archived.err")

    items: list[dict[str, Any]] = []
    counts = Counter()  # skip + candidate categories
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
                vid = f.stem
                meta = variant_map.get(vid)
                if meta is None:
                    err_log.warning(
                        "orphan_file variantId=%s path=%s reason=not_in_variant_map",
                        vid, f,
                    )
                    counts["orphan_file"] += 1
                    continue

                state = meta.get("state")
                status = meta.get("status")

                # Filter: state must be ARCHIVED.
                if state != "ARCHIVED":
                    counts["skip_not_archived"] += 1
                    continue

                # Within ARCHIVED, fork on status.
                if status == "PUBLISHED":
                    counts["candidate_published"] += 1
                elif status == "DRAFT":
                    counts["candidate_draft"] += 1
                elif status is None:
                    # No authoritative pre-state recorded; refuse to act.
                    log.info(
                        "skip_archived_null_status variantId=%s siteIds=%s",
                        vid, site_ids,
                    )
                    counts["skip_archived_null"] += 1
                    continue
                else:
                    # Defensive — unknown status value
                    log.warning(
                        "skip_archived_unknown_status variantId=%s status=%r",
                        vid, status,
                    )
                    counts["skip_unknown_status"] += 1
                    continue

                # Load payload now so the API phase stays pure I/O.
                try:
                    payload = json.loads(f.read_text())
                except json.JSONDecodeError as e:
                    err_log.error(
                        "bad_payload_file variantId=%s path=%s detail=%r",
                        vid, f, str(e),
                    )
                    counts["bad_payload"] += 1
                    continue

                items.append({
                    "variantId": vid,
                    "siteIds":   site_ids,     # full list from folder
                    "siteId":    primary_site, # primary (= site_ids[0])
                                               # used for batch grouping
                    "filePath":  f,
                    "payload":   payload,
                    "status":    status,       # PUBLISHED or DRAFT — drives the fork
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
# API primitives
# ---------------------------------------------------------------------------

def chunked(xs: list[str], n: int) -> list[list[str]]:
    """Yield xs in successive lists of up to n items."""
    if n <= 0:
        return [xs]
    return [xs[i:i + n] for i in range(0, len(xs), n)]


def _post_batch(
    session: requests.Session,
    url: str,
    site_id: str,
    body: dict[str, Any],
    timeout: float,
    dry_run: bool,
    label: str,
    delay_ms: int = 0,
) -> tuple[bool, str]:
    """
    Wrapper for the three batch endpoints we hit (recover, change-state,
    change-status). All have the same envelope: POST with ids list,
    X-UpStart-Site header, returns 2xx on success.

    `label` is used only for dry-run output (so the line is informative).
    `delay_ms` adds a sleep AFTER the request completes (success or
    failure) — see put_variant for the rationale.
    """
    if dry_run:
        return True, f"DRY_RUN POST {url} site={site_id} {label} ids_n={len(body.get('ids', []))}"
    try:
        r = session.post(
            url,
            headers={"X-UpStart-Site": site_id},
            json=body,
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


def batch_recover(
    session: requests.Session,
    api_host: str,
    site_id: str,
    variant_ids: list[str],
    timeout: float = 60.0,
    dry_run: bool = False,
    delay_ms: int = 0,
) -> tuple[bool, str]:
    """ POST /pim/batch/variants/recover -- ARCHIVED → NORMAL """
    url = f"{api_host.rstrip('/')}/pim/batch/variants/recover"
    return _post_batch(
        session, url, site_id,
        {"ids": variant_ids},
        timeout, dry_run, label="recover", delay_ms=delay_ms,
    )


def batch_change_status(
    session: requests.Session,
    api_host: str,
    site_id: str,
    variant_ids: list[str],
    new_status: str,           # "READY" or "COMMIT"
    timeout: float = 60.0,
    dry_run: bool = False,
    delay_ms: int = 0,
) -> tuple[bool, str]:
    """ POST /pim/batch/variants/change-status """
    url = f"{api_host.rstrip('/')}/pim/batch/variants/change-status"
    return _post_batch(
        session, url, site_id,
        {"ids": variant_ids, "status": new_status},
        timeout, dry_run, label=f"change-status={new_status}", delay_ms=delay_ms,
    )


def batch_change_state_archived(
    session: requests.Session,
    api_host: str,
    site_id: str,
    variant_ids: list[str],
    timeout: float = 60.0,
    dry_run: bool = False,
    delay_ms: int = 0,
) -> tuple[bool, str]:
    """
    POST /pim/batch/variants/change-state/ARCHIVED  -- NORMAL → ARCHIVED.

    Note: the target state is part of the URL path, unlike change-status
    where it's in the body. Mirror the platform's actual contract.
    """
    url = f"{api_host.rstrip('/')}/pim/batch/variants/change-state/ARCHIVED"
    return _post_batch(
        session, url, site_id,
        {"ids": variant_ids},
        timeout, dry_run, label="change-state=ARCHIVED", delay_ms=delay_ms,
    )


def put_variant(
    session: requests.Session,
    api_host: str,
    variant_id: str,
    site_ids: list[str],
    payload: dict[str, Any],
    timeout: float = 30.0,
    dry_run: bool = False,
    delay_ms: int = 0,
) -> tuple[bool, str]:
    """
    PUT /pim/variants/<id> with payload body.

    Multi-site contract (same as push_variants.put_variant):
      X-UpStart-Site: <site_ids[0]>           (header)
      ?sites=<site_ids[1]>,<site_ids[2]>,...  (query, omitted if single-site)

    `delay_ms` adds a sleep AFTER the request completes — mimics
    portal-paced traffic. With workers > 1 the delay is per-worker.
    """
    if not site_ids:
        return False, "no site_ids provided"

    primary_site = site_ids[0]
    other_sites = site_ids[1:]
    url = f"{api_host.rstrip('/')}/pim/variants/{variant_id}"

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


# ---------------------------------------------------------------------------
# Phase dispatchers
# ---------------------------------------------------------------------------
#
# Each phase takes a set of (siteId → list of variant ids) and runs the
# appropriate batch call in chunks, across the worker pool. Returns a new
# (siteId → list of ids) dict containing ONLY the variants whose batch
# succeeded — so the next phase naturally skips anything that failed.
#
# Why pass siteId → ids dicts between phases:
#   - sites are isolated (each batch is single-site), so the partitioning
#     never crosses sites
#   - if a batch fails partway, we lose a whole batch's worth of variants
#     for subsequent phases but no variants from other batches/sites are
#     affected. The "lose the batch" granularity matches how the API
#     itself fails.

def _run_batched_phase(
    work_by_site: dict[str, list[str]],
    runner,                # callable(site_id, batch_ids) -> (ok, info)
    phase_name: str,
    batch_size: int,
    workers: int,
) -> dict[str, list[str]]:
    """
    Generic batched-phase dispatcher.

    `runner(site_id, batch_ids)` should return (ok, info_string).
    On ok, the batch's variants are added to the returned dict so the
    NEXT phase processes them. On failure, they are logged and dropped
    from future phases (safer side of partial failure for the archive
    flow — better to leave a variant un-rearchived than to push it
    further along a broken pipeline).
    """
    ok_log = logging.getLogger("archived.ok")
    err_log = logging.getLogger("archived.err")
    log = logging.getLogger("archived")

    tasks = [
        (site, batch)
        for site, vids in work_by_site.items()
        for batch in chunked(vids, batch_size)
    ]
    total_variants = sum(len(v) for v in work_by_site.values())
    log.info(
        "[%s] dispatching %d batches across %d variants in %d sites (batch_size=%d)",
        phase_name, len(tasks), total_variants, len(work_by_site), batch_size,
    )

    surviving: dict[str, list[str]] = defaultdict(list)
    surviving_lock = threading.Lock()
    ok_total = fail_total = 0

    with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futs = {pool.submit(runner, site, batch): (site, batch) for site, batch in tasks}
        for fut in as_completed(futs):
            site, batch = futs[fut]
            ok, info = fut.result()
            if ok:
                ok_log.info(
                    "phase=%s verdict=ok siteId=%s n=%d detail=%s",
                    phase_name, site, len(batch), info,
                )
                with surviving_lock:
                    surviving[site].extend(batch)
                ok_total += len(batch)
            else:
                err_log.error(
                    "phase=%s verdict=failed siteId=%s n=%d ids=%s detail=%s",
                    phase_name, site, len(batch), batch[:5], info,
                )
                fail_total += len(batch)

    log.info(
        "[%s] done: variants_through=%d variants_lost=%d",
        phase_name, ok_total, fail_total,
    )
    return dict(surviving)


# Specialized PUT phase — per-variant (not batched) because PUT URLs are
# variant-scoped. Mirrors push_variants.py's PUT phase.

def _run_put_phase(
    session: requests.Session,
    api_host: str,
    work_by_site: dict[str, list[str]],
    items_by_id: dict[str, dict[str, Any]],
    workers: int,
    dry_run: bool,
    delay_ms: int = 0,
) -> dict[str, list[str]]:
    """
    PUT each variant. Returns siteId → list of variants whose PUT
    succeeded (eligible for the next phase).

    `delay_ms` is forwarded to put_variant — see its docstring.
    """
    ok_log = logging.getLogger("archived.ok")
    err_log = logging.getLogger("archived.err")
    log = logging.getLogger("archived")

    all_tasks: list[tuple[str, str]] = [
        (site, vid)
        for site, vids in work_by_site.items()
        for vid in vids
    ]
    log.info(
        "[PUT] dispatching %d variants across %d workers",
        len(all_tasks), workers,
    )

    surviving: dict[str, list[str]] = defaultdict(list)
    surviving_lock = threading.Lock()
    ok_total = fail_total = 0

    def _do_one(site_id: str, variant_id: str):
        """
        `site_id` here is the variant's PRIMARY site (used for batch grouping
        upstream). PUT itself needs the full siteIds list — primary in
        header, rest in ?sites= query. Look it up from items_by_id.
        """
        item = items_by_id[variant_id]
        return put_variant(
            session, api_host, variant_id,
            item["siteIds"],            # full list, not just primary
            item["payload"],
            dry_run=dry_run, delay_ms=delay_ms,
        )

    with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futs = {
            pool.submit(_do_one, site, vid): (site, vid)
            for site, vid in all_tasks
        }
        for fut in as_completed(futs):
            site, vid = futs[fut]
            ok, info = fut.result()
            if ok:
                ok_log.info(
                    "phase=PUT verdict=ok variantId=%s siteId=%s detail=%s",
                    vid, site, info,
                )
                with surviving_lock:
                    surviving[site].append(vid)
                ok_total += 1
            else:
                err_log.error(
                    "phase=PUT verdict=failed variantId=%s siteId=%s detail=%s",
                    vid, site, info,
                )
                fail_total += 1

    log.info("[PUT] done: ok=%d fail=%d", ok_total, fail_total)
    return dict(surviving)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--input", required=True, type=Path,
                    help="Output dir from repair_variants.py (must contain "
                         "variant_map.json and <tpl>/<site>/<vid>.json files)")
    ap.add_argument(
        "--api-host",
        default="https://nochannel-test-1-api.nochannel-test.upstart.team",
        help="API host (no trailing slash)",
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
    ap.add_argument("--workers", type=int, default=1,
                    help="Concurrent PUT/batch workers (default 1)")
    ap.add_argument("--batch-size", type=int, default=1,
                    help="Max variants per batch call (default 1)")
    ap.add_argument("--delay-ms", type=int, default=100,
                    help="Sleep this many milliseconds after each API request "
                         "(per worker), to mimic portal-paced traffic. Default 100. "
                         "Set to 0 to disable.")
    ap.add_argument("--limit", type=int, default=0,
                    help="Process only first N candidates (0=all). For testing.")
    ap.add_argument("--dry-run", action="store_true",
                    help="Log what would happen; no HTTP requests made.")
    ap.add_argument(
        "--stop-after",
        choices=["scan", "recover", "put", "publish", "archive", "all"],
        default="all",
        help="Stop the pipeline after the named phase. "
             "Useful for staged rollouts and debugging. Default: run everything.",
    )
    ap.add_argument("--log-dir", type=Path, default=None,
                    help="Where logs go. Default: <input>/logs")
    ap.add_argument("--verbose", "-v", action="store_true",
                    help="Enable DEBUG-level logging")
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

    # --- load variant_map ---
    map_path = args.input / "variant_map.json"
    if not map_path.exists():
        log.error("variant_map.json not found at %s", map_path)
        return 1
    variant_map = json.loads(map_path.read_text())
    log.info("loaded variant_map with %d entries", len(variant_map))

    # --- scan + filter ---
    log.info("[1/6] scan: looking for archived-PUBLISHED and archived-DRAFT candidates")
    items = collect_archived_items(args.input, variant_map)
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

    # Partition candidates by status. The DRAFT path is one step (just PUT);
    # the PUBLISHED path is the full 7-step pipeline. Both keep per-variant
    # site bindings so each call goes to the right site.
    published_items = [it for it in items if it["status"] == "PUBLISHED"]
    draft_items     = [it for it in items if it["status"] == "DRAFT"]
    items_by_id: dict[str, dict[str, Any]] = {it["variantId"]: it for it in items}
    log.info(
        "candidates partitioned: published=%d (full 7-step pipeline)  draft=%d (PUT only)",
        len(published_items), len(draft_items),
    )

    t_run = time.time()
    any_progress = False

    # ----------------------------------------------------------------------
    # DRAFT path — just PUT, no state changes.
    # An archived variant with a DRAFT status means there's an editable
    # draft revision attached to the archived record. The platform allows
    # PUTting directly to that draft; no recover/republish/rearchive needed.
    # ----------------------------------------------------------------------
    if draft_items:
        log.info("[draft/1] PUT phase: %d archived-DRAFT variants (no state changes)",
                 len(draft_items))
        draft_by_site: dict[str, list[str]] = defaultdict(list)
        for it in draft_items:
            draft_by_site[it["siteId"]].append(it["variantId"])
        draft_survivors = _run_put_phase(
            session, args.api_host, draft_by_site, items_by_id,
            workers=args.workers, dry_run=args.dry_run, delay_ms=args.delay_ms,
        )
        n = sum(len(v) for v in draft_survivors.values())
        log.info("[draft/1] done: %d/%d archived-DRAFT variants updated",
                 n, len(draft_items))
        if n:
            any_progress = True

    # ----------------------------------------------------------------------
    # PUBLISHED path — the full 7-step pipeline.
    # Each phase consumes the survivors of the previous one. A variant that
    # drops out at any step doesn't get touched again by this run.
    # ----------------------------------------------------------------------
    if not published_items:
        log.info("no archived-PUBLISHED candidates; skipping full pipeline")
        return _final_summary(log, run_id, t_run, success=any_progress)

    work_by_site: dict[str, list[str]] = defaultdict(list)
    for it in published_items:
        work_by_site[it["siteId"]].append(it["variantId"])
    log.info(
        "[published path] %d variants across %d sites",
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
        return _final_summary(log, run_id, t_run, success=any_progress or bool(survivors))

    # [3/6] PUT  (per-variant)
    log.info("[3/6] PUT phase: pushing repaired payloads")
    survivors = _run_put_phase(
        session, args.api_host, survivors, items_by_id,
        workers=args.workers, dry_run=args.dry_run, delay_ms=args.delay_ms,
    )
    if args.stop_after == "put" or not survivors:
        log.info("stopping after PUT (stop_after=%s, survivors=%d)",
                 args.stop_after, sum(len(v) for v in survivors.values()))
        return _final_summary(log, run_id, t_run, success=any_progress or bool(survivors))

    # [4/6] COMMIT (post-PUT) — publish the recovered+updated revision.
    # READY is removed per operator direction; we go straight from PUT to
    # COMMIT. A COMMIT failure here leaves the variant in NORMAL state
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
        return _final_summary(log, run_id, t_run, success=any_progress or bool(survivors))

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
            "no survivors after archive; variants are stuck in NORMAL state — "
            "needs manual rescue. See errors.log."
        )
        return _final_summary(log, run_id, t_run, success=any_progress)

    # [6/6] COMMIT (post-archive) — finalize the archived state.
    # READY-2 is removed per operator direction; we COMMIT directly. A
    # failure here leaves variants in ARCHIVED state with an uncommitted
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

    return _final_summary(log, run_id, t_run, success=any_progress or bool(survivors))


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