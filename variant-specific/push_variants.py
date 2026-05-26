#!/usr/bin/env python3
"""
Variant push script — companion to repair_variants.py.

Reads what repair_variants.py produced (variant_map.json + the
<productTemplateId>/<siteId>/<variantId>.json payload files) and pushes
those repaired payloads back to the PIM API.

Decision matrix (status from variant_map.json):

    status=PUBLISHED  →  PUT payload, then COMMIT  (republish, no READY)
    status=DRAFT      →  PUT payload only          (leave as draft)
    status=null       →  skip entirely             (fetch never worked)

Why this matrix:
    - A variant that was PUBLISHED needs its update to go live. We PUT
      then issue COMMIT directly (READY is intentionally skipped per
      operator direction — see note in run_publish_phase). A COMMIT
      failure leaves the variant in a stuck mid-write state; the operator
      must re-run or fix manually.
    - A variant that was DRAFT was deliberately kept in draft (someone is
      mid-edit); we don't want to surprise them by publishing changes.
      Just save the update as a draft revision.
    - A null status means the fetch never succeeded, so we have no
      authoritative pre-state and shouldn't be making changes.

Site handling:
    The payload file's folder name IS the site that returned the data
    during the fetch phase. PUT and the batch change-status calls go to
    that same site (X-UpStart-Site header). No multi-site retry on push —
    we know exactly where each variant lives.

Batch handling:
    PUT is per-variant (the URL is variant-scoped). COMMIT uses the batch
    endpoint `/pim/batch/variants/change-status` which accepts
    `{"ids":[...]}`. We batch COMMIT requests by site, in chunks of
    --batch-size (default 50). Cuts network round trips on the ~15k
    variant workload from ~15k+8k down to ~15k+160.

Auth:
    Same as repair_variants.py — Cookie + session-id headers. The script
    can also read API_COOKIE / API_SESSION_ID from env.

Usage:
    python push_variants.py \\
        --input ./out \\
        --workers 8 \\
        --dry-run            # see what would happen, no requests sent

    # for real:
    python push_variants.py --input ./out --workers 8
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import threading
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any

import requests


# ---------------------------------------------------------------------------
# Logging — same three-file pattern as repair_variants.py
# ---------------------------------------------------------------------------
LOG_FMT = "%(asctime)s | %(levelname)-7s | %(name)-10s | %(message)s"
LOG_DATEFMT = "%Y-%m-%d %H:%M:%S"


def setup_logging(log_dir: Path, verbose: bool = False) -> logging.Logger:
    """
    Three loggers / three files (run/success/errors), just like the
    repair script. See repair_variants.py setup_logging for rationale.
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

    main = logging.getLogger("push")
    main.setLevel(logging.DEBUG if verbose else logging.INFO)
    main.propagate = False
    main.handlers.clear()
    main.addHandler(_rotating(log_dir / "push.run.log", logging.DEBUG if verbose else logging.INFO))
    console = logging.StreamHandler(sys.stderr)
    console.setLevel(logging.INFO)
    console.setFormatter(formatter)
    main.addHandler(console)

    success = logging.getLogger("push.ok")
    success.setLevel(logging.INFO)
    success.propagate = False
    success.handlers.clear()
    success.addHandler(_rotating(log_dir / "push.success.log", logging.INFO))

    error = logging.getLogger("push.err")
    error.setLevel(logging.WARNING)
    error.propagate = True  # also lands in run.log and console
    error.handlers.clear()
    error.addHandler(_rotating(log_dir / "push.errors.log", logging.WARNING))

    main.info("push logging initialised: dir=%s verbose=%s", log_dir, verbose)
    return main


# ---------------------------------------------------------------------------
# HTTP session
# ---------------------------------------------------------------------------

def make_session(
    session_id: str,
    cookie_header: str | None = None,
) -> requests.Session:
    """
    Session shared across all workers. X-UpStart-Site is NOT set here —
    it goes per-request (variants live on different sites, and mutating
    session.headers across worker threads would be a race).
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

def collect_work_items(
    input_dir: Path,
    variant_map: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Scan input_dir for repaired payload files and pair each with its
    variant_map entry so we know the status (PUBLISHED / DRAFT / null).

    Expected layout (produced by repair_variants.py):
        <input_dir>/<productTemplateId>/<siteId>/<variantId>.json

    Returns one dict per variant:
        {
          "variantId":  ...,
          "siteId":     ...     (from the folder name)
          "status":     ...     (from variant_map.json)
          "filePath":   Path    (absolute path to the payload file)
          "payload":    dict    (already parsed)
        }

    Variants in variant_map but missing a file are reported and skipped.
    Files on disk but not in variant_map are also reported and skipped
    (probably stale from an earlier run).
    """
    log = logging.getLogger("push")
    err_log = logging.getLogger("push.err")

    items: list[dict[str, Any]] = []
    found_ids: set[str] = set()
    # Tally of archived variants we skipped (handled by repair_archived_variants
    # instead). Surfaced in the run log so it's obvious if a non-trivial chunk
    # of the input belongs to the other script.
    n_skipped_archived = 0

    # Walk <input_dir>/<tplId>/<siteIds-joined>/<vid>.json
    # The site folder name is the comma-joined siteIds from the variant's
    # original payload (in original order — see repair_variants.write_payload).
    # Single-site variants have a folder name like "siteA"; multi-site
    # variants have "siteA,siteB". We split on comma to recover the list.
    for tpl_dir in sorted(input_dir.iterdir()):
        if not tpl_dir.is_dir():
            continue
        if tpl_dir.name in {"logs"}:
            continue
        for site_dir in sorted(tpl_dir.iterdir()):
            if not site_dir.is_dir():
                continue
            # Parse folder name into site list. The first element goes in
            # the X-UpStart-Site header; the rest go in the ?sites= query.
            site_ids = [s for s in site_dir.name.split(",") if s]
            if not site_ids:
                err_log.warning(
                    "bad_site_folder name=%s reason=empty_after_split",
                    site_dir.name,
                )
                continue
            for f in sorted(site_dir.glob("*.json")):
                vid = f.stem
                found_ids.add(vid)
                if vid not in variant_map:
                    err_log.warning(
                        "orphan_file variantId=%s path=%s reason=not_in_variant_map",
                        vid, f,
                    )
                    continue

                meta = variant_map[vid]

                # Skip archived variants — those are handled by
                # repair_archived_variants.py, which has the
                # recover→PUT→republish→re-archive flow. Trying to PUT
                # directly to an archived variant in the PUBLISHED path
                # would fail at the platform level (you can't update an
                # archived record). For archived-DRAFT the PUT would
                # succeed but it's still not this script's job.
                if meta.get("state") == "ARCHIVED":
                    n_skipped_archived += 1
                    continue

                try:
                    payload = json.loads(f.read_text())
                except json.JSONDecodeError as e:
                    err_log.error(
                        "bad_payload_file variantId=%s path=%s detail=%r",
                        vid, f, str(e),
                    )
                    continue
                items.append({
                    "variantId": vid,
                    "siteIds":   site_ids,  # full list (folder-derived)
                    "siteId":    site_ids[0],  # the "primary" site for this variant
                    "status":    meta.get("status"),
                    "filePath":  f,
                    "payload":   payload,
                })

    if n_skipped_archived:
        log.info(
            "skipped %d archived variants (state==ARCHIVED) — "
            "these belong to repair_archived_variants.py, not this script",
            n_skipped_archived,
        )

    # Surface map entries that have no file on disk (probably fetch_failed
    # last time). Don't fail the run; just note them.
    map_missing_file = [
        vid for vid, meta in variant_map.items()
        if vid not in found_ids and meta.get("status") is not None
    ]
    if map_missing_file:
        log.warning(
            "%d variants in variant_map have non-null status but no payload "
            "file on disk; skipping them. First few: %s",
            len(map_missing_file), map_missing_file[:5],
        )

    return items


# ---------------------------------------------------------------------------
# Individual API calls
# ---------------------------------------------------------------------------

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
    PUT one variant's payload back to the API.

    Multi-site contract
    -------------------
    `site_ids` is the list of sites this variant lives on (from the folder
    name created by repair_variants.py). The platform expects:

      X-UpStart-Site: <site_ids[0]>            (header — the "primary"
                                                site for this request)
      ?sites=<site_ids[1]>,<site_ids[2]>,...   (query — additional sites
                                                this update should apply to)

    For a single-site variant, the query string just has nothing after `?`
    or — if we pass an empty list — we omit the `sites` param entirely.

    Auth: per-request X-UpStart-Site header (the same session works across
    sites and threads safely; mutating session-level headers from worker
    threads would race).

    `delay_ms` adds a sleep AFTER the request completes (success or
    failure). With workers > 1 the delay is per-worker, so effective
    rate is workers / (delay_ms/1000).

    Returns (ok, info_string). info is the path on success or the error
    summary on failure.
    """
    if not site_ids:
        return False, "no site_ids provided"

    primary_site = site_ids[0]
    other_sites = site_ids[1:]
    url = f"{api_host.rstrip('/')}/pim/variants/{variant_id}"

    # Build the request kwargs once so dry-run logging shows the same
    # things the real request would send.
    request_kwargs: dict[str, Any] = {
        "headers": {"X-UpStart-Site": primary_site},
        "json":    payload,
        "timeout": timeout,
    }
    if other_sites:
        # Single `sites=` param, comma-joined.
        request_kwargs["params"] = {"sites": ",".join(other_sites)}

    if dry_run:
        params_part = ""
        if other_sites:
            params_part = f"?sites={','.join(other_sites)}"
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
    # Truncate body so a runaway HTML page doesn't drown the log.
    return False, f"http {r.status_code} body={r.text[:300]!r}"


def batch_change_status(
    session: requests.Session,
    api_host: str,
    site_id: str,
    variant_ids: list[str],
    new_status: str,  # "READY" or "COMMIT"
    timeout: float = 60.0,
    dry_run: bool = False,
    delay_ms: int = 0,
) -> tuple[bool, str]:
    """
    POST `{"ids":[...], "status": new_status}` to the batch endpoint.

    The platform's contract: a successful batch call applies the status
    to ALL ids in the request, atomically per the site. (If that turns
    out not to be atomic, individual failures will need split-on-error
    retry — not implemented yet; flag if you see partial results.)

    `delay_ms` sleeps after the request completes — see put_variant.

    Returns (ok, info_string).
    """
    url = f"{api_host.rstrip('/')}/pim/batch/variants/change-status"
    body = {"ids": variant_ids, "status": new_status}

    if dry_run:
        return True, f"DRY_RUN POST {url} site={site_id} n={len(variant_ids)} status={new_status}"

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
        return True, f"http {r.status_code} n={len(variant_ids)}"
    return False, f"http {r.status_code} body={r.text[:300]!r}"


# ---------------------------------------------------------------------------
# Per-variant pipeline (PUT only — the batch state transitions happen later)
# ---------------------------------------------------------------------------

def push_variant(
    session: requests.Session,
    api_host: str,
    item: dict[str, Any],
    dry_run: bool = False,
    delay_ms: int = 0,
) -> dict[str, Any]:
    """
    Run the PUT phase for one variant. Returns a result dict; the caller
    aggregates these to decide which variants need the subsequent batch
    COMMIT calls.

    Status-driven flow:
      PUBLISHED → PUT, then mark variant for COMMIT (no READY anymore)
      DRAFT     → PUT only, no state transition
      null      → skip (shouldn't be in the input items, but defensive)

    `delay_ms` is forwarded to put_variant — see its docstring.
    """
    ok_log = logging.getLogger("push.ok")
    err_log = logging.getLogger("push.err")

    variant_id = item["variantId"]
    site_ids   = item["siteIds"]      # full list from folder name
    primary    = site_ids[0]           # X-UpStart-Site header value
    status     = item["status"]

    if status is None:
        # Defensive — collect_work_items already filters these, but if a
        # caller bypasses that we want a clean skip rather than wrong
        # behavior. ERROR (not just info) because something upstream
        # shouldn't have passed it through.
        err_log.error(
            "skipped variantId=%s reason=status_null", variant_id,
        )
        return {
            "variantId":    variant_id,
            "putOk":        False,
            "needsPublish": False,
            "reason":       "status_null",
        }

    t0 = time.monotonic()
    put_ok, put_info = put_variant(
        session, api_host, variant_id, site_ids, item["payload"],
        dry_run=dry_run, delay_ms=delay_ms,
    )
    elapsed_ms = int((time.monotonic() - t0) * 1000)

    if not put_ok:
        err_log.error(
            "verdict=put_failed variantId=%s siteIds=%s status=%s elapsed_ms=%d detail=%s",
            variant_id, site_ids, status, elapsed_ms, put_info,
        )
        return {
            "variantId":    variant_id,
            "siteIds":      site_ids,
            "primarySite":  primary,
            "status":       status,
            "putOk":        False,
            "needsPublish": False,
            "reason":       put_info,
        }

    needs_publish = (status == "PUBLISHED")
    ok_log.info(
        "verdict=put_ok variantId=%s siteIds=%s status=%s elapsed_ms=%d needs_publish=%s",
        variant_id, site_ids, status, elapsed_ms, needs_publish,
    )
    return {
        "variantId":    variant_id,
        "siteIds":      site_ids,
        "primarySite":  primary,
        "status":       status,
        "putOk":        True,
        "needsPublish": needs_publish,
        "reason":       put_info,
    }


# ---------------------------------------------------------------------------
# Batch publish phase
# ---------------------------------------------------------------------------

def chunked(xs: list[str], n: int) -> list[list[str]]:
    """Split xs into successive chunks of size <= n."""
    if n <= 0:
        return [xs]
    return [xs[i:i + n] for i in range(0, len(xs), n)]


def run_publish_phase(
    session: requests.Session,
    api_host: str,
    to_publish_by_site: dict[str, list[str]],
    batch_size: int,
    workers: int,
    dry_run: bool = False,
    delay_ms: int = 0,
) -> tuple[int, int, int, int]:
    """
    Apply COMMIT directly to every variant that needs publishing.

    READY is intentionally skipped (per operator direction). The platform
    accepts a COMMIT without a prior READY for these variants; the READY
    gate that previously protected against partial-write states is gone.

    Tradeoff
    --------
    Without READY, a COMMIT failure means the variant change is stuck
    mid-write (no rollback). Errors land in errors.log with the failing
    batch and ids; the operator must re-run or fix manually. We accepted
    this on the way in — see the run notes.

    `delay_ms` is forwarded to batch_change_status — see its docstring.

    Returns (ready_ok, ready_failed, commit_ok, commit_failed) for caller
    compatibility; ready_* are always 0 since no READY calls happen.
    """
    log = logging.getLogger("push")
    err_log = logging.getLogger("push.err")
    ok_log = logging.getLogger("push.ok")

    commit_ok = commit_fail = 0

    # Build COMMIT tasks directly from the work map. Each (site, batch)
    # tuple becomes one POST /pim/batch/variants/change-status call with
    # status=COMMIT.
    commit_tasks = [
        (site, batch)
        for site, vids in to_publish_by_site.items()
        for batch in chunked(vids, batch_size)
    ]
    log.info(
        "publish phase (COMMIT-only, no READY): %d sites, %d total variants, "
        "%d COMMIT batches (batch_size=%d)",
        len(to_publish_by_site),
        sum(len(v) for v in to_publish_by_site.values()),
        len(commit_tasks),
        batch_size,
    )

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
                    "verdict=commit_ok siteId=%s n=%d detail=%s",
                    site, len(batch), info,
                )
            else:
                commit_fail += len(batch)
                err_log.error(
                    "verdict=commit_failed siteId=%s n=%d ids=%s detail=%s",
                    site, len(batch), batch[:5], info,
                )

    # ready_ok and ready_fail are always 0 — READY is skipped entirely.
    return 0, 0, commit_ok, commit_fail


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--input", required=True, type=Path,
                    help="Output dir from repair_variants.py")
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
    ap.add_argument(
        "--workers", type=int, default=1,
        help="Concurrent PUT/batch workers (default 1)",
    )
    ap.add_argument(
        "--batch-size", type=int, default=1,
        help="Max variants per READY/COMMIT batch (default 1)",
    )
    ap.add_argument(
        "--delay-ms", type=int, default=100,
        help="Sleep this many milliseconds after each API request (per "
             "worker), to mimic portal-paced traffic. Default 100. Set to "
             "0 to disable.",
    )
    ap.add_argument(
        "--limit", type=int, default=0,
        help="Process only first N variants (0=all). For testing.",
    )
    ap.add_argument(
        "--dry-run", action="store_true",
        help="Don't actually call the API. Logs would-be requests instead.",
    )
    ap.add_argument(
        "--skip-publish", action="store_true",
        help="Run PUT phase only — skip READY+COMMIT for PUBLISHED variants. "
             "Useful for testing PUT before going all-in.",
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
        "run_start run_id=%s input=%s dry_run=%s workers=%d batch_size=%d delay_ms=%d",
        run_id, args.input, args.dry_run,
        args.workers, args.batch_size, args.delay_ms,
    )

    # --- load variant_map ---
    map_path = args.input / "variant_map.json"
    if not map_path.exists():
        log.error("variant_map.json not found at %s", map_path)
        return 1
    variant_map = json.loads(map_path.read_text())
    log.info("loaded variant_map with %d entries", len(variant_map))

    # --- discover work items ---
    log.info("[1/3] scanning %s for payload files", args.input)
    items = collect_work_items(args.input, variant_map)
    log.info("found %d payload files paired with variant_map entries", len(items))

    # Status breakdown — gives a clear "what's about to happen" preview
    from collections import Counter as _Counter
    status_counts = _Counter(it["status"] for it in items)
    log.info(
        "status breakdown: PUBLISHED=%d (will PUT+READY+COMMIT), "
        "DRAFT=%d (will PUT only), null=%d (will skip)",
        status_counts.get("PUBLISHED", 0),
        status_counts.get("DRAFT", 0),
        status_counts.get(None, 0),
    )

    # Drop null-status items here so push_variant doesn't have to defend.
    items = [it for it in items if it["status"] is not None]
    if args.limit > 0:
        items = items[: args.limit]
        log.info("--limit %d applied; processing first %d", args.limit, len(items))

    if not items:
        log.info("nothing to push. run_id=%s", run_id)
        return 0

    if not args.cookie and not args.dry_run:
        log.warning(
            "no Cookie provided — if the API gateway requires session cookies "
            "you'll get HTML responses. Pass --cookie or set API_COOKIE."
        )

    session = make_session(args.session_id, cookie_header=args.cookie or None)

    # --- PUT phase ---
    log.info("[2/3] PUT phase: %d variants across %d workers (dry_run=%s)",
             len(items), args.workers, args.dry_run)
    t0 = time.time()
    put_ok = put_fail = 0
    to_publish_by_site: dict[str, list[str]] = defaultdict(list)
    put_lock = threading.Lock()

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        futs = {
            pool.submit(push_variant, session, args.api_host, it, args.dry_run, args.delay_ms): it["variantId"]
            for it in items
        }
        for fut in as_completed(futs):
            result = fut.result()
            with put_lock:
                if result["putOk"]:
                    put_ok += 1
                    if result["needsPublish"]:
                        # Bucket by PRIMARY site (= site_ids[0]). The batch
                        # change-status endpoint takes only one site at a
                        # time via X-UpStart-Site, and the variant's primary
                        # is the one we sent its PUT to via the header, so
                        # READY/COMMIT for that variant go to the same site.
                        # The other siteIds were applied transitively via
                        # the ?sites= query on the PUT, so they don't need
                        # separate READY/COMMIT calls.
                        to_publish_by_site[result["primarySite"]].append(result["variantId"])
                else:
                    put_fail += 1
            done = put_ok + put_fail
            if done % 50 == 0:
                log.info(
                    "put progress %d/%d (ok=%d fail=%d)",
                    done, len(items), put_ok, put_fail,
                )

    elapsed_put = time.time() - t0
    log.info(
        "PUT phase done: ok=%d fail=%d to_publish=%d elapsed_s=%.1f",
        put_ok, put_fail,
        sum(len(v) for v in to_publish_by_site.values()),
        elapsed_put,
    )

    # --- READY + COMMIT phase ---
    ready_ok = ready_fail = commit_ok = commit_fail = 0
    if args.skip_publish:
        log.info("[3/3] --skip-publish set; not running READY/COMMIT")
    elif not to_publish_by_site:
        log.info("[3/3] nothing to publish (no PUT-succeeded PUBLISHED variants)")
    else:
        log.info("[3/3] publish phase (READY → COMMIT)")
        t1 = time.time()
        ready_ok, ready_fail, commit_ok, commit_fail = run_publish_phase(
            session, args.api_host,
            to_publish_by_site,
            batch_size=args.batch_size,
            workers=args.workers,
            dry_run=args.dry_run,
            delay_ms=args.delay_ms,
        )
        elapsed_pub = time.time() - t1
        log.info(
            "publish phase done: ready_ok=%d ready_fail=%d commit_ok=%d "
            "commit_fail=%d elapsed_s=%.1f",
            ready_ok, ready_fail, commit_ok, commit_fail, elapsed_pub,
        )

    elapsed_total = time.time() - t0
    log.info(
        "run_end run_id=%s put_ok=%d put_fail=%d ready_ok=%d ready_fail=%d "
        "commit_ok=%d commit_fail=%d elapsed_s=%.1f",
        run_id, put_ok, put_fail, ready_ok, ready_fail,
        commit_ok, commit_fail, elapsed_total,
    )

    # Non-zero exit if anything failed, so CI/cron wrappers can alert.
    any_fail = put_fail + ready_fail + commit_fail
    return 0 if any_fail == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())