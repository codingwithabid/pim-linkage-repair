#!/usr/bin/env python3
"""
push_products_negative.py
=========================

Negative-pipeline counterpart of push_products.py.

Same logic, same decision matrix — reads `product_map_negative.json`
(the output of repair_products_negative.py) and the repaired product
payloads, then pushes each non-archived product back to the platform.

The two pipelines stay isolated by reading different map files, so
operators can run them in either order or independently.

Filter
------
Only products in `product_map_negative.json` where `state != ARCHIVED`
are processed. Archived products are skipped at scan time — those are
handled by repair_archived_products_negative.py (a separate script with
the full recover→PUT→archive flow). The skipped-archived count is
surfaced in the run log.

Decision matrix (status from product_map_negative.json):

    status=PUBLISHED  →  PUT payload, then COMMIT  (republish, no READY)
    status=DRAFT      →  PUT payload only          (leave as draft)
    status=null       →  skip entirely             (fetch never worked)

Why this matrix:
    - A product that was PUBLISHED needs its update to go live. We PUT
      then issue COMMIT directly (READY is intentionally skipped per
      operator direction — see note in run_publish_phase). A COMMIT
      failure leaves the product in a stuck mid-write state; the operator
      must re-run or fix manually.
    - A product that was DRAFT was deliberately kept in draft (someone is
      mid-edit); we don't want to surprise them by publishing changes.
      Just save the update as a draft revision.
    - A null status means the fetch never succeeded, so we have no
      authoritative pre-state and shouldn't be making changes.

Batch handling:
    PUT is per-product (the URL is product-scoped). COMMIT uses the batch
    endpoint `/pim/batch/products/change-status` which accepts
    `{"ids":[...]}`. We batch COMMIT requests by site, in chunks of
    --batch-size (default 50). Cuts network round trips.

Concurrency:
    PUT phase uses a thread pool sized by --workers (default 8). One
    requests.Session is shared across workers; X-UpStart-Site is
    per-request (varies between products on different sites).

Idempotency:
    PUT replaces the product payload, so re-running this script on the
    same input directory is safe — the platform receives the same bytes
    each time. Re-running COMMIT on an already-COMMITTED product is a
    no-op on the platform side (or returns an error that's logged but
    doesn't break the run).

Multi-site PUT contract (matches push_variants.py):
    For products living on multiple sites:
      X-UpStart-Site: <first site>
      ?sites=<comma-joined rest>
    The platform applies the PUT to all sites in one call. Single-site
    products omit the ?sites= query parameter entirely.

Usage
-----
    export API_COOKIE='...'
    python3 push_products_negative.py --input ./out-products-negative --workers 4

Files read
----------
    <input>/product_map_negative.json
    <input>/<tpl>/<siteIds-joined>/<productId>.json

Files written
-------------
    <input>/logs/push-products-negative.run.log
    <input>/logs/push-products-negative.success.log
    <input>/logs/push-products-negative.errors.log
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
# Site configuration — same as repair_products_negative.py
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
    Periodic progress logger. Logs at every `step` items OR every
    `interval_s` seconds, whichever comes first. Step auto-computed as
    ~10% of total (clamped to 10..200). Thread-safe.
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
        self._last_log_t = time.monotonic()
        self._t0 = time.monotonic()
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

      push-products-negative.run.log      — everything (DEBUG if --verbose else INFO)
      push-products-negative.success.log  — per-product verdicts (INFO only)
      push-products-negative.errors.log   — WARNING+ (propagates to run.log)
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
    main.addHandler(_rotating(
        log_dir / "push-products-negative.run.log",
        logging.DEBUG if verbose else logging.INFO,
    ))
    console = logging.StreamHandler(sys.stderr)
    console.setLevel(logging.INFO)
    console.setFormatter(formatter)
    main.addHandler(console)

    ok = logging.getLogger("push.ok")
    ok.setLevel(logging.INFO)
    ok.propagate = False
    ok.handlers.clear()
    ok.addHandler(_rotating(log_dir / "push-products-negative.success.log", logging.INFO))

    err = logging.getLogger("push.err")
    err.setLevel(logging.WARNING)
    err.propagate = True
    err.handlers.clear()
    err.addHandler(_rotating(log_dir / "push-products-negative.errors.log", logging.WARNING))

    main.info("push-products logging initialised: dir=%s verbose=%s",
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

def collect_work_items(
    input_dir: Path,
    product_map: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Walk <input_dir>/<tpl>/<siteIds-joined>/<pid>.json and pair each with
    its product_map entry. Returns one work item per file:

        {
          "productId": str,
          "siteIds":   list[str],     # full list, folder-derived
          "siteId":    str,           # primary site = site_ids[0]
          "status":    str | None,    # from product_map
          "filePath":  Path,
          "payload":   dict,
        }

    Filtering at scan time (per operator direction):
      - state == ARCHIVED  → skip; count surfaced in log
      - product not in map → log as orphan_file; skip
      - status == null     → not filtered here (filtered in PUT phase);
                              they're still legitimate file entries

    Files with bad JSON are logged as bad_payload_file and skipped (rather
    than crashing the whole run).
    """
    log = logging.getLogger("push")
    err_log = logging.getLogger("push.err")

    items: list[dict[str, Any]] = []
    found_ids: set[str] = set()
    # Surface archived count to the run log so the operator knows how
    # many products were left for the archived-products script.
    n_skipped_archived = 0

    for tpl_dir in sorted(input_dir.iterdir()):
        if not tpl_dir.is_dir():
            continue
        if tpl_dir.name in {"logs"}:
            continue
        for site_dir in sorted(tpl_dir.iterdir()):
            if not site_dir.is_dir():
                continue
            # Folder name is comma-joined siteIds (in original payload order
            # — see repair_products.write_product_payload). Split on comma.
            site_ids = [s for s in site_dir.name.split(",") if s]
            if not site_ids:
                err_log.warning(
                    "bad_site_folder name=%s reason=empty_after_split",
                    site_dir.name,
                )
                continue
            for f in sorted(site_dir.glob("*.json")):
                pid = f.stem
                found_ids.add(pid)
                if pid not in product_map:
                    err_log.warning(
                        "orphan_file productId=%s path=%s reason=not_in_product_map",
                        pid, f,
                    )
                    continue

                meta = product_map[pid]

                # Skip archived — those are handled by repair_archived_products.py.
                # Trying to PUT directly to an archived record fails at the
                # platform level (it forbids writes to archived).
                if meta.get("state") == "ARCHIVED":
                    n_skipped_archived += 1
                    continue

                try:
                    payload = json.loads(f.read_text())
                except json.JSONDecodeError as e:
                    err_log.error(
                        "bad_payload_file productId=%s path=%s detail=%r",
                        pid, f, str(e),
                    )
                    continue
                items.append({
                    "productId": pid,
                    "siteIds":   site_ids,
                    "siteId":    site_ids[0],      # primary, for batch grouping
                    "status":    meta.get("status"),
                    "filePath":  f,
                    "payload":   payload,
                })

    if n_skipped_archived:
        log.info(
            "skipped %d archived products (state==ARCHIVED) — "
            "these belong to repair_archived_products.py, not this script",
            n_skipped_archived,
        )

    # Surface map entries that have no file on disk (probably fetch_failed
    # in repair_products_negative.py).
    map_missing_file = [
        pid for pid, meta in product_map.items()
        if pid not in found_ids and meta.get("status") is not None
    ]
    if map_missing_file:
        log.warning(
            "%d products in product_map have non-null status but no payload "
            "file on disk; skipping them. First few: %s",
            len(map_missing_file), map_missing_file[:5],
        )

    return items


# ---------------------------------------------------------------------------
# Individual API calls
# ---------------------------------------------------------------------------

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
    PUT one product's payload back to the API.

    Multi-site contract (same as push_variants):
      X-UpStart-Site: <site_ids[0]>            (header — primary site)
      ?sites=<site_ids[1]>,<site_ids[2]>,...   (query — additional sites)

    Single-site products omit the ?sites= query entirely.

    `delay_ms` adds a sleep AFTER the request completes (success or
    failure). This rate-limits the calling worker — mimics the timing
    of a human clicking through the portal so the API doesn't see a
    sustained burst from the script. With workers > 1 the delay is
    per-worker; effective rate is workers / (delay_ms / 1000).

    Returns (ok, info_string).
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

    Atomicity assumption (same as variants): a 2xx means the whole batch
    applied; anything else means none of it did. If real-world testing
    shows partial behavior, we'd need split-on-error retry.

    `delay_ms` sleeps after the request completes — see put_product.
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
# Per-product pipeline (PUT only — batch state transitions happen later)
# ---------------------------------------------------------------------------

def push_product(
    session: requests.Session,
    api_host: str,
    item: dict[str, Any],
    dry_run: bool = False,
    delay_ms: int = 0,
) -> dict[str, Any]:
    """
    Run the PUT phase for one product. Returns a result dict; the caller
    aggregates these to decide which products need the subsequent batch
    COMMIT calls.

    Status-driven flow:
      PUBLISHED → PUT, then mark product for COMMIT (no READY anymore)
      DRAFT     → PUT only, no state transition
      null      → skip (shouldn't be in the input items, but defensive)

    `delay_ms` is forwarded to put_product — see its docstring.
    """
    ok_log = logging.getLogger("push.ok")
    err_log = logging.getLogger("push.err")

    product_id = item["productId"]
    site_ids   = item["siteIds"]
    primary    = site_ids[0]
    status     = item["status"]

    if status is None:
        # Defensive — something upstream shouldn't have passed it through.
        err_log.error(
            "skipped productId=%s reason=status_null", product_id,
        )
        return {
            "productId":    product_id,
            "putOk":        False,
            "needsPublish": False,
            "reason":       "status_null",
        }

    t0 = time.monotonic()
    put_ok, put_info = put_product(
        session, api_host, product_id, site_ids, item["payload"],
        dry_run=dry_run, delay_ms=delay_ms,
    )
    elapsed_ms = int((time.monotonic() - t0) * 1000)

    if not put_ok:
        err_log.error(
            "verdict=put_failed productId=%s siteIds=%s status=%s elapsed_ms=%d detail=%s",
            product_id, site_ids, status, elapsed_ms, put_info,
        )
        return {
            "productId":    product_id,
            "siteIds":      site_ids,
            "primarySite":  primary,
            "status":       status,
            "putOk":        False,
            "needsPublish": False,
            "reason":       put_info,
        }

    needs_publish = (status == "PUBLISHED")
    ok_log.info(
        "verdict=put_ok productId=%s siteIds=%s status=%s elapsed_ms=%d needs_publish=%s",
        product_id, site_ids, status, elapsed_ms, needs_publish,
    )
    return {
        "productId":    product_id,
        "siteIds":      site_ids,
        "primarySite":  primary,
        "status":       status,
        "putOk":        True,
        "needsPublish": needs_publish,
        "reason":       put_info,
    }


# ---------------------------------------------------------------------------
# Publish phase — batched COMMIT (READY removed)
# ---------------------------------------------------------------------------

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
    Apply COMMIT directly to every product that needs publishing.

    READY is intentionally skipped (per operator direction). The platform
    accepts a COMMIT without a prior READY for these products; the READY
    gate that previously protected against partial-write states is gone.

    Tradeoff
    --------
    Without READY, a COMMIT failure means the product change is stuck
    mid-write (no rollback). Errors land in errors.log with the failing
    batch and ids; the operator must re-run or fix manually.

    `delay_ms` is forwarded to batch_change_status — see its docstring.

    Returns (ready_ok, ready_failed, commit_ok, commit_failed). The
    ready_* values are always 0 since no READY calls happen — kept for
    caller compatibility / symmetry with push_variants.
    """
    log = logging.getLogger("push")
    err_log = logging.getLogger("push.err")
    ok_log = logging.getLogger("push.ok")

    commit_ok = commit_fail = 0

    commit_tasks = [
        (site, batch)
        for site, ids in to_publish_by_site.items()
        for batch in chunked(ids, batch_size)
    ]
    log.info(
        "publish phase (COMMIT-only, no READY): %d sites, %d total products, "
        "%d COMMIT batches (batch_size=%d)",
        len(to_publish_by_site),
        sum(len(v) for v in to_publish_by_site.values()),
        len(commit_tasks),
        batch_size,
    )

    progress = ProgressReporter(log, "commit batches", total=len(commit_tasks))
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
            progress.tick(extra=f"ok={commit_ok} fail={commit_fail}")

    return 0, 0, commit_ok, commit_fail


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
        help="Output dir from repair_products_negative.py (contains product_map_negative.json + payloads)",
    )
    ap.add_argument(
        "--api-host",
        default="https://nochannel-test-1-api.nochannel-test.upstart.team",
        help="API host (no trailing slash)",
    )
    ap.add_argument(
        "--site-id", default=PRIMARY_SITE_ID,
        help="Default site to use if a payload's folder is missing siteIds",
    )
    ap.add_argument(
        "--fallback-site-ids",
        default=",".join(FALLBACK_SITE_IDS),
        help="Comma-separated fallback sites (informational; folder name "
             "wins for the actual PUT)",
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
        "--workers", type=int, default=8,
        help="Concurrent PUT workers (default 8)",
    )
    ap.add_argument(
        "--batch-size", type=int, default=50,
        help="Max products per COMMIT batch (default 50)",
    )
    ap.add_argument(
        "--delay-ms", type=int, default=100,
        help="Sleep this many milliseconds after each API request (per "
             "worker), to mimic portal-paced traffic. Default 100. Set to "
             "0 to disable.",
    )
    ap.add_argument(
        "--limit", type=int, default=0,
        help="Process only first N products (0=all). For testing.",
    )
    ap.add_argument(
        "--dry-run", action="store_true",
        help="Log what would happen; no HTTP requests made.",
    )
    ap.add_argument(
        "--skip-publish", action="store_true",
        help="Run the PUT phase only — skip COMMIT for PUBLISHED products. "
             "Useful for staged rollouts.",
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
        "run_start run_id=%s input=%s dry_run=%s skip_publish=%s "
        "workers=%d batch_size=%d delay_ms=%d",
        run_id, args.input, args.dry_run, args.skip_publish,
        args.workers, args.batch_size, args.delay_ms,
    )

    # --- load product_map.json ---
    map_path = args.input / "product_map_negative.json"
    if not map_path.exists():
        log.error("product_map_negative.json not found at %s", map_path)
        return 1
    product_map = json.loads(map_path.read_text())
    log.info("loaded product_map_negative.json with %d entries", len(product_map))

    # --- scan files ---
    log.info("[1/3] scanning %s for payload files", args.input)
    items = collect_work_items(args.input, product_map)
    if args.limit > 0:
        items = items[: args.limit]
        log.info("--limit %d applied; processing first %d", args.limit, len(items))
    if not items:
        log.info("no items to process. run_id=%s", run_id)
        return 0

    # status breakdown
    status_counts = Counter(it["status"] for it in items)
    log.info(
        "found %d payload files paired with product_map entries", len(items),
    )
    log.info(
        "status breakdown: PUBLISHED=%d (PUT+COMMIT), DRAFT=%d (PUT only), null=%d (skip)",
        status_counts.get("PUBLISHED", 0),
        status_counts.get("DRAFT", 0),
        status_counts.get(None, 0),
    )

    if not args.cookie and not args.dry_run:
        log.warning(
            "no Cookie provided — if the gateway requires session cookies "
            "you'll get HTML responses. Pass --cookie or API_COOKIE env."
        )

    session = make_session(args.session_id, cookie_header=args.cookie or None)
    t_run = time.time()

    # ---- PUT phase ----
    log.info("[2/3] PUT phase: %d products, %d workers", len(items), args.workers)
    to_publish_by_site: dict[str, list[str]] = defaultdict(list)
    put_ok = put_fail = 0
    put_lock = threading.Lock()
    progress = ProgressReporter(log, "PUT", total=len(items))

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        futs = {
            pool.submit(push_product, session, args.api_host, item, args.dry_run, args.delay_ms): item
            for item in items
        }
        for fut in as_completed(futs):
            result = fut.result()
            with put_lock:
                if result["putOk"]:
                    put_ok += 1
                    if result["needsPublish"]:
                        # Bucket by PRIMARY site (= site_ids[0]). The batch
                        # change-status endpoint takes only one site at a
                        # time via X-UpStart-Site; the variant's primary
                        # is the one we sent its PUT to via the header.
                        # Other siteIds were applied transitively via the
                        # ?sites= query on the PUT.
                        to_publish_by_site[result["primarySite"]].append(
                            result["productId"]
                        )
                else:
                    put_fail += 1
            progress.tick(extra=f"ok={put_ok} fail={put_fail}")

    log.info("PUT phase done: ok=%d fail=%d", put_ok, put_fail)

    # ---- COMMIT phase ----
    if args.skip_publish:
        log.info("[3/3] skip publish phase (--skip-publish)")
        return _final_summary(log, run_id, t_run, success=put_fail == 0)

    if not to_publish_by_site:
        log.info("[3/3] no products to publish (nothing was PUBLISHED + PUT-ok)")
        return _final_summary(log, run_id, t_run, success=put_fail == 0)

    log.info("[3/3] publish phase")
    _, _, commit_ok, commit_fail = run_publish_phase(
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

    success = put_fail == 0 and commit_fail == 0
    return _final_summary(log, run_id, t_run, success=success)


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