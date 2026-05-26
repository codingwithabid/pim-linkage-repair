# Variant Repair & Publish Pipeline

This repository contains a 4-step repair and publishing pipeline for repairing variant/product linkages, handling archived variants safely, and republishing affected products.

---

## Pipeline Overview

```text
INPUT: report-linkages.xlsx
        │
        ▼
1. repair_variants.py
        │
        ├─ variant_map.json
        ├─ duplicate_variants.json
        ├─ product_states_pre_repair.json
        └─ repaired variant payloads
        │
        ▼
2. push_variants.py
        │
        ▼
3. repair_archived_variants.py
        │
        ▼
4. publish_pending_products.py
```

---

## Script 1 — `repair_variants.py`

Repairs variant payloads and prepares all metadata required for downstream publishing.

### Responsibilities

- Read linkage report sheets
- Detect duplicate variants
- Fetch authoritative API data
- Generate repaired payloads
- Build variant metadata map
- Snapshot product publish state before repair

### Workflow

#### 1. Read Sheets

Parses the following sheets:

- `Variant Specific(+ve)`
- `Product Specific(+ve)`
- `Product to verify(-ve)`

Each row extracts variants from values like:

```text
CODE: id (name), id (name)
```

Only variants tagged with:

- `STALE_LATEST_EMPTY`
- `ORPHAN_BUT_REFERENCED`

are retained.

#### 2. Detect Duplicates

Variants are grouped by `variantId`.

A variant referenced by **2 or more products** is considered a duplicate.

Output:

```text
duplicate_variants.json
```

Contains:

- product associations
- sheet-derived statuses
- duplicate ownership information

#### 3. Enrich Duplicates (API)

For every duplicate variant:

- GET variant from API
- Retry across multiple sites if needed
- Fetch authoritative:
  - state
  - status
  - siteIds

The enriched data rewrites `duplicate_variants.json`.

> No repair payloads are generated for duplicates because they require manual review.

#### 4. Snapshot Product States

Captures original product states using sheet data only.

Rules:

- `draft_state` exists → `DRAFT`
- otherwise → `PUBLISHED`

Output:

```text
product_states_pre_repair.json
```

#### 5. Build Clean Variant Map

Builds a map excluding duplicate variants.

Initial fields:

```json
{
  "status": null,
  "state": null,
  "siteIds": null
}
```

Output:

```text
variant_map.json
```

#### 6. Fetch + Repair + Write Payloads

For each clean variant:

**API Fetch**

Attempts:

```text
Primary site
  → fallback sites
```

Views tried:

```text
LATEST_INCLUDE_DRAFT
→ LATEST
```

**Payload Repair**

Repairs payload by setting:

```json
{
  "productIds": ["<sheet-product-id>"]
}
```

Removes server-managed fields such as:

- links
- state
- status
- siteIds (encoded in folder path)
- timestamps
- metadata
- audit fields

(~19 fields stripped)

**Output Structure**

Payloads are written as:

```text
<output>/<productTemplateId>/<siteIds>/<variantId>.json
```

Folder name = comma-joined `siteIds` from the payload, in original order. The first site is the primary (used in `X-UpStart-Site` header during push); the rest go in the `?sites=` query parameter.

Example:

```text
output/
 └── tpl-123/
      └── siteA,siteB/
           └── variant-1.json
```

**Final Map Update**

After repair completes, `variant_map.json` is rewritten using authoritative API values:

- status
- state
- siteIds
- fetchedViaSite

---

## Script 2 — `push_variants.py`

Pushes repaired non-archived variants back into the platform.

### Responsibilities

- Scan repaired payloads
- Filter unsupported states
- PUT repaired variants
- Commit variants when required

### Workflow

#### 1. Scan

Scans:

```text
<input>/<tpl>/<siteIds>/<variantId>.json
```

For each file:

- `variantId` = filename
- `siteIds` = folder name split by `,`

Looks up metadata in `variant_map.json`.

##### Scan-Time Filtering

**Skip Archived**

```text
state = ARCHIVED
```

Skipped because archived variants are handled separately by `repair_archived_variants.py`.

The count of skipped archived variants is surfaced in the run log so the operator knows how many were handed off to script 3.

**Continue**

```text
state = NORMAL
```

moves to PUT phase.

#### 2. PUT Phase

Concurrent PUT execution per variant.

**Decision Matrix**

| Status | Action |
|---|---|
| PUBLISHED | PUT + mark for COMMIT |
| DRAFT | PUT only |
| null | Skip |

**PUT Request Shape**

```http
PUT /pim/variants/<id>?sites=<siteB,siteC>
X-UpStart-Site: <siteA>
```

Where:

- first site becomes primary site header
- remaining sites go in query parameter

For single-site variants the `?sites=` query parameter is omitted entirely.

#### 3. Publish Phase

Runs batched COMMIT requests grouped by primary site. **READY is intentionally skipped** (per operator direction) — the script goes straight from PUT to COMMIT.

**COMMIT**

```http
POST /pim/batch/variants/change-status
{
  "status": "COMMIT"
}
```

**Failure Rule**

If COMMIT fails:

```text
The batch is recorded as failed in errors.log.
Variants in that batch stay in DRAFT state.
```

No rollback gate. Recovery is by re-run or manual fix.

---

## Script 3 — `repair_archived_variants.py`

Handles archived variants separately because archived published variants require a recovery workflow.

### Responsibilities

- Repair archived variants safely
- Recover archived published variants
- Re-archive after publish

### Workflow

#### 1. Scan

Filters variants where:

```text
state == ARCHIVED
AND
status IN (PUBLISHED, DRAFT)
```

Skips:

- non-archived variants
- null statuses
- unknown statuses

### DRAFT Path

Runs first because it is fast.

**Archived Draft Variants**

Platform behavior allows direct PUT.

**Action**

```text
PUT only
```

Result:

```text
ARCHIVED state preserved
```

Done.

### PUBLISHED Path

Published archived variants require a 6-step recovery flow.

#### 2. RECOVER

```http
POST /batch/variants/recover
```

State transition:

```text
ARCHIVED → NORMAL
```

#### 3. PUT

Push repaired payload.

#### 4. COMMIT-1

```http
POST /batch/variants/change-status
{
  "status": "COMMIT"
}
```

Recovered draft becomes published. **READY is skipped** — direct COMMIT.

#### 5. ARCHIVE

```http
POST /batch/variants/change-state/ARCHIVED
```

State transition:

```text
NORMAL → ARCHIVED
```

#### 6. COMMIT-2

```http
POST /batch/variants/change-status
{
  "status": "COMMIT"
}
```

Finalizes the archive operation. Again, no READY beforehand.

### Failure Handling

If a batch fails at phase `N`:

```text
variants are excluded from phase N+1
```

Error logs capture the failed step and the offending ids. Without a READY gate, a COMMIT failure means the variant is in an intermediate state (NORMAL with uncommitted draft change, or ARCHIVED with uncommitted archive change) — manual rescue needed.

---

## Script 4 — `publish_pending_products.py`

Publishes products that received new drafts during variant repair.

### Responsibilities

- Detect pending product drafts
- Verify drafts exist
- Publish affected products

### Workflow

#### 1. Scan

Loads `product_states_pre_repair.json`.

Filters:

```text
state == PUBLISHED
```

Skips:

- DRAFT
- null

**Rationale**

Only products that were originally published require republishing.

#### 2. Check

Per-product API validation with multi-site retry.

**API Call**

```http
GET /pim/products/<id>?view=LATEST_INCLUDE_DRAFT&ignoreIfError=1
```

Retry order:

```text
primary site → fallback sites
```

**Outcomes**

| Result | Action |
|---|---|
| 200 on any site | Draft exists → queue for COMMIT on that site |
| 404 on all sites | Skip |
| Other errors | Skip + log |

#### 3. Publish

Single batched COMMIT per site. **READY is skipped** (per operator direction).

**COMMIT**

```http
POST /pim/batch/products/change-status
{
  "status": "COMMIT"
}
```

**Failure Rule**

If COMMIT fails:

```text
The batch is logged as failed.
Affected products stay in DRAFT.
```

No READY gate. Re-run or manual fix is the recovery story.

---

## End-to-End Flow

```text
report-linkages.xlsx
        │
        ▼
repair_variants.py
        │
        ├── variant_map.json
        ├── duplicate_variants.json
        ├── product_states_pre_repair.json
        └── repaired payloads
        │
        ▼
push_variants.py
        │
        ▼
repair_archived_variants.py
        │
        ▼
publish_pending_products.py
```

---

## Generated Artifacts

| Artifact | Description |
|---|---|
| `variant_map.json` | Variant metadata map with authoritative API state |
| `duplicate_variants.json` | Variants linked to multiple products |
| `product_states_pre_repair.json` | Product states before repair |
| `<tpl>/<siteIds>/<variantId>.json` | Repaired variant payloads (folder = comma-joined siteIds) |

---

## Safety Guarantees

- Duplicate variants are never auto-repaired
- Archived variants are never touched by `push_variants.py` (handed off to script 3)
- Archived published variants are restored to original archived state
- Unknown states/statuses are skipped
- Multi-site retries reduce false negatives
- Product republish only occurs when draft existence is confirmed
- Per operator direction: READY is **not** called anywhere — direct COMMIT is the contract for all publish operations

---

## Run Commands

- Build artifacts from sheet + API GETs 

``` python3 repair_variants.py --input ../report-linkages.xlsx --output ./out --workers 2```
- Push variant payloads (non-archived)

```python3 push_variants.py --input ./out --workers 2```
- Repair archived variants (with full recover/republish/rearchive flow)

```python3 repair_archived_variants.py --input ./out --workers 2```
- Publish products whose drafts were created by the variant changes

```python3 publish_auto_generated_product.py --input ./out --workers 2```


## Command Line Arguments

| Argument | Type | Default | Description |
|----------|------|----------|-------------|
| `--workers` | int | `1` | Concurrent PUT/batch workers. |
| `--batch-size` | int | `1` | Maximum variants per batch call. |
| `--delay-ms` | int | `100` | Sleep time in milliseconds after each API request (per worker) to mimic portal-paced traffic. Set `0` to disable. |
| `--limit` | int | `0` | Process only first `N` candidates. `0` means process all. Useful for testing. |
| `--dry-run` | flag | `false` | Log actions without making any HTTP requests. |
| `--stop-after` | enum | `all` | Stop pipeline after a specific phase. Options: `scan`, `recover`, `put`, `publish`, `archive`, `all`. |
| `--log-dir` | Path | `<input>/logs` | Directory where logs will be stored. |
| `--verbose`, `-v` | flag | `false` | Enable DEBUG-level logging. |
