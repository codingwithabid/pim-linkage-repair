# Product Repair & Publish Pipeline — Negative Workflow

This repository contains the **negative** product-repair pipeline. It runs separately from the positive pipeline and handles a different issue code from a different sheet, with its own input and output files so the two can run independently without cross-contamination.

| | Positive pipeline | Negative pipeline (this README) |
|---|---|---|
| Sheet | `Variant Specific(+ve)`, `Product Specific(+ve)`, `Product to verify(-ve)` | `Product to verify(-ve)` only |
| Issue code in scope | `PARENT_DOES_NOT_LINK_BACK` | `PRODUCT_DROPPED_VARIANT` |
| Map file | `product_map.json` | `product_map_negative.json` |
| Duplicates file | `duplicate_variants_across_product.json` | `duplicate_variants_across_product_negative.json` |
| Variant snapshot | `variant_states_pre_repair.json` | `variant_states_pre_repair_negative.json` |
| Log prefix | `products-repair.*` | `products-repair-negative.*` |

The mechanical fix is the same in both pipelines (append the listed variants to the product's `variantIds`, dedupe, preserve order). What's different is which rows get picked up and where the artifacts land.

What `PRODUCT_DROPPED_VARIANT` means: an earlier version of a product listed some variant in its `variantIds`, but the latest version no longer does — and the variant itself never claimed this product as a parent from its own side. The link existed only on the product's side, then got severed. We treat this as a dropped-by-mistake link and re-attach.

---

## Pipeline Overview

```text
INPUT: report-linkages.xlsx
        │
        ▼
1. repair_products_negative.py
        │
        ├─ product_map_negative.json
        ├─ duplicate_variants_across_product_negative.json
        ├─ variant_states_pre_repair_negative.json
        └─ repaired product payloads
        │
        ▼
2. push_products_negative.py
        │
        ▼
3. repair_archived_products_negative.py
        │
        ▼
4. publish_auto_generated_variants_negative.py
```

---

## Script 1 — `repair_products_negative.py`

Repairs product payloads and prepares all metadata required for downstream publishing.

### Responsibilities

- Read the negative linkage sheet
- Detect cross-product duplicate variants
- Fetch authoritative API data
- Generate repaired payloads
- Build product metadata map
- Snapshot variant publish state before repair

### Workflow

#### 1. Read Sheets

Reads **one sheet only**:

| Sheet | Variant-list column |
|---|---|
| `Product to verify(-ve)` | `links to be fixed` |

The other two sheets (`Variant Specific(+ve)`, `Product Specific(+ve)`) are intentionally ignored — they're handled by the positive pipeline. This narrows the input surface so a misclassified row in another sheet can't accidentally trigger a product-side repair through this script.

Cells can be either of two formats:

```text
# Format A — issue_label in its own column
issue_label   = "PRODUCT_DROPPED_VARIANT"
variants_cell = "BA-XYZ123 (Foo), BA-ABC456 (Bar)"

# Format B — code embedded as prefix
variants_cell = "PRODUCT_DROPPED_VARIANT: BA-XYZ123 (Foo), BA-ABC456 (Bar)"
```

Both work. Variant IDs are comma-separated within the cell; the optional `(Name)` label after each ID is discarded.

Compound rows (e.g. `issue_label = "PRODUCT_DROPPED_VARIANT | ORPHAN_BUT_REFERENCED"` with per-code sections inside the variants cell) are also handled — only the variants under `PRODUCT_DROPPED_VARIANT` are picked up; variants under other codes are skipped because they belong to other scripts.

Only rows tagged with:

- `PRODUCT_DROPPED_VARIANT`

are retained. Other codes (`STALE_LATEST_EMPTY`, `ORPHAN_BUT_REFERENCED`, `PARENT_DOES_NOT_LINK_BACK`) are handled by other scripts and silently dropped.

#### 2. Merge + Detect Duplicates

The same `productId` appearing in multiple rows has its variantIds combined into one deduplicated list.

A variant claimed by **2 or more different products** is considered a cross-product duplicate. These can't be auto-repaired (we don't know which product is the correct owner), so all products containing any duplicate variant are excluded from the clean repair map.

Output:

```text
duplicate_variants_across_product_negative.json
```

Shape:

```json
{
  "BA-INTWPWK": {
    "variantName": "Twin Oak Pewter King Panel Bed",
    "products": [
      {
        "productId":         "6da67707-...",
        "productName":       "Twin Oak 6 Drawer Storage Bed",
        "productTemplateId": "71d5cab7-...",
        "state":             "HIDDEN",
        "status":            "PUBLISHED",
        "issues":            ["PRODUCT_DROPPED_VARIANT"]
      },
      {
        "productId":         "f1259472-...",
        "productName":       "Twin Oak Panel Bed",
        ...
      }
    ],
    "state":   "ARCHIVED",
    "status":  "DRAFT",
    "siteId":  "201cb789-..."
  }
}
```

#### 3. Enrich Duplicates (API)

For every duplicate variant:

- GET variant from API
- Retry across multiple sites if needed
- Fetch authoritative:
  - state
  - status
  - siteIds
  - variantName

For every product claiming a duplicate variant:

- GET product from API
- Fetch authoritative state and status to populate the per-product fields

The enriched data rewrites `duplicate_variants_across_product_negative.json`.

> No repair payloads are generated for duplicates because they require manual review.

#### 4. Snapshot Variant States

For every unique variantId in the clean product map (across all clean products), GET from the variant API and record:

```json
{
  "BA-XYZ123": {
    "state":          "NORMAL",
    "status":         "PUBLISHED",
    "fetchedViaSite": "201cb789-..."
  }
}
```

Views tried: `LATEST_INCLUDE_DRAFT → LATEST`.

This is the authoritative "before" snapshot used by `publish_auto_generated_variants_negative.py` (script 4) to decide which variants need republishing after products get updated.

Output:

```text
variant_states_pre_repair_negative.json
```

#### 5. Build Clean Product Map

Builds a map excluding products that contain any cross-product duplicate variant.

Initial fields:

```json
{
  "<productId>": {
    "variantIds":        ["BA-XYZ", "BA-ABC"],
    "productName":       "Lounge 5 Pc. Theater Sectional",
    "productTemplateId": "71d5cab7-...",
    "status":            null,
    "state":             null,
    "siteIds":           null,
    "fetchedViaSite":    null,
    "issues":            ["PRODUCT_DROPPED_VARIANT"]
  }
}
```

`status`, `state`, `siteIds`, and `fetchedViaSite` are populated by step 6 from API truth.

Output:

```text
product_map_negative.json
```

#### 6. Fetch + Repair + Write Payloads

For each clean product:

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

Repairs payload by appending sheet variantIds to the existing `variantIds` list, deduplicated:

```json
{
  "variantIds": [
    "<existing IDs in their original order>",
    "<new IDs from the sheet, appended at the tail>"
  ]
}
```

Original payload order is preserved; new IDs not already present are appended. This is the smallest possible change that re-attaches the dropped link.

Removes server-managed fields such as:

- links
- state
- status
- siteIds (encoded in folder path)
- timestamps
- metadata
- audit fields

(~19 fields stripped — same list as the variant script)

**Output Structure**

Payloads are written as:

```text
<output>/<productTemplateId>/<siteIds>/<productId>.json
```

Folder name = comma-joined `siteIds` from the payload, in original order. The first site is the primary (used in `X-UpStart-Site` header during push); the rest go in the `?sites=` query parameter.

Example:

```text
output/
 └── tpl-123/
      └── siteA,siteB/
           └── prod-1.json
```

**Final Map Update**

After repair completes, `product_map_negative.json` is rewritten using authoritative API values:

- status
- state
- siteIds
- fetchedViaSite

---

## Script 2 — `push_products_negative.py`

Pushes repaired non-archived products back into the platform.

### Responsibilities

- Scan repaired payloads
- Filter unsupported states
- PUT repaired products
- Commit products when required

### Workflow

#### 1. Scan

Scans:

```text
<input>/<tpl>/<siteIds>/<productId>.json
```

For each file:

- `productId` = filename
- `siteIds` = folder name split by `,`

Looks up metadata in `product_map_negative.json`.

##### Scan-Time Filtering

**Skip Archived**

```text
state = ARCHIVED
```

Skipped because archived products are handled separately by `repair_archived_products_negative.py`.

The count of skipped archived products is surfaced in the run log so the operator knows how many were handed off to script 3.

**Continue**

```text
state = NORMAL
```

moves to PUT phase.

#### 2. PUT Phase

Concurrent PUT execution per product.

**Decision Matrix**

| Status | Action |
|---|---|
| PUBLISHED | PUT + mark for COMMIT |
| DRAFT | PUT only |
| null | Skip |

**PUT Request Shape**

```http
PUT /pim/products/<id>?sites=<siteB,siteC>
X-UpStart-Site: <siteA>
```

Where:

- first site becomes primary site header
- remaining sites go in query parameter

For single-site products the `?sites=` query parameter is omitted entirely.

#### 3. Publish Phase

Runs batched COMMIT requests grouped by primary site. **READY is intentionally skipped** (per operator direction) — the script goes straight from PUT to COMMIT.

**COMMIT**

```http
POST /pim/batch/products/change-status
{
  "status": "COMMIT"
}
```

**Failure Rule**

```text
The batch is recorded as failed in errors.log.
Products in that batch stay in DRAFT state.
```

No rollback gate. Recovery is by re-run or manual fix.

---

## Script 3 — `repair_archived_products_negative.py`

Handles archived products separately because archived published products require a recovery workflow.

### Responsibilities

- Repair archived products safely
- Recover archived published products
- Re-archive after publish

### Workflow

#### 1. Scan

Filters products in `product_map_negative.json` where:

```text
state == ARCHIVED
AND
status IN (PUBLISHED, DRAFT)
```

Skips:

- non-archived products
- null statuses
- unknown statuses

### DRAFT Path

Runs first because it is fast.

**Archived Draft Products**

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

Published archived products require a 6-step recovery flow.

#### 2. RECOVER

```http
POST /pim/batch/products/recover
```

State transition:

```text
ARCHIVED → NORMAL
```

#### 3. PUT

Push repaired payload.

#### 4. COMMIT-1

```http
POST /pim/batch/products/change-status
{
  "status": "COMMIT"
}
```

Recovered draft becomes published. **READY is skipped** — direct COMMIT.

#### 5. ARCHIVE

```http
POST /pim/batch/products/change-state/ARCHIVED
```

State transition:

```text
NORMAL → ARCHIVED
```

#### 6. COMMIT-2

```http
POST /pim/batch/products/change-status
{
  "status": "COMMIT"
}
```

Finalizes the archive operation. Again, no READY beforehand.

### Failure Handling

If a batch fails at phase `N`:

```text
products are excluded from phase N+1
```

Error logs capture the failed step and the offending ids. Without a READY gate, a COMMIT failure means the product is in an intermediate state (NORMAL with uncommitted draft change, or ARCHIVED with uncommitted archive change) — manual rescue needed.

---

## Script 4 — `publish_auto_generated_variants_negative.py`

Publishes variants that received auto-generated drafts during product repair.

"Auto-generated" here means drafts the PIM platform created as a side effect of the product PUTs — not drafts the operator created by hand. When a product's `variantIds` list is updated, the platform automatically spins up a DRAFT revision on each linked variant; this script promotes those drafts to live for variants that were already PUBLISHED before the run started.

### Responsibilities

- Detect pending variant drafts
- Verify drafts exist
- Publish affected variants

### Workflow

#### 1. Scan

Loads `variant_states_pre_repair_negative.json`.

Filters:

```text
status == PUBLISHED
```

Skips:

- DRAFT
- null

**Rationale**

Only variants that were originally published require republishing. A variant that was DRAFT before the run is meant to stay a draft — the product scripts may have piled additional draft content on top, but it's still a draft, no publish needed.

#### 2. Check

Per-variant API validation with multi-site retry.

**API Call**

```http
GET /pim/variants/<id>?view=LATEST_INCLUDE_DRAFT&ignoreIfError=1
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
POST /pim/batch/variants/change-status
{
  "status": "COMMIT"
}
```

**Failure Rule**

```text
The batch is logged as failed.
Affected variants stay in DRAFT.
```

No READY gate. Re-run or manual fix is the recovery story.

---

## End-to-End Flow

```text
report-linkages.xlsx
        │
        ▼
repair_products_negative.py
        │
        ├── product_map_negative.json
        ├── duplicate_variants_across_product_negative.json
        ├── variant_states_pre_repair_negative.json
        └── repaired payloads
        │
        ▼
push_products_negative.py
        │
        ▼
repair_archived_products_negative.py
        │
        ▼
publish_auto_generated_variants_negative.py
```

---

## Generated Artifacts

| Artifact | Description |
|---|---|
| `product_map_negative.json` | Product metadata map with authoritative API state |
| `duplicate_variants_across_product_negative.json` | Variants claimed by 2+ different products |
| `variant_states_pre_repair_negative.json` | Variant states (state, status) before repair |
| `<tpl>/<siteIds>/<productId>.json` | Repaired product payloads (folder = comma-joined siteIds) |

---

## Recommended Run Order

```bash
export API_COOKIE='paste-from-DevTools'
export API_SESSION_ID='paste-from-DevTools'

# Build artifacts
python3 repair_products_negative.py \
    --input ../report-linkages.xlsx \
    --output ./out-products-negative \
    --workers 4

# Preview each downstream script (no HTTP)
python3 push_products_negative.py                       --input ./out-products-negative --dry-run
python3 repair_archived_products_negative.py            --input ./out-products-negative --dry-run
python3 publish_auto_generated_variants_negative.py     --input ./out-products-negative --dry-run

# Real runs
python3 push_products_negative.py                       --input ./out-products-negative --workers 4
python3 repair_archived_products_negative.py            --input ./out-products-negative --workers 4
python3 publish_auto_generated_variants_negative.py     --input ./out-products-negative --workers 4
```

For first real runs on production data, consider `--stop-after recover` on the archived script and `--stop-after check` on the publish script to inspect intermediate state before continuing.

---

## Safety Guarantees

- Only the `Product to verify(-ve)` sheet is read — rows in the other two sheets cannot accidentally trigger negative-side repairs
- Only `PRODUCT_DROPPED_VARIANT` rows are picked up — other issue codes are silently dropped
- Cross-product duplicate variants are never auto-repaired — products containing them are excluded from the clean map and surfaced for human review
- Archived products are never touched by `push_products_negative.py` (handed off to script 3)
- Archived published products are restored to original archived state
- Unknown states/statuses are skipped
- Multi-site retries reduce false negatives
- Variant republish only occurs when draft existence is confirmed by API
- Per operator direction: READY is **not** called anywhere — direct COMMIT is the contract for all publish operations

---

## Run Commands

- Build artifacts from sheet + API GETs

```python3 repair_products_negative.py --input ../report-linkages.xlsx --output ./out --workers 2```

- Push product payloads (non-archived)

```python3 push_products_negative.py --input ./out --workers 2```

- Repair archived products (with full recover/republish/rearchive flow)

```python3 repair_archived_products_negative.py --input ./out --workers 2```

- Publish variants whose drafts were created by the product changes

```python3 publish_auto_generated_variants_negative.py --input ./out --workers 2```

---

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