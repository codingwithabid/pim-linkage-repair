# Production Activity — Linkage Repair Toolkit

A collection of Python scripts for detecting and repairing product/variant linkage mismatches in the PIM platform. The toolkit reads a linkage report (Excel), fetches authoritative data from the API, generates repaired payloads, and pushes them back — handling archived entities, cross-product duplicates, and dependent republishing automatically.

---

## Repository Structure

```text
production-activity/
├── report-linkages_1.xlsx              # Sample linkage report
│
├── build-linkages-report/              # Step 0 — Generate the linkage report
│   ├── linkages-bothside.py            #   Compare draft vs live exports
│   └── adding_name.py                  #   Enrich report with product names
│
├── variant-specific/                   # Step 1 — Repair from the variant side
│   ├── repair_variants.py              #   Build variant map + repaired payloads
│   ├── push_variants.py                #   PUT non-archived variants
│   ├── repair_archived_variants.py     #   Recover → PUT → COMMIT → archive flow
│   └── publish_auto_generated_product.py  # Republish affected products
│
├── product-specific/                   # Step 2 — Product-side repair (positive)
│   ├── repair_products.py              #   Handles PARENT_DOES_NOT_LINK_BACK
│   ├── push_products.py                #   PUT non-archived products
│   ├── repair_archived_products.py     #   Recover → PUT → COMMIT → archive flow
│   └── publish_auto_generated_variants.py # Republish affected variants
│
└── product-specific-negative/          # Step 3 — Product-side repair (negative)
    ├── repair_products_negative.py     #   Handles PRODUCT_DROPPED_VARIANT
    ├── push_products_negative.py       #   PUT non-archived products
    ├── repair_archived_products_negative.py            # Archived flow
    └── publish_auto_generated_variants_negative.py     # Republish affected variants
```

The **positive** and **negative** product pipelines read different sheets in the same report and produce separate artifact files (`product_map.json` vs `product_map_negative.json` etc.), so they can run independently without overwriting each other's outputs.

| | Positive | Negative |
|---|---|---|
| Sheet | All three sheets | `Product to verify(-ve)` only |
| Issue code | `PARENT_DOES_NOT_LINK_BACK` | `PRODUCT_DROPPED_VARIANT` |
| Fix | Add missing variants to product's `variantIds` | Re-attach dropped variants to product's `variantIds` |
| Folder | `product-specific/` | `product-specific-negative/` |

---

## End-to-End Workflow

### 0. Build Linkage Report

Generate the mismatch report from Cassandra exports, then enrich it with product names.

```bash
cd build-linkages-report

python3 linkages-bothside.py product_item_live.csv \
  --draft product_item.csv \
  -o ../report-linkages.xlsx

python3 adding_name.py ../report-linkages.xlsx \
  product_item_live.csv \
  --draft product_item.csv \
  -o ../report-patched.xlsx
```

### 1. Variant-Side Repair

Repairs variant payloads (fixes `productIds` on variants) and republishes affected products.

```bash
cd variant-specific
export API_COOKIE='paste-from-DevTools'
export API_SESSION_ID='paste-from-DevTools'

python3 repair_variants.py           --input ../report-linkages.xlsx --output ./out --workers 2
python3 push_variants.py             --input ./out --workers 2
python3 repair_archived_variants.py  --input ./out --workers 2
python3 publish_auto_generated_product.py --input ./out --workers 2
```

### 2. Product-Side Repair (Positive)

Repairs product payloads for `PARENT_DOES_NOT_LINK_BACK` rows across all three sheets, then republishes affected variants.

```bash
cd product-specific
export API_COOKIE='paste-from-DevTools'
export API_SESSION_ID='paste-from-DevTools'

python3 repair_products.py           --input ../report-linkages.xlsx --output ./out --workers 2
python3 push_products.py             --input ./out --workers 2
python3 repair_archived_products.py  --input ./out --workers 2
python3 publish_auto_generated_variants.py --input ./out --workers 2
```

### 3. Product-Side Repair (Negative)

Repairs product payloads for `PRODUCT_DROPPED_VARIANT` rows from the `Product to verify(-ve)` sheet, then republishes affected variants. Runs independently of step 2 — the negative scripts read their own `*_negative.json` artifacts so the two product-side pipelines can be run in either order or in parallel.

```bash
cd product-specific-negative
export API_COOKIE='paste-from-DevTools'
export API_SESSION_ID='paste-from-DevTools'

python3 repair_products_negative.py            --input ../report-linkages.xlsx --output ./out --workers 2
python3 push_products_negative.py              --input ./out --workers 2
python3 repair_archived_products_negative.py   --input ./out --workers 2
python3 publish_auto_generated_variants_negative.py --input ./out --workers 2
```

---

## Common CLI Arguments

All pipeline scripts share the same argument interface:

| Argument | Type | Default | Description |
|----------|------|---------|-------------|
| `--workers` | int | `1` | Concurrent PUT/batch workers |
| `--batch-size` | int | `1` | Max entities per batch call |
| `--delay-ms` | int | `100` | Sleep (ms) between API requests per worker |
| `--limit` | int | `0` | Process only first N candidates (0 = all) |
| `--dry-run` | flag | `false` | Log actions without making HTTP requests |
| `--stop-after` | enum | `all` | Stop after a phase: `scan`, `recover`, `put`, `publish`, `archive`, `all` |
| `--log-dir` | path | `<input>/logs` | Log output directory |
| `--verbose`, `-v` | flag | `false` | Enable DEBUG-level logging |

---

## Safety Guarantees

- Cross-entity duplicates (variant claimed by 2+ products) are excluded from auto-repair and surfaced for manual review
- Archived entities are handled by dedicated scripts with a full recover/repair/re-archive cycle
- Unknown or null states/statuses are skipped
- Multi-site retry reduces false negatives on API fetches
- Dependent republishing only occurs when a pending draft is confirmed via API
- READY is never called — direct COMMIT is the contract for all publish operations (per operator direction)
- `--dry-run` is available on every script for safe pre-flight checks
- The positive and negative product pipelines are isolated: each reads its own map and snapshot files, so neither can overwrite the other's artifacts or accidentally pick up rows belonging to the other workflow

---

## Sub-Module Documentation

Each directory contains its own `readme.md` with detailed per-script documentation:

- [`build-linkages-report/readme.md`](build-linkages-report/readme.md) — Report generation pipeline
- [`variant-specific/readme.md`](variant-specific/readme.md) — Variant-side repair pipeline
- [`product-specific/readme.md`](product-specific/readme.md) — Product-side repair pipeline (positive)
- [`product-specific-negative/readme.md`](product-specific-negative/readme.md) — Product-side repair pipeline (negative)