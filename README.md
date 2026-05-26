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
└── product-specific/                   # Step 2 — Repair from the product side
    ├── repair_products.py              #   Build product map + repaired payloads
    ├── push_products.py                #   PUT non-archived products
    ├── repair_archived_products.py     #   Recover → PUT → COMMIT → archive flow
    └── publish_auto_generated_variants.py # Republish affected variants
```

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

### 2. Product-Side Repair

Repairs product payloads (fixes `variantIds` on products) and republishes affected variants.

```bash
cd product-specific
export API_COOKIE='paste-from-DevTools'
export API_SESSION_ID='paste-from-DevTools'

python3 repair_products.py           --input ../report-linkages.xlsx --output ./out --workers 2
python3 push_products.py             --input ./out --workers 2
python3 repair_archived_products.py  --input ./out --workers 2
python3 publish_auto_generated_variants.py --input ./out --workers 2
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

---

## Sub-Module Documentation

Each directory contains its own `readme.md` with detailed per-script documentation:

- [`build-linkages-report/readme.md`](build-linkages-report/readme.md) — Report generation pipeline
- [`product-specific/readme.md`](product-specific/readme.md) — Product-side repair pipeline
- [`variant-specific/readme.md`](variant-specific/readme.md) — Variant-side repair pipeline
