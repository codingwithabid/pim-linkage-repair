## Tools

Utility scripts that aren't part of any specific pipeline. Run on demand
against any pipeline's output folder.

### `export_duplicates_to_excel.py`

Converts the `duplicate_variants.json` (variant side) and/or
`duplicate_variants_across_product.json` (product side) files produced
by the repair scripts into a single-sheet Excel workbook for human review.

One row per (variant, product) pairing, sorted by variant_id. Useful when
the toolchain flags cross-entity duplicates and someone has to decide
which product is the correct owner.

```bash
# Convert duplicates from one pipeline
python3 tools/export_duplicates_to_excel.py --input ./variant-specific/out

# Combine duplicates from multiple pipelines into one workbook
python3 tools/export_duplicates_to_excel.py \
    --variant-duplicates ./variant-specific/out/duplicate_variants.json \
    --product-duplicates ./product-specific/out/duplicate_variants_across_product.json \
    --output ./review/duplicates_combined.xlsx
```