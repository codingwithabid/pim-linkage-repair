# Linkages Report Pipeline

This repository contains scripts for generating linkage reports and enriching affected SKUs with product names using exported Cassandra data.

---

## Input Files

| File | Description |
|---|---|
| `product_item.csv` | Draft product items export |
| `product_item_live.csv` | Live product items export |

---

## Run Commands

### 1. Build Linkages Report

Generates a linkage mismatch report by comparing draft and live product items.

```bash
python3 linkages-bothside.py product_item_live.csv \
  --draft product_item.csv \
  -o ../report-linkages.xlsx
```

### Output

```text
../report-linkages.xlsx
```

---

### 2. Add Names to Affected SKUs

Before running this script make sure the file is in the format as shared by client. Adds product names/details to affected SKUs in the generated linkage report.

```bash
python3 adding_name.py ../report-linkages.xlsx \
  product_item_live.csv \
  --draft product_item.csv \
  -o ../report-patched.xlsx
```

### Output

```text
../report-patched.xlsx
```

---

## Example Workflow

```bash
# Step 1: Generate linkage report
python3 linkages-bothside.py product_item_live.csv \
  --draft product_item.csv \
  -o ../report-linkages.xlsx

# Step 2: Enrich report with product names
python3 adding_name.py ../report-linkages.xlsx \
  product_item_live.csv \
  --draft product_item.csv \
  -o ../report-patched.xlsx
```

---