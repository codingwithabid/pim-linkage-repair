# Linkages Report Pipeline

This repository contains scripts for generating linkage mismatch reports and enriching affected SKUs with product details using exported Cassandra data.

The pipeline compares draft and live product item records to identify incorrect variant/product linkages and generates an Excel report for further validation and patching.

---

# Pipeline Overview

```text
Cassandra Export
       │
       ▼
product_item_live.csv
product_item.csv
       │
       ▼
linkages-bothside.py
       │
       ▼
report-linkages.xlsx
       │
       ▼
adding_name.py
       │
       ▼
report-patched.xlsx
```

---

# Prerequisites

- Python 3.x
- Access to Cassandra cluster
- Kubernetes access (`kubectl`)
- Access to tools pod

---

# Export Cassandra Data

Run the following commands inside `cqlsh` to export the required tables.

## 1. Export Live Product Items

```sql
COPY pim_writeside.product_item_live(
    tenant_id,
    site_ids,
    item_id,
    version,
    request_schema,
    template_id,
    batch_id,
    date_modified,
    item_scope,
    parent_id,
    sku,
    state,
    variant_id
)
TO '/tmp/product_item_live.csv'
WITH HEADER = TRUE;
```

---

## 2. Export Draft Product Items

```sql
COPY pim_writeside.product_item(
    tenant_id,
    site_ids,
    item_id,
    request_schema,
    template_id,
    batch_id,
    date_modified,
    item_scope,
    product_refs,
    sku,
    state,
    variant_id
)
TO '/tmp/product_item.csv'
WITH HEADER = TRUE;
```

---

# Copy Exported Files from Kubernetes Pod

After export completes, copy the generated CSV files from the tools pod to your local machine.

## Copy Live Export

```bash
kubectl exec -n default tools-6c7766f6db-7tl5c -- cat /tmp/product_item_live.csv > product_item_live.csv
```

---

## Copy Draft Export

```bash
kubectl exec -n default tools-6c7766f6db-7tl5c -- cat /tmp/product_item.csv > product_item.csv
```

---

# Input Files

| File | Description |
|---|---|
| `product_item_live.csv` | Export of live product items |
| `product_item.csv` | Export of draft product items |

---

# Generate Linkages Report

This step compares draft and live product records and generates a linkage mismatch report.

## Command

```bash
python3 linkages-bothside.py product_item_live.csv \
  --draft product_item.csv \
  -o ../report-linkages.xlsx
```

---

## Output

```text
../report-linkages.xlsx
```

---

# Add Product Names to Affected SKUs

This step enriches the generated report by adding product names/details for affected SKUs.

Before running this script, make sure the report format matches the file shared by the client.

## Command

```bash
python3 adding_name.py ../report-linkages.xlsx \
  product_item_live.csv \
  --draft product_item.csv \
  -o ../report-patched.xlsx
```

---

## Output

```text
../report-patched.xlsx
```

---

# Example End-to-End Workflow

```bash
# Step 1: Generate linkage mismatch report
python3 linkages-bothside.py product_item_live.csv \
  --draft product_item.csv \
  -o ../report-linkages.xlsx


# Step 2: Add product names/details
python3 adding_name.py ../report-linkages.xlsx \
  product_item_live.csv \
  --draft product_item.csv \
  -o ../report-patched.xlsx
```

---

# Generated Files

| File | Description |
|---|---|
| `report-linkages.xlsx` | Initial linkage mismatch report |
| `report-patched.xlsx` | Enriched report with product details |

---

# Notes

- Ensure CSV exports are fully generated before copying from the pod.
- Large exports may take several minutes depending on table size.
- The scripts expect CSV headers to be present.
- Always validate the final patched report before sharing with the client.

---