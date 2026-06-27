# Connecting Excel to the Data Center (Power Query)

The build script publishes compact CSVs to a shared folder (default
`C:\Users\yzsun\Desktop\DD\_datastore\`). Excel reads them with **native Power
Query** — no driver, no add-in, no install.

## One-time wiring (per workbook)

1. **Data → Get Data → From File → From Folder** → pick `…\DD\_datastore`.
   (Or *From Text/CSV* for a single file.)
2. In the navigator choose **Transform Data**, keep the columns, then
   **Close & Load To… → Table** (or **Only Create Connection** + add to the
   Data Model if a table ever exceeds ~1,000,000 rows).
3. Each CSV becomes a refreshable Excel **Table**, e.g. `tam_by_indication_year`.

## Referencing the data downstream (by key, never by row)

```excel
=SUMIFS(tam_by_indication_year[tam_usd_m],
        tam_by_indication_year[indication_code], $B5,
        tam_by_indication_year[year], C$4)
```
or `XLOOKUP` on a concatenated `indication_code|year` key. Because lookups match
on **keys**, inserting a drug never shifts a reference.

## Refreshing after new DD data

1. (WSL) edit a JSON file and run `python datastore/build_datastore.py`.
2. (Excel) **Data → Refresh All** — or set
   **Query Properties → Usage → "Refresh data when opening the file"**
   (and optionally *Refresh every N minutes*).

> ⚠️ A refresh only happens while the workbook is **open**; rewriting the CSV
> from WSL does **not** push into an already-open workbook. One click of
> *Refresh All* (or reopen) picks up the new data. This is the price of zero
> driver installs.
