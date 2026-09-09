---
name: excel-range-cleanup
description: Clean a messy Excel range — headers, types, blanks, and convert to a table. Use when the user wants to tidy a sheet, fix headers, or table-ize a block of data.
license: Apache-2.0
compatibility: Requires OpenPlugin Excel host
metadata:
  openplugin/hosts: "excel"
  openplugin/tools: "excel.getSummary excel.readRange excel.writeRange excel.createTable excel.setFormulas"
  openplugin/version: "1.0"
---

# Excel range cleanup

1. Call `excel.getSummary`, then `excel.readRange` on the selection.
2. Propose a cleaned grid: one header row, no blank header cells, consistent types.
3. Queue `excel.writeRange` with the cleaned values.
4. Queue `excel.createTable` over the written range when there is a header row.
5. Do not invent rows that were not in the selection.
6. Leave formulas as formulas when they already exist; do not convert them to values.
