# Excel table rows

`append_table_rows`, `insert_table_rows`, and `delete_table_rows` operate through ListRows, not worksheet-row insertion. Indexes are 1-based data-row positions; headers and totals are excluded. A write batch is limited to 1,000 rows.

Row arrays must be rectangular and match the inspected table column count. Alternatively, sparse row objects use inspected column names; omitted keys are intentionally untouched so Excel calculated-column formulas remain in place. JSON scalar types are preserved: number, boolean, null/blank, and string. Insertion uses ListRows.Add with the default AlwaysInsert behavior, so occupied cells below a table may shift; callers must not assume surrounding worksheet addresses are fixed.


Sparse objects must use exact existing column names. Omitted columns are intended for calculated columns: readback requires their formulas to remain present. Supply explicit null for an ordinary blank value. Nested arrays/objects are rejected before mutation.
