# Excel conditional rules

Inspect `conditionalFormats` first. Inventory emits a 1-based collection `index` and a fingerprint. Indices are selection positions, not stable IDs.

`update_conditional_format` and `delete_conditional_format` require `range`, `index`, and `expectedFingerprint`. The adapter re-reads the selected rule before mutation and rejects a stale or unreadable target. Update supports only existing expression and cellValue rules, preserving AppliesTo and other rules. Requested formulas, operator, and supplied style are read back; unreadable values do not verify.
