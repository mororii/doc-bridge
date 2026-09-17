# Excel rich text

`set_rich_text` replaces the complete contents of one non-formula cell. It requires `cell`, and a complete `baseFont` with `name`, `size`, `bold`, `italic`, `underline`, and `color`, plus ordered non-empty `{text,font?}` runs. A run font is sparse and inherits only from baseFont, never a preceding run.

Input is limited to 256 runs and 32,767 characters. Surrogate characters are rejected because native character indexing is not verified for them. Strings, including values beginning `=`, are written as literal text; the original NumberFormat is restored. Existing fill, borders, and layout are not changed.

Inspect uses DataObjects scope `richText`, requiring explicit `sheet` and one-cell `range`. It scans at most 4,096 characters and reports `complete`, `truncated`, and `unreadable`. Apply readback checks only requested spans and effective fonts.

Incremental rich-text restore covers readable non-empty literal strings. Original formulas, numeric values, empty-cell font state, or incomplete font inventory are explicitly marked as unsupported by this narrow restore path; the existing workbook recovery artifact remains the fallback.
