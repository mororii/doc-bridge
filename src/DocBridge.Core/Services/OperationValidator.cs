using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// operation batch 구조/필드 검증기.
/// - 배치 스키마(operation-batch.schema.json) 핵심 규칙
/// - op별 필수 필드 규칙
/// 구조 오류는 dry-run이어도 즉시 실패시킨다.
/// </summary>
public sealed class OperationValidator
{
    private readonly PolicyEngine _policy;
    public OperationValidator(PolicyEngine policy) => _policy = policy;

    /// <summary>op 이름 → 필수 필드 검증 규칙 (필드명, 타입)</summary>
    private static readonly Dictionary<string, (string Field, string Type)[]> RequiredFields =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["set_values"] = new[] { ("range", "string"), ("values", "array") },
            ["set_formulas"] = new[] { ("range", "string"), ("formulas", "array") },
            ["insert_rows"] = new[] { ("row", "int"), ("count", "int") },
            ["insert_cols"] = new[] { ("col", "any"), ("count", "int") },
            ["format_range"] = new[] { ("range", "string"), ("style", "object") },
            ["find_replace"] = new[] { ("find", "string"), ("replace", "string") },
            ["copy_sheet"] = new[] { ("sourceWorkbook", "string"), ("sourceSheet", "string") },
            ["merge_cells"] = new[] { ("range", "string") },
            ["unmerge_cells"] = new[] { ("range", "string") },
            ["set_rows_hidden"] = new[] { ("row", "int"), ("count", "int"), ("hidden", "bool") },
            ["set_cols_hidden"] = new[] { ("col", "any"), ("count", "int"), ("hidden", "bool") },
            ["set_sheet_visibility"] = new[] { ("visibility", "string") },
            ["insert_text"] = new[] { ("text", "string") },
            ["append_text"] = new[] { ("text", "string") },
            ["insert_before_text"] = new[] { ("anchor", "string"), ("text", "string") },
            ["insert_after_text"] = new[] { ("anchor", "string"), ("text", "string") },
            ["replace_document_text"] = new[] { ("text", "string") },
            ["replace_selection"] = new[] { ("text", "string") },
            ["set_paragraph_style_basic"] = new[] { ("style", "object") },
            ["set_paragraph_format"] = new[] { ("style", "object") },
            ["format_paragraphs"] = new[] { ("items", "array") },
            ["set_page_setup"] = new[] { ("page", "object") },
            ["insert_break"] = new[] { ("type", "string") },
            ["insert_table"] = new[] { ("rows", "array") },
            // 셀 위치는 직사각 표의 row+col 또는 병합 표에 안전한 cellIndex 중 하나를 어댑터가 검증한다.
            ["table_cell_set_text"] = new[] { ("text", "string") },
            ["table_set_cells"] = new[] { ("cells", "array") },
            ["table_insert_rows"] = new[] { ("row", "int"), ("count", "int") },
            ["table_insert_columns"] = new[] { ("col", "int"), ("count", "int") },
            ["table_delete_rows"] = new[] { ("row", "int") },
            ["table_delete_columns"] = new[] { ("col", "int") },
            ["table_merge_cells"] = new[] { ("startRow", "int"), ("startCol", "int"), ("endRow", "int"), ("endCol", "int") },
            ["table_set_row_height"] = new[] { ("row", "int"), ("heightMm", "number") },
            ["table_set_row_heights"] = new[] { ("rows", "array") },
            ["set_field_text"] = new[] { ("name", "string"), ("text", "string") },
            ["insert_picture"] = new[] { ("path", "string") },
            ["insert_page_number"] = Array.Empty<(string Field, string Type)>(),
            ["set_header_footer_text"] = new[] { ("kind", "string"), ("text", "string") },
            ["export_pdf"] = new[] { ("output", "string") },
            ["set_layer_visibility"] = new[] { ("layer", "string"), ("visible", "bool") },
            ["set_layer_color"] = new[] { ("layer", "string"), ("color", "any") },
            ["activate_document"] = new[] { ("document", "string") },
            ["move_entities"] = new[] { ("handles", "array"), ("dx", "number"), ("dy", "number") },
            ["rotate_entities"] = new[] { ("handles", "array"), ("angleDeg", "number") },
            ["set_text_value"] = new[] { ("handle", "string"), ("text", "string") },
            ["delete_entities"] = new[] { ("handles", "array") },
            ["delete_entities_in_bounds"] = new[] { ("bounds", "object") },
            ["delete_entities_from_index"] = new[] { ("startIndex", "int") },
            ["run_script_template"] = new[] { ("template", "string") },
            ["copy_entities_between_documents"] = Array.Empty<(string Field, string Type)>(),
            ["insert_xref"] = new[] { ("sourceFile", "string"), ("insertionPoint", "object") },
            ["zoom_window"] = new[] { ("bounds", "object") },
            ["draw_entities"] = new[] { ("entities", "array") },
            ["copy_entities"] = new[] { ("handles", "array"), ("dx", "number"), ("dy", "number") },
            ["scale_entities"] = new[] { ("handles", "array"), ("basePoint", "array"), ("factor", "number") },
            ["mirror_entities"] = new[] { ("handles", "array"), ("axisStart", "array"), ("axisEnd", "array") },
            ["offset_entities"] = new[] { ("handles", "array"), ("distance", "number") },
            ["set_entity_properties"] = new[] { ("handles", "array"), ("properties", "object") },
            ["set_block_attributes"] = new[] { ("handle", "string"), ("attributes", "object") },
            ["configure_layout"] = new[] { ("name", "string") },
            ["create_viewport"] = new[] { ("layout", "string"), ("center", "array"), ("width", "number"), ("height", "number"), ("viewHeight", "number") },
            ["save_document"] = Array.Empty<(string Field, string Type)>(),
            ["plot_pdf"] = new[] { ("output", "string") },
            ["draw_taegeukgi"] = Array.Empty<(string Field, string Type)>(),
            ["draw_union_jack"] = Array.Empty<(string Field, string Type)>(),
            ["draw_block_wall_schematic"] = Array.Empty<(string Field, string Type)>(),
        };

    public sealed record ParsedBatch(
        List<JsonObject> Ops,
        bool DryRun,
        string? ConfirmToken,
        bool HighRiskConfirm,
        bool HasHighRiskOps);

    public ParsedBatch? Validate(JsonObject? batch, string app, List<string> errors)
    {
        if (batch is null)
        {
            errors.Add("batch body is required (object with 'ops' array)");
            return null;
        }

        var opsArr = Json.GetArr(batch, "ops");
        if (opsArr is null || opsArr.Count == 0)
        {
            errors.Add("'ops' must be a non-empty array");
            return null;
        }

        var ops = new List<JsonObject>();
        var hasHighRisk = false;
        var i = 0;
        foreach (var node in opsArr)
        {
            i++;
            if (node is not JsonObject op)
            {
                errors.Add($"ops[{i}] is not an object");
                continue;
            }
            var name = Json.GetString(op, "op");
            if (string.IsNullOrWhiteSpace(name))
            {
                errors.Add($"ops[{i}].op is required");
                continue;
            }

            switch (_policy.ClassifyOp(app, name))
            {
                case OpClass.Forbidden:
                    errors.Add($"ops[{i}].op '{name}' is FORBIDDEN by policy (app={app})");
                    break;
                case OpClass.Unknown:
                    errors.Add($"ops[{i}].op '{name}' is not in allowlist (app={app})");
                    break;
                case OpClass.HighRisk:
                    hasHighRisk = true;
                    break;
            }

            if (RequiredFields.TryGetValue(name, out var rules))
                foreach (var (field, type) in rules)
                    ValidateField(op, i, name, field, type, errors);

            if (app.Equals("excel", StringComparison.OrdinalIgnoreCase))
                ValidateExcelTarget(op, i, name, errors);

            ops.Add(op);
        }

        if (app.Equals("excel", StringComparison.OrdinalIgnoreCase) &&
            ops.Any(op => string.Equals(Json.GetString(op, "op"), "copy_sheet", StringComparison.OrdinalIgnoreCase)) &&
            ops.Any(op => !string.Equals(Json.GetString(op, "op"), "copy_sheet", StringComparison.OrdinalIgnoreCase)))
        {
            errors.Add(
                "Excel copy_sheet cannot be mixed with other operations in the same batch; " +
                "apply the sheet copy first, then run a new dry-run batch for follow-up edits");
        }

        if (app.Equals("excel", StringComparison.OrdinalIgnoreCase))
        {
            var mergeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "merge_cells", "unmerge_cells",
            };
            if (ops.Any(op => mergeNames.Contains(Json.GetString(op, "op") ?? "")) && ops.Count != 1)
            {
                errors.Add(
                    "Excel merge_cells/unmerge_cells must be the only operation in a batch; " +
                    "apply content edits and formatting in separate dry-run batches");
            }

            var visibilityNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "set_rows_hidden", "set_cols_hidden", "set_sheet_visibility",
            };
            if (ops.Any(op => visibilityNames.Contains(Json.GetString(op, "op") ?? "")) &&
                ops.Any(op => !visibilityNames.Contains(Json.GetString(op, "op") ?? "")))
            {
                errors.Add(
                    "Excel visibility operations cannot be mixed with content, format, copy, or merge operations; " +
                    "use a separate dry-run batch so visibility rollback remains exact");
            }
        }

        var dryRun = !batch.TryGetPropertyValue("dryRun", out var dv) || Json.GetBool(batch, "dryRun", true);
        var confirmToken = Json.GetString(batch, "confirmToken");
        var highRiskConfirm = Json.GetBool(batch, "highRiskConfirm", false);

        if (errors.Count > 0) return null;
        return new ParsedBatch(ops, dryRun, confirmToken, highRiskConfirm, hasHighRisk);
    }

    private static void ValidateField(JsonObject op, int index, string opName, string field, string type, List<string> errors)
    {
        if (!op.TryGetPropertyValue(field, out var v) || v is null)
        {
            errors.Add($"ops[{index}] '{opName}' requires field '{field}'");
            return;
        }
        var bad = type switch
        {
            "string" => v is not JsonValue jvs || !jvs.TryGetValue<string>(out _),
            "int" => v is not JsonValue jvi || !jvi.TryGetValue<int>(out _),
            "number" => v is not JsonValue jvn || !(jvn.TryGetValue<double>(out _) || jvn.TryGetValue<int>(out _)),
            "bool" => v is not JsonValue jvb || !jvb.TryGetValue<bool>(out _),
            "array" => v is not JsonArray,
            "object" => v is not JsonObject,
            _ => false, // "any"
        };
        if (bad) errors.Add($"ops[{index}] '{opName}' field '{field}' must be {type}");
    }

    private static void ValidateExcelTarget(JsonObject op, int index, string opName, List<string> errors)
    {
        var target = Json.GetObj(op, "target");
        var targetSheet = Json.GetString(target, "sheet");

        if (opName is "set_values" or "set_formulas" or "format_range" or "merge_cells" or "unmerge_cells")
        {
            var range = Json.GetString(op, "range");
            if (string.IsNullOrWhiteSpace(range)) return; // RequiredFields reports this.
            try
            {
                var parsed = ExcelRangeReference.Parse(range);
                if (string.IsNullOrWhiteSpace(targetSheet) && string.IsNullOrWhiteSpace(parsed.SheetName))
                    errors.Add(
                        $"ops[{index}] '{opName}' requires target.sheet or a sheet-qualified range; active sheet writes are not allowed");
                else if (!string.IsNullOrWhiteSpace(targetSheet) &&
                         !string.IsNullOrWhiteSpace(parsed.SheetName) &&
                         !string.Equals(targetSheet, parsed.SheetName, StringComparison.OrdinalIgnoreCase))
                    errors.Add(
                        $"ops[{index}] '{opName}' target.sheet '{targetSheet}' does not match range sheet '{parsed.SheetName}'");
            }
            catch (FormatException ex)
            {
                errors.Add($"ops[{index}] '{opName}' has invalid Excel range: {ex.Message}");
            }
            return;
        }

        if (opName is "insert_rows" or "insert_cols" or "set_rows_hidden" or "set_cols_hidden" or "set_sheet_visibility")
        {
            if (string.IsNullOrWhiteSpace(targetSheet))
                errors.Add($"ops[{index}] '{opName}' requires target.sheet; active sheet writes are not allowed");

            if (opName == "set_rows_hidden")
            {
                var row = Json.GetInt(op, "row") ?? 0;
                var count = Json.GetInt(op, "count") ?? 0;
                if (row < 1 || count < 1 || (long)row + count - 1 > 1_048_576)
                    errors.Add($"ops[{index}] 'set_rows_hidden' row/count must stay within Excel rows 1..1048576");
            }
            else if (opName == "set_cols_hidden")
            {
                var count = Json.GetInt(op, "count") ?? 0;
                if (!TryParseExcelColumn(op["col"], out var col) || count < 1 || (long)col + count - 1 > 16_384)
                    errors.Add($"ops[{index}] 'set_cols_hidden' col/count must stay within Excel columns A..XFD (1..16384)");
            }
            else if (opName == "set_sheet_visibility")
            {
                var visibility = Json.GetString(op, "visibility")?.ToLowerInvariant();
                if (visibility is not ("visible" or "hidden"))
                    errors.Add($"ops[{index}] 'set_sheet_visibility' visibility must be 'visible' or 'hidden'");
            }
            return;
        }

        if (opName == "find_replace")
        {
            var scope = (Json.GetString(target, "scope") ?? "sheet").ToLowerInvariant();
            if (scope is not ("sheet" or "workbook"))
                errors.Add($"ops[{index}] 'find_replace' target.scope must be 'sheet' or 'workbook'");
            else if (scope == "sheet" && string.IsNullOrWhiteSpace(targetSheet))
                errors.Add(
                    $"ops[{index}] 'find_replace' requires target.sheet when target.scope is 'sheet'; active sheet writes are not allowed");
        }
    }

    private static bool TryParseExcelColumn(JsonNode? node, out int column)
    {
        column = 0;
        if (node is not JsonValue value) return false;
        if (value.TryGetValue<int>(out column)) return column is >= 1 and <= 16_384;
        if (!value.TryGetValue<string>(out var text) || string.IsNullOrWhiteSpace(text)) return false;

        foreach (var ch in text.Trim().ToUpperInvariant())
        {
            if (ch is < 'A' or > 'Z') return false;
            try { column = checked(column * 26 + ch - 'A' + 1); }
            catch (OverflowException) { return false; }
            if (column > 16_384) return false;
        }
        return column >= 1;
    }
}
