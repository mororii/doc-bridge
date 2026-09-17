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
            ["set_row_heights"] = new[] { ("rows", "array") },
            ["set_column_widths"] = new[] { ("columns", "array") },
            ["freeze_panes"] = Array.Empty<(string Field, string Type)>(),
            ["rename_sheet"] = new[] { ("newName", "string") },
            ["clear_range"] = new[] { ("range", "string") },
            ["copy_range"] = new[] { ("range", "string"), ("destRange", "string") },
            ["delete_rows"] = new[] { ("row", "int"), ("count", "int") },
            ["delete_cols"] = new[] { ("col", "any"), ("count", "int") },
            ["add_sheet"] = new[] { ("name", "string") },
            ["move_sheet"] = new[] { ("position", "string") },
            ["protect_sheet"] = Array.Empty<(string Field, string Type)>(),
            ["unprotect_sheet"] = Array.Empty<(string Field, string Type)>(),
            ["create_workbook"] = Array.Empty<(string Field, string Type)>(),
            ["open_workbook"] = new[] { ("path", "string") },
            ["close_workbook"] = Array.Empty<(string Field, string Type)>(),
            ["save_workbook"] = Array.Empty<(string Field, string Type)>(),
            ["fill_range"] = new[] { ("range", "string") },
            ["auto_fill"] = new[] { ("range", "string"), ("destRange", "string") },
            ["calculate"] = Array.Empty<(string Field, string Type)>(),
            ["delete_sheet"] = Array.Empty<(string Field, string Type)>(),
            ["set_tab_color"] = new[] { ("color", "any") },
            ["set_outline"] = new[] { ("range", "string") },
            ["import_csv"] = new[] { ("path", "string") },
            ["export_csv"] = new[] { ("output", "string") },
            ["set_page_breaks"] = new[] { ("range", "string") },
            ["set_view"] = Array.Empty<(string Field, string Type)>(),
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
            ["regen_document"] = Array.Empty<(string Field, string Type)>(),
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

    /// <summary>
    /// 오류 응답에서 op별 선택 필드를 함께 안내하기 위한 최소 발견성 카탈로그.
    /// 실제 허용 여부는 어댑터가 계속 최종 검증하며, 이 목록은 모델이 첫 재시도에서
    /// 올바른 요청 모양을 만들 수 있도록 돕는 용도다.
    /// </summary>
    private static readonly Dictionary<string, string[]> OptionalFields =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["find_replace"] = new[] { "scope", "occurrence", "matchCase", "target", "file", "documentRef" },
            ["insert_text"] = new[] { "style", "preserveStyle", "styleSource", "file", "documentRef" },
            ["append_text"] = new[] { "startNewParagraph", "style", "preserveStyle", "styleSource", "file", "documentRef" },
            ["insert_before_text"] = new[] { "occurrence", "matchCase", "mode", "style", "preserveStyle", "styleSource", "file", "documentRef" },
            ["insert_after_text"] = new[] { "occurrence", "matchCase", "mode", "style", "preserveStyle", "styleSource", "file", "documentRef" },
            ["replace_document_text"] = new[] { "style", "preserveStyle", "file", "documentRef" },
            ["replace_selection"] = new[] { "style", "preserveStyle", "styleSource", "file", "documentRef" },
            ["set_paragraph_style_basic"] = new[] { "target", "file", "documentRef" },
            ["set_paragraph_format"] = new[] { "target", "file", "documentRef" },
            ["format_paragraphs"] = new[] { "file", "documentRef" },
            ["set_page_setup"] = new[] { "applyTo", "file", "documentRef" },
            ["insert_break"] = new[] { "file", "documentRef" },
            ["insert_table"] = new[]
            {
                "header", "headerFill", "firstColumnFill", "fontSize", "columnWidths",
                "cellStyles", "mergeCells", "verticalCenter", "hideAllBorders", "file", "documentRef",
            },
            ["table_cell_set_text"] = new[] { "tableIndex", "row", "col", "cellIndex", "preserveStyle", "style", "styleSource", "file", "documentRef" },
            ["table_set_cells"] = new[] { "tableIndex", "preserveStyle", "file", "documentRef" },
            ["table_insert_rows"] = new[] { "tableIndex", "col", "position", "file", "documentRef" },
            ["table_insert_columns"] = new[] { "tableIndex", "row", "position", "file", "documentRef" },
            ["table_delete_rows"] = new[] { "tableIndex", "col", "count", "file", "documentRef" },
            ["table_delete_columns"] = new[] { "tableIndex", "row", "count", "file", "documentRef" },
            ["table_merge_cells"] = new[] { "tableIndex", "file", "documentRef" },
            ["table_set_row_height"] = new[] { "tableIndex", "file", "documentRef" },
            ["table_set_row_heights"] = new[] { "tableIndex", "file", "documentRef" },
            ["set_field_text"] = new[] { "file", "documentRef" },
            ["insert_picture"] = new[]
            {
                "tableIndex", "row", "col", "cellIndex", "clearCell", "embedded", "sizeOption",
                "widthMm", "heightMm", "effect", "reverse", "watermark", "file", "documentRef",
            },
            ["insert_page_number"] = new[] { "position", "format", "startNumber", "file", "documentRef" },
            ["set_header_footer_text"] = new[] { "pages", "file", "documentRef" },
            ["set_values"] = new[] { "target", "targetWorkbook" },
            ["set_formulas"] = new[] { "target", "targetWorkbook" },
            ["format_range"] = new[] { "target", "targetWorkbook" },
            ["merge_cells"] = new[] { "target", "targetWorkbook" },
            ["unmerge_cells"] = new[] { "target", "targetWorkbook" },
            ["set_row_heights"] = new[] { "target", "targetWorkbook" },
            ["set_column_widths"] = new[] { "target", "targetWorkbook" },
            ["freeze_panes"] = new[] { "target", "targetWorkbook", "cell", "rows", "columns", "unfreeze" },
            ["rename_sheet"] = new[] { "target", "targetWorkbook" },
            ["clear_range"] = new[] { "target", "targetWorkbook", "what" },
            ["copy_range"] = new[] { "target", "targetWorkbook", "destSheet", "mode" },
            ["delete_rows"] = new[] { "target", "targetWorkbook" },
            ["delete_cols"] = new[] { "target", "targetWorkbook" },
            ["add_sheet"] = new[] { "target", "targetWorkbook", "afterSheet" },
            ["move_sheet"] = new[] { "target", "targetWorkbook", "relativeSheet" },
            ["protect_sheet"] = new[]
            {
                "target", "targetWorkbook", "options", "protection",
                "drawingObjects", "contents", "scenarios", "userInterfaceOnly",
                "allowFormattingCells", "allowFormattingColumns", "allowFormattingRows",
                "allowInsertingColumns", "allowInsertingRows", "allowInsertingHyperlinks",
                "allowDeletingColumns", "allowDeletingRows", "allowSorting", "allowFiltering",
                "allowUsingPivotTables",
            },
            ["unprotect_sheet"] = new[] { "target", "targetWorkbook" },
            ["create_workbook"] = new[] { "sheetName" },
            ["open_workbook"] = new[] { "targetWorkbook" },
            ["close_workbook"] = new[] { "saveChanges", "target", "targetWorkbook" },
            ["save_workbook"] = new[] { "output", "overwrite", "target", "targetWorkbook" },
            ["fill_range"] = new[] { "target", "targetWorkbook", "value", "values" },
            ["auto_fill"] = new[] { "target", "targetWorkbook", "type" },
            ["calculate"] = new[] { "target", "targetWorkbook", "range", "formula2" },
            ["delete_sheet"] = new[] { "target", "targetWorkbook" },
            ["set_tab_color"] = new[] { "target", "targetWorkbook" },
            ["set_outline"] = new[] { "target", "targetWorkbook", "level", "summaryBelow", "summaryRight", "show" },
            ["import_csv"] = new[] { "target", "targetWorkbook", "destination", "delimiter" },
            ["export_csv"] = new[] { "target", "targetWorkbook", "overwrite", "sheet" },
            ["set_page_breaks"] = new[] { "target", "targetWorkbook", "clear" },
            ["set_view"] = new[]
            {
                "target", "targetWorkbook", "zoom", "view",
                "displayGridlines", "displayHeadings", "displayZeros",
            },
            ["export_pdf"] = new[] { "file", "documentRef", "output", "overwrite", "target", "targetWorkbook", "sheet" },
            ["copy_sheet"] = new[] { "targetSheet", "targetWorkbook" },
        };

    private static readonly Dictionary<(string Op, string Field), string> FieldExpectations = new()
    {
        [("insert_table", "rows")] = "array of row arrays, e.g. [[\"A\",\"B\"],[\"C\",\"D\"]]",
        [("table_set_row_heights", "rows")] = "array of objects, e.g. [{\"row\":0,\"heightMm\":8.0}]",
        [("table_set_cells", "cells")] = "array of cell objects, e.g. [{\"row\":0,\"col\":0,\"text\":\"A\"}]",
        [("format_paragraphs", "items")] = "array of format objects with target and characterStyle and/or paragraphStyle",
        [("set_page_setup", "page")] = "object, e.g. Excel {\"paperSize\":\"A3\",\"orientation\":\"landscape\",\"scale\":55} or HWP {\"widthMm\":210,\"heightMm\":297,\"orientation\":\"portrait\"}",
        [("set_row_heights", "rows")] = "array of {row, count?, heightPoints?, autoFit?}",
        [("set_column_widths", "columns")] = "array of {col, count?, widthChars?, autoFit?}",
        [("clear_range", "what")] = "all|contents|formats|formulas",
        [("copy_range", "mode")] = "all|values|formulas|formats",
        [("move_sheet", "position")] = "before|after|first|last",
        [("draw_entities", "entities")] = "array of CAD entity objects",
        [("set_values", "values")] = "2D array of cell values, e.g. [[1,2],[3,4]]",
        [("set_formulas", "formulas")] = "2D array of formula strings, e.g. [[\"=SUM(A1:A2)\"]]",
        [("format_range", "style")] = ExcelStyleContract.StyleExpectation,
    };

    public const string ExecutionModeLegacy = "legacy";
    public const string ExecutionModeExecute = "execute";

    public sealed record ParsedBatch(
        List<JsonObject> Ops,
        bool DryRun,
        string? ConfirmToken,
        bool HighRiskConfirm,
        bool HasHighRiskOps,
        IReadOnlyList<string> OptimizationWarnings,
        string ExecutionMode = ExecutionModeLegacy,
        string? RequestId = null,
        string? ExpectedDocumentRef = null,
        string? ExplicitDocumentRef = null);

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

            if (!RequiredFields.TryGetValue(name, out var rules) &&
                app.Equals("excel", StringComparison.OrdinalIgnoreCase))
                ExcelDataOperationsContract.RequiredFields.TryGetValue(name, out rules);
            if (rules is not null)
                foreach (var (field, type) in rules)
                    ValidateField(op, i, name, field, type, errors);

            if (app.Equals("excel", StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelTarget(op, i, name, errors);
                if (string.Equals(name, "format_range", StringComparison.OrdinalIgnoreCase))
                    ExcelStyleContract.Validate(Json.GetObj(op, "style"), i, errors);
                if (string.Equals(name, "set_page_setup", StringComparison.OrdinalIgnoreCase))
                    ExcelPageSetupContract.TryNormalize(Json.GetObj(op, "page"), i, errors, out _);
                if (string.Equals(name, "rename_sheet", StringComparison.OrdinalIgnoreCase))
                    ExcelSheetNameContract.TryNormalize(Json.GetString(op, "newName"), i, "rename_sheet.newName", errors, out _);
                if (string.Equals(name, "add_sheet", StringComparison.OrdinalIgnoreCase))
                    ExcelSheetNameContract.TryNormalize(Json.GetString(op, "name"), i, "add_sheet.name", errors, out _);
                if (ExcelDataOperationsContract.IsDataOperation(name))
                    ExcelDataOperationsContract.ValidatePublicInput(op, i, errors);
            }

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
            ValidateExcelBatchFamilies(ops, errors);
            ValidateExcelMergeBatchGeometry(ops, errors);
            ValidateFormatBorderOrder(ops, errors);
        }

        var dryRun = !batch.TryGetPropertyValue("dryRun", out var dv) || Json.GetBool(batch, "dryRun", true);
        var confirmToken = Json.GetString(batch, "confirmToken");
        var highRiskConfirm = Json.GetBool(batch, "highRiskConfirm", false);
        var executionModeRaw = Json.GetString(batch, "executionMode");
        var isExecute = false;
        string? requestId = null;
        string? expectedDocumentRef = null;
        string? explicitDocumentRef = null;

        if (batch.ContainsKey("executionMode"))
        {
            if (!string.Equals(executionModeRaw, ExecutionModeExecute, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add("executionMode must be 'execute' when present; omit it for the legacy dry-run/token path");
            }
            else
            {
                isExecute = true;
                dryRun = false;
                if (batch.ContainsKey("dryRun"))
                    errors.Add("executionMode=execute cannot be combined with dryRun; use the legacy token path for previews");
                if (batch.ContainsKey("confirmToken"))
                    errors.Add("executionMode=execute cannot be combined with confirmToken");
                if (batch.ContainsKey("highRiskConfirm"))
                    errors.Add("executionMode=execute cannot be combined with highRiskConfirm; high-risk edits require the legacy review path");

                if (!ExecuteJournalService.TryNormalizeRequestId(Json.GetString(batch, "requestId"), out var normalizedId))
                    errors.Add("executionMode=execute requires requestId as a UUID");
                else
                    requestId = normalizedId;

                expectedDocumentRef = Json.GetString(batch, "expectedDocumentRef");
                if (string.IsNullOrWhiteSpace(expectedDocumentRef))
                    errors.Add("executionMode=execute requires expectedDocumentRef");

                foreach (var op in ops)
                {
                    var name = Json.GetString(op, "op") ?? "";
                    var classification = _policy.ClassifyOp(app, name);
                    if (classification is OpClass.HighRisk or OpClass.Forbidden or OpClass.Unknown)
                        errors.Add($"executionMode=execute cannot run '{name}' ({classification}); use the legacy dry-run review path");
                    else if (!_policy.IsAutoExecutable(app, name))
                        errors.Add(
                            $"executionMode=execute does not allow '{name}' on {app}; " +
                            "structural edits, deletion, save/export, cross-document, and activate_document stay on the token path");
                }

                if (!string.IsNullOrWhiteSpace(expectedDocumentRef) && errors.Count == 0)
                {
                    var bind = DocumentTargetBinder.BindForExecute(app, ops, expectedDocumentRef);
                    if (!bind.Ok)
                        errors.AddRange(bind.Errors);
                    else
                    {
                        expectedDocumentRef = bind.NormalizedExpected ?? expectedDocumentRef;
                        explicitDocumentRef = bind.ExplicitDocumentRef;
                    }
                }
            }
        }

        if (errors.Count > 0) return null;
        return new ParsedBatch(
            ops, dryRun, confirmToken, highRiskConfirm, hasHighRisk,
            BuildOptimizationWarnings(app, ops),
            isExecute ? ExecutionModeExecute : ExecutionModeLegacy,
            requestId, expectedDocumentRef, explicitDocumentRef);
    }

    /// <summary>
    /// 검증 실패 응답에 동봉할 op별 기계 판독 가능한 필드 목록.
    /// 인덱스는 사용자 오류 메시지와 맞추기 위해 1부터 시작한다.
    /// </summary>
    public JsonArray DescribeExpectedSchemas(JsonObject? batch, string app)
    {
        var result = new JsonArray();
        var ops = Json.GetArr(batch, "ops");
        if (ops is null) return result;

        for (var index = 0; index < ops.Count; index++)
        {
            if (ops[index] is not JsonObject op) continue;
            var opName = Json.GetString(op, "op");
            if (string.IsNullOrWhiteSpace(opName)) continue;
            if (!RequiredFields.TryGetValue(opName, out var rules) &&
                app.Equals("excel", StringComparison.OrdinalIgnoreCase))
                ExcelDataOperationsContract.RequiredFields.TryGetValue(opName, out rules);
            if (rules is null) continue;

            var required = new JsonObject();
            foreach (var (field, type) in rules)
                required[field] = ExpectedFieldDescription(opName, field, type);

            var optional = OptionalFields.TryGetValue(opName, out var fields)
                ? fields
                : ExcelDataOperationsContract.OptionalFields.TryGetValue(opName, out fields)
                    ? fields
                    : CommonOptionalFields(app);
            result.Add(new JsonObject
            {
                ["index"] = index + 1,
                ["op"] = opName,
                ["required"] = required,
                ["optional"] = Json.ToArray(optional),
            });
        }
        return result;
    }

    private static void ValidateField(JsonObject op, int index, string opName, string field, string type, List<string> errors)
    {
        if (!op.TryGetPropertyValue(field, out var v) || v is null)
        {
            errors.Add($"ops[{index}] '{opName}' requires field '{field}' ({ExpectedFieldDescription(opName, field, type)})");
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
        if (bad)
            errors.Add($"ops[{index}] '{opName}' field '{field}' must be {ExpectedFieldDescription(opName, field, type)}");
    }

    private static string ExpectedFieldDescription(string opName, string field, string type) =>
        FieldExpectations.TryGetValue((opName, field), out var expectation)
            ? expectation
            : ExcelDataOperationsContract.FieldExpectations.TryGetValue((opName, field), out expectation)
                ? expectation
            : type switch
            {
                "string" => "string",
                "int" => "integer",
                "number" => "number",
                "bool" => "boolean",
                "array" => "array",
                "object" => "object",
                _ => "value",
            };

    private static string[] CommonOptionalFields(string app) => app.ToLowerInvariant() switch
    {
        "hwp" => new[] { "file", "documentRef" },
        "excel" => new[] { "target", "targetWorkbook" },
        "cad" or "gstarcad" => new[] { "document" },
        _ => Array.Empty<string>(),
    };

    private static IReadOnlyList<string> BuildOptimizationWarnings(string app, IReadOnlyList<JsonObject> ops)
    {
        var warnings = new List<string>();
        if (!app.Equals("hwp", StringComparison.OrdinalIgnoreCase)) return warnings;

        for (var start = 0; start < ops.Count;)
        {
            var name = Json.GetString(ops[start], "op");
            var tableIndex = Json.GetInt(ops[start], "tableIndex") ?? 0;
            var end = start + 1;
            while (end < ops.Count &&
                   string.Equals(Json.GetString(ops[end], "op"), name, StringComparison.OrdinalIgnoreCase) &&
                   (Json.GetInt(ops[end], "tableIndex") ?? 0) == tableIndex)
                end++;

            var count = end - start;
            if (string.Equals(name, "table_set_row_height", StringComparison.OrdinalIgnoreCase) && count >= 3)
                warnings.Add(
                    $"ops[{start + 1}..{end}] contains {count} consecutive table_set_row_height operations for table {tableIndex}; " +
                    "use one table_set_row_heights op with rows:[{row,heightMm}, ...] to avoid repeated COM validation cycles");
            start = end;
        }
        return warnings;
    }

    internal static readonly HashSet<string> ExcelMergeOpNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "merge_cells", "unmerge_cells",
    };

    internal static readonly HashSet<string> ExcelVisibilityOpNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "set_rows_hidden", "set_cols_hidden", "set_sheet_visibility",
    };

    internal static readonly HashSet<string> ExcelSheetLayoutOpNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "set_row_heights", "set_column_widths", "freeze_panes", "set_page_setup", "set_view",
    };

    internal static readonly HashSet<string> ExcelRenameOpNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "rename_sheet",
    };

    internal static readonly HashSet<string> ExcelStructureOpNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "add_sheet", "move_sheet", "delete_sheet", "set_tab_color",
    };

    internal static readonly HashSet<string> ExcelDeleteOpNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "delete_rows", "delete_cols",
    };

    internal static readonly HashSet<string> ExcelRangeEditOpNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "clear_range", "copy_range",
    };

    internal static readonly HashSet<string> ExcelProtectOpNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "protect_sheet", "unprotect_sheet",
    };

    internal static readonly HashSet<string> ExcelLifecycleOpNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "create_workbook", "open_workbook", "close_workbook", "save_workbook", "export_pdf",
    };

    internal static readonly HashSet<string> ExcelExtendedOpNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "fill_range", "auto_fill", "calculate", "set_outline", "import_csv", "export_csv", "set_page_breaks",
    };

    private static void ValidateExcelBatchFamilies(IReadOnlyList<JsonObject> ops, List<string> errors)
    {
        var mergeNames = ops
            .Select(op => Json.GetString(op, "op") ?? "")
            .Where(name => ExcelMergeOpNames.Contains(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (mergeNames.Count > 0 && ops.Any(op => !ExcelMergeOpNames.Contains(Json.GetString(op, "op") ?? "")))
        {
            errors.Add(
                "Excel merge_cells/unmerge_cells cannot be mixed with other operations; " +
                "apply content edits and formatting in separate dry-run batches");
        }
        else if (mergeNames.Count > 1)
        {
            errors.Add(
                "Excel merge_cells and unmerge_cells cannot share a batch; " +
                "run one merge-only batch or one unmerge-only batch so snapshot restore stays exact");
        }

        RequireFamilyIsolation(ops, ExcelVisibilityOpNames,
            "Excel visibility operations cannot be mixed with content, format, copy, or merge operations; " +
            "use a separate dry-run batch so visibility rollback remains exact", errors);
        RequireFamilyIsolation(ops, ExcelSheetLayoutOpNames,
            "Excel row/column size, freeze_panes, and set_page_setup must stay in a sheet-layout-only batch", errors);
        RequireFamilyIsolation(ops, ExcelRenameOpNames,
            "Excel rename_sheet must be the only operation family in its batch", errors);
        RequireFamilyIsolation(ops, ExcelStructureOpNames,
            "Excel add_sheet/move_sheet cannot be mixed with cell edits; copy_sheet stays in its own batch", errors);
        RequireFamilyIsolation(ops, ExcelDeleteOpNames,
            "Excel delete_rows/delete_cols are high-risk and must be isolated from other operations", errors);
        RequireFamilyIsolation(ops, ExcelRangeEditOpNames,
            "Excel clear_range/copy_range must stay in a range-edit-only batch", errors);
        RequireFamilyIsolation(ops, ExcelProtectOpNames,
            "Excel protect_sheet/unprotect_sheet must stay in a protection-only batch", errors);
        RequireFamilyIsolation(ops, ExcelLifecycleOpNames,
            "Excel create_workbook/open_workbook/close_workbook/save_workbook/export_pdf must stay in a lifecycle-only batch", errors);
        RequireFamilyIsolation(ops, ExcelExtendedOpNames,
            "Excel fill/calculate/outline/csv/pagebreak operations must stay in an extended-only batch", errors);
        ExcelDataOperationsContract.ValidateDataBatchMixing(ops, errors);

        var lifecycle = ops.Select(op => Json.GetString(op, "op") ?? "")
            .Where(name => ExcelLifecycleOpNames.Contains(name))
            .ToList();
        if (lifecycle.Count > 1 &&
            lifecycle.Any(name => name is "create_workbook" or "open_workbook" or "close_workbook"))
        {
            errors.Add("Excel create_workbook, open_workbook, and close_workbook must be the only op in their batch");
        }
    }

    private static void RequireFamilyIsolation(
        IReadOnlyList<JsonObject> ops, HashSet<string> family, string message, List<string> errors)
    {
        if (ops.Any(op => family.Contains(Json.GetString(op, "op") ?? "")) &&
            ops.Any(op => !family.Contains(Json.GetString(op, "op") ?? "")))
            errors.Add(message);
    }

    private static void ValidateExcelMergeBatchGeometry(IReadOnlyList<JsonObject> ops, List<string> errors)
    {
        var seen = new List<(string Sheet, ExcelA1Box Box, int Index)>();
        for (var i = 0; i < ops.Count; i++)
        {
            var name = Json.GetString(ops[i], "op") ?? "";
            if (!ExcelMergeOpNames.Contains(name)) continue;
            var range = Json.GetString(ops[i], "range");
            if (string.IsNullOrWhiteSpace(range)) continue;
            try
            {
                var parsed = ExcelRangeReference.Parse(range);
                var sheet = Json.GetString(Json.GetObj(ops[i], "target"), "sheet") ?? parsed.SheetName ?? "";
                if (!ExcelA1Box.TryParse(parsed.Address, out var box))
                {
                    errors.Add($"ops[{i + 1}] '{name}' range must be one contiguous A1 rectangle");
                    continue;
                }

                foreach (var prior in seen)
                {
                    if (!string.Equals(prior.Sheet, sheet, StringComparison.OrdinalIgnoreCase)) continue;
                    if (prior.Box.Intersects(box))
                    {
                        errors.Add(
                            $"[EXCEL_MERGE_BATCH_OVERLAP] ops[{prior.Index}] and ops[{i + 1}] overlap on '{sheet}' " +
                            $"({prior.Box.Address} vs {box.Address}); none of the batch will be applied");
                    }
                }

                seen.Add((sheet, box, i + 1));
            }
            catch (FormatException)
            {
                // ValidateExcelTarget already reported the range error.
            }
        }

        if (ops.Count(op => ExcelMergeOpNames.Contains(Json.GetString(op, "op") ?? "")) > 400)
            errors.Add("Excel merge batch is limited to 400 merge/unmerge operations");
    }

    /// <summary>
    /// Adjacent ranges share one edge object: clearing E16:E18 also clears the
    /// left edge of an F16:H18 outline. A clear that runs after a draw on a
    /// shared edge deterministically destroys the drawing, so it is rejected
    /// here (same range and different sheets are exempt). Cross-batch
    /// interference is reported as a preview warning instead.
    /// </summary>
    private static void ValidateFormatBorderOrder(IReadOnlyList<JsonObject> ops, List<string> errors)
    {
        var formats = new List<(int Index, string Sheet, ExcelA1Box Box, HashSet<string> Clears, bool Draws)>();
        for (var i = 0; i < ops.Count; i++)
        {
            if (!string.Equals(Json.GetString(ops[i], "op"), "format_range", StringComparison.OrdinalIgnoreCase))
                continue;
            var rangeText = Json.GetString(ops[i], "range");
            if (string.IsNullOrWhiteSpace(rangeText)) continue;
            try
            {
                var parsed = ExcelRangeReference.Parse(rangeText);
                var sheet = Json.GetString(Json.GetObj(ops[i], "target"), "sheet") ?? parsed.SheetName ?? "";
                if (!ExcelA1Box.TryParse(parsed.Address, out var box)) continue;
                var borders = Json.GetObj(Json.GetObj(ops[i], "style"), "borders");
                formats.Add((i + 1, sheet, box, ClearedOuterEdges(borders), DrawsOuterEdges(borders)));
            }
            catch (FormatException)
            {
                // ValidateExcelTarget already reported the range error.
            }
        }

        for (var a = 0; a < formats.Count; a++)
        {
            for (var b = a + 1; b < formats.Count; b++)
            {
                var first = formats[a];
                var second = formats[b];
                if (string.IsNullOrWhiteSpace(first.Sheet) || string.IsNullOrWhiteSpace(second.Sheet) ||
                    !string.Equals(first.Sheet, second.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (first.Box.Address == second.Box.Address) continue;
                // Only a draw followed by a clear on a shared edge is rejected.
                // Clear-then-draw is the supported order: the outline is drawn last.
                if (first.Draws && second.Clears.Count > 0 &&
                    SharedEdges(first.Box, second.Box) is { Count: > 0 } shared)
                {
                    errors.Add(
                        $"[EXCEL_BORDER_ORDER] ops[{first.Index}] draws borders and ops[{second.Index}] clears " +
                        $"a shared edge on '{first.Sheet}' ({first.Box.Address} vs {second.Box.Address}: " +
                        $"{string.Join(",", shared)}); put clears before draws in one batch or split the " +
                        "batches so outlines are drawn last");
                }
            }
        }
    }

    private static readonly HashSet<string> BorderOuterEdgeKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "left", "right", "top", "bottom", "outline", "all",
    };

    private static HashSet<string> ClearedOuterEdges(JsonObject? borders)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (borders is null) return result;
        foreach (var (key, node) in borders)
        {
            if (!BorderOuterEdgeKeys.Contains(key)) continue;
            if (node is JsonValue value && value.TryGetValue<string>(out var text) &&
                string.Equals(text, "none", StringComparison.OrdinalIgnoreCase))
            {
                if (key is "all" or "outline")
                    result.UnionWith(new[] { "left", "right", "top", "bottom" });
                else
                    result.Add(key.ToLowerInvariant());
            }
        }
        return result;
    }

    private static bool DrawsOuterEdges(JsonObject? borders)
    {
        if (borders is null) return false;
        foreach (var (key, node) in borders)
        {
            if (!BorderOuterEdgeKeys.Contains(key) || node is null) continue;
            if (node is JsonValue value && value.TryGetValue<string>(out var text) &&
                string.Equals(text, "none", StringComparison.OrdinalIgnoreCase))
                continue;
            return true;
        }
        return false;
    }

    private static List<string> SharedEdges(ExcelA1Box first, ExcelA1Box second)
    {
        var shared = new List<string>();
        var rowsOverlap = first.Row < second.Row + second.Rows && second.Row < first.Row + first.Rows;
        var colsOverlap = first.Column < second.Column + second.Columns && second.Column < first.Column + first.Columns;
        if (first.Column + first.Columns == second.Column && rowsOverlap) shared.Add("right=left");
        if (second.Column + second.Columns == first.Column && rowsOverlap) shared.Add("left=right");
        if (first.Row + first.Rows == second.Row && colsOverlap) shared.Add("bottom=top");
        if (second.Row + second.Rows == first.Row && colsOverlap) shared.Add("top=bottom");
        return shared;
    }

    private static void ValidateExcelTarget(JsonObject op, int index, string opName, List<string> errors)
    {
        var target = Json.GetObj(op, "target");
        var targetSheet = Json.GetString(target, "sheet");

        if (op.ContainsKey("password") || Json.GetObj(op, "protection")?.ContainsKey("password") == true)
            errors.Add($"ops[{index}] '{opName}' must not include a password; password-protected sheets are out of scope");

        if (opName is "set_values" or "set_formulas" or "format_range" or "merge_cells" or "unmerge_cells"
            or "clear_range" or "copy_range" or "fill_range" or "auto_fill" or "set_outline" or "set_page_breaks")
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

        if (opName is "insert_rows" or "insert_cols" or "delete_rows" or "delete_cols"
            or "set_rows_hidden" or "set_cols_hidden" or "set_sheet_visibility"
            or "set_row_heights" or "set_column_widths" or "freeze_panes"
            or "rename_sheet" or "set_page_setup" or "set_view" or "protect_sheet" or "unprotect_sheet"
            or "move_sheet" or "delete_sheet" or "set_tab_color")
        {
            if (opName != "add_sheet" && string.IsNullOrWhiteSpace(targetSheet))
                errors.Add($"ops[{index}] '{opName}' requires target.sheet; active sheet writes are not allowed");

            if (opName is "set_rows_hidden" or "delete_rows")
            {
                var row = Json.GetInt(op, "row") ?? 0;
                var count = Json.GetInt(op, "count") ?? 0;
                if (row < 1 || count < 1 || (long)row + count - 1 > 1_048_576)
                    errors.Add($"ops[{index}] '{opName}' row/count must stay within Excel rows 1..1048576");
            }
            else if (opName is "set_cols_hidden" or "delete_cols")
            {
                var count = Json.GetInt(op, "count") ?? 0;
                if (!TryParseExcelColumn(op["col"], out var col) || count < 1 || (long)col + count - 1 > 16_384)
                    errors.Add($"ops[{index}] '{opName}' col/count must stay within Excel columns A..XFD (1..16384)");
            }
            else if (opName == "set_sheet_visibility")
            {
                var visibility = Json.GetString(op, "visibility")?.ToLowerInvariant();
                if (visibility is not ("visible" or "hidden" or "veryhidden"))
                    errors.Add($"ops[{index}] 'set_sheet_visibility' visibility must be 'visible', 'hidden', or 'veryHidden'");
            }
            else if (opName == "set_row_heights")
                ValidateRowHeightItems(op, index, errors);
            else if (opName == "set_column_widths")
                ValidateColumnWidthItems(op, index, errors);
            else if (opName == "freeze_panes")
                ValidateFreezePanes(op, index, errors);
            else if (opName == "set_view")
            {
                var hasAny = op.ContainsKey("zoom") || op.ContainsKey("view") ||
                             op.ContainsKey("displayGridlines") || op.ContainsKey("displayHeadings") ||
                             op.ContainsKey("displayZeros");
                if (!hasAny)
                    errors.Add($"ops[{index}] 'set_view' requires zoom, view, displayGridlines, displayHeadings, or displayZeros");
                if (Json.GetInt(op, "zoom") is int zoom && zoom is < 10 or > 400)
                    errors.Add($"ops[{index}] 'set_view' zoom must be 10..400");
                var view = Json.GetString(op, "view");
                if (!string.IsNullOrWhiteSpace(view) &&
                    view.ToLowerInvariant() is not ("normal" or "pagelayout" or "pagebreakpreview"))
                    errors.Add($"ops[{index}] 'set_view' view must be normal|pageLayout|pageBreakPreview");
            }
            else if (opName == "move_sheet")
            {
                var position = Json.GetString(op, "position")?.ToLowerInvariant();
                if (position is not ("before" or "after" or "first" or "last"))
                    errors.Add($"ops[{index}] 'move_sheet' position must be before|after|first|last");
                if (position is "before" or "after" && string.IsNullOrWhiteSpace(Json.GetString(op, "relativeSheet")))
                    errors.Add($"ops[{index}] 'move_sheet' requires relativeSheet when position is before/after");
            }
            return;
        }

        if (opName == "clear_range")
        {
            var what = (Json.GetString(op, "what") ?? "all").ToLowerInvariant();
            if (what is not ("all" or "contents" or "formats" or "formulas"))
                errors.Add($"ops[{index}] 'clear_range' what must be all|contents|formats|formulas");
        }

        if (opName == "copy_range")
        {
            var mode = (Json.GetString(op, "mode") ?? "all").ToLowerInvariant();
            if (mode is not ("all" or "values" or "formulas" or "formats"))
                errors.Add($"ops[{index}] 'copy_range' mode must be all|values|formulas|formats");
            var dest = Json.GetString(op, "destRange");
            if (!string.IsNullOrWhiteSpace(dest))
            {
                try { _ = ExcelRangeReference.Parse(dest); }
                catch (FormatException ex) { errors.Add($"ops[{index}] 'copy_range' has invalid destRange: {ex.Message}"); }
            }
        }

        if (opName is "save_workbook" or "export_pdf")
        {
            var output = Json.GetString(op, "output");
            if (opName == "export_pdf" && string.IsNullOrWhiteSpace(output))
                errors.Add($"ops[{index}] 'export_pdf' requires output");
            if (!string.IsNullOrWhiteSpace(output) && ExcelAuthoringPaths.IsProtectedSource(output))
                errors.Add($"ops[{index}] '{opName}' refuses to write the protected source workbook");
        }

        if (opName == "open_workbook")
        {
            var path = Json.GetString(op, "path");
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path))
                errors.Add($"ops[{index}] 'open_workbook' path must be an absolute existing-file path");
        }

        if (opName == "close_workbook")
        {
            var workbook = Json.GetString(target, "workbook") ?? Json.GetString(op, "targetWorkbook");
            if (string.IsNullOrWhiteSpace(workbook))
                errors.Add(
                    $"ops[{index}] 'close_workbook' requires target.workbook (or targetWorkbook); active workbook close is not allowed");
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

    private static void ValidateRowHeightItems(JsonObject op, int index, List<string> errors)
    {
        var rows = Json.GetArr(op, "rows");
        if (rows is null || rows.Count == 0)
        {
            errors.Add($"ops[{index}] 'set_row_heights' rows must be a non-empty array");
            return;
        }

        if (rows.Count > 500)
            errors.Add($"ops[{index}] 'set_row_heights' rows is limited to 500 items");
        for (var i = 0; i < rows.Count; i++)
        {
            if (rows[i] is not JsonObject item)
            {
                errors.Add($"ops[{index}] 'set_row_heights' rows[{i}] must be an object");
                continue;
            }

            var row = Json.GetInt(item, "row") ?? 0;
            var count = Json.GetInt(item, "count") ?? 1;
            var autoFit = Json.GetBool(item, "autoFit");
            var hasHeight = item.ContainsKey("heightPoints");
            if (row < 1 || count < 1 || (long)row + count - 1 > 1_048_576)
                errors.Add($"ops[{index}] 'set_row_heights' rows[{i}] row/count must stay within 1..1048576");
            if (autoFit == hasHeight)
                errors.Add($"ops[{index}] 'set_row_heights' rows[{i}] must set exactly one of heightPoints or autoFit:true");
            if (hasHeight && (!TryReadPositiveNumber(item["heightPoints"], out var height) || height is < 0.1 or > 409.5))
                errors.Add($"ops[{index}] 'set_row_heights' rows[{i}].heightPoints must be 0.1..409.5 (Excel points)");
        }
    }

    private static void ValidateColumnWidthItems(JsonObject op, int index, List<string> errors)
    {
        var columns = Json.GetArr(op, "columns");
        if (columns is null || columns.Count == 0)
        {
            errors.Add($"ops[{index}] 'set_column_widths' columns must be a non-empty array");
            return;
        }

        if (columns.Count > 500)
            errors.Add($"ops[{index}] 'set_column_widths' columns is limited to 500 items");
        for (var i = 0; i < columns.Count; i++)
        {
            if (columns[i] is not JsonObject item)
            {
                errors.Add($"ops[{index}] 'set_column_widths' columns[{i}] must be an object");
                continue;
            }

            var count = Json.GetInt(item, "count") ?? 1;
            var autoFit = Json.GetBool(item, "autoFit");
            var hasWidth = item.ContainsKey("widthChars");
            if (!TryParseExcelColumn(item["col"], out var col) || count < 1 || (long)col + count - 1 > 16_384)
                errors.Add($"ops[{index}] 'set_column_widths' columns[{i}] col/count must stay within A..XFD");
            if (autoFit == hasWidth)
                errors.Add($"ops[{index}] 'set_column_widths' columns[{i}] must set exactly one of widthChars or autoFit:true");
            if (hasWidth && (!TryReadPositiveNumber(item["widthChars"], out var width) || width is < 0 or > 255))
                errors.Add($"ops[{index}] 'set_column_widths' columns[{i}].widthChars must be 0..255 (Excel character units)");
        }
    }

    private static void ValidateFreezePanes(JsonObject op, int index, List<string> errors)
    {
        var unfreeze = Json.GetBool(op, "unfreeze");
        var hasCell = !string.IsNullOrWhiteSpace(Json.GetString(op, "cell"));
        var hasRows = op.ContainsKey("rows");
        var hasCols = op.ContainsKey("columns");
        var freezeSelectors = (unfreeze ? 1 : 0) + (hasCell ? 1 : 0) + ((hasRows || hasCols) ? 1 : 0);
        if (freezeSelectors != 1)
        {
            errors.Add($"ops[{index}] 'freeze_panes' must set exactly one of unfreeze:true, cell, or rows/columns");
            return;
        }

        if (hasCell && !ExcelA1Box.TryParseCell(Json.GetString(op, "cell")!, out _, out _))
            errors.Add($"ops[{index}] 'freeze_panes' cell must be a single A1 address such as G6");
        if (hasRows && (Json.GetInt(op, "rows") is not int rows || rows < 0 || rows > 1_048_575))
            errors.Add($"ops[{index}] 'freeze_panes' rows must be 0..1048575");
        if (hasCols && (Json.GetInt(op, "columns") is not int cols || cols < 0 || cols > 16_383))
            errors.Add($"ops[{index}] 'freeze_panes' columns must be 0..16383");
    }

    private static bool TryReadPositiveNumber(JsonNode? node, out double number)
    {
        number = 0;
        if (node is not JsonValue value) return false;
        if (value.TryGetValue<double>(out number) && double.IsFinite(number)) return true;
        if (value.TryGetValue<int>(out var i)) { number = i; return true; }
        return false;
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
