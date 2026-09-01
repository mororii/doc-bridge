using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>테스트마다 독립된 루트 디렉터리를 주어 파일 상태가 섞이지 않게 한다</summary>
public sealed class TestHome : IDisposable
{
    public string Dir { get; }
    public TestHome()
    {
        Dir = Path.Combine(Path.GetTempPath(), "docbridge-test-" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(Dir);
    }
    public DocBridgeOptions Options => new(Dir);
    public void Dispose()
    {
        try { Directory.Delete(Dir, recursive: true); } catch { }
    }
}

public class PolicyEngineTests : IClassFixture<TestHome>
{
    private readonly PolicyEngine _policy = new();

    [Theory]
    [InlineData("excel", "set_values", OpClass.Allowed)]
    [InlineData("excel", "find_replace", OpClass.Allowed)]
    [InlineData("excel", "copy_sheet", OpClass.Allowed)]
    [InlineData("excel", "merge_cells", OpClass.Allowed)]
    [InlineData("excel", "unmerge_cells", OpClass.Allowed)]
    [InlineData("excel", "set_rows_hidden", OpClass.Allowed)]
    [InlineData("excel", "set_cols_hidden", OpClass.Allowed)]
    [InlineData("excel", "set_sheet_visibility", OpClass.Allowed)]
    [InlineData("excel", "delete_sheet", OpClass.Forbidden)]
    [InlineData("excel", "run_macro", OpClass.Forbidden)]
    [InlineData("hwp", "insert_text", OpClass.Allowed)]
    [InlineData("hwp", "append_text", OpClass.Allowed)]
    [InlineData("hwp", "insert_before_text", OpClass.Allowed)]
    [InlineData("hwp", "insert_after_text", OpClass.Allowed)]
    [InlineData("hwp", "replace_document_text", OpClass.Allowed)]
    [InlineData("hwp", "table_cell_set_text", OpClass.Allowed)]
    [InlineData("hwp", "table_insert_rows", OpClass.Allowed)]
    [InlineData("hwp", "table_delete_rows", OpClass.HighRisk)]
    [InlineData("hwp", "table_merge_cells", OpClass.Allowed)]
    [InlineData("hwp", "table_set_row_height", OpClass.Allowed)]
    [InlineData("hwp", "table_set_row_heights", OpClass.Allowed)]
    [InlineData("hwp", "format_paragraphs", OpClass.Allowed)]
    [InlineData("hwp", "set_field_text", OpClass.Allowed)]
    [InlineData("hwp", "insert_picture", OpClass.Allowed)]
    [InlineData("hwp", "export_pdf", OpClass.HighRisk)]
    [InlineData("hwp", "run_external_macro", OpClass.Forbidden)]
    [InlineData("cad", "delete_entities", OpClass.HighRisk)]
    [InlineData("cad", "run_script_template", OpClass.HighRisk)]
    [InlineData("cad", "set_layer_color", OpClass.Allowed)]
    [InlineData("cad", "draw_entities", OpClass.Allowed)]
    [InlineData("cad", "copy_entities", OpClass.Allowed)]
    [InlineData("cad", "set_block_attributes", OpClass.Allowed)]
    [InlineData("cad", "configure_layout", OpClass.Allowed)]
    [InlineData("cad", "save_document", OpClass.HighRisk)]
    [InlineData("cad", "plot_pdf", OpClass.HighRisk)]
    [InlineData("excel", "totally_made_up_op", OpClass.Unknown)]
    public void ClassifyOp_follows_policy(string app, string op, OpClass expected)
        => Assert.Equal(expected, _policy.ClassifyOp(app, op));

    [Fact]
    public void Restore_snapshot_is_high_risk_tool()
        => Assert.True(_policy.IsToolHighRisk("core_restore_snapshot"));

    [Fact]
    public void Apply_tools_are_not_high_risk_tools()
        => Assert.False(_policy.IsToolHighRisk("excel_apply_ops"));
}

public class OperationValidatorTests
{
    private readonly OperationValidator _v = new(new PolicyEngine());

    [Fact]
    public void Valid_batch_passes()
    {
        var batch = Json.ParseObject("""
        { "ops": [ { "op": "set_values", "range": "Sheet1!A1", "values": [["x"]] } ], "dryRun": true }
        """);
        var errors = new List<string>();
        var parsed = _v.Validate(batch, "excel", errors);
        Assert.NotNull(parsed);
        Assert.Empty(errors);
        Assert.True(parsed!.DryRun);
    }

    [Fact]
    public void Copy_sheet_batch_requires_source_workbook_and_sheet()
    {
        var valid = Json.ParseObject("""
        { "ops": [ { "op": "copy_sheet", "sourceWorkbook": "source.xlsx", "sourceSheet": "Sheet1" } ], "dryRun": true }
        """);
        var validErrors = new List<string>();
        Assert.NotNull(_v.Validate(valid, "excel", validErrors));
        Assert.Empty(validErrors);

        var invalid = Json.ParseObject("""
        { "ops": [ { "op": "copy_sheet", "sourceWorkbook": "source.xlsx" } ], "dryRun": true }
        """);
        var invalidErrors = new List<string>();
        Assert.Null(_v.Validate(invalid, "excel", invalidErrors));
        Assert.Contains(invalidErrors, e => e.Contains("sourceSheet"));
    }

    [Fact]
    public void Copy_sheet_cannot_be_mixed_with_cell_edits_in_one_batch()
    {
        var batch = Json.ParseObject("""
        {
          "ops": [
            { "op": "copy_sheet", "sourceWorkbook": "source.xlsx", "sourceSheet": "Sheet1" },
            { "op": "set_values", "range": "Copied!A1", "values": [["x"]] }
          ],
          "dryRun": true
        }
        """);
        var errors = new List<string>();

        Assert.Null(_v.Validate(batch, "excel", errors));
        Assert.Contains(errors, e => e.Contains("cannot be mixed", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Excel_basic_layout_operations_validate_explicit_targets_and_bounds()
    {
        var valid = Json.ParseObject("""
        {
          "ops": [
            { "op": "set_rows_hidden", "target": { "sheet": "내역" }, "row": 2, "count": 3, "hidden": true },
            { "op": "set_cols_hidden", "target": { "sheet": "내역" }, "col": "XFC", "count": 2, "hidden": false },
            { "op": "set_sheet_visibility", "target": { "sheet": "보조" }, "visibility": "hidden" }
          ],
          "dryRun": true
        }
        """);
        var validErrors = new List<string>();
        Assert.NotNull(_v.Validate(valid, "excel", validErrors));
        Assert.Empty(validErrors);

        var invalid = Json.ParseObject("""
        {
          "ops": [
            { "op": "set_rows_hidden", "row": 0, "count": 1, "hidden": true },
            { "op": "set_cols_hidden", "target": { "sheet": "내역" }, "col": "XFD", "count": 2, "hidden": true },
            { "op": "set_sheet_visibility", "target": { "sheet": "보조" }, "visibility": "veryHidden" }
          ],
          "dryRun": true
        }
        """);
        var invalidErrors = new List<string>();
        Assert.Null(_v.Validate(invalid, "excel", invalidErrors));
        Assert.Contains(invalidErrors, error => error.Contains("target.sheet"));
        Assert.Contains(invalidErrors, error => error.Contains("1..1048576"));
        Assert.Contains(invalidErrors, error => error.Contains("A..XFD"));
        Assert.Contains(invalidErrors, error => error.Contains("visible' or 'hidden"));
    }

    [Fact]
    public void Excel_merge_is_range_targeted_and_isolated_in_its_own_batch()
    {
        var valid = Json.ParseObject("""
        { "ops": [ { "op": "merge_cells", "range": "'내 역'!A1:C1" } ], "dryRun": true }
        """);
        var validErrors = new List<string>();
        Assert.NotNull(_v.Validate(valid, "excel", validErrors));
        Assert.Empty(validErrors);

        var mixed = Json.ParseObject("""
        {
          "ops": [
            { "op": "unmerge_cells", "target": { "sheet": "내역" }, "range": "A1:C1" },
            { "op": "format_range", "target": { "sheet": "내역" }, "range": "A1:C1", "style": { "bold": true } }
          ],
          "dryRun": true
        }
        """);
        var mixedErrors = new List<string>();
        Assert.Null(_v.Validate(mixed, "excel", mixedErrors));
        Assert.Contains(mixedErrors, error => error.Contains("only operation", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Replace_document_text_requires_text()
    {
        var valid = Json.ParseObject("""
        { "ops": [ { "op": "replace_document_text", "text": "일일 계획서" } ], "dryRun": true }
        """);
        var validErrors = new List<string>();
        Assert.NotNull(_v.Validate(valid, "hwp", validErrors));
        Assert.Empty(validErrors);

        var invalid = Json.ParseObject("""
        { "ops": [ { "op": "replace_document_text" } ], "dryRun": true }
        """);
        var invalidErrors = new List<string>();
        Assert.Null(_v.Validate(invalid, "hwp", invalidErrors));
        Assert.Contains(invalidErrors, e => e.Contains("text"));
    }

    [Fact]
    public void Append_text_accepts_multiple_paragraphs_and_requires_text()
    {
        var valid = Json.ParseObject("""
        { "ops": [ { "op": "append_text", "text": "첫째 문단\n둘째 문단" } ], "dryRun": true }
        """);
        var validErrors = new List<string>();
        Assert.NotNull(_v.Validate(valid, "hwp", validErrors));
        Assert.Empty(validErrors);

        var invalid = Json.ParseObject("""
        { "ops": [ { "op": "append_text" } ], "dryRun": true }
        """);
        var invalidErrors = new List<string>();
        Assert.Null(_v.Validate(invalid, "hwp", invalidErrors));
        Assert.Contains(invalidErrors, e => e.Contains("text"));
    }

    [Theory]
    [InlineData("insert_before_text")]
    [InlineData("insert_after_text")]
    public void Relative_insert_requires_anchor_and_text(string opName)
    {
        var valid = Json.ParseObject($$"""
        { "ops": [ { "op": "{{opName}}", "anchor": "기준 문단", "text": "추가 문단", "mode": "paragraph" } ], "dryRun": true }
        """);
        var validErrors = new List<string>();
        Assert.NotNull(_v.Validate(valid, "hwp", validErrors));
        Assert.Empty(validErrors);

        var invalid = Json.ParseObject($$"""
        { "ops": [ { "op": "{{opName}}", "text": "추가 문단" } ], "dryRun": true }
        """);
        var invalidErrors = new List<string>();
        Assert.Null(_v.Validate(invalid, "hwp", invalidErrors));
        Assert.Contains(invalidErrors, e => e.Contains("anchor"));
    }

    [Fact]
    public void Forbidden_op_fails_even_in_dryrun()
    {
        var batch = Json.ParseObject("""
        { "ops": [ { "op": "run_macro", "name": "Evil" } ], "dryRun": true }
        """);
        var errors = new List<string>();
        Assert.Null(_v.Validate(batch, "excel", errors));
        Assert.Contains(errors, e => e.Contains("FORBIDDEN"));
    }

    [Fact]
    public void Missing_required_field_fails()
    {
        var batch = Json.ParseObject("""
        { "ops": [ { "op": "set_values", "range": "A1" } ], "dryRun": true }
        """);
        var errors = new List<string>();
        Assert.Null(_v.Validate(batch, "excel", errors));
        Assert.Contains(errors, e => e.Contains("values"));
    }

    [Fact]
    public void Empty_ops_fails()
    {
        var batch = Json.ParseObject("""{ "ops": [], "dryRun": true }""");
        var errors = new List<string>();
        Assert.Null(_v.Validate(batch, "excel", errors));
    }

    [Fact]
    public void Highrisk_op_is_flagged()
    {
        var batch = Json.ParseObject("""
        { "ops": [ { "op": "delete_entities", "handles": ["A", "B"] } ], "dryRun": true }
        """);
        var errors = new List<string>();
        var parsed = _v.Validate(batch, "cad", errors);
        Assert.NotNull(parsed);
        Assert.True(parsed!.HasHighRiskOps);
    }

    [Fact]
    public void Hwp_production_operations_are_validated()
    {
        var valid = Json.ParseObject("""
        {
          "ops": [
            { "op": "set_paragraph_format", "style": { "align": "center", "lineSpacingPercent": 160 } },
            { "op": "set_page_setup", "page": { "widthMm": 210, "heightMm": 297 } },
            { "op": "table_cell_set_text", "tableIndex": 0, "row": 1, "col": 2, "text": "완료" },
            { "op": "table_cell_set_text", "tableIndex": 0, "cellIndex": 5, "text": "병합표", "preserveStyle": true },
            { "op": "table_insert_rows", "tableIndex": 0, "row": 1, "col": 0, "count": 1 },
            { "op": "table_merge_cells", "tableIndex": 0, "startRow": 0, "startCol": 0, "endRow": 0, "endCol": 1 },
            { "op": "table_set_row_height", "tableIndex": 0, "row": 1, "heightMm": 10.5 },
            { "op": "table_set_row_heights", "tableIndex": 0, "rows": [{ "row": 0, "heightMm": 8 }, { "row": 1, "heightMm": 9 }] },
            { "op": "format_paragraphs", "items": [{ "target": { "text": "제목" }, "characterStyle": { "bold": true }, "paragraphStyle": { "align": "center" } }] },
            { "op": "set_field_text", "name": "담당자", "text": "홍길동" },
            { "op": "insert_picture", "path": "C:\\temp\\photo.png" },
            { "op": "set_header_footer_text", "kind": "footer", "text": "현장명" }
          ],
          "dryRun": true
        }
        """);
        var errors = new List<string>();
        var parsed = _v.Validate(valid, "hwp", errors);
        Assert.NotNull(parsed);
        Assert.Empty(errors);

        var invalid = Json.ParseObject("""{ "ops": [{ "op": "set_page_setup" }] }""");
        var invalidErrors = new List<string>();
        Assert.Null(_v.Validate(invalid, "hwp", invalidErrors));
        Assert.Contains(invalidErrors, e => e.Contains("page"));
    }

    [Fact]
    public void Activate_document_is_allowed_for_cad_and_requires_selector()
    {
        var valid = Json.ParseObject("""
        { "ops": [ { "op": "activate_document", "document": "target*.dwg" } ], "dryRun": true }
        """);
        var validErrors = new List<string>();
        Assert.NotNull(_v.Validate(valid, "cad", validErrors));
        Assert.Empty(validErrors);

        var invalid = Json.ParseObject("""
        { "ops": [ { "op": "activate_document" } ], "dryRun": true }
        """);
        var invalidErrors = new List<string>();
        Assert.Null(_v.Validate(invalid, "cad", invalidErrors));
        Assert.Contains(invalidErrors, e => e.Contains("document"));
    }

    [Fact]
    public void Cad_production_operations_are_validated()
    {
        var valid = Json.ParseObject("""
        { "ops": [
          { "op": "copy_entities", "handles": ["AB"], "dx": 10, "dy": 20 },
          { "op": "scale_entities", "handles": ["AB"], "basePoint": [0,0], "factor": 2 },
          { "op": "set_entity_properties", "handles": ["AB"], "properties": { "layer": "PLAN" } },
          { "op": "configure_layout", "name": "A1" },
          { "op": "create_viewport", "layout": "A1", "center": [100,70], "width": 160, "height": 100, "viewHeight": 50 }
        ], "dryRun": true }
        """);
        var errors = new List<string>();
        Assert.NotNull(_v.Validate(valid, "cad", errors));
        Assert.Empty(errors);

        var invalid = Json.ParseObject("""{ "ops": [{ "op": "plot_pdf" }] }""");
        var invalidErrors = new List<string>();
        Assert.Null(_v.Validate(invalid, "cad", invalidErrors));
        Assert.Contains(invalidErrors, e => e.Contains("output"));
    }

    [Fact]
    public void Draw_entities_requires_entities_array()
    {
        var batch = Json.ParseObject("""
        { "ops": [ { "op": "draw_entities" } ], "dryRun": true }
        """);
        var errors = new List<string>();
        Assert.Null(_v.Validate(batch, "cad", errors));
        Assert.Contains(errors, e => e.Contains("entities"));
    }

    [Fact]
    public void Draw_entities_valid_batch_passes()
    {
        var batch = Json.ParseObject("""
        {
          "ops": [ { "op": "draw_entities", "entities": [
            { "type": "lwpolyline", "points": [[0,0],[10,0],[10,10]], "closed": true },
            { "type": "circle", "center": [5,5], "radius": 2.5, "color": { "rgb": [255,0,0] } },
            { "type": "hatch", "loop": { "points": [[0,0],[10,0],[5,10]], "bulges": [0,0,0] }, "color": { "aci": 1 } }
          ] } ],
          "dryRun": true
        }
        """);
        var errors = new List<string>();
        var parsed = _v.Validate(batch, "cad", errors);
        Assert.NotNull(parsed);
        Assert.Empty(errors);
        Assert.False(parsed!.HasHighRiskOps);
    }
}

public class ConfirmTokenServiceTests : IDisposable
{
    private readonly TestHome _home = new();
    public void Dispose() => _home.Dispose();

    [Fact]
    public void Validate_without_consume_keeps_token_available()
    {
        var svc = new ConfirmTokenService(_home.Options, ttlSeconds: 300);
        var (token, _) = svc.Create("apply:hwp", "ops-hash", "snapshot-1");

        var first = svc.Validate(token, "apply:hwp", "ops-hash");
        var second = svc.ValidateAndConsume(token, "apply:hwp", "ops-hash");
        var third = svc.ValidateAndConsume(token, "apply:hwp", "ops-hash");

        Assert.True(first.Ok);
        Assert.True(second.Ok);
        Assert.False(third.Ok);
        Assert.Equal("snapshot-1", first.SnapshotId);
    }

    [Fact]
    public void Create_and_consume_roundtrip()
    {
        var svc = new ConfirmTokenService(_home.Options, ttlSeconds: 300);
        var (token, expires) = svc.Create("apply:fake", "hash123", "snap1");
        Assert.StartsWith("conf_", token);
        Assert.Equal(300, expires);

        var check = svc.ValidateAndConsume(token, "apply:fake", "hash123");
        Assert.True(check.Ok);
        Assert.Equal("snap1", check.SnapshotId);
    }

    [Fact]
    public void Token_is_single_use()
    {
        var svc = new ConfirmTokenService(_home.Options, ttlSeconds: 300);
        var (token, _) = svc.Create("apply:fake", "hash123", null);
        Assert.True(svc.ValidateAndConsume(token, "apply:fake", "hash123").Ok);
        var second = svc.ValidateAndConsume(token, "apply:fake", "hash123");
        Assert.False(second.Ok);
        Assert.Contains("single-use", second.Error);
    }

    [Fact]
    public void Token_rejects_ops_changed_after_dryrun()
    {
        var svc = new ConfirmTokenService(_home.Options, ttlSeconds: 300);
        var (token, _) = svc.Create("apply:fake", "hashAAA", null);
        var check = svc.ValidateAndConsume(token, "apply:fake", "hashBBB");
        Assert.False(check.Ok);
        Assert.Contains("ops changed", check.Error);
    }

    [Fact]
    public void Expired_token_fails()
    {
        var svc = new ConfirmTokenService(_home.Options, ttlSeconds: 1);
        var (token, _) = svc.Create("apply:fake", "hash123", null);
        Thread.Sleep(1500);
        var check = svc.ValidateAndConsume(token, "apply:fake", "hash123");
        Assert.False(check.Ok);
        Assert.Contains("expired", check.Error);
    }

    [Fact]
    public void Tampered_signature_fails()
    {
        var svc = new ConfirmTokenService(_home.Options, ttlSeconds: 300);
        var (token, _) = svc.Create("apply:fake", "hash123", null);
        var tampered = token[..^2] + (token[^2] == 'A' ? "BA" : "AA");
        var check = svc.ValidateAndConsume(tampered, "apply:fake", "hash123");
        Assert.False(check.Ok);
    }

    [Fact]
    public void Scope_mismatch_fails()
    {
        var svc = new ConfirmTokenService(_home.Options, ttlSeconds: 300);
        var (token, _) = svc.Create("apply:excel", "hash123", null);
        var check = svc.ValidateAndConsume(token, "apply:hwp", "hash123");
        Assert.False(check.Ok);
        Assert.Contains("scope", check.Error);
    }
}

public class SnapshotServiceTests : IDisposable
{
    private readonly TestHome _home = new();
    public void Dispose() => _home.Dispose();

    [Fact]
    public void Create_list_get_roundtrip()
    {
        var svc = new SnapshotService(_home.Options);
        var info = svc.Create("fake", "unit-test", "fake://doc",
            (dir, meta) => File.WriteAllText(Path.Combine(dir, "state.json"), "{}"));

        Assert.True(Directory.Exists(info.Dir));
        Assert.True(File.Exists(Path.Combine(info.Dir, "metadata.json")));

        var list = svc.List("fake", 10);
        Assert.Contains(list, s => s.SnapshotId == info.SnapshotId);

        var got = svc.Get(info.SnapshotId);
        Assert.NotNull(got);
        Assert.Equal("fake", got!.Value.Info.App);
    }

    [Fact]
    public void Unknown_snapshot_returns_null()
    {
        var svc = new SnapshotService(_home.Options);
        Assert.Null(svc.Get("no-such-snapshot"));
    }
}
