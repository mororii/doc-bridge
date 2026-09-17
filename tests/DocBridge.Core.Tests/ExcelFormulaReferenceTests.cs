using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelFormulaReferenceTests
{
    [Fact]
    public void Quoted_and_unquoted_Korean_sheet_same_relativity_match()
    {
        Assert.True(ExcelFormulaReference.SemanticEquals("='기성입력'!B2", "=기성입력!B2"));
        Assert.True(ExcelFormulaReference.SemanticEquals("='기성입력'!$B$2", "=기성입력!$B$2"));
        Assert.Equal("=기성입력!b2", ExcelFormulaReference.Normalize("='기성입력'!B2"));
        Assert.Equal("=기성입력!$b$2", ExcelFormulaReference.Normalize("='기성입력'!$B$2"));
        Assert.False(ExcelFormulaReference.NeedsSheetQuotes("기성입력"));
    }

    [Fact]
    public void Relative_and_absolute_A1_are_not_equal_on_name_readback()
    {
        Assert.False(ExcelFormulaReference.SemanticEquals("='기성입력'!$B$2", "=기성입력!B2"));
        Assert.False(ExcelFormulaReference.SemanticEquals("='자재대장'!$G$2:$G$8", "=자재대장!G2:G8"));
        Assert.False(ExcelFormulaReference.SemanticEquals("=$A$1", "=A1"));
        Assert.False(ExcelFormulaReference.SemanticEquals("=$A1", "=A$1"));
        Assert.False(ExcelFormulaReference.SemanticEquals("=$A1", "=A1"));
        Assert.True(ExcelFormulaReference.HasRelativeCellRef("=월별기성!B2"));
        Assert.False(ExcelFormulaReference.HasRelativeCellRef("='월별기성'!$B$2"));
        Assert.Equal("='월별기성'!$B$2", ExcelFormulaReference.AbsoluteQuotedRefersTo("월별기성", "B2"));
        Assert.Equal("='월별기성'!$B$5", ExcelFormulaReference.AbsoluteQuotedRefersTo("월별기성", "B5"));
    }

    [Fact]
    public void Sheet_identifier_case_folds_but_string_literals_do_not()
    {
        Assert.True(ExcelFormulaReference.SemanticEquals("=Sheet1!A1", "=sheet1!A1"));
        Assert.True(ExcelFormulaReference.SemanticEquals("=$B$2", "=$b$2"));
        Assert.False(ExcelFormulaReference.SemanticEquals("=\"Hello\"", "=\"hello\""));
        Assert.False(ExcelFormulaReference.SemanticEquals("=\"IT'S\"", "=\"it's\""));
        Assert.Equal("=\"Hello\"", ExcelFormulaReference.Normalize("=\"Hello\""));
        Assert.Equal("=\"$100\"", ExcelFormulaReference.Normalize("=\"$100\""));
        Assert.False(ExcelFormulaReference.SemanticEquals("=\"$100\"", "=\"100\""));
        Assert.False(ExcelFormulaReference.SemanticEquals("=\"기성입력\"", "=기성입력!B2"));
        Assert.False(ExcelFormulaReference.SemanticEquals("=\"'기성입력'!B2\"", "=기성입력!B2"));
    }

    [Fact]
    public void Sheet_with_space_keeps_quotes_and_does_not_match_broken_unquoted()
    {
        Assert.Equal("='sheet 1'!a1", ExcelFormulaReference.Normalize("='Sheet 1'!A1"));
        Assert.False(ExcelFormulaReference.SemanticEquals("='Sheet 1'!A1", "=Sheet 1!A1"));
        Assert.True(ExcelFormulaReference.NeedsSheetQuotes("Sheet 1"));
    }

    [Fact]
    public void Workbook_Korean_names_stay_workbook_scope()
    {
        Assert.True(ExcelFormulaReference.TryParseDefinedName("ContractAmount", out var scope, out var sheet, out var local));
        Assert.Equal("workbook", scope);
        Assert.Null(sheet);
        Assert.Equal("ContractAmount", local);

        Assert.True(ExcelFormulaReference.TryParseDefinedName("'기성입력'!ContractAmount", out scope, out sheet, out local));
        Assert.Equal("sheet", scope);
        Assert.Equal("기성입력", sheet);
        Assert.Equal("ContractAmount", local);

        Assert.True(ExcelFormulaReference.TryParseDefinedName("기성입력!JanAmt", out scope, out sheet, out local));
        Assert.Equal("sheet", scope);
        Assert.Equal("기성입력", sheet);
        Assert.Equal("JanAmt", local);
    }

    [Fact]
    public void E5_saved_relative_names_are_not_labeled_stale_cache_only()
    {
        var names = new[]
        {
            new ExcelDefinedNameDependencies.NameBinding("ContractAmount", "workbook", null, "=월별기성!B2", 100_000_000),
            new ExcelDefinedNameDependencies.NameBinding("JanAmt", "workbook", null, "=월별기성!B3", 12_000_000),
            new ExcelDefinedNameDependencies.NameBinding("FebAmt", "workbook", null, "=월별기성!B4", 20_000_000),
            new ExcelDefinedNameDependencies.NameBinding("MarAmt", "workbook", null, "=월별기성!B5", 15_000_000),
        };
        Assert.Equal(new[] { "JanAmt", "FebAmt", "MarAmt" },
            ExcelDefinedNameDependencies.EnumerateNameTokens("JanAmt+FebAmt+MarAmt"));
        Assert.Empty(ExcelDefinedNameDependencies.ExplainMissingWorkbookBindings("=JanAmt+FebAmt+MarAmt", names));
        Assert.True(ExcelDefinedNameDependencies.SourceRefersToExpected(names[0], "='월별기성'!B2"));
        Assert.False(ExcelDefinedNameDependencies.SourceRefersToExpected(names[0], "='월별기성'!$B$2"));
        Assert.Contains(
            ExcelDefinedNameDependencies.ExplainRelativeRefersTo(names),
            text => text.Contains("selection-dependent") && text.Contains("$B$2"));
        Assert.Empty(ExcelDefinedNameDependencies.ExplainStaleCachedZero("=JanAmt+FebAmt+MarAmt", 0, names));

        var absolute = names.Select(item => item with
        {
            RefersTo = ExcelFormulaReference.AbsoluteQuotedRefersTo("월별기성", item.RefersTo!.Split('!')[^1]),
        }).ToArray();
        Assert.Contains(
            ExcelDefinedNameDependencies.ExplainStaleCachedZero("=JanAmt+FebAmt+MarAmt", 0, absolute),
            text => text.Contains("calculate"));

        var local = names.Select(item => item with { Scope = "sheet", Sheet = "월별기성" }).ToArray();
        Assert.Contains(
            ExcelDefinedNameDependencies.ExplainMissingWorkbookBindings("=ContractAmount-B2", local),
            text => text.Contains("sheet-local"));
    }

    [Fact]
    public void Inspect_objects_must_clone_parented_note_nodes()
    {
        var captured = new JsonArray
        {
            new JsonObject { ["range"] = "B2", ["text"] = "note" },
        };
        var note = Assert.IsType<JsonObject>(captured[0]);
        Assert.NotNull(note.Parent);
        Assert.Throws<InvalidOperationException>(() =>
        {
            var other = new JsonArray();
            other.Add(note);
        });

        var detached = ExcelJsonOwnership.DetachObject(note);
        var items = new JsonArray();
        ExcelJsonOwnership.AddDetached(items, note);
        detached["type"] = "note";
        Assert.Equal("B2", items[0]!["range"]!.GetValue<string>());
        Assert.Equal("note", Assert.Single(captured)!["text"]!.GetValue<string>());
    }

    [Fact]
    public void References_sheet_skips_literals_and_backup_links()
    {
        Assert.True(ExcelFormulaReference.ReferencesSheet("=Data!A1", "Data"));
        Assert.True(ExcelFormulaReference.ReferencesSheet("='Data'!$A$1", "Data"));
        Assert.False(ExcelFormulaReference.ReferencesSheet("=Data10!A1", "Data"));
        Assert.False(ExcelFormulaReference.ReferencesSheet("=\"Data\"", "Data"));
        Assert.False(ExcelFormulaReference.ReferencesSheet("=[backup.xlsx]Data!A1", "Data"));
        Assert.True(ExcelFormulaReference.SemanticEquals("=$A$1", "=$A$1"));
        Assert.False(ExcelFormulaReference.SemanticEquals("=$A$1", "=A1"));
        Assert.False(ExcelFormulaReference.SemanticEquals("=\"Hello\"", "=\"hello\""));
    }
}
