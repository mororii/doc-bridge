using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelDataOperationsRestoreTests
{
    [Fact]
    public void Table_create_rollback_is_unlist_not_delete()
    {
        Assert.Equal("Unlist", ExcelDataSnapshotEquality.TableCreateRollbackMethod);
        Assert.Equal("Unlist", Json.GetString(ExcelDataOperationsContract.DescribeSchema(), "tableCreateRollback"));
    }

    [Fact]
    public void Verified_is_false_when_equality_is_not_run()
    {
        var result = ExcelDataSnapshotEquality.Evaluate(CompleteSortSnapshot(), liveActualsByEntryIndex: null);
        Assert.False(Json.GetBool(result, "verified"));
        Assert.Contains(MismatchTexts(result), text => text.Contains("equality comparison was not performed"));
    }

    [Fact]
    public void Version_1_snapshot_cannot_be_verified()
    {
        var state = CompleteSortSnapshot();
        state["snapshotVersion"] = 1;
        var result = ExcelDataSnapshotEquality.Evaluate(state, MatchingSortActuals());
        Assert.False(Json.GetBool(result, "verified"));
        Assert.Contains(MismatchTexts(result), text => text.Contains("snapshotVersion 1"));
    }

    [Fact]
    public void Incomplete_sort_without_values_cannot_be_verified()
    {
        var entry = new JsonObject
        {
            ["kind"] = "sortRange",
            ["sheet"] = "자재대장",
            ["range"] = "A1:B2",
            ["formulas"] = Grid("=1", "=2"),
        };
        Assert.Contains(ExcelDataSnapshotEquality.ExplainIncomplete(entry), text => text.Contains("values"));
        var result = ExcelDataSnapshotEquality.Evaluate(Wrap(entry), new JsonArray { entry.DeepClone() });
        Assert.False(Json.GetBool(result, "verified"));
    }

    [Fact]
    public void Sort_value_mismatch_forces_verified_false()
    {
        var expected = CompleteSortEntry();
        var actual = CompleteSortEntry();
        actual["values"] = Grid("changed", "B");
        var result = ExcelDataSnapshotEquality.Evaluate(Wrap(expected), new JsonArray { actual });
        Assert.False(Json.GetBool(result, "verified"));
        Assert.Contains(MismatchTexts(result), text => text.Contains("values"));
    }

    [Fact]
    public void Matching_complete_sort_surface_can_verify()
    {
        var result = ExcelDataSnapshotEquality.Evaluate(CompleteSortSnapshot(), MatchingSortActuals());
        Assert.True(Json.GetBool(result, "verified"), string.Join("; ", MismatchTexts(result)));
    }

    [Fact]
    public void Filter_without_criteria_array_is_incomplete()
    {
        var entry = new JsonObject
        {
            ["kind"] = "autoFilter",
            ["sheet"] = "자재대장",
            ["state"] = new JsonObject { ["enabled"] = true, ["filtered"] = true },
        };
        Assert.Contains(ExcelDataSnapshotEquality.ExplainIncomplete(entry), text => text.Contains("criteria"));
    }

    [Fact]
    public void Filter_criteria_mismatch_forces_verified_false()
    {
        var expected = FilterEntry("입고");
        var actual = new JsonObject
        {
            ["kind"] = "autoFilter",
            ["state"] = FilterState("출고"),
        };
        var result = ExcelDataSnapshotEquality.Evaluate(Wrap(expected), new JsonArray { actual });
        Assert.False(Json.GetBool(result, "verified"));
        Assert.Contains(MismatchTexts(result), text => text.Contains("criteria1"));
    }

    [Fact]
    public void Deleted_name_without_refersTo_cannot_be_recreated()
    {
        var entry = new JsonObject
        {
            ["kind"] = "definedName",
            ["existed"] = true,
            ["name"] = "StockOnHand",
            ["scope"] = "sheet",
            ["state"] = new JsonObject { ["name"] = "StockOnHand" },
        };
        Assert.Contains(ExcelDataSnapshotEquality.ExplainIncomplete(entry), text => text.Contains("refersTo"));
    }

    [Fact]
    public void Quoted_and_unquoted_Korean_RefersTo_restore_as_equal()
    {
        var expected = new JsonObject
        {
            ["kind"] = "definedName",
            ["existed"] = true,
            ["name"] = "ContractAmount",
            ["scope"] = "workbook",
            ["state"] = new JsonObject { ["name"] = "ContractAmount", ["refersTo"] = "='기성입력'!B2" },
        };
        var actual = new JsonObject
        {
            ["kind"] = "definedName",
            ["present"] = true,
            ["state"] = new JsonObject { ["name"] = "ContractAmount", ["refersTo"] = "=기성입력!B2" },
        };
        var result = ExcelDataSnapshotEquality.Evaluate(Wrap(expected), new JsonArray { actual });
        Assert.True(Json.GetBool(result, "verified"));
    }

    [Fact]
    public void Absolute_and_relative_RefersTo_restore_as_unequal()
    {
        var expected = new JsonObject
        {
            ["kind"] = "definedName",
            ["existed"] = true,
            ["name"] = "ContractAmount",
            ["scope"] = "workbook",
            ["state"] = new JsonObject { ["name"] = "ContractAmount", ["refersTo"] = "='월별기성'!$B$2" },
        };
        var actual = new JsonObject
        {
            ["kind"] = "definedName",
            ["present"] = true,
            ["state"] = new JsonObject { ["name"] = "ContractAmount", ["refersTo"] = "=월별기성!B2" },
        };
        var result = ExcelDataSnapshotEquality.Evaluate(Wrap(expected), new JsonArray { actual });
        Assert.False(Json.GetBool(result, "verified"));
        Assert.Contains(MismatchTexts(result), text => text.Contains("RefersTo"));
    }

    [Fact]
    public void Name_missing_after_restore_is_not_verified()
    {
        var expected = new JsonObject
        {
            ["kind"] = "definedName",
            ["existed"] = true,
            ["name"] = "StockOnHand",
            ["scope"] = "sheet",
            ["state"] = new JsonObject { ["name"] = "StockOnHand", ["refersTo"] = "='자재대장'!$G$2:$G$8" },
        };
        var actual = new JsonObject { ["kind"] = "definedName", ["present"] = false };
        var result = ExcelDataSnapshotEquality.Evaluate(Wrap(expected), new JsonArray { actual });
        Assert.False(Json.GetBool(result, "verified"));
        Assert.Contains(MismatchTexts(result), text => text.Contains("not recreated"));
    }

    [Fact]
    public void Validation_missing_dropdown_and_alerts_is_incomplete()
    {
        var entry = new JsonObject
        {
            ["kind"] = "validation",
            ["range"] = "C2:C8",
            ["state"] = new JsonObject
            {
                ["present"] = true,
                ["validationType"] = 3,
                ["formula1"] = "입고,출고,이동",
            },
        };
        var incomplete = ExcelDataSnapshotEquality.ExplainIncomplete(entry);
        Assert.Contains(incomplete, text => text.Contains("inCellDropdown"));
        Assert.Contains(incomplete, text => text.Contains("showError"));
    }

    [Fact]
    public void Unreadable_conditional_rule_cannot_be_verified()
    {
        var entry = new JsonObject
        {
            ["kind"] = "conditionalFormat",
            ["range"] = "G2:G8",
            ["state"] = new JsonObject
            {
                ["count"] = 1,
                ["rules"] = new JsonArray { new JsonObject { ["index"] = 1, ["unreadable"] = true } },
            },
        };
        Assert.Contains(ExcelDataSnapshotEquality.ExplainIncomplete(entry), text => text.Contains("unreadable"));
    }

    [Fact]
    public void Chart_without_series_is_incomplete()
    {
        var entry = new JsonObject
        {
            ["kind"] = "chart",
            ["existed"] = true,
            ["name"] = "MonthlyOutput",
            ["state"] = new JsonObject { ["chartType"] = "columnClustered", ["title"] = "월별" },
        };
        Assert.Contains(ExcelDataSnapshotEquality.ExplainIncomplete(entry), text => text.Contains("series"));
    }

    [Fact]
    public void Chart_series_mismatch_forces_verified_false()
    {
        var expected = ChartEntry("=SERIES(월간보고!$B$1,,월간보고!$B$2:$B$6,1)");
        var actual = new JsonObject
        {
            ["kind"] = "chart",
            ["present"] = true,
            ["state"] = ChartState("=SERIES(월간보고!$C$1,,월간보고!$C$2:$C$6,1)"),
        };
        var result = ExcelDataSnapshotEquality.Evaluate(Wrap(expected), new JsonArray { actual });
        Assert.False(Json.GetBool(result, "verified"));
        Assert.Contains(MismatchTexts(result), text => text.Contains("formula"));
    }

    [Fact]
    public void Combo_series_type_mismatch_forces_verified_false()
    {
        var expected = new JsonObject
        {
            ["kind"] = "chart",
            ["existed"] = true,
            ["name"] = "계획대비실적",
            ["state"] = new JsonObject
            {
                ["combination"] = true,
                ["chartType"] = "columnClustered",
                ["seriesCount"] = 2,
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["index"] = 1,
                        ["formula"] = "=SERIES(\"금액\",A2:A7,B2:B7,1)",
                        ["chartType"] = "columnClustered",
                        ["axisGroup"] = "primary",
                    },
                    new JsonObject
                    {
                        ["index"] = 2,
                        ["formula"] = "=SERIES(\"진척률\",A2:A7,C2:C7,2)",
                        ["chartType"] = "lineMarkers",
                        ["axisGroup"] = "secondary",
                    },
                },
            },
        };
        var actual = new JsonObject
        {
            ["kind"] = "chart",
            ["present"] = true,
            ["state"] = new JsonObject
            {
                ["combination"] = true,
                ["chartType"] = "raw:73",
                ["seriesCount"] = 2,
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["index"] = 1,
                        ["formula"] = "=SERIES(\"금액\",A2:A7,B2:B7,1)",
                        ["chartType"] = "columnClustered",
                        ["axisGroup"] = "primary",
                    },
                    new JsonObject
                    {
                        ["index"] = 2,
                        ["formula"] = "=SERIES(\"진척률\",A2:A7,C2:C7,2)",
                        ["chartType"] = "line",
                        ["axisGroup"] = "primary",
                    },
                },
            },
        };
        var mismatch = ExcelDataSnapshotEquality.Evaluate(Wrap(expected), new JsonArray { actual });
        Assert.False(Json.GetBool(mismatch, "verified"));
        Assert.Contains(MismatchTexts(mismatch), text => text.Contains("chartType"));
        Assert.Contains(MismatchTexts(mismatch), text => text.Contains("axisGroup"));

        var restored = (JsonObject)actual.DeepClone();
        var series = Json.GetArr(Json.GetObj(restored, "state"), "series")!;
        ((JsonObject)series[1]!)["chartType"] = "lineMarkers";
        ((JsonObject)series[1]!)["axisGroup"] = "secondary";
        var ok = ExcelDataSnapshotEquality.Evaluate(Wrap(expected), new JsonArray { restored });
        Assert.True(Json.GetBool(ok, "verified"));
    }

    [Fact]
    public void Combo_axes_mismatch_forces_verified_false()
    {
        var expected = new JsonObject
        {
            ["kind"] = "chart",
            ["existed"] = true,
            ["name"] = "계획대비실적",
            ["state"] = new JsonObject
            {
                ["combination"] = true,
                ["chartType"] = "columnClustered",
                ["seriesCount"] = 2,
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["index"] = 1,
                        ["formula"] = "=SERIES(\"금액\",A2:A7,B2:B7,1)",
                        ["chartType"] = "columnClustered",
                    },
                    new JsonObject
                    {
                        ["index"] = 2,
                        ["formula"] = "=SERIES(\"진척률\",A2:A7,C2:C7,2)",
                        ["chartType"] = "lineMarkers",
                    },
                },
                ["axes"] = new JsonObject
                {
                    ["value"] = new JsonObject { ["numberFormat"] = "#,##0", ["title"] = "금액" },
                    ["valueSecondary"] = new JsonObject
                    {
                        ["minimum"] = 0,
                        ["maximum"] = 100,
                        ["numberFormat"] = "0\\%",
                        ["title"] = "진척률",
                    },
                },
            },
        };
        static JsonObject Live(bool unreadable) => new()
        {
            ["kind"] = "chart",
            ["present"] = true,
            ["state"] = new JsonObject
            {
                ["combination"] = true,
                ["chartType"] = "columnClustered",
                ["seriesCount"] = 2,
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["index"] = 1,
                        ["formula"] = "=SERIES(\"금액\",A2:A7,B2:B7,1)",
                        ["chartType"] = "columnClustered",
                    },
                    new JsonObject
                    {
                        ["index"] = 2,
                        ["formula"] = "=SERIES(\"진척률\",A2:A7,C2:C7,2)",
                        ["chartType"] = "lineMarkers",
                    },
                },
                ["axes"] = new JsonObject
                {
                    ["value"] = new JsonObject { ["numberFormat"] = "#,##0", ["title"] = "금액" },
                    ["valueSecondary"] = unreadable
                        ? new JsonObject { ["unreadable"] = true }
                        : new JsonObject
                        {
                            ["minimum"] = 0,
                            ["maximum"] = 100,
                            ["numberFormat"] = "0\\%",
                            ["title"] = "진척률",
                        },
                },
            },
        };
        var mismatch = ExcelDataSnapshotEquality.Evaluate(Wrap(expected), new JsonArray { Live(true) });
        Assert.False(Json.GetBool(mismatch, "verified"));
        Assert.Contains(MismatchTexts(mismatch), text => text.Contains("axes.", StringComparison.Ordinal));
        var ok = ExcelDataSnapshotEquality.Evaluate(Wrap(expected), new JsonArray { Live(false) });
        Assert.True(Json.GetBool(ok, "verified"));
    }

    [Fact]
    public void Hyperlink_without_cell_content_or_font_is_incomplete()
    {
        var entry = new JsonObject
        {
            ["kind"] = "hyperlink",
            ["range"] = "D2",
            ["state"] = new JsonObject
            {
                ["present"] = true,
                ["address"] = "https://example.invalid",
            },
        };
        var incomplete = ExcelDataSnapshotEquality.ExplainIncomplete(entry);
        Assert.Contains(incomplete, text => text.Contains("cellFormula") || text.Contains("cellValue"));
        Assert.Contains(incomplete, text => text.Contains("font"));
    }

    [Fact]
    public void Hyperlink_font_mismatch_forces_verified_false()
    {
        var expected = HyperlinkEntry();
        var actual = HyperlinkEntry();
        Json.GetObj(actual, "state")!["font"] = new JsonObject { ["bold"] = true, ["color"] = 255 };
        var result = ExcelDataSnapshotEquality.Evaluate(Wrap(expected), new JsonArray { actual });
        Assert.False(Json.GetBool(result, "verified"));
        Assert.Contains(MismatchTexts(result), text => text.Contains("font"));
    }

    [Fact]
    public void Created_table_still_present_after_unlist_is_not_verified()
    {
        var expected = new JsonObject
        {
            ["kind"] = "table",
            ["existed"] = false,
            ["name"] = "Materials",
            ["range"] = "A1:B2",
            ["values"] = Grid("H1", "H2"),
            ["formulas"] = Grid("H1", "H2"),
        };
        var actual = new JsonObject
        {
            ["kind"] = "table",
            ["present"] = true,
            ["name"] = "Materials",
            ["values"] = Grid("H1", "H2"),
            ["formulas"] = Grid("H1", "H2"),
        };
        var result = ExcelDataSnapshotEquality.Evaluate(Wrap(expected), new JsonArray { actual });
        Assert.False(Json.GetBool(result, "verified"));
        Assert.Contains(MismatchTexts(result), text => text.Contains("still present"));
    }

    [Fact]
    public void Invalid_restore_mode_cannot_be_verified()
    {
        var state = CompleteSortSnapshot();
        state["restoreMode"] = "values";
        var result = ExcelDataSnapshotEquality.Evaluate(state, MatchingSortActuals());
        Assert.False(Json.GetBool(result, "verified"));
    }

    private static JsonObject CompleteSortSnapshot() => Wrap(CompleteSortEntry());

    private static JsonArray MatchingSortActuals() => new() { CompleteSortEntry() };

    private static JsonObject CompleteSortEntry() => new()
    {
        ["kind"] = "sortRange",
        ["sheet"] = "자재대장",
        ["range"] = "A1:B2",
        ["formulas"] = Grid("A", "B"),
        ["values"] = Grid("A", "B"),
        ["numberFormats"] = Grid("General", "General"),
        ["notes"] = new JsonArray(),
        ["hyperlinks"] = new JsonArray(),
    };

    private static JsonObject FilterEntry(string value) => new()
    {
        ["kind"] = "autoFilter",
        ["sheet"] = "자재대장",
        ["state"] = FilterState(value),
    };

    private static JsonObject FilterState(string value) => new()
    {
        ["enabled"] = true,
        ["filtered"] = true,
        ["filters"] = new JsonArray
        {
            new JsonObject { ["field"] = 1, ["on"] = true, ["criteria1"] = value },
        },
    };

    private static JsonObject ChartEntry(string formula) => new()
    {
        ["kind"] = "chart",
        ["existed"] = true,
        ["name"] = "MonthlyOutput",
        ["state"] = ChartState(formula),
    };

    private static JsonObject ChartState(string formula) => new()
    {
        ["chartType"] = "columnClustered",
        ["seriesCount"] = 1,
        ["series1Formula"] = formula,
        ["series"] = new JsonArray { new JsonObject { ["index"] = 1, ["formula"] = formula } },
    };

    private static JsonObject HyperlinkEntry() => new()
    {
        ["kind"] = "hyperlink",
        ["range"] = "D2",
        ["state"] = new JsonObject
        {
            ["present"] = true,
            ["address"] = "https://example.invalid",
            ["cellFormula"] = "관부속",
            ["cellValue"] = "관부속",
            ["font"] = new JsonObject { ["bold"] = false, ["color"] = 16711680 },
        },
    };

    private static JsonObject Wrap(JsonObject entry) => new()
    {
        ["snapshotVersion"] = 2,
        ["restoreMode"] = "data-objects",
        ["entries"] = new JsonArray { entry.DeepClone() },
    };

    private static JsonArray Grid(params string[] cells) =>
        new() { new JsonArray(cells.Select(cell => (JsonNode)JsonValue.Create(cell)!).ToArray()) };

    private static IReadOnlyList<string> MismatchTexts(JsonObject result) =>
        (Json.GetArr(result, "mismatches") ?? new JsonArray())
        .Select(node => node?.ToString() ?? "")
        .ToArray();
}
