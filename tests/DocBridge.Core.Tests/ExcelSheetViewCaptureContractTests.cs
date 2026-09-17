using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelSheetViewCaptureContractTests
{
    [Fact]
    public void Two_sheet_views_do_not_cross_assign_and_original_active_is_restored()
    {
        var surface = new FakeSheetViews();
        surface.Activate("A");
        surface.Views["A"] = 80;
        surface.Views["B"] = 140;
        surface.Selection = "A1";
        surface.ScrollRow = 3;
        surface.Log.Clear();

        var capturedB = ExcelSheetViewCaptureContract.CaptureTargetView(surface, "B");
        Assert.True(ExcelSheetViewCaptureContract.CapturedViewBelongsTo(capturedB, "B"));
        Assert.Equal(140, Json.GetInt(capturedB, "zoom"));
        Assert.Equal("A", surface.ActiveSheet);
        Assert.Equal("A1", surface.Selection);
        Assert.Equal(3, surface.ScrollRow);
        Assert.Equal(new[] { "activate:B", "restore:A" }, surface.Log);

        var capturedA = ExcelSheetViewCaptureContract.CaptureTargetView(surface, "A");
        Assert.Equal(80, Json.GetInt(capturedA, "zoom"));
        Assert.False(ExcelSheetViewCaptureContract.MustActivateTarget("A", "A"));
        Assert.True(ExcelSheetViewCaptureContract.MustActivateTarget("B", "A"));
    }

    [Fact]
    public void Restore_failure_and_active_mismatch_are_not_readable()
    {
        var restoreFail = new FakeSheetViews { FailRestore = true };
        restoreFail.Activate("A");
        restoreFail.Views["B"] = 140;
        var restored = ExcelSheetViewCaptureContract.CaptureTargetView(restoreFail, "B");
        Assert.False(Json.GetBool(restored, "readable"));
        Assert.Contains("restore failed", Json.GetString(restored, "error") ?? "", StringComparison.OrdinalIgnoreCase);

        var stuck = new FakeSheetViews();
        stuck.Activate("A");
        stuck.IgnoreActivate = true;
        stuck.Views["B"] = 90;
        var mismatch = ExcelSheetViewCaptureContract.CaptureTargetView(stuck, "B");
        Assert.False(Json.GetBool(mismatch, "readable"));
        Assert.Equal("A", stuck.ActiveSheet);
    }

    private sealed class FakeSheetViews : ExcelSheetViewCaptureContract.ISheetViewSurface
    {
        public Dictionary<string, int> Views { get; } = new(StringComparer.OrdinalIgnoreCase);
        public List<string> Log { get; } = [];
        public string? ActiveSheet { get; private set; }
        public string? Selection { get; set; }
        public int? ScrollRow { get; set; }
        public int? ScrollColumn { get; set; }
        public bool FailRestore { get; set; }
        public bool IgnoreActivate { get; set; }

        public ExcelSheetViewCaptureContract.WindowBookmark CaptureOriginal() =>
            new(ActiveSheet, Selection, ScrollRow, ScrollColumn);

        public void Activate(string sheet)
        {
            if (!IgnoreActivate)
                ActiveSheet = sheet;
            Log.Add($"activate:{sheet}");
        }

        public JsonObject ReadActiveView() =>
            new()
            {
                ["zoom"] = ActiveSheet is not null && Views.TryGetValue(ActiveSheet, out var zoom) ? zoom : 100,
            };

        public void Restore(ExcelSheetViewCaptureContract.WindowBookmark original)
        {
            if (FailRestore)
                throw new InvalidOperationException("window bookmark restore failed");
            ActiveSheet = original.ActiveSheet;
            Selection = original.Selection;
            ScrollRow = original.ScrollRow;
            ScrollColumn = original.ScrollColumn;
            Log.Add($"restore:{original.ActiveSheet}");
        }
    }
}
