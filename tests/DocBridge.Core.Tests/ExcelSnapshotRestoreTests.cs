using System.Collections;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelSnapshotRestoreTests
{
    [Fact]
    public void Copy_sheet_snapshot_captures_topology_without_reading_existing_ranges()
    {
        using var home = new TestHome();
        var workbook = new FakeWorkbook(@"C:\fixtures\target.xlsx", "Base", "Formula-heavy");
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        var metadata = new JsonObject();

        adapter.CaptureSnapshot(snapshotDir, metadata, CopyOps("Copied-1", "Copied-2"));

        var state = JsonNode.Parse(File.ReadAllText(Path.Combine(snapshotDir, "state.json")))!.AsObject();
        Assert.Equal(2, Json.GetInt(state, "snapshotVersion"));
        Assert.Equal("copy-sheet-topology", Json.GetString(state, "restoreMode"));
        Assert.Equal(workbook.FullName, Json.GetString(state, "documentRef"));
        Assert.Equal(new[] { "Base", "Formula-heavy" }, ReadStrings(state, "originalSheets"));
        Assert.Equal("Base", Json.GetString(state, "originalActiveSheet"));
        Assert.Equal(new[] { "Copied-1", "Copied-2" }, ReadStrings(state, "targetSheets"));
        Assert.Equal(0, workbook.RangeAccessCount);
        Assert.Equal(0, workbook.FormulaSetterCount);
        Assert.Equal(0, workbook.ValueSetterCount);
    }

    [Fact]
    public void Copy_sheet_restore_deletes_targets_in_reverse_and_never_writes_existing_cells()
    {
        using var home = new TestHome();
        var workbook = new FakeWorkbook(@"C:\fixtures\target.xlsx", "Base", "Formula-heavy");
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        var metadata = new JsonObject();
        adapter.CaptureSnapshot(snapshotDir, metadata, CopyOps("Copied-1", "Copied-2"));

        workbook.AddSheet("Copied-1", activate: true);
        workbook.AddSheet("Copied-2", activate: true);
        Assert.Equal("Copied-2", workbook.ActiveSheetName);
        // Production ExcelWorker in 0.4.14 did not return metadata mutations to the host.
        // The structural snapshot must therefore remain self-identifying in state.json.
        var restored = adapter.RestoreSnapshot(snapshotDir, new JsonObject());

        Assert.True(Json.GetBool(restored, "ok"), restored.ToJsonString());
        Assert.True(Json.GetBool(restored, "restored"));
        Assert.Equal("copy-sheet-topology", Json.GetString(restored, "restoreMode"));
        Assert.Equal(new[] { "Copied-2", "Copied-1" }, workbook.DeletedSheets);
        Assert.Equal(new[] { "Base", "Formula-heavy" }, workbook.SheetNames);
        Assert.Equal("Base", workbook.ActiveSheetName);
        Assert.Equal(0, workbook.RangeAccessCount);
        Assert.Equal(0, workbook.FormulaSetterCount);
        Assert.Equal(0, workbook.ValueSetterCount);
        var readback = Json.GetObj(restored, "readback")!;
        Assert.True(Json.GetBool(readback, "verified"));
        Assert.Equal(5, Json.GetInt(readback, "checked"));
        Assert.Equal(0, Json.GetInt(readback, "totalMismatchCount"));
        Assert.False(Json.GetBool(readback, "mismatchesTruncated"));
    }

    [Fact]
    public void Structural_restore_caps_mismatch_samples_and_never_claims_restored_when_unverified()
    {
        using var home = new TestHome();
        var originals = Enumerable.Range(1, 150).Select(index => $"Original-{index:000}").ToArray();
        var workbook = new FakeWorkbook(@"C:\fixtures\target.xlsx", originals);
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        var metadata = new JsonObject();
        adapter.CaptureSnapshot(snapshotDir, metadata, CopyOps("Copied"));

        workbook.RenameAll(index => $"Changed-{index:000}");
        workbook.AddSheet("Copied");
        var restored = adapter.RestoreSnapshot(snapshotDir, metadata);

        Assert.False(Json.GetBool(restored, "ok"));
        Assert.False(Json.GetBool(restored, "restored"));
        var readback = Json.GetObj(restored, "readback")!;
        Assert.False(Json.GetBool(readback, "verified"));
        Assert.Equal(152, Json.GetInt(readback, "totalMismatchCount"));
        Assert.Equal(100, Json.GetInt(readback, "mismatchSampleCount"));
        Assert.True(Json.GetBool(readback, "mismatchesTruncated"));
        Assert.Equal(100, Json.GetArr(readback, "mismatches")!.Count);
        Assert.Equal(100, Json.GetArr(restored, "errors")!.Count);
        Assert.Equal(0, workbook.FormulaSetterCount);
        Assert.Equal(0, workbook.ValueSetterCount);
    }

    [Fact]
    public void Versionless_snapshot_retains_legacy_full_range_restore_contract()
    {
        using var home = new TestHome();
        var workbook = new FakeWorkbook(@"C:\fixtures\target.xlsx", "Base");
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        File.WriteAllText(Path.Combine(snapshotDir, "state.json"), new JsonObject
        {
            ["sheets"] = new JsonObject
            {
                ["Base"] = new JsonObject
                {
                    ["address"] = "A1",
                    ["values"] = new JsonArray(new JsonArray(1)),
                    ["formulas"] = new JsonArray(new JsonArray("=1")),
                    ["truncated"] = false,
                },
            },
            ["ops"] = new JsonArray(),
            ["formatStates"] = new JsonArray(),
        }.ToJsonString());

        var restored = adapter.RestoreSnapshot(snapshotDir, new JsonObject
        {
            ["documentRef"] = workbook.FullName,
        });

        Assert.True(Json.GetBool(restored, "ok"), restored.ToJsonString());
        Assert.True(Json.GetBool(restored, "restored"));
        Assert.Equal("legacy-full-range", Json.GetString(restored, "restoreMode"));
        Assert.Equal(1, workbook.FormulaSetterCount);
        Assert.Equal(0, workbook.ValueSetterCount);
    }

    [Fact]
    public void Failed_copy_sheet_apply_uses_structural_restore_for_automatic_rollback()
    {
        using var home = new TestHome();
        var workbook = new FakeWorkbook(@"C:\fixtures\target.xlsx", "Base", "Formula-heavy");
        using var host = new DocBridgeHost(home.Options);
        var harness = new FailingCopySheetAdapter(workbook);
        host.Router.Register("excel", harness);
        var ops = CopyOps("Copied");

        var dryRun = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = ToJsonArray(ops),
            ["dryRun"] = true,
        });
        Assert.True(Json.GetBool(dryRun, "ok"), dryRun.ToJsonString());

        var applied = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = ToJsonArray(ops),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(dryRun, "confirmToken"),
        });

        Assert.False(Json.GetBool(applied, "ok"));
        var rollback = Json.GetObj(applied, "rollback")!;
        Assert.True(Json.GetBool(rollback, "attempted"));
        Assert.True(Json.GetBool(rollback, "verified"), rollback.ToJsonString());
        Assert.Equal("copy-sheet-topology",
            Json.GetString(Json.GetObj(rollback, "result"), "restoreMode"));
        Assert.Equal(new[] { "Base", "Formula-heavy" }, workbook.SheetNames);
        Assert.Equal("Base", workbook.ActiveSheetName);
        Assert.Equal(new[] { "Copied" }, workbook.DeletedSheets);
        Assert.Equal(0, workbook.FormulaSetterCount);
        Assert.Equal(0, workbook.ValueSetterCount);
    }

    private static ExcelAdapter CreateAdapter(FakeWorkbook workbook) =>
        new(() => new FakeExcelApplication(workbook));

    private static string CreateSnapshotDir(TestHome home)
    {
        var result = Path.Combine(home.Dir, "snapshot");
        Directory.CreateDirectory(result);
        return result;
    }

    private static IReadOnlyList<JsonObject> CopyOps(params string[] targetSheets) =>
        targetSheets.Select(target => new JsonObject
        {
            ["op"] = "copy_sheet",
            ["sourceWorkbook"] = "source.xlsx",
            ["sourceSheet"] = "Source",
            ["targetSheet"] = target,
        }).ToArray();

    private static JsonArray ToJsonArray(IEnumerable<JsonObject> values)
    {
        var result = new JsonArray();
        foreach (var value in values) result.Add(value.DeepClone());
        return result;
    }

    private static string[] ReadStrings(JsonObject value, string property) =>
        Json.GetArr(value, property)!.Select(item => item!.GetValue<string>()).ToArray();

    public sealed class FakeExcelApplication
    {
        public FakeExcelApplication(FakeWorkbook workbook) => ActiveWorkbook = workbook;
        public FakeWorkbook ActiveWorkbook { get; }
        public long Hwnd => 1;
        public bool ScreenUpdating { get; set; } = true;
        public bool DisplayAlerts { get; set; } = true;
    }

    public sealed class FakeWorkbook
    {
        private readonly List<FakeWorksheet> _sheets = new();
        private FakeWorksheet? _activeSheet;

        public FakeWorkbook(string fullName, params string[] sheetNames)
        {
            FullName = fullName;
            Name = Path.GetFileName(fullName);
            Worksheets = new FakeWorksheets(this);
            foreach (var sheetName in sheetNames) AddSheet(sheetName);
        }

        public string FullName { get; }
        public string Name { get; }
        public FakeWorksheets Worksheets { get; }
        public FakeWorksheet ActiveSheet => _activeSheet ??
            throw new InvalidOperationException("workbook has no active worksheet");
        public string ActiveSheetName => ActiveSheet.Name;
        public int RangeAccessCount { get; private set; }
        public int FormulaSetterCount { get; private set; }
        public int ValueSetterCount { get; private set; }
        public List<string> DeletedSheets { get; } = new();
        public string[] SheetNames => _sheets.Select(sheet => sheet.Name).ToArray();

        public void AddSheet(string name, bool activate = false)
        {
            var sheet = new FakeWorksheet(this, name);
            _sheets.Add(sheet);
            if (_activeSheet is null || activate) _activeSheet = sheet;
        }

        public void RenameAll(Func<int, string> nameFactory)
        {
            for (var index = 0; index < _sheets.Count; index++) _sheets[index].Name = nameFactory(index + 1);
        }

        internal int SheetCount => _sheets.Count;
        internal FakeWorksheet SheetAt(int oneBasedIndex) => _sheets[oneBasedIndex - 1];
        internal FakeWorksheet SheetNamed(string name) =>
            _sheets.First(sheet => string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase));

        internal void Delete(FakeWorksheet sheet)
        {
            DeletedSheets.Add(sheet.Name);
            var deletedActiveSheet = ReferenceEquals(_activeSheet, sheet);
            var deletedIndex = _sheets.IndexOf(sheet);
            _sheets.Remove(sheet);
            if (deletedActiveSheet)
            {
                if (_sheets.Count == 0) _activeSheet = null;
                else _activeSheet = _sheets[Math.Min(deletedIndex, _sheets.Count - 1)];
            }
        }

        internal void Activate(FakeWorksheet sheet)
        {
            if (!_sheets.Contains(sheet)) throw new InvalidOperationException("worksheet is not in this workbook");
            _activeSheet = sheet;
        }

        internal void RecordRangeAccess() => RangeAccessCount++;
        internal void RecordFormulaSetter() => FormulaSetterCount++;
        internal void RecordValueSetter() => ValueSetterCount++;
    }

    public sealed class FakeWorksheets : IEnumerable<FakeWorksheet>
    {
        private readonly FakeWorkbook _workbook;
        public FakeWorksheets(FakeWorkbook workbook) => _workbook = workbook;
        public int Count => _workbook.SheetCount;
        public FakeWorksheet Item(object key) => key switch
        {
            int index => _workbook.SheetAt(index),
            string name => _workbook.SheetNamed(name),
            _ => throw new ArgumentOutOfRangeException(nameof(key)),
        };
        public IEnumerator<FakeWorksheet> GetEnumerator() =>
            ((IEnumerable<FakeWorksheet>)_workbook.SheetNames.Select(_workbook.SheetNamed).ToArray()).GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    public sealed class FakeWorksheet
    {
        private readonly FakeWorkbook _workbook;
        private readonly FakeRange _range;

        public FakeWorksheet(FakeWorkbook workbook, string name)
        {
            _workbook = workbook;
            Name = name;
            _range = new FakeRange(workbook);
        }

        public string Name { get; set; }
        public FakeRange UsedRange
        {
            get { _workbook.RecordRangeAccess(); return _range; }
        }
        public FakeRange Range(string address)
        {
            _workbook.RecordRangeAccess();
            return _range;
        }
        public void Delete() => _workbook.Delete(this);
        public void Activate() => _workbook.Activate(this);
    }

    public sealed class FakeRange
    {
        private readonly FakeWorkbook _workbook;
        private object? _formula;
        private object? _value;

        public FakeRange(FakeWorkbook workbook) => _workbook = workbook;
        public int Row => 1;
        public int Column => 1;
        public object? Formula
        {
            get => _formula;
            set { _workbook.RecordFormulaSetter(); _formula = AsComArray(value); }
        }
        public object? Value2
        {
            get => _value;
            set { _workbook.RecordValueSetter(); _value = AsComArray(value); }
        }

        private static object? AsComArray(object? value)
        {
            if (value is not object[,] source) return value;
            var rows = source.GetLength(0);
            var columns = source.GetLength(1);
            var result = Array.CreateInstance(typeof(object), new[] { rows, columns }, new[] { 1, 1 });
            for (var row = 0; row < rows; row++)
                for (var column = 0; column < columns; column++)
                    result.SetValue(source[row, column], row + 1, column + 1);
            return result;
        }
    }

    private sealed class FailingCopySheetAdapter : IAppAdapter
    {
        private readonly FakeWorkbook _workbook;
        private readonly ExcelAdapter _inner;

        public FailingCopySheetAdapter(FakeWorkbook workbook)
        {
            _workbook = workbook;
            _inner = CreateAdapter(workbook);
        }

        public string App => "excel";
        public AdapterStatus GetStatus() => new(true, true, "excel", "test", _workbook.FullName, "test");
        public JsonObject GetCapabilities() => new();
        public ContextResult GetActiveContext() => new() { Ok = true, App = App, DocumentRef = _workbook.FullName };
        public JsonObject Read(JsonObject args) => new() { ["ok"] = true };
        public ApplyPreview Preview(IReadOnlyList<JsonObject> ops) => new();

        public ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId)
        {
            foreach (var op in ops)
                _workbook.AddSheet(
                    Json.GetString(op, "targetSheet") ?? Json.GetString(op, "sourceSheet")!,
                    activate: true);
            var result = new ApplyExecution
            {
                Ok = false,
                Readback = new JsonObject { ["verified"] = false },
            };
            result.Errors.Add("simulated failure after copy_sheet");
            return result;
        }

        public void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null) =>
            _inner.CaptureSnapshot(snapshotDir, metadata, ops);
        public JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata) =>
            _inner.RestoreSnapshot(snapshotDir, metadata);
        public void Dispose() => _inner.Dispose();
    }
}
