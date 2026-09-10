using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// CadAdapter.CadOperationStateHash used to reread EntityName, Layer, TextString,
/// Height, Rotation, InsertionPoint and GetBoundingBox for its legacy suffix even
/// though EntityJson had just read the same properties. The suffix now reuses those
/// captured values.
///
/// These tests do two things the review asked for. They pin the emitted v2 digest
/// against <see cref="LegacyOperationStateHash"/>, a reimplementation of the
/// pre-deduplication algorithm, so the new code is never compared only with itself;
/// and they count native-like property accesses on a fake entity to prove the
/// duplicate reads are actually gone rather than merely reordered.
///
/// No AutoCAD or GstarCAD process is involved. The fakes are plain objects reached
/// through the same late-bound `dynamic` call sites production uses.
/// </summary>
public class CadFingerprintReadPerformanceTests
{
    // Pinned so a future change to EntityJson, to the suffix, or to numeric
    // formatting shows up as a byte-level digest change and not just as
    // "old oracle and new code still agree with each other".
    private const string GoldenTextStateHash =
        "b8f818b1d0b6e4fa2c9f8c21f80a832db4e76b71af1322ba3d5a19749781c7ef";
    private const string GoldenMTextStateHash =
        "73c7d80d9ccd982a88306222b6844a60c2016f3eb4841f99f8723a0a3dc070e9";

    private static readonly string[] DeduplicatedMembers =
    {
        "EntityName", "Layer", "TextString", "Height", "Rotation", "InsertionPoint", "GetBoundingBox",
    };

    // ---------------------------------------------------------------- equality

    [Theory]
    [InlineData("AcDbText")]
    [InlineData("AcDbMText")]
    public void Deduplicated_suffix_matches_legacy_algorithm_for_text_and_mtext(string entityName)
    {
        var ops = new[] { TextOp("A1"), LayerColorOp("PLAN") };

        var legacy = LegacyOperationStateHash(NewDocument(entityName, new ReadCounter()), ops);
        var current = CadAdapter.CadOperationStateHash(NewDocument(entityName, new ReadCounter()), ops);

        Assert.NotNull(legacy);
        Assert.Equal(legacy, current);
    }

    [Fact]
    public void Deduplicated_suffix_matches_pinned_golden_hashes()
    {
        var ops = new[] { TextOp("A1"), LayerColorOp("PLAN") };

        Assert.Equal(GoldenTextStateHash,
            CadAdapter.CadOperationStateHash(NewDocument("AcDbText", new ReadCounter()), ops));
        Assert.Equal(GoldenMTextStateHash,
            CadAdapter.CadOperationStateHash(NewDocument("AcDbMText", new ReadCounter()), ops));

        // The oracle must agree with the pin too, so a golden refresh cannot quietly
        // bless a behavior change in the new implementation.
        Assert.Equal(GoldenTextStateHash,
            LegacyOperationStateHash(NewDocument("AcDbText", new ReadCounter()), ops));
        Assert.Equal(GoldenMTextStateHash,
            LegacyOperationStateHash(NewDocument("AcDbMText", new ReadCounter()), ops));
    }

    [Theory]
    // Values whose shortest round-trippable ("R") rendering is not the literal text,
    // so a reused double that had silently gone through a string or a float would
    // produce different bytes here.
    [InlineData(0.1, 0.2, 0.30000000000000004)]
    [InlineData(1e-7, -1e-7, 1.7976931348623157E+308)]
    [InlineData(4.9406564584124654E-324, 1234567.891011, -0.0)]
    [InlineData(3.141592653589793, 6.283185307179586, 2.220446049250313E-16)]
    [InlineData(0, 0, 0)]
    public void Deduplicated_suffix_preserves_invariant_numeric_bytes(double height, double rotation, double bound)
    {
        var ops = new[] { TextOp("A1") };
        var counter = new ReadCounter();

        var legacyEntity = NewEntity(counter, height: height, rotation: rotation,
            insertionPoint: new[] { bound, -bound, 0d }, minX: bound, minY: -bound, maxX: bound + 1, maxY: 2 * bound);
        var currentEntity = NewEntity(new ReadCounter(), height: height, rotation: rotation,
            insertionPoint: new[] { bound, -bound, 0d }, minX: bound, minY: -bound, maxX: bound + 1, maxY: 2 * bound);

        var legacy = LegacySuffixText(legacyEntity);
        var current = CurrentSuffixText(currentEntity);

        Assert.Equal(legacy, current);
        // Guard against both sides degenerating to empty separators only.
        Assert.Contains(height.ToString("R", CultureInfo.InvariantCulture), current, StringComparison.Ordinal);
    }

    public static TheoryData<string, Action<FakeEntity>> StateChanges() => new()
    {
        { "text", e => e.TextString = "CHANGED" },
        { "layer", e => e.Layer = "OTHER" },
        { "height", e => e.Height = 2.5 },
        { "rotation", e => e.Rotation = 0.75 },
        { "insertionPoint", e => e.InsertionPoint = new[] { 9d, 9d, 0d } },
        { "boundingBox", e => e.MaxX = 99 },
        { "visible", e => e.Visible = false },
        { "color", e => e.Color = 7 },
        { "transparency", e => e.EntityTransparency = "50" },
    };

    [Theory]
    [MemberData(nameof(StateChanges))]
    public void Relevant_state_changes_still_change_the_digest_identically(string label, Action<FakeEntity> mutate)
    {
        var ops = new[] { TextOp("A1"), LayerColorOp("PLAN") };

        var baselineLegacy = LegacyOperationStateHash(NewDocument("AcDbText", new ReadCounter()), ops);
        var baselineCurrent = CadAdapter.CadOperationStateHash(NewDocument("AcDbText", new ReadCounter()), ops);

        var mutatedLegacyDoc = NewDocument("AcDbText", new ReadCounter());
        mutate(mutatedLegacyDoc.Entity);
        var mutatedCurrentDoc = NewDocument("AcDbText", new ReadCounter());
        mutate(mutatedCurrentDoc.Entity);

        var mutatedLegacy = LegacyOperationStateHash(mutatedLegacyDoc, ops);
        var mutatedCurrent = CadAdapter.CadOperationStateHash(mutatedCurrentDoc, ops);

        Assert.Equal(baselineLegacy, baselineCurrent);
        Assert.Equal(mutatedLegacy, mutatedCurrent);
        Assert.True(mutatedCurrent != baselineCurrent, $"{label} did not change the digest");
    }

    [Fact]
    public void Layer_and_document_state_changes_still_change_the_digest_identically()
    {
        var ops = new[] { TextOp("A1"), LayerColorOp("PLAN") };
        var baseline = CadAdapter.CadOperationStateHash(NewDocument("AcDbText", new ReadCounter()), ops);

        var layerChanged = NewDocument("AcDbText", new ReadCounter());
        layerChanged.Layers.Entry.Color = 42;
        var countChanged = NewDocument("AcDbText", new ReadCounter());
        countChanged.ModelSpace.Count = 4;
        var renamed = NewDocument("AcDbText", new ReadCounter());
        renamed.FullName = @"C:\owned\other.dwg";

        foreach (var changed in new[] { layerChanged, countChanged, renamed })
        {
            Assert.Equal(LegacyOperationStateHash(changed, ops), CadAdapter.CadOperationStateHash(changed, ops));
            Assert.NotEqual(baseline, CadAdapter.CadOperationStateHash(changed, ops));
        }
    }

    [Theory]
    [InlineData("Height")]
    [InlineData("Rotation")]
    [InlineData("InsertionPoint")]
    [InlineData("GetBoundingBox")]
    public void Unreadable_geometry_fails_closed_exactly_as_before(string unreadableMember)
    {
        var ops = new[] { TextOp("A1") };

        var legacyDoc = NewDocument("AcDbText", new ReadCounter());
        legacyDoc.Entity.Unreadable.Add(unreadableMember);
        var currentDoc = NewDocument("AcDbText", new ReadCounter());
        currentDoc.Entity.Unreadable.Add(unreadableMember);

        Assert.Null(LegacyOperationStateHash(legacyDoc, ops));
        Assert.Null(CadAdapter.CadOperationStateHash(currentDoc, ops));
    }

    [Theory]
    [InlineData("AcDbLine")]
    [InlineData("AcDbPolyline")]
    [InlineData("AcDbBlockReference")]
    public void Unsupported_entity_types_still_refuse_before_reading_geometry(string entityName)
    {
        var ops = new[] { TextOp("A1") };
        var counter = new ReadCounter();
        var doc = NewDocument(entityName, counter);

        Assert.Null(LegacyOperationStateHash(NewDocument(entityName, new ReadCounter()), ops));
        Assert.Null(CadAdapter.CadOperationStateHash(doc, ops));
        // The type guard runs before EntityJson, so nothing beyond EntityName is read.
        Assert.Equal(1, counter["EntityName"]);
        Assert.Equal(0, counter["GetBoundingBox"]);
    }

    [Fact]
    public void Unresolvable_handle_and_unsupported_op_fail_closed_identically()
    {
        var missingHandleOps = new[] { TextOp("NOPE") };
        var docA = NewDocument("AcDbText", new ReadCounter());
        var docB = NewDocument("AcDbText", new ReadCounter());
        Assert.Null(LegacyOperationStateHash(docA, missingHandleOps));
        Assert.Null(CadAdapter.CadOperationStateHash(docB, missingHandleOps));

        var unsupportedOps = new[] { new JsonObject { ["op"] = "delete_entities", ["handle"] = "A1" } };
        Assert.Null(LegacyOperationStateHash(NewDocument("AcDbText", new ReadCounter()), unsupportedOps));
        Assert.Null(CadAdapter.CadOperationStateHash(NewDocument("AcDbText", new ReadCounter()), unsupportedOps));

        var missingLayerOps = new[] { LayerColorOp("GHOST") };
        Assert.Null(LegacyOperationStateHash(NewDocument("AcDbText", new ReadCounter()), missingLayerOps));
        Assert.Null(CadAdapter.CadOperationStateHash(NewDocument("AcDbText", new ReadCounter()), missingLayerOps));
    }

    // ------------------------------------------------------------ read counting

    [Theory]
    [InlineData("AcDbText")]
    [InlineData("AcDbMText")]
    public void Suffix_no_longer_rereads_the_seven_properties_entityjson_already_read(string entityName)
    {
        var ops = new[] { TextOp("A1"), LayerColorOp("PLAN") };

        var legacyCounter = new ReadCounter();
        var legacyHash = LegacyOperationStateHash(NewDocument(entityName, legacyCounter), ops);
        var currentCounter = new ReadCounter();
        var currentHash = CadAdapter.CadOperationStateHash(NewDocument(entityName, currentCounter), ops);

        Assert.Equal(legacyHash, currentHash);

        foreach (var member in DeduplicatedMembers)
        {
            Assert.True(legacyCounter[member] > 0, $"oracle never read {member}");
            Assert.Equal(legacyCounter[member] - 1, currentCounter[member]);
        }

        // Exactly one duplicate read of each of the seven is removed, and nothing else
        // about the traversal changes.
        Assert.Equal(legacyCounter.Total - DeduplicatedMembers.Length, currentCounter.Total);
    }

    [Fact]
    public void Removed_reads_scale_with_the_number_of_fingerprinted_entities()
    {
        var handles = new[] { "A1", "A2", "A3", "A4" };
        var ops = handles.Select(TextOp).ToArray();

        var legacyCounter = new ReadCounter();
        var legacyDoc = NewDocument("AcDbText", legacyCounter, handles);
        var currentCounter = new ReadCounter();
        var currentDoc = NewDocument("AcDbText", currentCounter, handles);

        Assert.Equal(LegacyOperationStateHash(legacyDoc, ops), CadAdapter.CadOperationStateHash(currentDoc, ops));
        Assert.Equal(legacyCounter.Total - DeduplicatedMembers.Length * handles.Length, currentCounter.Total);
    }

    [Fact]
    public void Reused_suffix_is_byte_identical_to_the_legacy_suffix_for_one_entity()
    {
        var legacyEntity = NewEntity(new ReadCounter());
        var currentEntity = NewEntity(new ReadCounter());

        Assert.Equal(LegacySuffixText(legacyEntity), CurrentSuffixText(currentEntity));
    }

    // ------------------------------------------------------- suffix-level oracle

    private static string CurrentSuffixText(FakeEntity entity)
    {
        var geometry = CadAdapter.CadEntityGeometryForFingerprint(entity);
        var state = new StringBuilder();
        CadAdapter.AppendEntityStateSuffix(state, geometry, entity);
        return state.ToString();
    }

    /// <summary>Byte-for-byte reimplementation of the pre-deduplication suffix.</summary>
    private static string LegacySuffixText(FakeEntity entityObject)
    {
        var state = new StringBuilder();
        // The legacy code built the geometry payload first and then reread everything.
        _ = CadAdapter.CadEntityGeometryForFingerprint(entityObject);
        dynamic entity = entityObject;
        try { state.Append((string)entity.EntityName); } catch { }
        state.Append('|');
        try { state.Append((string)entity.Layer); } catch { }
        state.Append('|').Append(LegacyTextOf(entity)).Append('|');
        try { state.Append(Convert.ToDouble(entity.Height, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture)); } catch { }
        state.Append('|');
        try { state.Append(Convert.ToDouble(entity.Rotation, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture)); } catch { }
        state.Append('|');
        try { state.Append(LegacyPointJson((object?)entity.InsertionPoint).ToJsonString()); } catch { }
        state.Append('|');
        if (LegacyTryBoundingBox(entityObject, out var minX, out var minY, out var maxX, out var maxY))
            state.Append(minX.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                .Append(minY.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                .Append(maxX.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                .Append(maxY.ToString("R", CultureInfo.InvariantCulture));
        return state.ToString();
    }

    /// <summary>
    /// Byte-for-byte reimplementation of CadOperationStateHash as it stood before the
    /// deduplication, covering the op kinds these tests exercise. activate_document is
    /// out of scope here and is treated as unsupported.
    /// </summary>
    private static string? LegacyOperationStateHash(FakeDocument docObject, IReadOnlyList<JsonObject> ops)
    {
        dynamic doc = docObject;
        var state = new StringBuilder(4096);
        string documentName = "";
        try { documentName = (string)(doc.FullName ?? doc.Name ?? ""); } catch { }
        state.Append("document=").Append(documentName).Append('\n');
        try { state.Append("modelSpaceCount=").Append((int)doc.ModelSpace.Count).Append('\n'); } catch { }

        var handles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var layers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var op in ops)
        {
            var opName = Json.GetString(op, "op");
            if (opName is not ("regen_document" or "set_layer_visibility" or "set_layer_color" or
                "move_entities" or "rotate_entities" or "scale_entities" or "set_text_value")) return null;
            state.Append("op=").Append(opName ?? "").Append('\n');
            var handle = Json.GetString(op, "handle");
            if (!string.IsNullOrWhiteSpace(handle)) handles.Add(handle);
            if (Json.GetArr(op, "handles") is { } handleArray)
                foreach (var node in handleArray)
                {
                    var value = node?.GetValue<string>();
                    if (!string.IsNullOrWhiteSpace(value)) handles.Add(value);
                }
            var layer = Json.GetString(op, "layer");
            if (!string.IsNullOrWhiteSpace(layer)) layers.Add(layer);
        }

        foreach (var layerName in layers.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            state.Append("layer=").Append(layerName).Append('|');
            try
            {
                dynamic layer = doc.Layers.Item(layerName);
                state.Append((bool)layer.LayerOn).Append('|')
                    .Append((bool)layer.Freeze).Append('|')
                    .Append((bool)layer.Lock).Append('|')
                    .Append((int)layer.Color);
            }
            catch { return null; }
            state.Append('\n');
        }

        foreach (var handle in handles.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            state.Append("entity=").Append(handle).Append('|');
            try
            {
                dynamic entity = doc.HandleToObject(handle);
                if ((string)entity.EntityName is not ("AcDbText" or "AcDbMText")) return null;
                var geometry = CadAdapter.CadEntityGeometryForFingerprint((object)entity);
                if (geometry["bounds"] is null || geometry["height"] is null || geometry["rotation"] is null ||
                    geometry["insertionPoint"] is null) return null;
                state.Append(geometry.ToJsonString()).Append('|');
                try { state.Append((string)entity.EntityName); } catch { }
                state.Append('|');
                try { state.Append((string)entity.Layer); } catch { }
                state.Append('|').Append(LegacyTextOf(entity)).Append('|');
                try { state.Append(Convert.ToDouble(entity.Height, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture)); } catch { }
                state.Append('|');
                try { state.Append(Convert.ToDouble(entity.Rotation, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture)); } catch { }
                state.Append('|');
                try { state.Append(LegacyPointJson((object?)entity.InsertionPoint).ToJsonString()); } catch { }
                state.Append('|');
                if (LegacyTryBoundingBox((object)entity, out var minX, out var minY, out var maxX, out var maxY))
                    state.Append(minX.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                        .Append(minY.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                        .Append(maxX.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                        .Append(maxY.ToString("R", CultureInfo.InvariantCulture));
            }
            catch { return null; }
            state.Append('\n');
        }

        return Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(Encoding.UTF8.GetBytes(state.ToString()))).ToLowerInvariant();
    }

    private static string LegacyTextOf(dynamic e)
    {
        try { return (string)(e.TextString ?? ""); } catch { return ""; }
    }

    private static JsonArray LegacyPointJson(object? value)
    {
        var result = new JsonArray();
        if (value is not Array array) return result;
        foreach (var item in array) result.Add(Convert.ToDouble(item, CultureInfo.InvariantCulture));
        return result;
    }

    private static bool LegacyTryBoundingBox(object entityObject, out double minX, out double minY, out double maxX, out double maxY)
    {
        minX = minY = maxX = maxY = 0;
        try
        {
            dynamic entity = entityObject;
            object minValue = null!;
            object maxValue = null!;
            entity.GetBoundingBox(out minValue, out maxValue);
            if (minValue is not Array min || maxValue is not Array max) return false;
            minX = Convert.ToDouble(min.GetValue(0), CultureInfo.InvariantCulture);
            minY = Convert.ToDouble(min.GetValue(1), CultureInfo.InvariantCulture);
            maxX = Convert.ToDouble(max.GetValue(0), CultureInfo.InvariantCulture);
            maxY = Convert.ToDouble(max.GetValue(1), CultureInfo.InvariantCulture);
            return true;
        }
        catch { return false; }
    }

    // -------------------------------------------------------------------- fakes

    private static JsonObject TextOp(string handle) =>
        new() { ["op"] = "set_text_value", ["handle"] = handle, ["value"] = "x" };

    private static JsonObject LayerColorOp(string layer) =>
        new() { ["op"] = "set_layer_color", ["layer"] = layer, ["color"] = 3 };

    private static FakeEntity NewEntity(
        ReadCounter counter,
        string entityName = "AcDbText",
        string handle = "A1",
        double height = 1.5,
        double rotation = 0.25,
        double[]? insertionPoint = null,
        double minX = -1,
        double minY = -2,
        double maxX = 3.5,
        double maxY = 4.25) => new(counter)
        {
            EntityName = entityName,
            Handle = handle,
            Layer = "PLAN",
            TextString = "ROOM 101",
            Height = height,
            Rotation = rotation,
            InsertionPoint = insertionPoint ?? new[] { 10.5, 20.25, 0d },
            MinX = minX,
            MinY = minY,
            MaxX = maxX,
            MaxY = maxY,
        };

    private static FakeDocument NewDocument(string entityName, ReadCounter counter, params string[] handles)
    {
        if (handles.Length == 0) handles = new[] { "A1" };
        var doc = new FakeDocument();
        foreach (var handle in handles)
            doc.Entities[handle] = NewEntity(counter, entityName, handle);
        return doc;
    }

    public sealed class ReadCounter
    {
        private readonly Dictionary<string, int> _counts = new(StringComparer.Ordinal);

        public int this[string member] => _counts.TryGetValue(member, out var n) ? n : 0;

        public int Total { get; private set; }

        public void Hit(string member)
        {
            _counts[member] = this[member] + 1;
            Total++;
        }
    }

    /// <summary>
    /// Plain object reached through the same late-bound call sites production uses.
    /// Every member the fingerprint can touch is counted, and any member listed in
    /// <see cref="Unreadable"/> throws the way an unavailable native property would.
    /// </summary>
    public sealed class FakeEntity
    {
        private readonly ReadCounter _counter;

        public FakeEntity(ReadCounter counter) => _counter = counter;

        public HashSet<string> Unreadable { get; } = new(StringComparer.Ordinal);

        private string _entityName = "AcDbText";
        private string _layer = "PLAN";
        private string _handle = "A1";
        private string _textString = "ROOM 101";
        private double _height = 1.5;
        private double _rotation = 0.25;
        private double[] _insertionPoint = { 10.5, 20.25, 0d };
        private bool _visible = true;
        private int _color = 256;
        private string _transparency = "ByLayer";

        public string EntityName { get => Read(nameof(EntityName), _entityName); set => _entityName = value; }
        public string Layer { get => Read(nameof(Layer), _layer); set => _layer = value; }
        public string Handle { get => Read(nameof(Handle), _handle); set => _handle = value; }
        public string TextString { get => Read(nameof(TextString), _textString); set => _textString = value; }
        public double Height { get => Read(nameof(Height), _height); set => _height = value; }
        public double Rotation { get => Read(nameof(Rotation), _rotation); set => _rotation = value; }
        public double[] InsertionPoint { get => Read(nameof(InsertionPoint), _insertionPoint); set => _insertionPoint = value; }
        public bool Visible { get => Read(nameof(Visible), _visible); set => _visible = value; }
        public int Color { get => Read(nameof(Color), _color); set => _color = value; }
        public string EntityTransparency { get => Read(nameof(EntityTransparency), _transparency); set => _transparency = value; }
        public bool HasAttributes => Read(nameof(HasAttributes), false);

        public double MinX { get; set; } = -1;
        public double MinY { get; set; } = -2;
        public double MaxX { get; set; } = 3.5;
        public double MaxY { get; set; } = 4.25;

        public void GetBoundingBox(out object min, out object max)
        {
            _counter.Hit(nameof(GetBoundingBox));
            if (Unreadable.Contains(nameof(GetBoundingBox)))
            {
                min = null!;
                max = null!;
                throw new InvalidOperationException("bounding box unavailable");
            }

            min = new object[] { MinX, MinY, 0d };
            max = new object[] { MaxX, MaxY, 0d };
        }

        private T Read<T>(string member, T value)
        {
            _counter.Hit(member);
            if (Unreadable.Contains(member)) throw new InvalidOperationException($"{member} unavailable");
            return value;
        }
    }

    public sealed class FakeDocument
    {
        public string FullName { get; set; } = @"C:\owned\probe.dwg";

        public string Name { get; set; } = "probe.dwg";

        public FakeModelSpace ModelSpace { get; } = new();

        public FakeLayers Layers { get; } = new();

        public Dictionary<string, FakeEntity> Entities { get; } = new(StringComparer.OrdinalIgnoreCase);

        public FakeEntity Entity => Entities.Values.First();

        public object HandleToObject(string handle) =>
            Entities.TryGetValue(handle, out var entity)
                ? entity
                : throw new KeyNotFoundException($"handle {handle} is not in this drawing");
    }

    public sealed class FakeModelSpace
    {
        public int Count { get; set; } = 3;
    }

    public sealed class FakeLayers
    {
        public FakeLayer Entry { get; } = new();

        public FakeLayer Item(string name) =>
            string.Equals(name, "PLAN", StringComparison.OrdinalIgnoreCase)
                ? Entry
                : throw new KeyNotFoundException($"layer {name} is not in this drawing");
    }

    public sealed class FakeLayer
    {
        public bool LayerOn { get; set; } = true;

        public bool Freeze { get; set; }

        public bool Lock { get; set; }

        public int Color { get; set; } = 3;
    }
}
