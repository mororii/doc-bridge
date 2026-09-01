using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>DXF fallback 파서 단위 테스트 (AutoCAD 불필요)</summary>
public class DxfReaderTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "dxf-test-" + Guid.NewGuid().ToString("N")[..6]);

    private const string SampleDxf = """
  0
SECTION
  2
TABLES
  0
TABLE
  2
LAYER
  0
LAYER
  2
0
 70
0
 62
7
  6
CONTINUOUS
  0
LAYER
  2
WALLS
 70
0
 62
3
  6
CONTINUOUS
  0
ENDTAB
  0
ENDSEC
  0
SECTION
  2
ENTITIES
  0
TEXT
  5
1A2B
  8
WALLS
  1
Hello DXF
  0
LINE
  5
1A2C
  8
0
 10
0.0
 20
0.0
 11
100.0
 21
0.0
  0
MTEXT
  5
1A2D
  8
WALLS
  1
Multi line text
  0
ENDSEC
  0
EOF
""";

    public DxfReaderTests() => Directory.CreateDirectory(_dir);
    public void Dispose() { try { Directory.Delete(_dir, true); } catch { } }

    [Fact]
    public void Analyze_extracts_layers_and_entities()
    {
        var path = Path.Combine(_dir, "sample.dxf");
        File.WriteAllText(path, SampleDxf);

        var result = DxfReader.Analyze(path);
        Assert.True(Json.GetBool(result, "ok"), $"analyze failed: {result}");
        Assert.Equal("dxf-fallback (read-only)", Json.GetString(result, "mode"));

        var summary = Json.GetObj(result, "summary")!;
        var layers = Json.GetArr(summary, "layers")!;
        Assert.Contains(layers, l => Json.GetString(l as JsonObject, "name") == "WALLS");
        Assert.Equal(3, Json.GetInt(summary, "entityCount").GetValueOrDefault());

        var entities = Json.GetArr(result, "entities")!;
        var text = entities.FirstOrDefault(e => Json.GetString(e as JsonObject, "type") == "TEXT") as JsonObject;
        Assert.NotNull(text);
        Assert.Equal("Hello DXF", Json.GetString(text, "text"));
        Assert.Equal("WALLS", Json.GetString(text, "layer"));
        Assert.Equal("1A2B", Json.GetString(text, "handle"));
    }

    [Fact]
    public void Analyze_missing_file_returns_error()
    {
        var result = DxfReader.Analyze(Path.Combine(_dir, "nope.dxf"));
        Assert.False(Json.GetBool(result, "ok"));
    }
}
