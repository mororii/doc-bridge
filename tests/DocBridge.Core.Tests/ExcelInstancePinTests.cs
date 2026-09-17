using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelInstancePinTests : IDisposable
{
    private readonly string _home;
    private readonly string? _previous;

    public ExcelInstancePinTests()
    {
        _previous = Environment.GetEnvironmentVariable("DOCBRIDGE_HOME");
        _home = Path.Combine(Path.GetTempPath(), "docbridge-pin-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_home);
        Environment.SetEnvironmentVariable("DOCBRIDGE_HOME", _home);
    }

    public void Dispose()
    {
        Environment.SetEnvironmentVariable("DOCBRIDGE_HOME", _previous);
        try { Directory.Delete(_home, recursive: true); } catch { }
    }

    [Fact]
    public void Pin_round_trips_and_clears()
    {
        Assert.Null(ExcelInstancePin.TryLoadPin());
        Assert.True(ExcelInstancePin.TrySavePin(1234, 5678, "excel_launch", out var error));
        Assert.Null(error);
        var loaded = ExcelInstancePin.TryLoadPin();
        Assert.NotNull(loaded);
        Assert.Equal(1234, loaded.ProcessId);
        Assert.Equal(5678, loaded.Hwnd);
        ExcelInstancePin.ClearPin();
        Assert.Null(ExcelInstancePin.TryLoadPin());
    }

    [Fact]
    public void Pin_rejects_non_positive_identity()
    {
        Assert.False(ExcelInstancePin.TrySavePin(0, 10, "x", out _));
        Assert.False(ExcelInstancePin.TrySavePin(10, 0, "x", out _));
        Assert.Null(ExcelInstancePin.TryLoadPin());
    }

    [Fact]
    public void Corrupt_pin_reads_as_missing()
    {
        Directory.CreateDirectory(_home);
        File.WriteAllText(ExcelInstancePin.PinPath(), "{not json");
        Assert.Null(ExcelInstancePin.TryLoadPin());
        File.WriteAllText(ExcelInstancePin.PinPath(), """{ "version": 999, "processId": 1, "hwnd": 2 }""");
        Assert.Null(ExcelInstancePin.TryLoadPin());
    }

    [Fact]
    public void Selector_validation_accepts_each_shape()
    {
        AssertValid("""{ "processId": 1234 }""");
        AssertValid("""{ "hwnd": 5678 }""");
        AssertValid("""{ "processId": 1234, "hwnd": 5678 }""");
        AssertValid("""{ "activeWindow": true }""");
        AssertValid("""{ "clearPin": true }""");
        AssertValid("""{ "dedicatedInstance": true }""");
        AssertValid("""{}""");
    }

    [Fact]
    public void Selector_validation_rejects_conflicts_and_bad_types()
    {
        AssertInvalid("""{ "processId": 1, "activeWindow": true }""", "one instance selector");
        AssertInvalid("""{ "hwnd": 2, "activeWindow": true }""", "one instance selector");
        AssertInvalid("""{ "dedicatedInstance": true, "processId": 1 }""", "dedicatedInstance");
        AssertInvalid("""{ "clearPin": true, "hwnd": 2 }""", "clearPin");
        AssertInvalid("""{ "processId": 0 }""", "positive integer");
        AssertInvalid("""{ "processId": 1.5 }""", "positive integer");
        AssertInvalid("""{ "hwnd": -3 }""", "positive integer");
        AssertInvalid("""{ "activeWindow": "yes" }""", "boolean");
        AssertInvalid("""{ "clearPin": 1 }""", "boolean");
    }

    private static void AssertValid(string json)
    {
        var errors = new List<string>();
        Assert.True(ExcelInstancePin.TryValidateSelectors(Json.ParseObject(json), errors),
            string.Join("; ", errors));
    }

    private static void AssertInvalid(string json, string token)
    {
        var errors = new List<string>();
        Assert.False(ExcelInstancePin.TryValidateSelectors(Json.ParseObject(json), errors));
        Assert.Contains(errors, error => error.Contains(token, StringComparison.OrdinalIgnoreCase));
    }
}
