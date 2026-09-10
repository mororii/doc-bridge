using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// GetDocText used to catch every exception and return "", so a failed whole-document
/// read was reported as a successful read of an empty document: hwp_read_text answered
/// ok=true/length=0 and GetActiveContext answered Ok=true/textLength=0. A genuinely
/// empty document and a broken GetTextFile call were indistinguishable to every caller.
///
/// These tests pin the two halves of the fix. A real exception fails closed through the
/// existing error plumbing as HWP_DOCUMENT_TEXT_READ_FAILED, and a legitimately empty
/// or null return is still a success. Selection reads keep their tolerant behavior,
/// because an empty selection is the normal case.
///
/// The fakes are plain objects handed to the adapter's injected app factory and reached
/// through the same late-bound call sites production uses. No HWP process is started.
/// </summary>
public sealed class HwpTextReadFailureTests
{
    private const string FailureCode = "HWP_DOCUMENT_TEXT_READ_FAILED";

    // ------------------------------------------------------------ read: failure

    [Fact]
    public void Document_read_failure_is_reported_as_an_error_not_an_empty_document()
    {
        var hwp = new FakeHwp { DocumentTextThrows = "COM call failed (0x800A03EC)" };
        var result = NewAdapter(hwp).Read(new JsonObject { ["scope"] = "document" });

        Assert.False(Json.GetBool(result, "ok"), result.ToJsonString());
        Assert.Equal(FailureCode, Json.GetString(result, "errorCode"));
        // The defect being pinned: no success shape may survive a failed read.
        Assert.Null(result["text"]);
        Assert.Null(result["length"]);
        Assert.Contains("빈 문서로 보고하지 않습니다",
            string.Join(" ", (Json.GetArr(result, "errors") ?? new JsonArray()).Select(n => n?.ToString())));
        Assert.Equal(1, hwp.DocumentTextCalls);
    }

    [Theory]
    [InlineData("bundle")]
    [InlineData("document_map")]
    public void Bundle_and_map_reads_fail_closed_on_the_same_failure(string scope)
    {
        var hwp = new FakeHwp { DocumentTextThrows = "GetTextFile refused" };
        var args = new JsonObject { ["scope"] = scope };
        if (scope == "bundle") args["sections"] = new JsonArray("text");

        var result = NewAdapter(hwp).Read(args);

        Assert.False(Json.GetBool(result, "ok"), result.ToJsonString());
        Assert.Equal(FailureCode, Json.GetString(result, "errorCode"));
        Assert.Null(result["bundle"]);
        Assert.Null(result["paragraphs"]);
    }

    [Fact]
    public void Context_cannot_report_ok_when_the_document_read_fails()
    {
        var hwp = new FakeHwp { DocumentTextThrows = "RPC server unavailable" };
        var context = NewAdapter(hwp).GetActiveContext();

        // Ok is set true as soon as an app and a document are found, before the text
        // read. It must not survive the failure.
        Assert.False(context.Ok);
        Assert.Contains(context.Errors, e => e.Contains(FailureCode, StringComparison.Ordinal));
        Assert.False(context.Summary.ContainsKey("textLength"));
    }

    [Fact]
    public void Preview_fails_closed_rather_than_planning_against_an_empty_document()
    {
        var hwp = new FakeHwp { DocumentTextThrows = "document is busy" };
        var preview = NewAdapter(hwp).Preview(new[]
        {
            new JsonObject { ["op"] = "insert_text", ["text"] = "X" },
        });

        Assert.NotEmpty(preview.Errors);
        Assert.Contains(preview.Errors, e => e.Contains("본문을 읽지 못했습니다", StringComparison.Ordinal));
    }

    // ------------------------------------------------------------ read: success

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void A_genuinely_empty_document_is_still_a_successful_read(string? returned)
    {
        var hwp = new FakeHwp { DocumentText = returned };
        var result = NewAdapter(hwp).Read(new JsonObject { ["scope"] = "document" });

        Assert.True(Json.GetBool(result, "ok"), result.ToJsonString());
        Assert.Equal("", Json.GetString(result, "text"));
        Assert.Equal(0, Json.GetInt(result, "length"));
        Assert.Null(Json.GetString(result, "errorCode"));
    }

    [Fact]
    public void A_genuinely_empty_document_still_reports_context_ok()
    {
        var context = NewAdapter(new FakeHwp { DocumentText = "" }).GetActiveContext();

        Assert.True(context.Ok, string.Join("; ", context.Errors));
        Assert.Equal(0, Json.GetInt(context.Summary, "textLength"));
    }

    [Fact]
    public void Serialized_entities_are_still_decoded_on_the_success_path()
    {
        var hwp = new FakeHwp { DocumentText = "a&#8722;b" };
        var result = NewAdapter(hwp).Read(new JsonObject { ["scope"] = "document" });

        Assert.True(Json.GetBool(result, "ok"), result.ToJsonString());
        Assert.Equal("a−b", Json.GetString(result, "text"));
    }

    // --------------------------------------------------------- selection intact

    [Fact]
    public void Selection_read_keeps_its_tolerant_behavior_while_the_document_read_does_not()
    {
        var hwp = new FakeHwp
        {
            DocumentText = "body",
            SelectionTextThrows = "no block selected",
        };
        var adapter = NewAdapter(hwp);

        // Selection failure stays tolerant: an empty selection is the normal state.
        var selection = adapter.Read(new JsonObject { ["scope"] = "selection" });
        Assert.True(Json.GetBool(selection, "ok"), selection.ToJsonString());
        Assert.Equal("", Json.GetString(selection, "text"));

        // The document read on the same instance is unaffected by that tolerance.
        var document = adapter.Read(new JsonObject { ["scope"] = "document" });
        Assert.True(Json.GetBool(document, "ok"), document.ToJsonString());
        Assert.Equal("body", Json.GetString(document, "text"));
    }

    [Fact]
    public void Context_stays_ok_when_only_the_selection_read_fails()
    {
        var context = NewAdapter(new FakeHwp
        {
            DocumentText = "body",
            SelectionTextThrows = "no block selected",
        }).GetActiveContext();

        Assert.True(context.Ok, string.Join("; ", context.Errors));
        Assert.Equal(4, Json.GetInt(context.Summary, "textLength"));
        Assert.False(Json.GetBool(context.Selection, "hasSelection"));
    }

    // -------------------------------------------------------------------- fakes

    private static HwpAdapter NewAdapter(FakeHwp hwp) => new(() => hwp);

    /// <summary>
    /// Minimal shape the adapter reaches through `dynamic`. Members it probes and this
    /// fake does not have (window handles, controls, actions) throw and are already
    /// tolerated by the adapter, which is what a headless fake should look like.
    /// </summary>
    public sealed class FakeHwp
    {
        public string? DocumentText { get; set; } = "";

        public string? DocumentTextThrows { get; set; }

        public string? SelectionTextThrows { get; set; }

        public int DocumentTextCalls { get; private set; }

        public int SelectionTextCalls { get; private set; }

        public FakeDocuments XHwpDocuments { get; } = new();

        public string GetTextFile(string format, string option)
        {
            if (!string.Equals(format, "TEXT", StringComparison.Ordinal))
                throw new InvalidOperationException($"unexpected format '{format}'");

            if (string.Equals(option, "saveblock:true", StringComparison.Ordinal))
            {
                SelectionTextCalls++;
                if (SelectionTextThrows is { } selectionMessage)
                    throw new InvalidOperationException(selectionMessage);
                return "";
            }

            DocumentTextCalls++;
            if (DocumentTextThrows is { } message) throw new InvalidOperationException(message);
            return DocumentText!;
        }
    }

    public sealed class FakeDocuments
    {
        public FakeDocument Active_XHwpDocument { get; } = new();

        public int Count => 1;

        public FakeDocument Item(int index) => Active_XHwpDocument;
    }

    public sealed class FakeDocument
    {
        public string FullName => "";

        public int DocumentID => 1;
    }
}
