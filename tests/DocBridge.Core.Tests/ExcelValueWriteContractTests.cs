using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelValueWriteContractTests
{
    [Fact]
    public void Iso_date_string_is_not_excel_serial_46275()
    {
        var written = ExcelValueWriteContract.Classify(JsonValue.Create("2026-09-10"));
        Assert.Equal(ExcelValueWriteContract.JsonKind.String, written.Kind);
        Assert.True(ExcelValueWriteContract.IsIsoDateText("2026-09-10"));
        Assert.False(ExcelValueWriteContract.TypedEqual(written, ExcelValueWriteContract.IsoDateFixtureSerial));
        Assert.True(ExcelValueWriteContract.TypedEqual(written, "2026-09-10"));
        Assert.False(ExcelValueWriteContract.TypedEqual(written, 46275d));
    }

    [Fact]
    public void Empty_string_may_normalize_to_blank_without_keeping_at_format()
    {
        var empty = ExcelValueWriteContract.Classify(JsonValue.Create(""));
        var missing = ExcelValueWriteContract.Classify(null);
        Assert.Equal(ExcelValueWriteContract.JsonKind.EmptyString, empty.Kind);
        Assert.Equal(ExcelValueWriteContract.JsonKind.Null, missing.Kind);
        Assert.True(ExcelValueWriteContract.TypedEqual(missing, null));
        Assert.True(ExcelValueWriteContract.TypedEqual(empty, null));
        Assert.True(ExcelValueWriteContract.TypedEqual(empty, ""));
        Assert.True(ExcelValueWriteContract.TypedEqual(missing, ""));
        Assert.Null(empty.ComValue);
        Assert.Null(empty.Text);
        Assert.False(ExcelValueWriteContract.RequiresTextNumberFormat(missing.Kind));
        Assert.False(ExcelValueWriteContract.RequiresTextNumberFormat(empty.Kind));
    }

    [Fact]
    public void String_123_is_not_number_123()
    {
        var text = ExcelValueWriteContract.Classify(JsonValue.Create("123"));
        var number = ExcelValueWriteContract.Classify(JsonValue.Create(123));
        Assert.Equal(ExcelValueWriteContract.JsonKind.String, text.Kind);
        Assert.Equal(ExcelValueWriteContract.JsonKind.Number, number.Kind);
        Assert.True(ExcelValueWriteContract.DisplayEqualIsNotTypedEqual("123", 123d));
        Assert.False(ExcelValueWriteContract.TypedEqual(text, 123d));
        Assert.True(ExcelValueWriteContract.TypedEqual(text, "123"));
        Assert.True(ExcelValueWriteContract.TypedEqual(number, 123d));
        Assert.False(ExcelValueWriteContract.TypedEqual(number, "123"));
    }

    [Fact]
    public void Bool_and_leading_zero_text_keep_their_json_kind()
    {
        Assert.Equal(ExcelValueWriteContract.JsonKind.Boolean, ExcelValueWriteContract.Classify(JsonValue.Create(true)).Kind);
        Assert.True(ExcelValueWriteContract.TypedEqual(ExcelValueWriteContract.Classify(JsonValue.Create(true)), true));
        Assert.False(ExcelValueWriteContract.TypedEqual(ExcelValueWriteContract.Classify(JsonValue.Create(true)), "TRUE"));
        var zeros = ExcelValueWriteContract.Classify(JsonValue.Create("001500"));
        Assert.Equal(ExcelValueWriteContract.JsonKind.String, zeros.Kind);
        Assert.True(ExcelValueWriteContract.TypedEqual(zeros, "001500"));
        var writes = new ExcelValueWriteContract.CellWrite[1, 4];
        writes[0, 0] = ExcelValueWriteContract.Classify(JsonValue.Create(1));
        writes[0, 1] = ExcelValueWriteContract.Classify(JsonValue.Create("123"));
        writes[0, 2] = zeros;
        writes[0, 3] = ExcelValueWriteContract.Classify(JsonValue.Create(2));
        Assert.Equal(new[] { (0, 1, 2) }, ExcelValueWriteContract.ContiguousTextRuns(writes, 1, 4));
    }
}
