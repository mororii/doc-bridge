using System.Text.Json.Nodes;
using DocBridge.Core.Services;
using Xunit;

namespace DocBridge.Core.Tests;

public class ExcelDrawingDetailsContractTests
{
    [Fact]
    public void Full_chart_details_payload_is_publicly_validated()
    {
        var op = Json.ParseObject("""{ "op":"update_chart","target":{"sheet":"S"},"name":"C","dataLabels":{"show":true,"value":true},"series":[{"index":1,"dataLabels":{"category":true},"trendline":{"action":"add","type":"polynomial","order":2},"points":[{"index":1,"fillColor":"#112233","lineColor":"#445566"}]}]}""")!;
        Assert.Empty(ExcelDataOperationsContract.ValidatePublicInput(op));
        Assert.Empty(ExcelDataOperationSchema.Validate(op));
        var items = new JsonObject { ["type"] = "object", ["required"] = new JsonArray("op"), ["properties"] = new JsonObject { ["op"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("update_chart") } } };
        ExcelDataOperationSchema.AttachToApplyItems(items);
        var errors = new List<string>();
        Assert.True(ExcelDataOperationSchema.TryValidatePublishedItems(items, op, errors), string.Join(";", errors));
    }
    [Fact]
    public void Invalid_chart_and_picture_details_are_rejected()
    {
        var chart = Json.ParseObject("""{ "op":"update_chart","target":{"sheet":"S"},"name":"C","series":[{"trendline":{"action":"add","type":"polynomial","order":7}}]}""")!;
        var picture = Json.ParseObject("""{ "op":"update_picture","target":{"sheet":"S"},"name":"P","crop":{"left":-1}}""")!;
        Assert.NotEmpty(ExcelDataOperationsContract.ValidatePublicInput(chart));
        Assert.NotEmpty(ExcelDataOperationsContract.ValidatePublicInput(picture));
    }
}
