using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private static void PreviewCreateConnector(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var sheet = BindSheet(workbook, op); EnsureSheetWritable(sheet.Sheet);
        preview.Affected.Add(new AffectedRef("connector", $"{sheet.SheetName}!{Json.GetString(op, "name") ?? "new"}"));
        preview.Diff.Add(new DiffEntry { Ref = $"{sheet.SheetName}:connector", Before = null, After = op.DeepClone() });
    }
    private static void PreviewUpdateConnector(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var sheet = BindSheet(workbook, op); var shape = FindShape(sheet.Sheet, Json.GetString(op, "name")!);
        try { if (shape is null) { preview.Errors.Add("[EXCEL_CONNECTOR_NOT_FOUND] connector was not found"); return; }
            if (((dynamic)shape).Connector != ExcelDataObjectCatalog.MsoTrue) { preview.Errors.Add("[EXCEL_CONNECTOR_TYPE] named shape is not a connector"); return; }
            preview.Affected.Add(new AffectedRef("connector", $"{sheet.SheetName}!{Json.GetString(op,"name")}")); }
        finally { RotHelper.ReleaseComReference(shape); }
    }
    private static void ApplyCreateConnector(object workbook, JsonObject op, ApplyExecution execution, List<string> mismatches, ref int checkedCells)
    {
        using var sheet = BindSheet(workbook, op); EnsureSheetWritable(sheet.Sheet); object? shapes = null; object? connector = null;
        try
        {
            var begin = ExcelConnectorApply.Parse(Json.GetObj(op, "begin")!);
            var end = ExcelConnectorApply.Parse(Json.GetObj(op, "end")!);
            var requestedName = Json.GetString(op, "name");
            EnsureConnectorNameAvailable(sheet.Sheet, requestedName);
            ExcelConnectorApply.ValidateBeforeMutation(sheet.Sheet, requestedName, begin, end, FindShape, shape => ReadComString(shape, "Name"));

            shapes = (object)((dynamic)sheet.Sheet).Shapes;
            var type = Json.GetString(op,"connectorType") == "elbow" ? 2 : Json.GetString(op,"connectorType") == "curve" ? 3 : 1;
            var coordinates = ExcelConnectorApply.CreationCoordinates(begin, end);
            connector = (object)((dynamic)shapes).AddConnector(type, (float)coordinates.Begin.X, (float)coordinates.Begin.Y, (float)coordinates.End.X, (float)coordinates.End.Y);
            if (!string.IsNullOrWhiteSpace(requestedName)) ((dynamic)connector).Name = requestedName;
            ExcelConnectorApply.ApplyEndpoints(sheet.Sheet, connector, begin, end, FindShape, shape => ReadComString(shape, "Name"));
            ApplyShapeFormatting(connector, op); checkedCells++;
            VerifyConnectorRequested(ReadShapeState(connector, sheet.SheetName), op, mismatches, $"{sheet.SheetName}!{((dynamic)connector).Name}");
            execution.Affected.Add(new AffectedRef("connector", $"{sheet.SheetName}!{((dynamic)connector).Name}"));
        }
        finally { RotHelper.ReleaseComReference(connector); RotHelper.ReleaseComReference(shapes); }
    }
    private static void ApplyUpdateConnector(object workbook, JsonObject op, ApplyExecution execution, List<string> mismatches, ref int checkedCells)
    {
        using var sheet = BindSheet(workbook, op); EnsureSheetWritable(sheet.Sheet); var connector = FindShape(sheet.Sheet, Json.GetString(op,"name")!);
        try { if (connector is null) throw new InvalidOperationException("[EXCEL_CONNECTOR_NOT_FOUND] connector was not found");
            if (((dynamic)connector).Connector != ExcelDataObjectCatalog.MsoTrue) throw new InvalidOperationException("[EXCEL_CONNECTOR_TYPE] named shape is not a connector");
            var hasBegin = op.ContainsKey("begin");
            var hasEnd = op.ContainsKey("end");
            var begin = hasBegin ? ExcelConnectorApply.Parse(Json.GetObj(op, "begin")!) : (ExcelConnectorApply.EndpointRequest?)null;
            var end = hasEnd ? ExcelConnectorApply.Parse(Json.GetObj(op, "end")!) : (ExcelConnectorApply.EndpointRequest?)null;
            if (hasBegin || hasEnd)
                ExcelConnectorApply.ApplyEndpoints(sheet.Sheet, connector, begin, end, FindShape, shape => ReadComString(shape, "Name"));
            ApplyShapeFormatting(connector, op); checkedCells++;
            VerifyConnectorRequested(ReadShapeState(connector, sheet.SheetName), op, mismatches, $"{sheet.SheetName}!{Json.GetString(op,"name")}");
            execution.Affected.Add(new AffectedRef("connector", $"{sheet.SheetName}!{Json.GetString(op,"name")}")); }
        finally { RotHelper.ReleaseComReference(connector); }
    }
    private static void VerifyConnectorRequested(JsonObject actual, JsonObject op, List<string> mismatches, string label)
    {
        if (op.ContainsKey("connectorType") && !string.Equals(Json.GetString(actual,"connectorType"), Json.GetString(op,"connectorType"), StringComparison.OrdinalIgnoreCase)) mismatches.Add($"{label}: connectorType readback mismatch");
        foreach (var side in new[] { "begin", "end" })
        {
            if (Json.GetObj(op, side) is not JsonObject requested) continue;
            var got = Json.GetObj(actual, side);
            if (Json.GetString(requested,"shape") is string target)
            {
                if (got is null || got["connected"] is not JsonValue connected || !connected.TryGetValue<bool>(out var isConnected) || !isConnected ||
                    !string.Equals(Json.GetString(got,"shape"),target,StringComparison.Ordinal) || Json.GetInt(got,"site") != Json.GetInt(requested,"site"))
                    mismatches.Add($"{label}: {side} target/site readback mismatch");
            }
            else
            {
                var x = ExcelShapeFormatContract.ReadFiniteNumber(requested["x"]!);
                var y = ExcelShapeFormatContract.ReadFiniteNumber(requested["y"]!);
                if (got is null || got["connected"] is not JsonValue connected || !connected.TryGetValue<bool>(out var isConnected) || isConnected)
                    mismatches.Add($"{label}: {side} should be detached");
                else if (!ExcelDataOperationsContract.TryGetFiniteNumber(got["x"], out var actualX) ||
                         !ExcelDataOperationsContract.TryGetFiniteNumber(got["y"], out var actualY))
                    mismatches.Add($"{label}: {side} free coordinate readback unavailable");
                else if (Math.Abs(actualX - x) > ExcelConnectorGeometry.Epsilon || Math.Abs(actualY - y) > ExcelConnectorGeometry.Epsilon)
                    mismatches.Add($"{label}: {side} free coordinate readback mismatch");
            }
        }
        ExcelShapeFormatContract.CompareRequestedReadback(actual, op, label, mismatches);
    }

    // Retained as the narrow endpoint hook for callers compiled against this partial adapter.
    private static void ApplyConnectorEndpoint(object sheet, object connector, JsonObject? endpoint, bool begin)
    {
        if (endpoint is null) return;
        var request = ExcelConnectorApply.Parse(endpoint);
        ExcelConnectorApply.ApplyEndpoints(sheet, connector, begin ? request : null, begin ? null : request,
            FindShape, shape => ReadComString(shape, "Name"));
    }

    private static void EnsureConnectorNameAvailable(object sheet, string? requestedName)
    {
        if (string.IsNullOrWhiteSpace(requestedName)) return;
        var existing = FindShape(sheet, requestedName);
        try
        {
            if (existing is not null)
                throw new InvalidOperationException($"[EXCEL_CONNECTOR_NAME_DUPLICATE] shape '{requestedName}' already exists");
        }
        finally { RotHelper.ReleaseComReference(existing); }
    }
}
