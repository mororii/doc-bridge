using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>Validation shared by native connector endpoints and PictureFormat crops.</summary>
public static class ExcelDrawingDetailsContract
{
    public static void ValidatePictureDetails(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        if (op.ContainsKey("path"))
        {
            var path = Json.GetString(op, "path");
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathFullyQualified(path) ||
                !ExcelDataObjectCatalog.PictureExtensions.Contains(Path.GetExtension(path)))
                errors.Add($"ops[{index}] '{opName}' path must be an absolute supported local image path");
        }
        if (!op.ContainsKey("crop")) return;
        if (Json.GetObj(op, "crop") is not JsonObject crop) { errors.Add($"ops[{index}] '{opName}' crop must be an object"); return; }
        foreach (var (side, value) in crop)
            if (side is not ("left" or "top" or "right" or "bottom") ||
                !ExcelDataOperationsContract.TryGetFiniteNumber(value, out var points) || points < 0 || points > 20_000)
                errors.Add($"ops[{index}] '{opName}' crop.{side} must be a finite 0..20000 point value");
    }

    public static void ValidateConnector(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        if (opName == ExcelDataOperationsContract.CreateConnector &&
            Json.GetString(op, "connectorType") is not ("straight" or "elbow" or "curve"))
            errors.Add($"ops[{index}] '{opName}' connectorType must be straight|elbow|curve");
        foreach (var field in new[] { "begin", "end" })
        {
            if (!op.ContainsKey(field)) continue;
            if (Json.GetObj(op, field) is not JsonObject end) { errors.Add($"ops[{index}] '{opName}' {field} must be an endpoint object"); continue; }
            var shape = Json.GetString(end, "shape");
            if (!string.IsNullOrWhiteSpace(shape))
            {
                if (!ExcelDataOperationsContract.IsValidObjectName(shape) || !ExcelDataOperationsContract.TryGetFiniteNumber(end["site"], out var site) || site < 1 || site != Math.Truncate(site))
                    errors.Add($"ops[{index}] '{opName}' {field} shape requires valid shape and 1-based integer site");
            }
            else if (!ExcelDataOperationsContract.TryGetFiniteNumber(end["x"], out var x) || !ExcelDataOperationsContract.TryGetFiniteNumber(end["y"], out var y) || x < 0 || y < 0 || x > 20_000 || y > 20_000)
                errors.Add($"ops[{index}] '{opName}' {field} requires x,y points or shape,site");
        }
    }
}
