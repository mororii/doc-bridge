using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private static void ReadConnectorState(object shape, JsonObject state)
    {
        object? format = null;
        try
        {
            format = (object)((dynamic)shape).ConnectorFormat;
            state["connectorType"] = ConnectorTypeToken(TryComInt(format, "Type"));
            ReadConnectorEnd(format, state, "begin", "BeginConnected", "BeginConnectedShape", "BeginConnectionSite");
            ReadConnectorEnd(format, state, "end", "EndConnected", "EndConnectedShape", "EndConnectionSite");
            ReadConnectorGeometry(shape, state);
        }
        catch { state["connectorUnreadable"] = true; }
        finally { RotHelper.ReleaseComReference(format); }
    }

    private static string? ConnectorTypeToken(int? value) => value switch { 1 => "straight", 2 => "elbow", 3 => "curve", _ => value is null ? null : $"raw:{value}" };

    private static void ReadConnectorEnd(object format, JsonObject state, string name, string connectedProperty, string shapeProperty, string siteProperty)
    {
        var end = new JsonObject();
        int connected;
        try { connected = Convert.ToInt32(connectedProperty == "BeginConnected" ? (object)((dynamic)format).BeginConnected : (object)((dynamic)format).EndConnected, CultureInfo.InvariantCulture); }
        catch
        {
            end["connected"] = null;
            end["connectedState"] = "unreadable";
            end["connectedUnreadable"] = true;
            state[name] = end;
            return;
        }

        if (connected is -1 or 1)
        {
            end["connected"] = true;
            end["connectedState"] = "connected";
            object? target = null;
            try
            {
                target = shapeProperty == "BeginConnectedShape" ? (object)((dynamic)format).BeginConnectedShape : (object)((dynamic)format).EndConnectedShape;
                end["shape"] = ReadComString(target, "Name");
                end["site"] = Js(ReadConnectorInt(format, siteProperty));
            }
            catch { end["targetUnreadable"] = true; }
            finally { RotHelper.ReleaseComReference(target); }
        }
        else if (connected == 0)
        {
            end["connected"] = false;
            end["connectedState"] = "disconnected";
        }
        else
        {
            end["connected"] = null;
            end["connectedState"] = "mixed";
            end["connectedMixed"] = true;
        }
        state[name] = end;
    }

    private static void ReadConnectorGeometry(object shape, JsonObject state)
    {
        if (!ExcelDataOperationsContract.TryGetFiniteNumber(state["left"], out var left) ||
            !ExcelDataOperationsContract.TryGetFiniteNumber(state["top"], out var top) ||
            !ExcelDataOperationsContract.TryGetFiniteNumber(state["width"], out var width) ||
            !ExcelDataOperationsContract.TryGetFiniteNumber(state["height"], out var height) ||
            !ExcelDataOperationsContract.TryGetFiniteNumber(state["rotation"], out var rotation))
        {
            state["connectorGeometryUnavailable"] = true;
            return;
        }

        try
        {
            var horizontalFlip = ReadConnectorFlip((object)((dynamic)shape).HorizontalFlip);
            var verticalFlip = ReadConnectorFlip((object)((dynamic)shape).VerticalFlip);
            state["horizontalFlip"] = horizontalFlip;
            state["verticalFlip"] = verticalFlip;
            var box = new ExcelConnectorGeometry.Box(left, top, width, height, horizontalFlip, verticalFlip, rotation);
            if (!ExcelConnectorGeometry.TryGetEndpoints(box, out var begin, out var end))
            {
                state["connectorGeometryUnavailable"] = true;
                if (!ExcelConnectorGeometry.IsZeroRotation(rotation)) state["connectorGeometryRotationUnsupported"] = true;
                return;
            }
            if (Json.GetObj(state, "begin") is JsonObject beginEnd) { beginEnd["x"] = begin.X; beginEnd["y"] = begin.Y; }
            if (Json.GetObj(state, "end") is JsonObject endEnd) { endEnd["x"] = end.X; endEnd["y"] = end.Y; }
        }
        catch { state["connectorGeometryUnavailable"] = true; }
    }

    private static int? ReadConnectorInt(object format, string property)
    {
        try
        {
            return Convert.ToInt32(property == "BeginConnectionSite" ? (object)((dynamic)format).BeginConnectionSite : (object)((dynamic)format).EndConnectionSite,
                CultureInfo.InvariantCulture);
        }
        catch { return null; }
    }

    private static bool ReadConnectorFlip(object raw)
    {
        if (raw is bool boolValue) return boolValue;
        var numericValue = Convert.ToInt32(raw, CultureInfo.InvariantCulture);
        if (numericValue is not (0 or -1 or 1)) throw new InvalidOperationException("Flip is mixed.");
        return numericValue != 0;
    }
}
