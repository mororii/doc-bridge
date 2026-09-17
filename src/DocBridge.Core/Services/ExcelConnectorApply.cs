using System.Globalization;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>Native Excel ConnectorFormat endpoint mutations. The ConnectorFormat, never Shape, owns connect/disconnect methods.</summary>
internal static class ExcelConnectorApply
{
    internal readonly record struct EndpointRequest(string? ShapeName, int Site, ExcelConnectorGeometry.Point Point)
    {
        public bool IsAttachment => !string.IsNullOrWhiteSpace(ShapeName);
    }

    private readonly record struct SavedEndpoint(ExcelConnectorGeometry.Point Point, string? ShapeName, int Site)
    {
        public bool IsAttachment => !string.IsNullOrWhiteSpace(ShapeName);
    }

    internal static EndpointRequest Parse(JsonObject endpoint)
    {
        var shape = Json.GetString(endpoint, "shape");
        if (!string.IsNullOrWhiteSpace(shape))
            return new EndpointRequest(shape, checked((int)ExcelShapeFormatContract.ReadFiniteNumber(endpoint["site"]!)), default);
        return new EndpointRequest(null, 0, new ExcelConnectorGeometry.Point(
            ExcelShapeFormatContract.ReadFiniteNumber(endpoint["x"]!),
            ExcelShapeFormatContract.ReadFiniteNumber(endpoint["y"]!)));
    }

    internal static (ExcelConnectorGeometry.Point Begin, ExcelConnectorGeometry.Point End) CreationCoordinates(
        EndpointRequest begin, EndpointRequest end)
    {
        var beginPoint = begin.IsAttachment ? new ExcelConnectorGeometry.Point(0, 0) : begin.Point;
        var endPoint = end.IsAttachment
            ? new ExcelConnectorGeometry.Point(beginPoint.X + 1, beginPoint.Y + 1)
            : end.Point;
        return (beginPoint, endPoint);
    }

    internal static void ValidateBeforeMutation(object sheet, string? connectorName,
        EndpointRequest? begin, EndpointRequest? end, Func<object, string, object?> findShape, Func<object, string?> readName)
    {
        ValidateEndpoint(sheet, connectorName, begin, findShape, readName);
        ValidateEndpoint(sheet, connectorName, end, findShape, readName);
    }

    internal static void ApplyEndpoints(object sheet, object connector, EndpointRequest? begin, EndpointRequest? end,
        Func<object, string, object?> findShape, Func<object, string?> readName)
    {
        ValidateBeforeMutation(sheet, readName(connector), begin, end, findShape, readName);

        object? format = null;
        try
        {
            format = (object)((dynamic)connector).ConnectorFormat;
            var movesFreeEndpoint = begin is { IsAttachment: false } || end is { IsAttachment: false };
            // Excel can adjust the connector while connecting either end. Preserve an omitted named end explicitly.
            var savedBegin = ReadSavedEndpoint(format, true, default);
            var savedEnd = ReadSavedEndpoint(format, false, default);
            if (movesFreeEndpoint)
            {
                var geometry = ReadGeometry(connector);
                savedBegin = savedBegin with { Point = geometry.Begin };
                savedEnd = savedEnd with { Point = geometry.End };

                // Disconnect only endpoints explicitly made free. The opposite named attachment is saved and reapplied below.
                if (begin is { IsAttachment: false }) ((dynamic)format).BeginDisconnect();
                if (end is { IsAttachment: false }) ((dynamic)format).EndDisconnect();

                var wantedBegin = begin is { IsAttachment: false } requestedFreeBegin ? requestedFreeBegin.Point : savedBegin.Point;
                var wantedEnd = end is { IsAttachment: false } requestedFreeEnd ? requestedFreeEnd.Point : savedEnd.Point;
                var planned = ExcelConnectorGeometry.PlanBox(wantedBegin, wantedEnd,
                    geometry.Box.HorizontalFlip, geometry.Box.VerticalFlip, geometry.Box.Rotation);
                ApplyFlips(connector, geometry.Box, planned);
                ((dynamic)connector).Left = planned.Left;
                ((dynamic)connector).Top = planned.Top;
                ((dynamic)connector).Width = planned.Width;
                ((dynamic)connector).Height = planned.Height;
            }

            ApplyAttachment(sheet, format, true, begin is { IsAttachment: true } requestedAttachmentBegin ? requestedAttachmentBegin :
                begin is null && savedBegin.IsAttachment
                    ? new EndpointRequest(savedBegin.ShapeName, savedBegin.Site, default) : null, findShape);
            ApplyAttachment(sheet, format, false, end is { IsAttachment: true } requestedAttachmentEnd ? requestedAttachmentEnd :
                end is null && savedEnd.IsAttachment
                    ? new EndpointRequest(savedEnd.ShapeName, savedEnd.Site, default) : null, findShape);
        }
        finally { RotHelper.ReleaseComReference(format); }
    }

    private static void ValidateEndpoint(object sheet, string? connectorName, EndpointRequest? endpoint,
        Func<object, string, object?> findShape, Func<object, string?> readName)
    {
        if (endpoint is not { IsAttachment: true } requested) return;
        if (requested.Site < 1)
            throw new InvalidOperationException("[EXCEL_CONNECTOR_SITE] connection site must be 1-based");
        if (!string.IsNullOrWhiteSpace(connectorName) &&
            string.Equals(requested.ShapeName, connectorName, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("[EXCEL_CONNECTOR_SELF] connector cannot attach to itself");

        var target = findShape(sheet, requested.ShapeName!);
        try
        {
            if (target is null)
                throw new InvalidOperationException($"[EXCEL_CONNECTOR_TARGET_NOT_FOUND] '{requested.ShapeName}' was not found");
            if (string.Equals(readName(target), connectorName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("[EXCEL_CONNECTOR_SELF] connector cannot attach to itself");
            var count = ReadConnectionSiteCount(target);
            if (requested.Site > count)
                throw new InvalidOperationException($"[EXCEL_CONNECTOR_SITE] '{requested.ShapeName}' has no connection site {requested.Site}");
        }
        finally { RotHelper.ReleaseComReference(target); }
    }

    private static int ReadConnectionSiteCount(object target)
    {
        try
        {
            var count = Convert.ToInt32((object)((dynamic)target).ConnectionSiteCount, CultureInfo.InvariantCulture);
            if (count < 1)
                throw new InvalidOperationException("[EXCEL_CONNECTOR_SITE] target has no readable ConnectionSiteCount");
            return count;
        }
        catch (InvalidOperationException) { throw; }
        catch (Exception ex) when (ex is FormatException or InvalidCastException or OverflowException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        {
            throw new InvalidOperationException("[EXCEL_CONNECTOR_SITE] target has no readable ConnectionSiteCount", ex);
        }
    }

    private static (ExcelConnectorGeometry.Box Box, ExcelConnectorGeometry.Point Begin, ExcelConnectorGeometry.Point End) ReadGeometry(object connector)
    {
        try
        {
            var box = new ExcelConnectorGeometry.Box(
                Convert.ToDouble((object)((dynamic)connector).Left, CultureInfo.InvariantCulture),
                Convert.ToDouble((object)((dynamic)connector).Top, CultureInfo.InvariantCulture),
                Convert.ToDouble((object)((dynamic)connector).Width, CultureInfo.InvariantCulture),
                Convert.ToDouble((object)((dynamic)connector).Height, CultureInfo.InvariantCulture),
                ReadFlip((object)((dynamic)connector).HorizontalFlip),
                ReadFlip((object)((dynamic)connector).VerticalFlip),
                Convert.ToDouble((object)((dynamic)connector).Rotation, CultureInfo.InvariantCulture));
            if (!ExcelConnectorGeometry.TryGetEndpoints(box, out var begin, out var end))
                throw new InvalidOperationException("[EXCEL_CONNECTOR_ROTATION_GEOMETRY] coordinate editing is unavailable for a rotated or unreadable connector");
            return (box, begin, end);
        }
        catch (InvalidOperationException) { throw; }
        catch (Exception ex) when (ex is FormatException or InvalidCastException or OverflowException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        {
            throw new InvalidOperationException("[EXCEL_CONNECTOR_GEOMETRY] connector coordinates or flips are unreadable", ex);
        }
    }

    private static SavedEndpoint ReadSavedEndpoint(object format, bool begin, ExcelConnectorGeometry.Point point)
    {
        try
        {
            var connected = Convert.ToInt32(begin ? (object)((dynamic)format).BeginConnected : (object)((dynamic)format).EndConnected,
                CultureInfo.InvariantCulture);
            if (connected is not (-1 or 1)) return new SavedEndpoint(point, null, 0);
            object? target = null;
            try
            {
                target = begin ? (object)((dynamic)format).BeginConnectedShape : (object)((dynamic)format).EndConnectedShape;
                var site = Convert.ToInt32(begin ? (object)((dynamic)format).BeginConnectionSite : (object)((dynamic)format).EndConnectionSite,
                    CultureInfo.InvariantCulture);
                var name = Convert.ToString((object)((dynamic)target).Name, CultureInfo.InvariantCulture);
                return string.IsNullOrWhiteSpace(name) || site < 1
                    ? new SavedEndpoint(point, null, 0)
                    : new SavedEndpoint(point, name, site);
            }
            finally { RotHelper.ReleaseComReference(target); }
        }
        catch { return new SavedEndpoint(point, null, 0); }
    }

    private static void ApplyAttachment(object sheet, object format, bool begin, EndpointRequest? request,
        Func<object, string, object?> findShape)
    {
        if (request is not { IsAttachment: true } attachment) return;
        var target = findShape(sheet, attachment.ShapeName!);
        try
        {
            if (target is null)
                throw new InvalidOperationException($"[EXCEL_CONNECTOR_TARGET_NOT_FOUND] '{attachment.ShapeName}' was not found");
            if (begin) ((dynamic)format).BeginConnect(target, attachment.Site);
            else ((dynamic)format).EndConnect(target, attachment.Site);
        }
        finally { RotHelper.ReleaseComReference(target); }
    }

    private static bool ReadFlip(object raw)
    {
        if (raw is bool value) return value;
        var flag = Convert.ToInt32(raw, CultureInfo.InvariantCulture);
        if (flag is not (0 or -1 or 1)) throw new InvalidOperationException("Connector flip is mixed or unreadable.");
        return flag != 0;
    }

    private static void ApplyFlips(object connector, ExcelConnectorGeometry.Box current, ExcelConnectorGeometry.Box planned)
    {
        // MsoFlipCmd: msoFlipHorizontal = 0, msoFlipVertical = 1. Flip before assigning the final box.
        if (current.HorizontalFlip != planned.HorizontalFlip) ((dynamic)connector).Flip(0);
        if (current.VerticalFlip != planned.VerticalFlip) ((dynamic)connector).Flip(1);
    }
}
