namespace DocBridge.Core.Services;

/// <summary>Axis-aligned connector endpoint geometry. Excel represents line direction through flips, not negative dimensions.</summary>
public static class ExcelConnectorGeometry
{
    public const double Epsilon = 0.0001;

    public readonly record struct Point(double X, double Y);

    public readonly record struct Box(double Left, double Top, double Width, double Height,
        bool HorizontalFlip, bool VerticalFlip, double Rotation = 0);

    public static bool TryGetEndpoints(Box box, out Point begin, out Point end)
    {
        begin = default;
        end = default;
        if (!double.IsFinite(box.Left) || !double.IsFinite(box.Top) || !double.IsFinite(box.Width) ||
            !double.IsFinite(box.Height) || !double.IsFinite(box.Rotation) || !IsZeroRotation(box.Rotation))
            return false;

        var left = box.Left;
        var top = box.Top;
        var right = left + Math.Abs(box.Width);
        var bottom = top + Math.Abs(box.Height);
        begin = new Point(box.HorizontalFlip ? right : left, box.VerticalFlip ? bottom : top);
        end = new Point(box.HorizontalFlip ? left : right, box.VerticalFlip ? top : bottom);
        return true;
    }

    /// <summary>Plans a non-negative Excel bounding box and derives direction flips from the requested endpoints.
    /// A zero-length axis retains its prior flip because that axis has no observable direction.</summary>
    public static Box PlanBox(Point begin, Point end, bool horizontalFlip, bool verticalFlip, double rotation = 0)
    {
        if (!double.IsFinite(begin.X) || !double.IsFinite(begin.Y) || !double.IsFinite(end.X) || !double.IsFinite(end.Y))
            throw new ArgumentOutOfRangeException(nameof(begin), "Connector endpoints must be finite.");
        var horizontalDirection = Math.Abs(begin.X - end.X) <= Epsilon ? horizontalFlip : begin.X > end.X;
        var verticalDirection = Math.Abs(begin.Y - end.Y) <= Epsilon ? verticalFlip : begin.Y > end.Y;
        return new Box(Math.Min(begin.X, end.X), Math.Min(begin.Y, end.Y),
            Math.Abs(end.X - begin.X), Math.Abs(end.Y - begin.Y), horizontalDirection, verticalDirection, rotation);
    }

    /// <summary>Changes one endpoint while retaining the other endpoint; flips change when needed to retain direction.</summary>
    public static bool TryPlanEndpoint(Box current, bool begin, Point requested, out Box planned)
    {
        planned = default;
        if (!TryGetEndpoints(current, out var existingBegin, out var existingEnd) ||
            !double.IsFinite(requested.X) || !double.IsFinite(requested.Y))
            return false;
        planned = PlanBox(begin ? requested : existingBegin, begin ? existingEnd : requested,
            current.HorizontalFlip, current.VerticalFlip, current.Rotation);
        return true;
    }

    public static bool IsZeroRotation(double rotation)
    {
        var normalized = rotation % 360;
        if (normalized < 0) normalized += 360;
        return Math.Abs(normalized) <= Epsilon || Math.Abs(normalized - 360) <= Epsilon;
    }
}
