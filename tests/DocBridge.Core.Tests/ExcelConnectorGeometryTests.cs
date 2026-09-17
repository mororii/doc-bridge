using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public sealed class ExcelConnectorGeometryTests
{
    [Fact]
    public void Reverse_horizontal_and_vertical_flips_define_actual_endpoints()
    {
        var horizontal = new ExcelConnectorGeometry.Box(10, 20, 30, 40, HorizontalFlip: true, VerticalFlip: false);
        Assert.True(ExcelConnectorGeometry.TryGetEndpoints(horizontal, out var horizontalBegin, out var horizontalEnd));
        Assert.Equal(new ExcelConnectorGeometry.Point(40, 20), horizontalBegin);
        Assert.Equal(new ExcelConnectorGeometry.Point(10, 60), horizontalEnd);

        var vertical = new ExcelConnectorGeometry.Box(10, 20, 30, 40, HorizontalFlip: false, VerticalFlip: true);
        Assert.True(ExcelConnectorGeometry.TryGetEndpoints(vertical, out var verticalBegin, out var verticalEnd));
        Assert.Equal(new ExcelConnectorGeometry.Point(10, 60), verticalBegin);
        Assert.Equal(new ExcelConnectorGeometry.Point(40, 20), verticalEnd);
    }

    [Fact]
    public void Zero_width_or_height_retains_prior_axis_flip_and_has_exact_endpoints()
    {
        var vertical = ExcelConnectorGeometry.PlanBox(new(15, 70), new(15, 20), horizontalFlip: true, verticalFlip: false);
        Assert.True(vertical.HorizontalFlip);
        Assert.True(vertical.VerticalFlip);
        Assert.Equal(0, vertical.Width);
        Assert.True(ExcelConnectorGeometry.TryGetEndpoints(vertical, out var begin, out var end));
        Assert.Equal(new ExcelConnectorGeometry.Point(15, 70), begin);
        Assert.Equal(new ExcelConnectorGeometry.Point(15, 20), end);

        var horizontal = ExcelConnectorGeometry.PlanBox(new(70, 15), new(20, 15), horizontalFlip: false, verticalFlip: true);
        Assert.True(horizontal.HorizontalFlip);
        Assert.True(horizontal.VerticalFlip);
        Assert.Equal(0, horizontal.Height);
        Assert.True(ExcelConnectorGeometry.TryGetEndpoints(horizontal, out begin, out end));
        Assert.Equal(new ExcelConnectorGeometry.Point(70, 15), begin);
        Assert.Equal(new ExcelConnectorGeometry.Point(20, 15), end);
    }

    [Fact]
    public void Moving_one_endpoint_preserves_the_omitted_endpoint_and_uses_nonnegative_box()
    {
        var original = new ExcelConnectorGeometry.Box(10, 20, 30, 40, HorizontalFlip: false, VerticalFlip: false);
        Assert.True(ExcelConnectorGeometry.TryPlanEndpoint(original, begin: true, new ExcelConnectorGeometry.Point(100, 200), out var planned));
        Assert.True(ExcelConnectorGeometry.TryGetEndpoints(planned, out var begin, out var end));
        Assert.Equal(new ExcelConnectorGeometry.Point(100, 200), begin);
        Assert.Equal(new ExcelConnectorGeometry.Point(40, 60), end);
        Assert.True(planned.Width >= 0);
        Assert.True(planned.Height >= 0);
        Assert.True(planned.HorizontalFlip);
        Assert.True(planned.VerticalFlip);
    }

    [Fact]
    public void Apply_uses_connector_format_not_shape_for_connect()
    {
        var target = new FakeShape("target", connectionSiteCount: 2);
        var connector = new FakeConnector("connector");
        var request = new ExcelConnectorApply.EndpointRequest("target", 2, default);

        ExcelConnectorApply.ApplyEndpoints(new object(), connector, request, null,
            (_, name) => name == "target" ? target : null, shape => ((FakeShape)shape).Name);

        Assert.Equal(1, connector.Format.BeginConnectCalls);
        Assert.Same(target, connector.Format.BeginTarget);
        Assert.Equal(2, connector.Format.BeginSite);
        Assert.Equal(0, connector.ShapeConnectCalls);
    }

    [Fact]
    public void Apply_free_begin_preserves_and_reapplies_omitted_named_end()
    {
        var target = new FakeShape("target", connectionSiteCount: 1);
        var connector = new FakeConnector("connector")
        {
            Left = 10, Top = 20, Width = 30, Height = 40,
        };
        connector.Format.EndConnected = -1;
        connector.Format.EndConnectedShape = target;
        connector.Format.EndConnectionSite = 1;

        ExcelConnectorApply.ApplyEndpoints(new object(), connector,
            new ExcelConnectorApply.EndpointRequest(null, 0, new ExcelConnectorGeometry.Point(100, 200)), null,
            (_, name) => name == "target" ? target : null, shape => ((FakeShape)shape).Name);

        var box = new ExcelConnectorGeometry.Box(connector.Left, connector.Top, connector.Width, connector.Height,
            connector.HorizontalFlip != 0, connector.VerticalFlip != 0, connector.Rotation);
        Assert.True(ExcelConnectorGeometry.TryGetEndpoints(box, out var begin, out var end));
        Assert.Equal(new ExcelConnectorGeometry.Point(100, 200), begin);
        Assert.Equal(new ExcelConnectorGeometry.Point(40, 60), end);
        Assert.Equal(1, connector.Format.BeginDisconnectCalls);
        Assert.Equal(1, connector.Format.EndConnectCalls);
        Assert.Same(target, connector.Format.EndTarget);
    }

    [Fact]
    public void Apply_validates_both_named_endpoints_before_any_format_mutation()
    {
        var beginTarget = new FakeShape("begin", connectionSiteCount: 1);
        var invalidEndTarget = new FakeShape("end", connectionSiteCount: 1);
        var connector = new FakeConnector("connector");

        var error = Assert.Throws<InvalidOperationException>(() => ExcelConnectorApply.ApplyEndpoints(new object(), connector,
            new ExcelConnectorApply.EndpointRequest("begin", 1, default),
            new ExcelConnectorApply.EndpointRequest("end", 2, default),
            (_, name) => name == "begin" ? beginTarget : invalidEndTarget, shape => ((FakeShape)shape).Name));

        Assert.Contains("[EXCEL_CONNECTOR_SITE]", error.Message);
        Assert.Equal(0, connector.Format.BeginConnectCalls + connector.Format.EndConnectCalls + connector.Format.BeginDisconnectCalls + connector.Format.EndDisconnectCalls);
    }

    public class FakeShape
    {
        public FakeShape(string name, int connectionSiteCount) { Name = name; ConnectionSiteCount = connectionSiteCount; }
        public string Name { get; }
        public int ConnectionSiteCount { get; }
    }

    public sealed class FakeConnector : FakeShape
    {
        public FakeConnector(string name) : base(name, 2) { Format = new FakeConnectorFormat(); }
        public FakeConnectorFormat ConnectorFormat => Format;
        public FakeConnectorFormat Format { get; }
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; } = 1;
        public double Height { get; set; } = 1;
        public int HorizontalFlip { get; set; }
        public int VerticalFlip { get; set; }
        public double Rotation { get; set; }
        public int ShapeConnectCalls { get; private set; }
        public void BeginConnect(object _, int __) { ShapeConnectCalls++; throw new Xunit.Sdk.XunitException("Shape must not own BeginConnect."); }
        public void EndConnect(object _, int __) { ShapeConnectCalls++; throw new Xunit.Sdk.XunitException("Shape must not own EndConnect."); }
        public void BeginDisconnect() { ShapeConnectCalls++; throw new Xunit.Sdk.XunitException("Shape must not own BeginDisconnect."); }
        public void EndDisconnect() { ShapeConnectCalls++; throw new Xunit.Sdk.XunitException("Shape must not own EndDisconnect."); }
        public void Flip(int direction)
        {
            if (direction == 0) HorizontalFlip = HorizontalFlip == 0 ? -1 : 0;
            else if (direction == 1) VerticalFlip = VerticalFlip == 0 ? -1 : 0;
            else throw new ArgumentOutOfRangeException(nameof(direction));
        }
    }

    public sealed class FakeConnectorFormat
    {
        public int BeginConnected { get; set; }
        public int EndConnected { get; set; }
        public FakeShape? BeginConnectedShape { get; set; }
        public FakeShape? EndConnectedShape { get; set; }
        public int BeginConnectionSite { get; set; }
        public int EndConnectionSite { get; set; }
        public int BeginConnectCalls { get; private set; }
        public int EndConnectCalls { get; private set; }
        public int BeginDisconnectCalls { get; private set; }
        public int EndDisconnectCalls { get; private set; }
        public FakeShape? BeginTarget { get; private set; }
        public FakeShape? EndTarget { get; private set; }
        public int BeginSite { get; private set; }
        public int EndSite { get; private set; }
        public void BeginConnect(object target, int site) { var shape = Assert.IsType<FakeShape>(target); BeginConnectCalls++; BeginConnected = -1; BeginConnectedShape = BeginTarget = shape; BeginConnectionSite = BeginSite = site; }
        public void EndConnect(object target, int site) { var shape = Assert.IsType<FakeShape>(target); EndConnectCalls++; EndConnected = -1; EndConnectedShape = EndTarget = shape; EndConnectionSite = EndSite = site; }
        public void BeginDisconnect() { BeginDisconnectCalls++; BeginConnected = 0; BeginConnectedShape = null; BeginConnectionSite = 0; }
        public void EndDisconnect() { EndDisconnectCalls++; EndConnected = 0; EndConnectedShape = null; EndConnectionSite = 0; }
    }
}
