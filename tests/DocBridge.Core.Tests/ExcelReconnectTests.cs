using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelReconnectTests
{
    [Fact]
    public void Worker_discovery_timeout_precedes_common_client_deadline()
    {
        Assert.True(ExcelWorkerAdapter.TimeoutForMethod("status") < TimeSpan.FromSeconds(60));
        Assert.True(ExcelWorkerAdapter.TimeoutForMethod("context") < TimeSpan.FromSeconds(60));
        Assert.True(ExcelWorkerAdapter.TimeoutForMethod("read") < TimeSpan.FromSeconds(60));
        Assert.Equal(TimeSpan.FromSeconds(150), ExcelWorkerAdapter.TimeoutForMethod("apply"));
    }

    [Fact]
    public void Adapter_reacquires_application_after_cached_COM_reference_disconnects()
    {
        var first = new FakeExcelApplication(101, "first.xlsx");
        var second = new FakeExcelApplication(202, "second.xlsx");
        var factoryCalls = 0;

        using var adapter = new ExcelAdapter(() => ++factoryCalls == 1 ? first : second);

        var before = adapter.GetStatus();
        Assert.True(before.Connected);
        Assert.Equal("first.xlsx", before.Document);
        Assert.Equal(1, factoryCalls);

        first.Disconnected = true;

        var after = adapter.GetStatus();
        Assert.True(after.Connected);
        Assert.Equal("second.xlsx", after.Document);
        Assert.Equal(2, factoryCalls);
    }

    [Theory]
    [InlineData(unchecked((int)0x80010108))] // RPC_E_DISCONNECTED
    [InlineData(unchecked((int)0x800706BA))] // RPC_S_SERVER_UNAVAILABLE
    [InlineData(unchecked((int)0x800401FD))] // CO_E_OBJNOTCONNECTED
    [InlineData(unchecked((int)0x80010007))] // RPC_E_SERVER_DIED
    [InlineData(unchecked((int)0x80010012))] // RPC_E_SERVER_DIED_DNE
    public void Adapter_reacquires_for_known_server_disconnect_hresult(int hresult)
    {
        var first = new FakeExcelApplication(101, "first.xlsx") { DisconnectHResult = hresult };
        var second = new FakeExcelApplication(202, "second.xlsx");
        var factoryCalls = 0;
        using var adapter = new ExcelAdapter(() => ++factoryCalls == 1 ? first : second);

        Assert.True(adapter.GetStatus().Connected);
        first.Disconnected = true;

        var status = adapter.GetStatus();
        Assert.True(status.Connected);
        Assert.Equal("second.xlsx", status.Document);
        Assert.Equal(2, factoryCalls);
    }

    [Fact]
    public void Existing_instance_is_never_quit_and_status_identifies_user_connection()
    {
        var application = new LifecycleExcelApplication(saved: true);
        using var adapter = new ExcelAdapter(() => application, appFactoryOwnsInstance: false);

        var status = adapter.GetStatus();
        Assert.True(status.Connected);
        Assert.Equal("사용자가 열어 둔 엑셀 창에 연결됨", status.Detail);

        var disconnected = adapter.Disconnect();
        Assert.True(disconnected["ok"]!.GetValue<bool>());
        Assert.False(disconnected["ownedInstance"]!.GetValue<bool>());
        Assert.Equal(0, application.QuitCalls);
    }

    [Fact]
    public void Owned_instance_is_visible_and_quit_when_all_workbooks_are_saved()
    {
        var application = new LifecycleExcelApplication(saved: true);
        using var adapter = new ExcelAdapter(() => application, appFactoryOwnsInstance: true);

        var status = adapter.GetStatus();
        Assert.True(application.Visible);
        Assert.Equal("DocBridge가 생성한 인스턴스", status.Detail);

        var disconnected = adapter.Disconnect();
        Assert.True(disconnected["quitCalled"]!.GetValue<bool>());
        Assert.Equal(1, application.QuitCalls);
    }

    [Fact]
    public void Owned_instance_with_unsaved_workbook_is_released_without_quit()
    {
        var application = new LifecycleExcelApplication(saved: false);
        using var adapter = new ExcelAdapter(() => application, appFactoryOwnsInstance: true);
        Assert.True(adapter.GetStatus().Connected);

        var disconnected = adapter.Disconnect();
        Assert.False(disconnected["quitCalled"]!.GetValue<bool>());
        Assert.Single(disconnected["unsavedWorkbooks"]!.AsArray());
        Assert.Equal(0, application.QuitCalls);
    }

    [Fact]
    public void Core_disconnect_routes_to_excel_lifecycle_without_quitting_user_instance()
    {
        using var home = new TestHome();
        using var host = new DocBridgeHost(home.Options);
        var application = new LifecycleExcelApplication(saved: true);
        var adapter = new ExcelAdapter(() => application, appFactoryOwnsInstance: false);
        host.Router.Register("excel", adapter);
        Assert.True(adapter.GetStatus().Connected);

        var result = host.CoreDisconnect(new JsonObject { ["app"] = "excel" });

        Assert.True(Json.GetBool(result, "ok"));
        Assert.True(Json.GetBool(result, "disconnected"));
        Assert.False(Json.GetBool(result, "ownedInstance"));
        Assert.Equal(0, application.QuitCalls);
    }

    [Fact]
    public void Active_context_failure_never_reports_partial_success()
    {
        using var adapter = new ExcelAdapter(() => new ContextFailureExcelApplication());

        var context = adapter.GetActiveContext();

        Assert.False(context.Ok);
        Assert.Contains(context.Errors, error => error.Contains("synthetic worksheet failure", StringComparison.Ordinal));
    }

    [Fact]
    public void Workbook_path_without_explicit_opt_in_never_calls_Workbooks_Open()
    {
        var path = Path.Combine(Path.GetTempPath(), $"docbridge-path-probe-{Guid.NewGuid():N}.xlsx");
        File.WriteAllBytes(path, new byte[] { 0x50, 0x4B, 0x03, 0x04 });
        try
        {
            var application = new EmptyOwnedExcelApplication(openThrows: true);
            using var adapter = new ExcelAdapter(() => application, appFactoryOwnsInstance: true);

            var result = adapter.Read(new JsonObject
            {
                ["scope"] = "scan",
                ["workbook"] = path,
            });

            Assert.False(Json.GetBool(result, "ok"));
            Assert.Equal(0, application.Workbooks.OpenCalls);
            // The injected owned fixture is cleaned as a zero-workbook shell.  Production does
            // not create it at all because AttachExcel(false) performs ROT discovery only.
            Assert.Equal(1, application.QuitCalls);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Explicit_file_open_failure_immediately_quits_owned_empty_instance()
    {
        var path = Path.Combine(Path.GetTempPath(), $"docbridge-open-failure-{Guid.NewGuid():N}.xlsx");
        File.WriteAllBytes(path, new byte[] { 0x50, 0x4B, 0x03, 0x04 });
        try
        {
            var application = new EmptyOwnedExcelApplication(openThrows: true);
            using var adapter = new ExcelAdapter(() => application, appFactoryOwnsInstance: true);

            var result = adapter.Read(new JsonObject
            {
                ["scope"] = "scan",
                ["workbook"] = path,
                ["allowOpenFile"] = true,
            });

            Assert.False(Json.GetBool(result, "ok"));
            Assert.Equal(1, application.Workbooks.OpenCalls);
            Assert.Equal(1, application.QuitCalls);
            Assert.True(application.Visible);
            Assert.False(Json.GetBool(adapter.Disconnect(), "disconnected"));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    [Trait("Category", "E2E")]
    public void Discovery_and_path_probe_do_not_launch_real_Excel()
    {
        if (!string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal))
            return;
        using var existing = new ProcessSet(Process.GetProcessesByName("EXCEL"));

        var path = Path.Combine(Path.GetTempPath(), $"docbridge-noncreating-probe-{Guid.NewGuid():N}.xlsx");
        File.WriteAllBytes(path, new byte[] { 0x50, 0x4B, 0x03, 0x04 });
        try
        {
            using var adapter = new ExcelAdapter();
            _ = adapter.GetStatus();
            _ = adapter.GetActiveContext();
            var probe = adapter.Read(new JsonObject
            {
                ["scope"] = "scan",
                ["workbook"] = path,
            });
            Assert.False(Json.GetBool(probe, "ok"));

            Thread.Sleep(500);
            using var after = new ProcessSet(Process.GetProcessesByName("EXCEL"));
            Assert.Empty(after.ProcessIds.Except(existing.ProcessIds));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    [Trait("Category", "E2E")]
    public void Adapter_reconnects_after_real_Excel_process_is_restarted()
    {
        if (!string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal))
            return;

        object? latestApplication = null;
        var factoryCalls = 0;
        using var adapter = new ExcelAdapter(() =>
        {
            var type = Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException("Excel is not installed.");
            dynamic application = Activator.CreateInstance(type)!;
            application.Visible = false;
            object? workbooks = null;
            object? workbook = null;
            object? sheet = null;
            object? range = null;
            try
            {
                workbooks = (object)application.Workbooks;
                workbook = (object)((dynamic)workbooks).Add();
                sheet = (object)((dynamic)workbook).ActiveSheet;
                range = (object)((dynamic)sheet).Range("A1");
                ((dynamic)range).Value2 = ++factoryCalls;
            }
            finally
            {
                DocBridge.Core.Services.RotHelper.ReleaseComObject(range);
                DocBridge.Core.Services.RotHelper.ReleaseComObject(sheet);
                DocBridge.Core.Services.RotHelper.ReleaseComObject(workbook);
                DocBridge.Core.Services.RotHelper.ReleaseComObject(workbooks);
            }
            latestApplication = application;
            return (object)application;
        }, appFactoryOwnsInstance: true);

        try
        {
            var first = adapter.GetStatus();
            Assert.True(first.Connected, first.Detail);
            Assert.Equal(1, factoryCalls);

            adapter.RunOnAdapterThread<object?>(() =>
            {
                dynamic application = latestApplication!;
                CloseAllTestWorkbooks(application);
                application.Quit();
                return null;
            });

            var second = adapter.GetStatus();
            Assert.True(second.Connected, second.Detail);
            Assert.Equal(2, factoryCalls);
        }
        finally
        {
            try
            {
                adapter.RunOnAdapterThread<object?>(() =>
                {
                    if (latestApplication is null) return null;
                    dynamic application = latestApplication;
                    try { CloseAllTestWorkbooks(application); } catch { }
                    return null;
                });
                latestApplication = null;
                _ = adapter.Disconnect();
            }
            catch { }
        }
    }

    private static void CloseAllTestWorkbooks(object application)
    {
        object? workbooks = null;
        try
        {
            workbooks = (object)((dynamic)application).Workbooks;
            for (var index = Convert.ToInt32(((dynamic)workbooks).Count); index >= 1; index--)
            {
                object? workbook = null;
                try
                {
                    workbook = (object)((dynamic)workbooks).Item(index);
                    ((dynamic)workbook).Close(false);
                }
                finally { DocBridge.Core.Services.RotHelper.ReleaseComObject(workbook); }
            }
        }
        finally { DocBridge.Core.Services.RotHelper.ReleaseComObject(workbooks); }
    }

    public sealed class FakeExcelApplication
    {
        private readonly long _hwnd;

        public FakeExcelApplication(long hwnd, string workbook)
        {
            _hwnd = hwnd;
            ActiveWorkbook = new FakeWorkbook(workbook);
        }

        public bool Disconnected { get; set; }
        public int DisconnectHResult { get; set; } = unchecked((int)0x80010108);
        public string Version => "16.0";
        public FakeWorkbook ActiveWorkbook { get; }

        public long Hwnd => Disconnected
            ? throw new COMException("The COM server is disconnected.", DisconnectHResult)
            : _hwnd;
    }

    public sealed class FakeWorkbook
    {
        public FakeWorkbook(string fullName) => FullName = fullName;
        public string FullName { get; }
    }

    public sealed class LifecycleExcelApplication
    {
        public LifecycleExcelApplication(bool saved)
        {
            ActiveWorkbook = new LifecycleWorkbook("lifecycle.xlsx", saved);
            Workbooks = new LifecycleWorkbooks(ActiveWorkbook);
        }

        public bool Visible { get; set; }
        public int QuitCalls { get; private set; }
        public long Hwnd => 303;
        public string Version => "16.0";
        public LifecycleWorkbook ActiveWorkbook { get; }
        public LifecycleWorkbooks Workbooks { get; }
        public void Quit() => QuitCalls++;
    }

    public sealed class LifecycleWorkbooks
    {
        private readonly LifecycleWorkbook _workbook;
        public LifecycleWorkbooks(LifecycleWorkbook workbook) => _workbook = workbook;
        public int Count => 1;
        public LifecycleWorkbook Item(int index) => index == 1 ? _workbook : throw new ArgumentOutOfRangeException(nameof(index));
    }

    public sealed class LifecycleWorkbook
    {
        public LifecycleWorkbook(string fullName, bool saved)
        {
            FullName = fullName;
            Name = Path.GetFileName(fullName);
            Saved = saved;
        }
        public string FullName { get; }
        public string Name { get; }
        public bool Saved { get; }
    }

    public sealed class ContextFailureExcelApplication
    {
        public ContextFailureWorkbook ActiveWorkbook { get; } = new();
    }

    public sealed class ContextFailureWorkbook
    {
        public string FullName => "context-failure.xlsx";
        public ContextFailureSheet ActiveSheet { get; } = new();
        public ContextFailureWorksheets Worksheets { get; } = new();
    }

    public sealed class ContextFailureSheet
    {
        public string Name => "Sheet1";
    }

    public sealed class ContextFailureWorksheets
    {
        public int Count => 1;
        public object Item(int index) => throw new InvalidOperationException("synthetic worksheet failure");
    }

    public sealed class EmptyOwnedExcelApplication
    {
        public EmptyOwnedExcelApplication(bool openThrows) => Workbooks = new EmptyOwnedWorkbooks(openThrows);
        public bool Visible { get; set; }
        public int QuitCalls { get; private set; }
        public long Hwnd => 404;
        public EmptyOwnedWorkbooks Workbooks { get; }
        public object? ActiveWorkbook => null;
        public void Quit() => QuitCalls++;
    }

    public sealed class EmptyOwnedWorkbooks
    {
        private readonly bool _openThrows;
        public EmptyOwnedWorkbooks(bool openThrows) => _openThrows = openThrows;
        public int Count => 0;
        public int OpenCalls { get; private set; }
        public object Open(string path, int updateLinks, bool readOnly)
        {
            OpenCalls++;
            if (_openThrows) throw new InvalidOperationException("synthetic Workbooks.Open failure");
            throw new NotSupportedException();
        }
        public object Item(int index) => throw new ArgumentOutOfRangeException(nameof(index));
    }

    private sealed class ProcessSet : IDisposable
    {
        private readonly Process[] _processes;
        public ProcessSet(Process[] processes)
        {
            _processes = processes;
            ProcessIds = processes.Select(process => process.Id).ToHashSet();
        }
        public HashSet<int> ProcessIds { get; }
        public void Dispose()
        {
            foreach (var process in _processes) process.Dispose();
        }
    }
}
