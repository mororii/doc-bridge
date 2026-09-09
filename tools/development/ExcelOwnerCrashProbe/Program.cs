using System.Diagnostics;
using System.Runtime.InteropServices;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace ExcelOwnerCrashProbe;

internal enum ProbeWorkbookMode
{
    /// <summary>CreateInstance only. Does not Add or SaveAs. Compares original empty-shell leftover semantics.</summary>
    Empty,

    /// <summary>Workbooks.Add a new owned book and leave it unsaved. Does not SaveAs startup/add-in/user books.</summary>
    UnsavedOwned,

    /// <summary>Workbooks.Add a new owned book and SaveAs only that book.</summary>
    SavedOwned,
}

internal static class Program
{
    private static readonly object StaLock = new();
    private static int _ownedExcelPid;
    private static long _ownedHwnd;

    public static int Main(string[] args)
    {
        if (ExcelOwnerWatchdog.TryRun(args, out var watchdogExit))
            return watchdogExit;

        if (args.Any(static a => a.Equals("--list-excel", StringComparison.OrdinalIgnoreCase)))
            return ListExcel();

        if (args.Any(static a => a.Equals("--probe-worker", StringComparison.OrdinalIgnoreCase)))
            return RunProbeWorker();

        var pipeOwner = ReadOption(args, "--pipe-owner");
        if (pipeOwner is not null)
            return RunPipeOwner(pipeOwner);

        var graceful = HasFlag(args, "--graceful");
        var crashEmpty = HasFlag(args, "--crash-empty");
        var crashUnsaved = HasFlag(args, "--crash-unsaved");
        var crashSaved = HasFlag(args, "--crash");
        var pidFile = args.FirstOrDefault(static a =>
            !a.StartsWith("--", StringComparison.Ordinal) && !a.Equals("excel-owner-crash-probe", StringComparison.OrdinalIgnoreCase));

        if (graceful)
            return RunAdapterOwned(ProbeWorkbookMode.SavedOwned, pidFile, disposeWhenReady: true);

        if (crashEmpty)
            return RunWatchdogHang(ProbeWorkbookMode.Empty, pidFile);

        if (crashUnsaved)
            return RunWatchdogHang(ProbeWorkbookMode.UnsavedOwned, pidFile);

        if (crashSaved || pidFile is not null)
            return RunWatchdogHang(ProbeWorkbookMode.SavedOwned, pidFile);

        Console.Error.WriteLine(
            "usage: ExcelOwnerCrashProbe --list-excel | --probe-worker | --pipe-owner <pidFile> | --graceful [pidFile] | --crash [pidFile] | --crash-empty [pidFile] | --crash-unsaved [pidFile]");
        return 2;
    }

    private static bool HasFlag(string[] args, string name) =>
        args.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase));

    private static string? ReadOption(string[] args, string name)
    {
        for (var i = 0; i < args.Length - 1; i++)
        {
            if (args[i].Equals(name, StringComparison.OrdinalIgnoreCase))
                return args[i + 1];
        }

        return null;
    }

    private static int ListExcel()
    {
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            using (process)
                Console.WriteLine($"{process.Id}\t{process.ProcessName}\t{process.MainWindowTitle}");
        }

        return 0;
    }

    /// <summary>
    /// Production-shaped worker: ExcelAdapter owns the factory instance. HWND/PID are recorded
    /// inside the creating STA factory before GetActiveContext continues. Host thread never
    /// touches the RCW. stdin EOF runs ExcelAdapter.Dispose (DisconnectExcelCore), not a raw Quit.
    /// </summary>
    private static int RunProbeWorker() =>
        RunAdapterOwned(ProbeWorkbookMode.SavedOwned, pidFile: null, disposeWhenReady: false);

    private static int RunAdapterOwned(ProbeWorkbookMode mode, string? pidFile, bool disposeWhenReady)
    {
        if (mode == ProbeWorkbookMode.Empty)
        {
            Console.Error.WriteLine(
                "empty adapter mode is unsupported: GetActiveContext finally Quits an empty owned shell. Use --crash-empty.");
            return 2;
        }

        ExcelAdapter? adapter = null;
        try
        {
            adapter = new ExcelAdapter(
                () => CreateOwnedExcelOnFactoryThread(mode),
                appFactoryOwnsInstance: true);

            try
            {
                var ctx = adapter.GetActiveContext();
                Phase(
                    "get-active-context",
                    $"ok={ctx.Ok} app={ctx.App} documentRef={ctx.DocumentRef} errors={string.Join(";", ctx.Errors)} warnings={string.Join(";", ctx.Warnings)}");
                if (!ctx.Ok)
                    Phase("get-active-context-fail", $"result-not-ok errors={string.Join(";", ctx.Errors)}");
            }
            catch (Exception ex)
            {
                Phase("get-active-context-fail", FormatException(ex));
                throw;
            }

            PublishOwnedPid(pidFile);
            if (disposeWhenReady)
                return 0;

            Console.In.ReadToEnd();
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
        finally
        {
            adapter?.Dispose();
        }
    }

    private static int RunWatchdogHang(ProbeWorkbookMode mode, string? pidFile)
    {
        Exception? fault = null;
        ExcelOwnerWatchdog.Lease? watchdog = null;
        var ready = new ManualResetEventSlim(false);
        var thread = new Thread(() =>
        {
            try
            {
                lock (StaLock)
                {
                    var application = CreateOwnedExcelOnFactoryThread(mode);
                    if (_ownedExcelPid <= 0)
                        throw new InvalidOperationException("owned Excel PID was not recorded inside the STA factory.");
                    watchdog = ExcelOwnerWatchdog.Start(_ownedExcelPid);
                    if (watchdog is null)
                        throw new InvalidOperationException("ExcelOwnerWatchdog.Start returned null (host exe name must be doc-bridge-cli).");
                    Phase("watchdog-started", $"pid={_ownedExcelPid}");
                    ready.Set();
                    WaitUntilKilled();
                    GC.KeepAlive(application);
                    GC.KeepAlive(watchdog);
                }
            }
            catch (Exception ex)
            {
                fault = ex;
                try { ready.Set(); } catch { }
            }
        })
        {
            IsBackground = false,
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        ready.Wait();
        if (fault is not null)
        {
            Console.Error.WriteLine(fault);
            return 1;
        }

        try
        {
            PublishOwnedPid(pidFile);
            thread.Join();
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static int RunPipeOwner(string pidFile)
    {
        try
        {
            return RunPipeOwnerCore(pidFile);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static int RunPipeOwnerCore(string pidFile)
    {
        var exe = Environment.ProcessPath ?? throw new InvalidOperationException("missing ProcessPath");
        using var child = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = exe,
                Arguments = "--probe-worker",
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            },
        };

        if (!child.Start())
            throw new InvalidOperationException("failed to start --probe-worker");

        var line = child.StandardOutput.ReadLine();
        if (string.IsNullOrWhiteSpace(line) || !int.TryParse(line.Trim(), out var excelPid) || excelPid <= 0)
        {
            Console.Error.WriteLine(child.StandardError.ReadToEnd());
            throw new InvalidOperationException($"--probe-worker did not publish a PID. stdout='{line}'");
        }

        File.WriteAllText(pidFile, excelPid.ToString());
        Console.WriteLine(excelPid);

        child.StandardInput.Close();
        if (!child.WaitForExit(20000))
            throw new TimeoutException("--probe-worker did not exit after stdin EOF / ExcelAdapter.Dispose");

        Console.Error.WriteLine(child.StandardError.ReadToEnd());
        return child.ExitCode;
    }

    /// <summary>
    /// Must run on the creating STA thread (adapter factory or parked crash STA).
    /// Records HWND/PID before Visible or any save. Never SaveAs preexisting books.
    /// Does not FinalRelease a reused preexisting Excel RCW.
    /// </summary>
    private static object CreateOwnedExcelOnFactoryThread(ProbeWorkbookMode mode)
    {
        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            throw new InvalidOperationException("Excel factory must run on STA.");

        var preexisting = CurrentExcelPids();
        Phase("preexisting-excel", preexisting.Count == 0 ? "(none)" : string.Join(",", preexisting.OrderBy(static id => id)));

        object? application = null;
        var provedNewOwned = false;
        try
        {
            application = RotHelper.CreateInstance("Excel.Application")
                ?? throw new InvalidOperationException("Excel.Application CreateInstance returned null.");
            dynamic app = application;
            _ownedHwnd = Convert.ToInt64(app.Hwnd, System.Globalization.CultureInfo.InvariantCulture);
            _ownedExcelPid = RotHelper.ProcessIdFromWindowHandle(_ownedHwnd);
            Phase("pid-from-hwnd", $"hwnd={_ownedHwnd} pid={_ownedExcelPid}");

            if (_ownedExcelPid <= 0)
                throw new InvalidOperationException("CreateInstance produced hwnd/pid=0; refusing to mutate this Excel.");
            if (preexisting.Contains(_ownedExcelPid))
                throw new InvalidOperationException(
                    $"CreateInstance reused preexisting Excel PID {_ownedExcelPid}; refusing Visible/SaveAs.");

            provedNewOwned = true;
            Phase("pid-proven-new", _ownedExcelPid.ToString());
            TraceWorkbooks(application, "after-create-before-prefs");

            app.Visible = true;
            Phase("visible-set", "true");

            if (mode != ProbeWorkbookMode.Empty)
                CreateAndSaveOnlyOwnedProbeWorkbook(application, save: mode == ProbeWorkbookMode.SavedOwned);
            else
                Phase("empty-mode", "no Workbooks.Add and no SaveAs");

            TraceWorkbooks(application, "ready");
            return application;
        }
        catch (Exception ex)
        {
            Phase("factory-fail", FormatException(ex));
            if (provedNewOwned)
                RotHelper.ReleaseComObject(application);
            _ownedExcelPid = 0;
            _ownedHwnd = 0;
            throw;
        }
    }

    private static string FormatException(Exception ex)
    {
        var parts = new List<string>();
        for (var current = ex; current is not null; current = current.InnerException)
        {
            var code = current is COMException com ? com.ErrorCode : current.HResult;
            parts.Add($"{current.GetType().FullName} HRESULT=0x{unchecked((uint)code):X8} {current.Message}");
        }

        return string.Join(" | ", parts);
    }

    private static void CreateAndSaveOnlyOwnedProbeWorkbook(object application, bool save)
    {
        object? workbooks = null;
        object? created = null;
        try
        {
            dynamic app = application;
            workbooks = app.Workbooks;
            created = ((dynamic)workbooks).Add();
            var name = ReadWorkbookName(created) ?? "(unnamed)";
            Phase("owned-workbook-added", name);

            if (!save)
            {
                Phase("owned-workbook-left-unsaved", name);
                return;
            }

            var path = Path.Combine(
                Path.GetTempPath(),
                $"docbridge-excel-probe-owned-{Guid.NewGuid():N}.xlsx");
            ((dynamic)created).SaveAs(path);
            Phase("owned-workbook-saved", path);
        }
        finally
        {
            RotHelper.ReleaseComObject(created);
            RotHelper.ReleaseComObject(workbooks);
        }
    }

    private static void TraceWorkbooks(object application, string stage)
    {
        object? workbooks = null;
        try
        {
            dynamic app = application;
            workbooks = app.Workbooks;
            var count = (int)((dynamic)workbooks).Count;
            Phase("workbook-count", $"{stage} count={count}");
            for (var i = 1; i <= count; i++)
            {
                object? book = null;
                try
                {
                    book = ((dynamic)workbooks).Item(i);
                    var name = ReadWorkbookName(book) ?? "?";
                    var saved = ReadSaved(book);
                    var fullName = ReadFullName(book) ?? "";
                    Phase(
                        "workbook-item",
                        $"{stage} index={i} name={name} saved={FormatSaved(saved)} path={fullName}");
                }
                finally
                {
                    RotHelper.ReleaseComObject(book);
                }
            }
        }
        catch (Exception ex)
        {
            Phase("workbook-trace-fail", $"{stage} {ex.GetType().Name}: {ex.Message}");
        }
        finally
        {
            RotHelper.ReleaseComObject(workbooks);
        }
    }

    private static string? ReadWorkbookName(object? workbook)
    {
        if (workbook is null) return null;
        try { return (string)((dynamic)workbook).Name; }
        catch { return null; }
    }

    private static string? ReadFullName(object? workbook)
    {
        if (workbook is null) return null;
        try { return (string)((dynamic)workbook).FullName; }
        catch { return null; }
    }

    private static bool? ReadSaved(object? workbook)
    {
        if (workbook is null) return null;
        try { return (bool)((dynamic)workbook).Saved; }
        catch { return null; }
    }

    private static string FormatSaved(bool? saved) => saved switch
    {
        true => "true",
        false => "false",
        null => "unknown",
    };

    private static void PublishOwnedPid(string? pidFile)
    {
        if (_ownedExcelPid <= 0)
            throw new InvalidOperationException("owned Excel PID is not available on the host thread.");
        Console.WriteLine(_ownedExcelPid);
        if (!string.IsNullOrWhiteSpace(pidFile))
            File.WriteAllText(pidFile, _ownedExcelPid.ToString());
        Phase("published-pid", _ownedExcelPid.ToString());
    }

    private static void Phase(string name, string detail)
    {
        Console.Error.WriteLine($"probe-phase {name} {detail}");
    }

    private static HashSet<int> CurrentExcelPids()
    {
        var ids = new HashSet<int>();
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            using (process)
                ids.Add(process.Id);
        }

        return ids;
    }

    private static void WaitUntilKilled()
    {
        Thread.Sleep(Timeout.Infinite);
    }
}
