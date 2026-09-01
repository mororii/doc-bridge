using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

namespace DocBridge.Core.Services;

/// <summary>
/// Out-of-process lifetime guard for an Excel instance created by DocBridge. Managed shutdown
/// hooks cannot run after TerminateProcess/taskkill /F, so the guard acquires its own exact-PID
/// COM reference while the owner is alive. If the owner disappears without signalling release,
/// it calls Quit only when every workbook is saved (or there are no workbooks).
/// </summary>
public static class ExcelOwnerWatchdog
{
    public const string ModeArgument = "--excel-owner-watchdog";

    public sealed class Lease : IDisposable
    {
        private readonly EventWaitHandle _release;
        private readonly Process _process;
        private int _disposed;

        internal Lease(EventWaitHandle release, Process process)
        {
            _release = release;
            _process = process;
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
            try { _release.Set(); } catch { }
            try { _process.WaitForExit(5000); } catch { }
            _release.Dispose();
            _process.Dispose();
        }
    }

    public static Lease? Start(int excelProcessId)
    {
        if (excelProcessId <= 0) return null;
        var executable = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(executable)) return null;
        var hostName = Path.GetFileNameWithoutExtension(executable);
        if (!hostName.Equals("doc-bridge-mcp", StringComparison.OrdinalIgnoreCase) &&
            !hostName.Equals("doc-bridge-cli", StringComparison.OrdinalIgnoreCase))
            return null;

        var token = Guid.NewGuid().ToString("N");
        var readyName = $"Local\\DocBridge.ExcelWatchdog.Ready.{token}";
        var releaseName = $"Local\\DocBridge.ExcelWatchdog.Release.{token}";
        using var ready = new EventWaitHandle(false, EventResetMode.AutoReset, readyName);
        var release = new EventWaitHandle(false, EventResetMode.ManualReset, releaseName);
        Process? process = null;
        try
        {
            var start = new ProcessStartInfo
            {
                FileName = executable,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            };
            start.ArgumentList.Add(ModeArgument);
            start.ArgumentList.Add(Environment.ProcessId.ToString(CultureInfo.InvariantCulture));
            start.ArgumentList.Add(excelProcessId.ToString(CultureInfo.InvariantCulture));
            start.ArgumentList.Add(readyName);
            start.ArgumentList.Add(releaseName);
            process = Process.Start(start);
            if (process is null || !ready.WaitOne(TimeSpan.FromSeconds(5)))
            {
                try { release.Set(); } catch { }
                try { process?.WaitForExit(2000); } catch { }
                process?.Dispose();
                release.Dispose();
                return null;
            }
            return new Lease(release, process);
        }
        catch
        {
            try { release.Set(); } catch { }
            process?.Dispose();
            release.Dispose();
            return null;
        }
    }

    public static bool TryRun(string[] args, out int exitCode)
    {
        exitCode = 0;
        if (args.Length == 0 || !string.Equals(args[0], ModeArgument, StringComparison.Ordinal))
            return false;
        if (args.Length != 5 ||
            !int.TryParse(args[1], NumberStyles.None, CultureInfo.InvariantCulture, out var parentProcessId) ||
            !int.TryParse(args[2], NumberStyles.None, CultureInfo.InvariantCulture, out var excelProcessId))
        {
            exitCode = 2;
            return true;
        }
        exitCode = RunOnSta(parentProcessId, excelProcessId, args[3], args[4]);
        return true;
    }

    private static int RunOnSta(int parentProcessId, int excelProcessId, string readyName, string releaseName)
    {
        var result = 2;
        var thread = new Thread(() => result = Run(parentProcessId, excelProcessId, readyName, releaseName))
        {
            IsBackground = false,
            Name = "doc-bridge-excel-owner-watchdog",
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        return result;
    }

    private static int Run(int parentProcessId, int excelProcessId, string readyName, string releaseName)
    {
        EventWaitHandle? ready = null;
        EventWaitHandle? release = null;
        object? application = null;
        var quitAttempted = false;
        var result = 5;
        try
        {
            try { ready = EventWaitHandle.OpenExisting(readyName); }
            catch (WaitHandleCannotBeOpenedException) { }
            try { release = EventWaitHandle.OpenExisting(releaseName); }
            catch (WaitHandleCannotBeOpenedException) { }

            var deadline = DateTime.UtcNow.AddSeconds(5);
            do
            {
                application = FindApplicationByProcessId(excelProcessId);
                if (application is not null) break;
                Thread.Sleep(100);
            } while (DateTime.UtcNow < deadline);
            if (application is null)
            {
                result = 3;
            }
            else
            {
                ready?.Set();
                var releasedNormally = false;
                while (ParentIsAlive(parentProcessId))
                {
                    if (release?.WaitOne(250) == true)
                    {
                        releasedNormally = true;
                        break;
                    }
                    Thread.Sleep(50);
                }

                if (releasedNormally)
                {
                    result = 0;
                }
                else
                {
                    // Owner vanished without the normal release signal. Never suppress prompts
                    // and never close an unsaved workbook. A clean owned instance can exit.
                    quitAttempted = TryQuitSafely(application);
                    result = quitAttempted ? 0 : 4;
                }
            }
        }
        catch { result = 5; }
        finally
        {
            RotHelper.ReleaseComObject(application);
            release?.Dispose();
            ready?.Dispose();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        if (quitAttempted && !WaitForProcessExit(excelProcessId, TimeSpan.FromSeconds(15)))
            return 6;
        return result;
    }

    private static object? FindApplicationByProcessId(int processId)
    {
        object? selected = null;
        foreach (var candidate in RotHelper.GetExcelApplications())
        {
            var keep = false;
            try
            {
                var hwnd = Convert.ToInt64(((dynamic)candidate).Hwnd, CultureInfo.InvariantCulture);
                if (RotHelper.ProcessIdFromWindowHandle(hwnd) != processId) continue;
                selected = candidate;
                keep = true;
                break;
            }
            catch { }
            finally
            {
                if (!keep) RotHelper.ReleaseComObject(candidate);
            }
        }
        return selected;
    }

    private static bool TryQuitSafely(object application)
    {
        object? workbooks = null;
        try
        {
            workbooks = (object)((dynamic)application).Workbooks;
            var count = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? workbook = null;
                try
                {
                    workbook = (object)((dynamic)workbooks).Item(index);
                    if (!Convert.ToBoolean(((dynamic)workbook).Saved, CultureInfo.InvariantCulture))
                        return false;
                }
                catch { return false; }
                finally { RotHelper.ReleaseComObject(workbook); }
            }
            ((dynamic)application).Quit();
            return true;
        }
        finally { RotHelper.ReleaseComObject(workbooks); }
    }

    private static bool ParentIsAlive(int processId)
    {
        if (processId <= 0) return false;
        try
        {
            using var process = Process.GetProcessById(processId);
            return !process.HasExited;
        }
        catch { return false; }
    }

    private static bool WaitForProcessExit(int processId, TimeSpan timeout)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return process.WaitForExit((int)timeout.TotalMilliseconds);
        }
        catch { return true; }
    }
}
