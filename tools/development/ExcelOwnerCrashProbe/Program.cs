using System.Globalization;
using System.Diagnostics;
using System.Text;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

if (args.Length == 1 && args[0] == "--list-excel")
{
    foreach (var candidate in RotHelper.GetExcelApplications())
    {
        object? workbooks = null;
        try
        {
            var hwnd = Convert.ToInt64(((dynamic)candidate).Hwnd, CultureInfo.InvariantCulture);
            var processId = RotHelper.ProcessIdFromWindowHandle(hwnd);
            workbooks = (object)((dynamic)candidate).Workbooks;
            var count = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
            var descriptions = new List<string>();
            for (var index = 1; index <= count; index++)
            {
                object? workbook = null;
                try
                {
                    workbook = (object)((dynamic)workbooks).Item(index);
                    descriptions.Add($"{((dynamic)workbook).Name}|saved={((dynamic)workbook).Saved}");
                }
                finally { RotHelper.ReleaseComObject(workbook); }
            }
            Console.WriteLine($"pid={processId}; hwnd={hwnd}; workbooks={count}; {string.Join("; ", descriptions)}");
        }
        catch (Exception ex) { Console.WriteLine($"unreadable: {ex.Message}"); }
        finally
        {
            RotHelper.ReleaseComObject(workbooks);
            RotHelper.ReleaseComObject(candidate);
        }
    }
    return 0;
}

if (args.Length == 1 && args[0] == "--probe-worker")
{
    object? workerApplication = null;
    using var workerAdapter = new ExcelAdapter(() =>
    {
        workerApplication = RotHelper.CreateInstance("Excel.Application");
        return workerApplication;
    }, appFactoryOwnsInstance: true);
    _ = workerAdapter.GetActiveContext();
    var workerHwnd = Convert.ToInt64(((dynamic)workerApplication!).Hwnd, CultureInfo.InvariantCulture);
    Console.WriteLine(RotHelper.ProcessIdFromWindowHandle(workerHwnd).ToString(CultureInfo.InvariantCulture));
    Console.Out.Flush();
    while (Console.In.ReadLine() is not null) { }
    return 0;
}

if (ExcelOwnerWatchdog.TryRun(args, out var watchdogExitCode))
    return watchdogExitCode;

if (args.Length == 2 && args[0] == "--pipe-owner")
{
    var start = new ProcessStartInfo
    {
        FileName = Environment.ProcessPath!,
        UseShellExecute = false,
        CreateNoWindow = true,
        RedirectStandardInput = true,
        RedirectStandardOutput = true,
    };
    start.ArgumentList.Add("--probe-worker");
    using var worker = Process.Start(start)!;
    var workerExcelPid = await worker.StandardOutput.ReadLineAsync();
    if (string.IsNullOrWhiteSpace(workerExcelPid)) return 5;
    File.WriteAllText(args[1], workerExcelPid);
    Thread.Sleep(Timeout.Infinite);
    return 0;
}

var graceful = args.Length == 2 && args[0] == "--graceful";
var pidFile = graceful ? args[1] : args.Length == 1 ? args[0] : null;
if (pidFile is null)
    return 2;

object? application = null;
ExcelOwnerWatchdog.Lease? watchdog = null;
try
{
    application = RotHelper.CreateInstance("Excel.Application");
    if (application is null) return 3;
    ((dynamic)application).Visible = true;
    var hwnd = Convert.ToInt64(((dynamic)application).Hwnd, CultureInfo.InvariantCulture);
    var excelProcessId = RotHelper.ProcessIdFromWindowHandle(hwnd);
    watchdog = ExcelOwnerWatchdog.Start(excelProcessId);
    if (watchdog is null) return 4;
    File.WriteAllText(pidFile, excelProcessId.ToString(CultureInfo.InvariantCulture));
    if (graceful)
    {
        ((dynamic)application).Quit();
        return 0;
    }
    Thread.Sleep(Timeout.Infinite);
    return 0;
}
finally
{
    watchdog?.Dispose();
    RotHelper.ReleaseComObject(application);
}
