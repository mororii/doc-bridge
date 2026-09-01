param([Parameter(Mandatory=$true)][int]$ExpectedProcessId)
$ErrorActionPreference = 'Stop'
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class DocBridgeE2EWindowProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

$application = $null
$workbooks = $null
try {
    $application = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    [uint32]$actualProcessId = 0
    [void][DocBridgeE2EWindowProcess]::GetWindowThreadProcessId(
        [IntPtr]([long]$application.Hwnd), [ref]$actualProcessId)
    if ([int]$actualProcessId -ne $ExpectedProcessId) {
        throw "Refusing cleanup: expected Excel PID $ExpectedProcessId but COM returned PID $actualProcessId"
    }

    $testRoot = [IO.Path]::GetFullPath((Join-Path ([IO.Path]::GetTempPath()) 'docbridge-test-'))
    $workbooks = $application.Workbooks
    for ($index = [int]$workbooks.Count; $index -ge 1; $index--) {
        $workbook = $null
        try {
            $workbook = $workbooks.Item($index)
            $fullName = [IO.Path]::GetFullPath([string]$workbook.FullName)
            if (-not $fullName.StartsWith($testRoot, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Refusing cleanup: workbook is not a DocBridge E2E artifact: $fullName"
            }
            $workbook.Close($false)
        }
        finally {
            if ($null -ne $workbook) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
        }
    }
    $application.Quit()
}
finally {
    if ($null -ne $workbooks) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbooks) }
    if ($null -ne $application) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
