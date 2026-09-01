param([Parameter(Mandatory=$true)][string]$PidFile)
$ErrorActionPreference = 'Stop'
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class DocBridgeOwnedExcelWindowProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

$application = New-Object -ComObject Excel.Application
$application.Visible = $true
[uint32]$excelProcessId = 0
[void][DocBridgeOwnedExcelWindowProcess]::GetWindowThreadProcessId(
    [IntPtr]([long]$application.Hwnd), [ref]$excelProcessId)
[IO.File]::WriteAllText($PidFile, [string]$excelProcessId)
while ($true) { Start-Sleep -Seconds 1 }
