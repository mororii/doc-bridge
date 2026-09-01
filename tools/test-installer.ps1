<#
.SYNOPSIS
  Runs install -> doctor -> uninstall against isolated fake user folders.
#>
[CmdletBinding()]
param(
    [string]$PackageRoot,
    [switch]$KeepSandbox
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($PackageRoot)) {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $manifest = Get-Content -LiteralPath (Join-Path $repoRoot '.codex-plugin\plugin.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $baseVersion = ([string]$manifest.version).Split('+')[0]
    $PackageRoot = Join-Path $repoRoot ("releases\DocBridge-$baseVersion-win-x64")
}
$PackageRoot = [System.IO.Path]::GetFullPath($PackageRoot).TrimEnd('\')
$sandboxParent = Join-Path (Split-Path -Parent $PSScriptRoot) '.installer-tests'
# Keep the isolated root short enough that the staged self-contained executable
# remains below legacy Win32 process-launch path limits on Windows PowerShell 5.1.
$testRoot = Join-Path $sandboxParent ('t-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$userRoot = Join-Path $testRoot 'User'
$appDataRoot = Join-Path $testRoot 'AppData'
$localAppDataRoot = Join-Path $testRoot 'LocalAppData'
$installRoot = Join-Path $localAppDataRoot 'DocBridge'
$claudeConfig = Join-Path $appDataRoot 'Claude\claude_desktop_config.json'
$kimiConfig = Join-Path $userRoot '.kimi\mcp.json'
$cursorConfig = Join-Path $userRoot '.cursor\mcp.json'
$cursorProjectConfig = Join-Path $testRoot 'Project\.cursor\mcp.json'
$runningProbeProcess = $null
$runningWorkerProcess = $null

function Write-Utf8NoBom([string]$Path, [string]$Text) {
    New-Item -ItemType Directory -Path (Split-Path -Parent $Path) -Force | Out-Null
    [System.IO.File]::WriteAllText($Path, $Text, (New-Object System.Text.UTF8Encoding($false)))
}

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
    Write-Host "  [OK] $Message" -ForegroundColor Green
}

try {
    if (-not (Test-Path -LiteralPath (Join-Path $PackageRoot 'Install-DocBridge.ps1'))) {
        throw "Package root is invalid: $PackageRoot"
    }
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $PackageRoot 'payload\codex-marketplace\plugins\doc-bridge\dist\clients'))) `
        'Packaged dist does not contain build-machine client configurations'
    New-Item -ItemType Directory -Path $testRoot -Force | Out-Null

    Write-Host "`n=== Running DocBridge process guard" -ForegroundColor Cyan
    $runningProbeDir = Join-Path $testRoot 'RunningProcessGuard'
    $runningProbeExe = Join-Path $runningProbeDir 'doc-bridge-mcp.exe'
    $runningWorkerExe = Join-Path $runningProbeDir 'doc-bridge-hwp-worker.exe'
    $runningGuardInstall = Join-Path $runningProbeDir 'LocalAppData\DocBridge'
    $runningGuardUser = Join-Path $runningProbeDir 'User'
    $runningGuardAppData = Join-Path $runningProbeDir 'AppData'
    New-Item -ItemType Directory -Path $runningProbeDir -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $env:WINDIR 'System32\PING.EXE') -Destination $runningProbeExe -Force
    Copy-Item -LiteralPath (Join-Path $env:WINDIR 'System32\PING.EXE') -Destination $runningWorkerExe -Force
    $runningProbeProcess = Start-Process -FilePath $runningProbeExe `
        -ArgumentList @('-t', '127.0.0.1') -WindowStyle Hidden -PassThru
    $runningWorkerProcess = Start-Process -FilePath $runningWorkerExe `
        -ArgumentList @('-t', '127.0.0.1') -WindowStyle Hidden -PassThru
    Start-Sleep -Milliseconds 300
    Assert-True (-not $runningProbeProcess.HasExited -and $runningProbeProcess.ProcessName -eq 'doc-bridge-mcp') `
        'A deterministic doc-bridge-mcp process is running for the installer guard test'
    Assert-True (-not $runningWorkerProcess.HasExited -and $runningWorkerProcess.ProcessName -eq 'doc-bridge-hwp-worker') `
        'A deterministic doc-bridge-hwp-worker process is running for the installer guard test'

    # This child process is expected to fail. With the outer ErrorActionPreference
    # set to Stop, Windows PowerShell promotes native stderr to NativeCommandError
    # before the assertions can inspect it, so temporarily capture it as data.
    $savedErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $guardOutput = (& powershell.exe -NoProfile -ExecutionPolicy Bypass `
            -File (Join-Path $PackageRoot 'Install-DocBridge.ps1') `
            -Clients Kimi `
            -InstallRoot $runningGuardInstall `
            -UserProfileRoot $runningGuardUser `
            -AppDataRoot $runningGuardAppData `
            -LocalAppDataRoot (Join-Path $runningProbeDir 'LocalAppData') `
            -SkipHwpSecurity `
            -SkipDoctor `
            -SkipClientCommands 2>&1 | Out-String)
        $guardExitCode = $LASTEXITCODE
    }
    finally { $ErrorActionPreference = $savedErrorActionPreference }
    # Native PowerShell error formatting can wrap a long executable path at the
    # console width. Compare whitespace-insensitive text so the assertion still
    # verifies the complete path rather than a shortened fragment.
    $guardOutputCompacted = $guardOutput -replace '\s', ''
    $runningProbeExeCompacted = $runningProbeExe -replace '\s', ''
    $runningWorkerExeCompacted = $runningWorkerExe -replace '\s', ''
    Assert-True ($guardExitCode -ne 0) 'Install is blocked while a DocBridge MCP process is running'
    Assert-True ($guardOutput -match ('PID=' + $runningProbeProcess.Id + '\b')) 'Running-process error includes the exact PID'
    Assert-True ($guardOutput -match ('PID=' + $runningWorkerProcess.Id + '\b')) 'Running-process error includes the exact worker PID'
    Assert-True ($guardOutputCompacted.IndexOf($runningProbeExeCompacted, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) `
        'Running-process error includes the executable path'
    Assert-True ($guardOutputCompacted.IndexOf($runningWorkerExeCompacted, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) `
        'Running-process error includes the worker executable path'
    Assert-True (-not (Test-Path -LiteralPath $runningGuardInstall)) `
        'Running-process guard fails before creating the installation root'

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PackageRoot 'Install-DocBridge.ps1') `
        -Clients Kimi `
        -InstallRoot $runningGuardInstall `
        -UserProfileRoot $runningGuardUser `
        -AppDataRoot $runningGuardAppData `
        -LocalAppDataRoot (Join-Path $runningProbeDir 'LocalAppData') `
        -SkipHwpSecurity `
        -SkipDoctor `
        -SkipClientCommands `
        -AllowRunningDocBridge
    Assert-True ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath (Join-Path $runningGuardInstall 'installation.json'))) `
        'Explicit AllowRunningDocBridge bypass is honored without stopping the running process'
    Assert-True (-not $runningProbeProcess.HasExited -and -not $runningWorkerProcess.HasExited) `
        'Installer bypass never terminates running MCP or worker processes'
    foreach ($probeProcess in @($runningProbeProcess, $runningWorkerProcess)) {
        Stop-Process -Id $probeProcess.Id -Force
        $probeProcess.WaitForExit(5000)
        $probeProcess.Dispose()
    }
    $runningProbeProcess = $null
    $runningWorkerProcess = $null

    Write-Host "`n=== Deployment CMD forwarding" -ForegroundColor Cyan
    foreach ($cmdName in @('1-INSTALL.cmd', '2-TEST.cmd', '2-EXCEL-LIVE-TEST.cmd', '3-UNINSTALL.cmd')) {
        $cmdText = Get-Content -LiteralPath (Join-Path $PackageRoot $cmdName) -Raw -Encoding UTF8
        Assert-True ($cmdText -notmatch '(?i)%LOCALAPPDATA%\\DocBridge') "$cmdName has no hard-coded default-root gate"
        Assert-True ($cmdText -match '(?i)exit\s+/b\s+%DOCBRIDGE_EXIT%') "$cmdName propagates the PowerShell exit code"
    }
    $excelLiveCmd = Get-Content -LiteralPath (Join-Path $PackageRoot '2-EXCEL-LIVE-TEST.cmd') -Raw -Encoding UTF8
    Assert-True ($excelLiveCmd -match '(?i)-RequireExcelRuntime') 'Excel live wrapper requires the explicit runtime check'
    $verifyMcp = Get-Content -LiteralPath (Join-Path $PackageRoot 'support\verify-mcp.ps1') -Raw -Encoding UTF8
    Assert-True ($verifyMcp -match 'if \(\$runExcelRuntime\)') 'Default MCP verification does not invoke Excel context unless explicitly requested'

    Write-Host "`n=== Claude Code ownership parser" -ForegroundColor Cyan
    Import-Module (Join-Path $PackageRoot 'support\DocBridge.Deployment.psm1') -Force
    $notFound = ConvertFrom-DocBridgeClaudeMcpOutput -ExitCode 1 -Output 'No MCP server found with name: doc-bridge'
    Assert-True ($notFound.Exists -eq $false -and $null -eq $notFound.Error) 'Claude legacy missing-entry wording is recognized'
    $notFoundUser = ConvertFrom-DocBridgeClaudeMcpOutput -ExitCode 1 -Output 'No user-scoped MCP server found with name: doc-bridge'
    Assert-True ($notFoundUser.Exists -eq $false -and $null -eq $notFoundUser.Error) 'Claude current user-scoped missing-entry wording is recognized'
    $claudeOwnedOutput = @'
doc-bridge:
  Scope: User config (available in all your projects)
  Status: connected
  Type: stdio
  Command: C:\DocBridge\doc-bridge-mcp.exe
  Args: --stdio
  Environment:
'@
    $claudeOwned = ConvertFrom-DocBridgeClaudeMcpOutput -ExitCode 0 -Output $claudeOwnedOutput
    Assert-True (Test-DocBridgeClaudeMcpStateMatches -State $claudeOwned -Command 'C:\DocBridge\doc-bridge-mcp.exe' -Arguments @('--stdio')) 'Claude ownership requires matching command, args, and empty environment'
    $claudeChangedArgs = ConvertFrom-DocBridgeClaudeMcpOutput -ExitCode 0 -Output ($claudeOwnedOutput.Replace('Args: --stdio', 'Args: --stdio --user-change'))
    Assert-True (-not (Test-DocBridgeClaudeMcpStateMatches -State $claudeChangedArgs -Command 'C:\DocBridge\doc-bridge-mcp.exe' -Arguments @('--stdio'))) 'Claude user-changed arguments are not treated as owned'
    $claudeChangedEnvironment = ConvertFrom-DocBridgeClaudeMcpOutput -ExitCode 0 -Output ($claudeOwnedOutput.Replace('Environment:', 'Environment: SECRET=value'))
    Assert-True (-not (Test-DocBridgeClaudeMcpStateMatches -State $claudeChangedEnvironment -Command 'C:\DocBridge\doc-bridge-mcp.exe' -Arguments @('--stdio'))) 'Claude non-empty environment is not treated as owned'
    $claudeUnknownEnvironment = ConvertFrom-DocBridgeClaudeMcpOutput -ExitCode 0 -Output ($claudeOwnedOutput -replace '(?m)^\s*Environment:\s*\r?\n?', '')
    Assert-True ($null -ne $claudeUnknownEnvironment.Error) 'Claude entry is preserved when args or environment cannot be parsed'

    Write-Host "`n=== Codex marketplace ownership" -ForegroundColor Cyan
    $foreignCodexUser = Join-Path $testRoot 'CodexForeign\User'
    $foreignCodexAppData = Join-Path $testRoot 'CodexForeign\AppData'
    $foreignCodexLocal = Join-Path $testRoot 'CodexForeign\LocalAppData'
    $foreignCodexInstall = Join-Path $foreignCodexLocal 'DocBridge'
    $foreignCodexConfig = Join-Path $foreignCodexUser '.codex\config.toml'
    $foreignMarketplaceSource = Join-Path $testRoot 'ForeignOwner\marketplace'
    Write-Utf8NoBom $foreignCodexConfig @"
[marketplaces.docbridge-local]
source_type = "local"
source = '$foreignMarketplaceSource'

[windows]
sandbox = "standard"
"@
    $foreignConfigHash = (Get-FileHash -LiteralPath $foreignCodexConfig -Algorithm SHA256).Hash
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PackageRoot 'Install-DocBridge.ps1') `
        -Clients Codex `
        -InstallRoot $foreignCodexInstall `
        -UserProfileRoot $foreignCodexUser `
        -AppDataRoot $foreignCodexAppData `
        -LocalAppDataRoot $foreignCodexLocal `
        -SkipHwpSecurity `
        -SkipDoctor `
        -SkipClientCommands `
        -AllowRunningDocBridge
    Assert-True ($LASTEXITCODE -ne 0) 'Install fails safely when docbridge-local belongs to a foreign source'
    Assert-True ((Get-FileHash -LiteralPath $foreignCodexConfig -Algorithm SHA256).Hash -eq $foreignConfigHash) 'Foreign Codex marketplace config is not modified'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $foreignCodexInstall 'installation.json'))) 'Foreign Codex marketplace is never recorded as managed'

    $changedCodexUser = Join-Path $testRoot 'CodexChanged\User'
    $changedCodexAppData = Join-Path $testRoot 'CodexChanged\AppData'
    $changedCodexLocal = Join-Path $testRoot 'CodexChanged\LocalAppData'
    $changedCodexInstall = Join-Path $changedCodexLocal 'DocBridge'
    $changedCodexHome = Join-Path $testRoot 'CodexChanged\CodexHome'
    $fakeCodexSource = Join-Path $testRoot 'CodexChanged\FakeCodex.cs'
    $fakeCodexExe = Join-Path $testRoot 'CodexChanged\codex.exe'
    Write-Utf8NoBom $fakeCodexSource @'
using System;
using System.IO;
using System.Text;

public static class FakeCodex
{
    public static int Main(string[] args)
    {
        if (args.Length == 1 && args[0] == "--version") {
            Console.WriteLine("codex-fake 1.0");
            return 0;
        }
        if (args.Length == 2 && args[0] == "plugin" && args[1] == "--help") {
            Console.WriteLine("  marketplace  Manage marketplaces");
            return 0;
        }
        if (args.Length >= 4 && args[0] == "plugin" && args[1] == "marketplace" && args[2] == "add") {
            string home = Environment.GetEnvironmentVariable("CODEX_HOME");
            Directory.CreateDirectory(home);
            string config = Path.Combine(home, "config.toml");
            string source = Path.GetFullPath(args[3]);
            string text = "[marketplaces.docbridge-local]\r\nsource_type = \"local\"\r\nsource = '" + source + "'\r\n";
            File.WriteAllText(config, text, new UTF8Encoding(false));
            return 0;
        }
        if (args.Length >= 2 && args[0] == "mcp" && args[1] == "list") {
            Console.WriteLine("doc-bridge enabled");
            return 0;
        }
        if (args.Length >= 4 && args[0] == "plugin" && args[1] == "marketplace" && args[2] == "remove") {
            string home = Environment.GetEnvironmentVariable("CODEX_HOME");
            string config = Path.Combine(home, "config.toml");
            if (File.Exists(config)) File.Delete(config);
            return 0;
        }
        Console.Error.WriteLine("unsupported fake codex command: " + string.Join(" ", args));
        return 2;
    }
}
'@
    $frameworkCsc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
    & $frameworkCsc /nologo /target:exe "/out:$fakeCodexExe" $fakeCodexSource
    Assert-True ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $fakeCodexExe -PathType Leaf)) 'Isolated fake Codex CLI is compiled for marketplace lifecycle testing'
    & (Join-Path $PackageRoot 'Install-DocBridge.ps1') `
        -Clients Codex `
        -InstallRoot $changedCodexInstall `
        -UserProfileRoot $changedCodexUser `
        -AppDataRoot $changedCodexAppData `
        -LocalAppDataRoot $changedCodexLocal `
        -CodexCliPath $fakeCodexExe `
        -CodexHomeRoot $changedCodexHome `
        -SkipHwpSecurity `
        -SkipDoctor `
        -AllowRunningDocBridge
    $changedRecordPath = Join-Path $changedCodexInstall 'installation.json'
    $changedRecord = Get-Content -LiteralPath $changedRecordPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $expectedChangedSource = Join-Path $changedCodexInstall 'codex-marketplace'
    Assert-True ($changedRecord.codexMarketplaceOwnership.sourceType -eq 'local' -and
        $changedRecord.codexMarketplaceOwnership.normalizedSource -eq $expectedChangedSource) 'Successful Codex registration records normalized marketplace ownership'
    $changedCodexConfig = Join-Path $changedCodexHome 'config.toml'
    $userChangedSource = Join-Path $testRoot 'UserChangedOwner\marketplace'
    Write-Utf8NoBom $changedCodexConfig @"
[marketplaces.docbridge-local]
source_type = "local"
source = '$userChangedSource'
"@
    $changedConfigHash = (Get-FileHash -LiteralPath $changedCodexConfig -Algorithm SHA256).Hash
    & (Join-Path $PackageRoot 'Uninstall-DocBridge.ps1') `
        -Clients Codex `
        -InstallRoot $changedCodexInstall `
        -UserProfileRoot $changedCodexUser `
        -AppDataRoot $changedCodexAppData `
        -LocalAppDataRoot $changedCodexLocal `
        -CodexCliPath $fakeCodexExe `
        -CodexHomeRoot $changedCodexHome
    Assert-True ((Get-FileHash -LiteralPath $changedCodexConfig -Algorithm SHA256).Hash -eq $changedConfigHash) 'Uninstall preserves a Codex marketplace source changed after installation'
    $changedReport = Get-Content -LiteralPath (Join-Path $changedCodexInstall 'uninstall-report.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True (($changedReport.warnings -join ' ') -match 'source was changed or is foreign') 'Changed Codex marketplace source produces an explicit ownership warning'

    Write-Host "`n=== Exact checksum manifest rejection" -ForegroundColor Cyan
    $integrityTool = Join-Path $PackageRoot 'support\Test-PackageIntegrity.ps1'
    $integrityFixture = Join-Path $testRoot 'integrity-fixture'
    foreach ($relative in @(
        '0-VERIFY.cmd',
        '1-INSTALL.cmd',
        '2-TEST.cmd',
        '2-EXCEL-LIVE-TEST.cmd',
        '3-UNINSTALL.cmd',
        'Install-DocBridge.ps1',
        'Test-DocBridge.ps1',
        'Uninstall-DocBridge.ps1',
        'support\verify-mcp.ps1'
    )) {
        $fixtureTarget = Join-Path $integrityFixture $relative
        New-Item -ItemType Directory -Path (Split-Path -Parent $fixtureTarget) -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $PackageRoot $relative) -Destination $fixtureTarget -Force
    }
    foreach ($relative in @(
        'payload\codex-marketplace\plugins\doc-bridge\clients\cursor\mcp.example.json',
        'payload\codex-marketplace\plugins\doc-bridge\clients\cursor\CURSOR_USAGE.md',
        'payload\codex-marketplace\plugins\doc-bridge\clients\cursor\docbridge-user-rule.txt',
        'payload\codex-marketplace\plugins\doc-bridge\clients\cursor\rules\docbridge-safe-automation.mdc'
    )) {
        Write-Utf8NoBom (Join-Path $integrityFixture $relative) "fixture: $relative"
    }
    $fixtureHashLines = @(Get-ChildItem -LiteralPath $integrityFixture -Recurse -File | Sort-Object FullName | ForEach-Object {
        $relative = $_.FullName.Substring($integrityFixture.Length + 1).Replace('\', '/')
        "{0}  {1}" -f (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant(), $relative
    })
    Write-Utf8NoBom (Join-Path $integrityFixture 'SHA256SUMS.txt') ($fixtureHashLines -join "`r`n")
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $integrityTool -PackageRoot $integrityFixture
    Assert-True ($LASTEXITCODE -eq 0) 'Exact checksum manifest is accepted'
    Write-Utf8NoBom (Join-Path $integrityFixture 'unlisted.txt') 'not covered'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $integrityTool -PackageRoot $integrityFixture
    Assert-True ($LASTEXITCODE -eq 1) 'An unlisted package file is rejected'
    Remove-Item -LiteralPath (Join-Path $integrityFixture 'unlisted.txt') -Force
    Write-Utf8NoBom (Join-Path $integrityFixture 'SHA256SUMS.txt') (($fixtureHashLines + $fixtureHashLines[0]) -join "`r`n")
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $integrityTool -PackageRoot $integrityFixture
    Assert-True ($LASTEXITCODE -eq 1) 'A duplicate checksum path is rejected'

    Write-Host "`n=== Text encoding compatibility" -ForegroundColor Cyan
    $encodingTool = Join-Path $PackageRoot 'support\Convert-DocBridgeTextEncoding.ps1'
    Assert-True (Test-Path -LiteralPath $encodingTool -PathType Leaf) 'Encoding compatibility tool is included'
    $basUtf8 = Join-Path $testRoot 'sample-utf8.bas'
    $basCp949 = Join-Path $testRoot 'sample-cp949.bas'
    $basText = "Attribute VB_Name = `"안전점검`"`nSub 점검()`n    MsgBox `"한글 모듈 정상`"`nEnd Sub`n"
    [System.IO.File]::WriteAllText($basUtf8, $basText, (New-Object System.Text.UTF8Encoding($false)))
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $encodingTool `
        -Mode ConvertBasToCp949 -Path $basUtf8 -OutputPath $basCp949
    Assert-True ($LASTEXITCODE -eq 0) 'UTF-8 VBA module converts to CP949/CRLF'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $encodingTool -Mode CheckBas -Path $basCp949
    Assert-True ($LASTEXITCODE -eq 0) 'Converted VBA module passes CP949 validation'
    $basBytes = [System.IO.File]::ReadAllBytes($basCp949)
    Assert-True (-not ($basBytes.Length -ge 3 -and $basBytes[0] -eq 0xEF -and $basBytes[1] -eq 0xBB -and $basBytes[2] -eq 0xBF)) 'Converted VBA module has no UTF-8 BOM'
    $decodedBas = [System.Text.Encoding]::GetEncoding(949).GetString($basBytes)
    Assert-True ($decodedBas -eq ($basText -replace '\r\n|\r|\n', "`r`n")) 'Korean VBA text round-trips without loss'

    Write-Utf8NoBom $claudeConfig @'
{
  "theme": "dark",
  "mcpServers": {
    "doc-bridge": { "command": "user-owned-claude.exe", "args": ["--custom"] },
    "existing-server": { "command": "existing.exe", "args": ["--keep"] }
  }
}
'@
    Write-Utf8NoBom $kimiConfig @'
{
  "disabledMcpServers": ["disabled-example"],
  "mcpServers": {
    "existing-server": { "command": "existing.exe", "args": [] }
  }
}
'@
    Write-Utf8NoBom $cursorConfig @'
{
  "telemetry": { "enabled": false },
  "mcpServers": {
    "existing-server": { "command": "existing.exe", "args": ["--keep-cursor"] }
  }
}
'@
    Write-Utf8NoBom $cursorProjectConfig @'
{
  "projectSetting": "must-not-change",
  "mcpServers": {
    "project-only": { "command": "project.exe", "args": [] }
  }
}
'@
    $cursorProjectHashBefore = (Get-FileHash -LiteralPath $cursorProjectConfig -Algorithm SHA256).Hash

    Write-Host "`n=== Doctor before install" -ForegroundColor Cyan
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PackageRoot 'Test-DocBridge.ps1') `
        -ClientsCsv 'Codex,ClaudeDesktop,Kimi,Cursor' `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -SkipHwpSecurity `
        -SkipHwpRuntimeDoctor `
        -SkipClientCommands
    Assert-True ($LASTEXITCODE -eq 2) 'Doctor clearly reports not-installed state'
    Assert-True (-not (Test-Path -LiteralPath $installRoot)) 'Doctor does not create a fake installation folder'

    Write-Host "`n=== Isolated install" -ForegroundColor Cyan
    & (Join-Path $PackageRoot 'Install-DocBridge.ps1') `
        -Clients @('Codex', 'ClaudeDesktop', 'Kimi', 'Cursor') `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -CodexCliPath $fakeCodexExe `
        -SkipHwpSecurity `
        -SkipHwpRuntimeDoctor `
        -SkipClientCommands `
        -AllowRunningDocBridge

    $claude = Get-Content -LiteralPath $claudeConfig -Raw -Encoding UTF8 | ConvertFrom-Json
    $kimi = Get-Content -LiteralPath $kimiConfig -Raw -Encoding UTF8 | ConvertFrom-Json
    $cursor = Get-Content -LiteralPath $cursorConfig -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ($claude.theme -eq 'dark') 'Claude unrelated top-level setting is preserved'
    Assert-True ($null -ne $claude.mcpServers.PSObject.Properties['existing-server']) 'Claude unrelated MCP is preserved'
    Assert-True ($null -ne $claude.mcpServers.PSObject.Properties['doc-bridge']) 'Claude doc-bridge MCP is added'
    Assert-True ($kimi.disabledMcpServers[0] -eq 'disabled-example') 'Kimi unrelated top-level setting is preserved'
    Assert-True ($null -ne $kimi.mcpServers.PSObject.Properties['existing-server']) 'Kimi unrelated MCP is preserved'
    Assert-True ($null -ne $kimi.mcpServers.PSObject.Properties['doc-bridge']) 'Kimi doc-bridge MCP is added'
    Assert-True (-not [bool]$cursor.telemetry.enabled) 'Cursor unrelated nested setting is preserved'
    Assert-True ($null -ne $cursor.mcpServers.PSObject.Properties['existing-server']) 'Cursor unrelated MCP is preserved'
    Assert-True ($null -ne $cursor.mcpServers.PSObject.Properties['doc-bridge']) 'Cursor doc-bridge MCP is added'
    Assert-True ($cursor.mcpServers.'doc-bridge'.args.Count -eq 1 -and $cursor.mcpServers.'doc-bridge'.args[0] -eq '--stdio') 'Cursor doc-bridge uses stdio arguments'
    Assert-True ((Get-FileHash -LiteralPath $cursorProjectConfig -Algorithm SHA256).Hash -eq $cursorProjectHashBefore) 'Cursor project-level MCP config is not modified'
    Assert-True (Test-Path -LiteralPath (Join-Path $installRoot 'codex-marketplace\plugins\doc-bridge\dist\coreclr.dll')) 'Self-contained runtime is installed'
    $installedMcp = Join-Path $installRoot 'codex-marketplace\plugins\doc-bridge\dist\doc-bridge-mcp.exe'
    foreach ($generatedClient in @('mcp.json', 'claude_desktop_config.json', 'claude-code.mcp.json', 'kimi-mcp.json', 'cursor-mcp.json')) {
        $generatedConfig = Get-Content -LiteralPath (Join-Path $installRoot ('generated-configs\' + $generatedClient)) -Raw -Encoding UTF8 | ConvertFrom-Json
        Assert-True ($generatedConfig.mcpServers.'doc-bridge'.command -eq $installedMcp) "Generated $generatedClient uses the target PC installation path"
    }
    $generatedCodex = Get-Content -LiteralPath (Join-Path $installRoot 'generated-configs\codex-config.toml') -Raw -Encoding UTF8
    Assert-True ($generatedCodex -match [regex]::Escape($installedMcp.Replace('\', '/'))) 'Generated Codex TOML uses the target PC installation path'
    Assert-True (Test-Path -LiteralPath (Join-Path $installRoot 'generated-configs\install-codex.txt')) 'Codex manual registration commands are generated'
    $installedMarketplace = Get-Content -LiteralPath (Join-Path $installRoot 'codex-marketplace\.agents\plugins\marketplace.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ($installedMarketplace.plugins[0].policy.installation -eq 'INSTALLED_BY_DEFAULT') 'Current Codex installs DocBridge by marketplace policy'
    $codexManual = Get-Content -LiteralPath (Join-Path $installRoot 'generated-configs\install-codex.txt') -Raw -Encoding UTF8
    Assert-True ($codexManual -notmatch '(?m)^.*\splugin\s+add\s') 'Generated current Codex instructions do not use removed plugin add command'
    Assert-True (Test-Path -LiteralPath (Join-Path $installRoot 'codex-marketplace\plugins\doc-bridge\skills\hwp-production-workflows\SKILL.md')) 'HWP production skill is installed'
    Assert-True (Test-Path -LiteralPath (Join-Path $installRoot 'codex-marketplace\plugins\doc-bridge\skills\cad-production-workflows\SKILL.md')) 'CAD production skill is installed'
    Assert-True (Test-Path -LiteralPath (Join-Path $installRoot 'generated-configs\cursor\rules\docbridge-safe-automation.mdc')) 'Cursor project rule template is installed'
    Assert-True (Test-Path -LiteralPath (Join-Path $installRoot 'generated-configs\cursor\docbridge-user-rule.txt')) 'Cursor user rule template is installed'
    Assert-True (Test-Path -LiteralPath (Join-Path $installRoot 'generated-configs\cursor\CURSOR_USAGE.md')) 'Cursor usage guide is installed'
    $ownership = Get-Content -LiteralPath (Join-Path $installRoot 'installation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ([bool]$ownership.clientMcpProvenance.ClaudeDesktop.previousEntryExisted) 'Existing same-name Claude Desktop entry provenance is recorded'
    Assert-True ($ownership.clientMcpProvenance.ClaudeDesktop.previousEntry.command -eq 'user-owned-claude.exe') 'Existing same-name entry can be restored'
    Assert-True ((Get-ChildItem -LiteralPath (Join-Path $installRoot 'backups') -File).Count -ge 3) 'Original client configs are backed up'

    Write-Host "`n=== Isolated doctor" -ForegroundColor Cyan
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PackageRoot 'Test-DocBridge.ps1') `
        -ClientsCsv 'Codex,ClaudeDesktop,Kimi,Cursor' `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -SkipHwpSecurity `
        -SkipHwpRuntimeDoctor `
        -SkipClientCommands
    Assert-True ($LASTEXITCODE -eq 0) 'Doctor reports no critical failures'

    Write-Host "`n=== Partial Cursor uninstall" -ForegroundColor Cyan
    & (Join-Path $PackageRoot 'Uninstall-DocBridge.ps1') `
        -Clients Cursor `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -RemoveHwpSecurity `
        -SkipClientCommands
    $partialCursor = Get-Content -LiteralPath $cursorConfig -Raw -Encoding UTF8 | ConvertFrom-Json
    $partialClaude = Get-Content -LiteralPath $claudeConfig -Raw -Encoding UTF8 | ConvertFrom-Json
    $partialKimi = Get-Content -LiteralPath $kimiConfig -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ($null -eq $partialCursor.mcpServers.PSObject.Properties['doc-bridge']) 'Partial uninstall removes only the Cursor registration'
    Assert-True ($partialClaude.mcpServers.'doc-bridge'.command -eq $installedMcp) 'Partial Cursor uninstall leaves Claude Desktop registration active'
    Assert-True ($partialKimi.mcpServers.'doc-bridge'.command -eq $installedMcp) 'Partial Cursor uninstall leaves the managed Kimi registration active'
    Assert-True (Test-Path -LiteralPath $installedMcp -PathType Leaf) 'Partial uninstall retains the shared MCP executable'
    $partialOwnership = Get-Content -LiteralPath (Join-Path $installRoot 'installation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ($partialOwnership.installedClients -notcontains 'Cursor') 'Partial uninstall removes Cursor from managed clients'
    Assert-True ($partialOwnership.installedClients -contains 'ClaudeDesktop') 'Partial uninstall retains other managed clients'
    Assert-True ($null -eq $partialOwnership.clientMcpProvenance.PSObject.Properties['Cursor']) 'Partial uninstall removes only Cursor ownership provenance'
    $partialReport = Get-Content -LiteralPath (Join-Path $installRoot 'uninstall-report.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True (($partialReport.warnings -join ' ') -match 'HWP security registration was retained') 'Partial uninstall retains shared HWP security registration'

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PackageRoot 'Test-DocBridge.ps1') `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -SkipHwpSecurity `
        -SkipHwpRuntimeDoctor `
        -SkipClientCommands
    Assert-True ($LASTEXITCODE -eq 0) 'Default doctor follows remaining clients after partial uninstall'

    $kimiChanged = Get-Content -LiteralPath $kimiConfig -Raw -Encoding UTF8 | ConvertFrom-Json
    $kimiChanged.mcpServers.'doc-bridge' = [pscustomobject]@{ command = $installedMcp; args = @('--mine') }
    Write-Utf8NoBom $kimiConfig ($kimiChanged | ConvertTo-Json -Depth 20)
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PackageRoot 'Test-DocBridge.ps1') `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -SkipHwpSecurity `
        -SkipHwpRuntimeDoctor `
        -SkipClientCommands
    Assert-True ($LASTEXITCODE -eq 1) 'Doctor rejects a user-changed MCP argument even when the command still matches'

    Write-Host "`n=== Re-add Cursor and default final uninstall" -ForegroundColor Cyan
    & (Join-Path $PackageRoot 'Install-DocBridge.ps1') `
        -Clients Cursor `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -SkipHwpSecurity `
        -SkipHwpRuntimeDoctor `
        -SkipClientCommands `
        -AllowRunningDocBridge
    & (Join-Path $PackageRoot 'Uninstall-DocBridge.ps1') `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -SkipClientCommands

    $claudeAfter = Get-Content -LiteralPath $claudeConfig -Raw -Encoding UTF8 | ConvertFrom-Json
    $kimiAfter = Get-Content -LiteralPath $kimiConfig -Raw -Encoding UTF8 | ConvertFrom-Json
    $cursorAfter = Get-Content -LiteralPath $cursorConfig -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ($claudeAfter.mcpServers.'doc-bridge'.command -eq 'user-owned-claude.exe') 'Final uninstall restores a pre-existing same-name Claude Desktop entry'
    Assert-True ($null -ne $claudeAfter.mcpServers.PSObject.Properties['existing-server']) 'Claude unrelated MCP remains after uninstall'
    Assert-True ($claudeAfter.theme -eq 'dark') 'Claude unrelated setting remains after uninstall'
    Assert-True ($kimiAfter.mcpServers.'doc-bridge'.command -eq $installedMcp -and $kimiAfter.mcpServers.'doc-bridge'.args[0] -eq '--mine') 'Final uninstall preserves a user-changed Kimi argument'
    Assert-True ($null -ne $kimiAfter.mcpServers.PSObject.Properties['existing-server']) 'Kimi unrelated MCP remains after uninstall'
    Assert-True ($null -eq $cursorAfter.mcpServers.PSObject.Properties['doc-bridge']) 'Final uninstall removes the owned Cursor entry'
    Assert-True ($null -ne $cursorAfter.mcpServers.PSObject.Properties['existing-server']) 'Cursor unrelated MCP remains after uninstall'
    Assert-True (-not [bool]$cursorAfter.telemetry.enabled) 'Cursor unrelated setting remains after uninstall'
    Assert-True ((Get-FileHash -LiteralPath $cursorProjectConfig -Algorithm SHA256).Hash -eq $cursorProjectHashBefore) 'Cursor project-level MCP config remains untouched after uninstall'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $installRoot 'codex-marketplace'))) 'Last managed client removal archives the shared payload'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $installRoot 'installation.json'))) 'Final uninstall archives the active ownership record'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PackageRoot 'Test-DocBridge.ps1') `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -SkipHwpSecurity `
        -SkipHwpRuntimeDoctor `
        -SkipClientCommands
    Assert-True ($LASTEXITCODE -eq 2) 'Default doctor reports NOT INSTALLED after final removal'

    Write-Host "`n=== Cursor-only legacy upgrade and default uninstall" -ForegroundColor Cyan
    $claudeHashBeforeCursorOnly = (Get-FileHash -LiteralPath $claudeConfig -Algorithm SHA256).Hash
    $kimiHashBeforeCursorOnly = (Get-FileHash -LiteralPath $kimiConfig -Algorithm SHA256).Hash
    & (Join-Path $PackageRoot 'Install-DocBridge.ps1') `
        -Clients Cursor `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -SkipHwpSecurity `
        -SkipHwpRuntimeDoctor `
        -SkipClientCommands `
        -AllowRunningDocBridge
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PackageRoot 'Test-DocBridge.ps1') `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -SkipHwpSecurity `
        -SkipHwpRuntimeDoctor `
        -SkipClientCommands
    Assert-True ($LASTEXITCODE -eq 0) 'Default doctor checks only Cursor for a Cursor-only installation'

    $claudeBeforeUnmanagedProbe = Get-Content -LiteralPath $claudeConfig -Raw -Encoding UTF8
    $unmanagedClaude = $claudeBeforeUnmanagedProbe | ConvertFrom-Json
    $unmanagedClaude.mcpServers.'doc-bridge' = [pscustomobject]@{ command = $installedMcp; args = @('--stdio') }
    Write-Utf8NoBom $claudeConfig ($unmanagedClaude | ConvertTo-Json -Depth 20)
    & (Join-Path $PackageRoot 'Uninstall-DocBridge.ps1') `
        -Clients ClaudeDesktop `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -SkipClientCommands
    $unmanagedClaudeAfter = Get-Content -LiteralPath $claudeConfig -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ($unmanagedClaudeAfter.mcpServers.'doc-bridge'.command -eq $installedMcp) 'Explicit uninstall preserves an unowned client entry even when command and args resemble DocBridge'
    Write-Utf8NoBom $claudeConfig $claudeBeforeUnmanagedProbe
    $legacyRecordPath = Join-Path $installRoot 'installation.json'
    $legacyRecord = Get-Content -LiteralPath $legacyRecordPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $legacyRecord.PSObject.Properties.Remove('clientMcpProvenance')
    Write-Utf8NoBom $legacyRecordPath ($legacyRecord | ConvertTo-Json -Depth 50)
    & (Join-Path $PackageRoot 'Install-DocBridge.ps1') `
        -Clients Cursor `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -SkipHwpSecurity `
        -SkipHwpRuntimeDoctor `
        -SkipClientCommands `
        -AllowRunningDocBridge
    $upgradedLegacy = Get-Content -LiteralPath $legacyRecordPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ([bool]$upgradedLegacy.clientMcpProvenance.Cursor.upgradedFromLegacyRecord) 'Legacy ownership is safely upgraded without treating the owned entry as user data'
    & (Join-Path $PackageRoot 'Uninstall-DocBridge.ps1') `
        -InstallRoot $installRoot `
        -UserProfileRoot $userRoot `
        -AppDataRoot $appDataRoot `
        -LocalAppDataRoot $localAppDataRoot `
        -SkipClientCommands
    $cursorLegacyAfter = Get-Content -LiteralPath $cursorConfig -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ($null -eq $cursorLegacyAfter.mcpServers.PSObject.Properties['doc-bridge']) 'Cursor-only default uninstall removes the owned legacy-upgraded entry'
    Assert-True ((Get-FileHash -LiteralPath $claudeConfig -Algorithm SHA256).Hash -eq $claudeHashBeforeCursorOnly) 'Cursor-only lifecycle does not modify Claude Desktop config'
    Assert-True ((Get-FileHash -LiteralPath $kimiConfig -Algorithm SHA256).Hash -eq $kimiHashBeforeCursorOnly) 'Cursor-only lifecycle does not modify Kimi config'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $installRoot 'codex-marketplace'))) 'Cursor-only default uninstall archives the payload as the last client'
    Assert-True ((Get-ChildItem -LiteralPath (Join-Path $installRoot 'backups') -Directory -Filter 'uninstalled-codex-marketplace-*').Count -eq 2) 'Each final uninstall creates a distinct recoverable payload archive'

    Write-Host "`nInstaller lifecycle test passed." -ForegroundColor Green
} finally {
    foreach ($probeProcess in @($runningProbeProcess, $runningWorkerProcess)) {
        if ($null -eq $probeProcess) { continue }
        try {
            if (-not $probeProcess.HasExited) { Stop-Process -Id $probeProcess.Id -Force }
            $probeProcess.WaitForExit(5000)
        } catch { }
        try { $probeProcess.Dispose() } catch { }
    }
    if ($KeepSandbox) {
        Write-Host "Sandbox retained: $testRoot"
    } elseif (Test-Path -LiteralPath $testRoot) {
        $resolved = [System.IO.Path]::GetFullPath($testRoot)
        $allowedPrefix = [System.IO.Path]::GetFullPath($sandboxParent).TrimEnd('\') + '\t-'
        if (-not $resolved.StartsWith($allowedPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing unsafe test cleanup: $resolved"
        }
        Remove-Item -LiteralPath $resolved -Recurse -Force
        if ((Test-Path -LiteralPath $sandboxParent) -and (Get-ChildItem -LiteralPath $sandboxParent -Force).Count -eq 0) {
            Remove-Item -LiteralPath $sandboxParent -Force
        }
    }
}
