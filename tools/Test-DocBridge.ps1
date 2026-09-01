[CmdletBinding()]
param(
    [ValidateSet('All', 'Codex', 'ClaudeCode', 'ClaudeDesktop', 'Kimi', 'Cursor')]
    [string[]]$Clients = @('All'),
    [string]$ClientsCsv,
    [string]$InstallRoot,
    [string]$UserProfileRoot = $env:USERPROFILE,
    [string]$AppDataRoot = $env:APPDATA,
    [string]$LocalAppDataRoot = $env:LOCALAPPDATA,
    [string]$CodexCliPath,
    [string]$CodexHomeRoot,
    [switch]$SkipHwpSecurity,
    # Test harness only. Real installation diagnostics must exercise hwp_doctor.
    [switch]$SkipHwpRuntimeDoctor,
    [switch]$SkipClientCommands,
    [switch]$AllowCodexRestartPending,
    [switch]$RequireExcelRuntime
)

$ErrorActionPreference = 'Stop'
$modulePath = Join-Path $PSScriptRoot 'support\DocBridge.Deployment.psm1'
if (-not (Test-Path -LiteralPath $modulePath)) {
    $modulePath = Join-Path $PSScriptRoot 'deployment\DocBridge.Deployment.psm1'
}
Import-Module $modulePath -Force

if ([string]::IsNullOrWhiteSpace($InstallRoot)) {
    $InstallRoot = Get-DocBridgeDefaultRoot -LocalAppDataRoot $LocalAppDataRoot
}
$InstallRoot = [System.IO.Path]::GetFullPath($InstallRoot).TrimEnd('\')
$installationPath = Join-Path $InstallRoot 'installation.json'
$installation = $null
if (Test-Path -LiteralPath $installationPath -PathType Leaf) {
    try {
        $installation = Get-Content -LiteralPath $installationPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        throw "Cannot read the active installation record: $installationPath`n$($_.Exception.Message)"
    }
}
if (-not [string]::IsNullOrWhiteSpace($ClientsCsv)) {
    $Clients = @($ClientsCsv.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
} elseif (-not $PSBoundParameters.ContainsKey('Clients') -and $null -ne $installation) {
    if ($null -ne $installation.PSObject.Properties['installedClients']) {
        $Clients = @($installation.installedClients | ForEach-Object { [string]$_ })
    } elseif ($null -ne $installation.PSObject.Properties['clients']) {
        $Clients = @($installation.clients | ForEach-Object { [string]$_ })
    }
}
$selectedClients = Get-DocBridgeClientSet -Clients $Clients
$marketplaceRoot = Join-Path $InstallRoot 'codex-marketplace'
$pluginRoot = Join-Path $marketplaceRoot 'plugins\doc-bridge'
$mcpExe = Join-Path $pluginRoot 'dist\doc-bridge-mcp.exe'
$cliExe = Join-Path $pluginRoot 'dist\doc-bridge-cli.exe'
$hwpWorkerExe = Join-Path $pluginRoot 'dist\doc-bridge-hwp-worker.exe'
$results = New-Object System.Collections.Generic.List[object]
$expectedVersion = $null
$pluginManifestPath = Join-Path $pluginRoot '.codex-plugin\plugin.json'
if (Test-Path -LiteralPath $pluginManifestPath -PathType Leaf) {
    try { $expectedVersion = ([string](Get-Content -LiteralPath $pluginManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json).version).Split('+')[0] } catch { }
}
$marketplaceName = 'docbridge-local'
$marketplaceJson = Join-Path $marketplaceRoot '.agents\plugins\marketplace.json'
if (Test-Path -LiteralPath $marketplaceJson -PathType Leaf) {
    try { $marketplaceName = [string](Get-Content -LiteralPath $marketplaceJson -Raw -Encoding UTF8 | ConvertFrom-Json).name } catch { }
}
$codexEnvironment = $null
if (-not [string]::IsNullOrWhiteSpace($CodexHomeRoot)) {
    $CodexHomeRoot = [System.IO.Path]::GetFullPath($CodexHomeRoot).TrimEnd('\')
    $codexEnvironment = @{ CODEX_HOME = $CodexHomeRoot }
}
$docBridgeRuntimeEnvironment = @{ DOCBRIDGE_HOME = $InstallRoot }

function Get-ExpectedJsonMcpEntry([string]$ClientName) {
    $command = $mcpExe
    $arguments = @('--stdio')
    if ($null -ne $installation -and
        $null -ne $installation.PSObject.Properties['clientMcpProvenance'] -and
        $null -ne $installation.clientMcpProvenance) {
        $property = $installation.clientMcpProvenance.PSObject.Properties[$ClientName]
        if ($null -ne $property -and $null -ne $property.Value) {
            $record = $property.Value
            if ($null -ne $record.PSObject.Properties['installedCommand'] -and
                -not [string]::IsNullOrWhiteSpace([string]$record.installedCommand)) {
                $command = [string]$record.installedCommand
            }
            if ($null -ne $record.PSObject.Properties['installedArgs'] -and $null -ne $record.installedArgs) {
                $arguments = @($record.installedArgs | ForEach-Object { [string]$_ })
            }
        }
    }
    return [pscustomobject]@{ Command = $command; Arguments = @($arguments) }
}

function Add-Check {
    param(
        [string]$Name,
        [bool]$Ok,
        [string]$Detail,
        [bool]$Critical = $true,
        [ValidateSet('OK', 'FAIL', 'WARN', 'SKIP')]
        [string]$Status
    )
    if ([string]::IsNullOrWhiteSpace($Status)) {
        $Status = if ($Ok) { 'OK' } elseif ($Critical) { 'FAIL' } else { 'WARN' }
    }
    [void]$results.Add([pscustomobject]@{
        Check = $Name
        Status = $Status
        Ok = ($Status -eq 'OK')
        Critical = ($Status -eq 'FAIL' -and $Critical)
        Detail = $Detail
    })
}

if (-not (Test-Path -LiteralPath $installationPath -PathType Leaf)) {
    Write-Host "`nDocBridge doctor results" -ForegroundColor Cyan
    Write-Host '[NOT INSTALLED] DocBridge installation.json was not found.' -ForegroundColor Yellow
    Write-Host "Expected path: $installationPath"
    Write-Host ''
    Write-Host '1. Extract the entire ZIP file first. Do not run files inside the ZIP.' -ForegroundColor Yellow
    Write-Host '2. Run 1-INSTALL.cmd and wait for INSTALLATION SUCCESS.' -ForegroundColor Yellow
    Write-Host '3. Run 2-TEST.cmd only after installation succeeds.' -ForegroundColor Yellow
    Write-Host ''
    Write-Host 'No doctor report was created because DocBridge is not installed.'
    exit 2
}

Add-Check 'installation.json' $true $installationPath
Add-Check 'plugin manifest' (Test-Path -LiteralPath $pluginManifestPath) $pluginRoot
Add-Check 'self-contained runtime' (Test-Path -LiteralPath (Join-Path $pluginRoot 'dist\coreclr.dll')) 'coreclr.dll'
Add-Check 'MCP executable' (Test-Path -LiteralPath $mcpExe) $mcpExe
Add-Check 'CLI executable' (Test-Path -LiteralPath $cliExe) $cliExe
Add-Check 'HWP isolation worker' (Test-Path -LiteralPath $hwpWorkerExe) $hwpWorkerExe
Add-Check 'marketplace' (Test-Path -LiteralPath (Join-Path $marketplaceRoot '.agents\plugins\marketplace.json')) $marketplaceRoot
$requiredSkills = @('document-automation','hwp-production-workflows','cad-production-workflows','cad-profile-sheet-pipeline')
$missingSkills = @($requiredSkills | Where-Object { -not (Test-Path -LiteralPath (Join-Path $pluginRoot ("skills\" + $_ + "\SKILL.md")) -PathType Leaf) })
Add-Check 'workflow skills' ($missingSkills.Count -eq 0) $(if ($missingSkills.Count -eq 0) { $requiredSkills -join ', ' } else { 'missing: ' + ($missingSkills -join ', ') })

if (Test-Path -LiteralPath $mcpExe) {
    try {
        $version = Invoke-DocBridgeNative -FilePath $mcpExe -Arguments @('--version') -Environment $docBridgeRuntimeEnvironment
        $versionOk = $version.ExitCode -eq 0 -and
            -not [string]::IsNullOrWhiteSpace($expectedVersion) -and
            $version.Output.Trim() -eq $expectedVersion
        Add-Check 'MCP version matches plugin' $versionOk "expected=$expectedVersion actual=$($version.Output.Trim())"
    } catch { Add-Check 'MCP version' $false $_.Exception.Message }
}

if (Test-Path -LiteralPath $cliExe) {
    try {
        $pingResult = Invoke-DocBridgeNative -FilePath $cliExe -Arguments @('core_ping') -Environment $docBridgeRuntimeEnvironment
        $ping = $pingResult.StdOut | ConvertFrom-Json
        Add-Check 'core_ping' ($pingResult.ExitCode -eq 0 -and [bool]$ping.ok) $pingResult.Output
    } catch { Add-Check 'core_ping' $false $_.Exception.Message }

    if ($SkipHwpRuntimeDoctor) {
        Add-Check 'HWP TypeLib doctor' $false 'Skipped explicitly by the installer lifecycle test harness.' -Critical $false -Status 'SKIP'
    } else {
        try {
            $doctorResult = Invoke-DocBridgeNative -FilePath $cliExe -Arguments @('hwp_doctor') -Environment $docBridgeRuntimeEnvironment
            $doctor = $doctorResult.StdOut | ConvertFrom-Json
            $doctorState = [string]$doctor.state
            $doctorOk = $doctorResult.ExitCode -eq 0 -and [bool]$doctor.ok -and $doctorState -eq 'CHECK_PASSED'
            Add-Check 'HWP TypeLib doctor' $doctorOk "state=$doctorState; $($doctorResult.Output)"
            if ($doctorOk -and ([bool]$doctor.updateRecommended -or [bool]$doctor.ownedAutomationBlocked)) {
                $runtimeDetail = "installed=$([string]$doctor.installedVersion); recommended=$([string]$doctor.recommendedHwp2024Version); " +
                    "ownedAutomationBlocked=$([bool]$doctor.ownedAutomationBlocked); action=$([string]$doctor.userAction)"
                Add-Check 'HWP runtime update recommended' $false $runtimeDetail -Critical $false -Status 'WARN'
            }
        } catch { Add-Check 'HWP TypeLib doctor' $false $_.Exception.Message }
    }
}

$smoke = Join-Path $PSScriptRoot 'support\verify-mcp.ps1'
if (-not (Test-Path -LiteralPath $smoke)) {
    $smoke = Join-Path $PSScriptRoot 'verify-mcp.ps1'
}
if ((Test-Path -LiteralPath $smoke) -and (Test-Path -LiteralPath $mcpExe)) {
    try {
        $powerShellExe = (Get-Command powershell.exe -ErrorAction Stop).Source
        $smokeArguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $smoke, '-Exe', $mcpExe)
        if ($RequireExcelRuntime) { $smokeArguments += '-RequireExcelRuntime' }
        $smokeResult = Invoke-DocBridgeNative -FilePath $powerShellExe `
            -Arguments $smokeArguments `
            -Environment $docBridgeRuntimeEnvironment
        $smokeEncodingOk = $smokeResult.Output.IndexOf([char]0xFFFD) -lt 0
        Add-Check 'MCP handshake' ($smokeResult.ExitCode -eq 0 -and $smokeEncodingOk) $smokeResult.Output
    } catch { Add-Check 'MCP handshake' $false $_.Exception.Message }
}

if (-not $SkipHwpSecurity) {
    $expectedModule = Join-Path $pluginRoot 'dist\hwp-security\FilePathCheckerModuleExample.dll'
    try {
        $query = (& reg.exe query 'HKCU\SOFTWARE\HNC\HwpAutomation\Modules' /v DocBridgeFilePathChecker /reg:32 2>&1 | Out-String)
        $registered = $LASTEXITCODE -eq 0 -and $query.IndexOf($expectedModule, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
        Add-Check 'HWP security module' $registered $query.Trim()
    } catch { Add-Check 'HWP security module' $false $_.Exception.Message }
}

if ($selectedClients -contains 'ClaudeDesktop' -and -not [string]::IsNullOrWhiteSpace($AppDataRoot)) {
    $path = Join-Path $AppDataRoot 'Claude\claude_desktop_config.json'
    try {
        $config = Read-DocBridgeJsonObject -Path $path
        $entry = if ($null -ne $config.PSObject.Properties['mcpServers']) { $config.mcpServers.PSObject.Properties['doc-bridge'] } else { $null }
        $expected = Get-ExpectedJsonMcpEntry 'ClaudeDesktop'
        $ok = $null -ne $entry -and (Test-DocBridgeMcpJsonEntryMatches -Entry $entry.Value -Command $expected.Command -Arguments $expected.Arguments)
        Add-Check 'Claude Desktop config' $ok "$path; expectedCommand=$($expected.Command); expectedArgs=$($expected.Arguments -join ' ')"
    } catch { Add-Check 'Claude Desktop config' $false $_.Exception.Message }
}

if ($selectedClients -contains 'Kimi' -and -not [string]::IsNullOrWhiteSpace($UserProfileRoot)) {
    $path = Join-Path $UserProfileRoot '.kimi\mcp.json'
    try {
        $config = Read-DocBridgeJsonObject -Path $path
        $entry = if ($null -ne $config.PSObject.Properties['mcpServers']) { $config.mcpServers.PSObject.Properties['doc-bridge'] } else { $null }
        $expected = Get-ExpectedJsonMcpEntry 'Kimi'
        $ok = $null -ne $entry -and (Test-DocBridgeMcpJsonEntryMatches -Entry $entry.Value -Command $expected.Command -Arguments $expected.Arguments)
        Add-Check 'Kimi config' $ok "$path; expectedCommand=$($expected.Command); expectedArgs=$($expected.Arguments -join ' ')"
    } catch { Add-Check 'Kimi config' $false $_.Exception.Message }
}

if ($selectedClients -contains 'Cursor' -and -not [string]::IsNullOrWhiteSpace($UserProfileRoot)) {
    $path = Join-Path $UserProfileRoot '.cursor\mcp.json'
    try {
        $config = Read-DocBridgeJsonObject -Path $path
        $entry = if ($null -ne $config.PSObject.Properties['mcpServers']) { $config.mcpServers.PSObject.Properties['doc-bridge'] } else { $null }
        $expected = Get-ExpectedJsonMcpEntry 'Cursor'
        $cursorArguments = if ($null -ne $entry -and $null -ne $entry.Value.PSObject.Properties['args']) {
            @($entry.Value.args | ForEach-Object { [string]$_ })
        } else { @() }
        $cursorCommand = if ($null -ne $entry) { [string]$entry.Value.command } else { '' }
        $ok = $null -ne $entry -and (Test-DocBridgeMcpJsonEntryMatches -Entry $entry.Value -Command $expected.Command -Arguments $expected.Arguments)
        Add-Check 'Cursor global config' $ok "$path; command=$cursorCommand; args=$($cursorArguments -join ' ')"
    } catch { Add-Check 'Cursor global config' $false $_.Exception.Message }

    $cursorGuideRoot = Join-Path $InstallRoot 'generated-configs\cursor'
    $templates = @(
        (Join-Path $cursorGuideRoot 'mcp.example.json'),
        (Join-Path $cursorGuideRoot 'rules\docbridge-safe-automation.mdc'),
        (Join-Path $cursorGuideRoot 'docbridge-user-rule.txt'),
        (Join-Path $cursorGuideRoot 'CURSOR_USAGE.md')
    )
    $missingTemplates = @($templates | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
    Add-Check 'Cursor guidance templates' ($missingTemplates.Count -eq 0) $(if ($missingTemplates.Count -eq 0) { $cursorGuideRoot } else { 'missing: ' + ($missingTemplates -join ', ') })
}

if ($SkipClientCommands) {
    if ($selectedClients -contains 'Codex') {
        $manual = Join-Path $InstallRoot 'generated-configs\install-codex.txt'
        Add-Check 'Codex manual registration' (Test-Path -LiteralPath $manual -PathType Leaf) $manual
    }
    if ($selectedClients -contains 'ClaudeCode') {
        $manual = Join-Path $InstallRoot 'generated-configs\install-claude-code.txt'
        Add-Check 'Claude Code manual registration' (Test-Path -LiteralPath $manual -PathType Leaf) $manual
    }
} else {
    if ($selectedClients -contains 'ClaudeCode') {
        $claude = Get-Command claude -ErrorAction SilentlyContinue
        if ($null -eq $claude) { Add-Check 'Claude Code CLI' $false 'Not installed on this PC; test skipped.' $false 'SKIP' }
        else {
            $expectedCommand = $mcpExe
            $expectedArguments = @('--stdio')
            if ($null -ne $installation -and
                $null -ne $installation.PSObject.Properties['claudeCodeOwnership'] -and
                $null -ne $installation.claudeCodeOwnership) {
                if ($null -ne $installation.claudeCodeOwnership.PSObject.Properties['installedCommand'] -and
                    -not [string]::IsNullOrWhiteSpace([string]$installation.claudeCodeOwnership.installedCommand)) {
                    $expectedCommand = [string]$installation.claudeCodeOwnership.installedCommand
                }
                if ($null -ne $installation.claudeCodeOwnership.PSObject.Properties['installedArgs'] -and
                    $null -ne $installation.claudeCodeOwnership.installedArgs) {
                    $expectedArguments = @($installation.claudeCodeOwnership.installedArgs | ForEach-Object { [string]$_ })
                }
            }
            $claudeState = Get-DocBridgeClaudeMcpState -FilePath $claude.Source
            $claudeOk = Test-DocBridgeClaudeMcpStateMatches -State $claudeState -Command $expectedCommand -Arguments $expectedArguments
            Add-Check 'Claude Code MCP ownership' $claudeOk "command=$($claudeState.Command); args=$($claudeState.ArgsText); environment=$($claudeState.Environment); error=$($claudeState.Error)" $true
        }
    }
    if ($selectedClients -contains 'Kimi') {
        $kimi = Get-Command kimi -ErrorAction SilentlyContinue
        if ($null -eq $kimi) { Add-Check 'Kimi CLI' $false 'Not installed on this PC; test skipped.' $false 'SKIP' }
        else {
            $kimiTest = Invoke-DocBridgeNative -FilePath $kimi.Source -Arguments @('mcp', 'test', 'doc-bridge')
            Add-Check 'Kimi MCP test' ($kimiTest.ExitCode -eq 0) $kimiTest.Output $false
        }
    }
    if ($selectedClients -contains 'Codex') {
        $codexExecutable = Get-DocBridgeCodexCli -UserProfileRoot $UserProfileRoot -LocalAppDataRoot $LocalAppDataRoot -ExplicitPath $CodexCliPath
        $codexDetected = Test-DocBridgeCodexInstalled -UserProfileRoot $UserProfileRoot -LocalAppDataRoot $LocalAppDataRoot
        if ([string]::IsNullOrWhiteSpace($codexExecutable)) {
            if ($codexDetected) {
                Add-Check 'Codex CLI discovery' $false 'Codex is installed, but no working CLI was found. Open Codex once, close it, and reinstall.' $true
            } else {
                Add-Check 'Codex CLI' $false 'Not installed on this PC; test skipped.' $false 'SKIP'
            }
        }
        else {
            Add-Check 'Codex CLI discovery' $true $codexExecutable
            $cliMode = Get-DocBridgeCodexPluginCliMode -FilePath $codexExecutable -Environment $codexEnvironment
            Add-Check 'Codex plugin CLI mode' ($cliMode -ne 'Unknown') $cliMode
            if ($cliMode -eq 'LegacyPluginCommands') {
                $listResult = Invoke-DocBridgeNative -FilePath $codexExecutable -Arguments @('plugin', 'list') -Environment $codexEnvironment
                $enabled = $listResult.ExitCode -eq 0 -and
                    (Test-DocBridgeCodexPluginEnabled -Output $listResult.Output -MarketplaceName $marketplaceName)
                $matchingLine = @($listResult.Output -split "`r?`n" | Where-Object { $_ -match 'doc-bridge@' }) -join ' | '
                if ([string]::IsNullOrWhiteSpace($matchingLine)) { $matchingLine = 'doc-bridge was not listed.' }
                Add-Check 'Codex plugin installed and enabled' $enabled $matchingLine $true
            } else {
                $marketplaceRegistered = Test-DocBridgeCodexMarketplaceRegistered -MarketplaceName $marketplaceName -UserProfileRoot $UserProfileRoot -CodexHomeRoot $CodexHomeRoot
                $configPath = Get-DocBridgeCodexConfigPath -UserProfileRoot $UserProfileRoot -CodexHomeRoot $CodexHomeRoot
                Add-Check 'Codex marketplace registered' $marketplaceRegistered "$marketplaceName in $configPath" $true
            }
            $mcpListResult = Invoke-DocBridgeNative -FilePath $codexExecutable -Arguments @('mcp', 'list') -Environment $codexEnvironment
            $mcpEnabled = $mcpListResult.ExitCode -eq 0 -and (Test-DocBridgeCodexMcpEnabled -Output $mcpListResult.Output)
            $matchingMcpLine = @($mcpListResult.Output -split "`r?`n" | Where-Object { $_ -match '^\s*doc-bridge\s+' }) -join ' | '
            if ([string]::IsNullOrWhiteSpace($matchingMcpLine)) { $matchingMcpLine = 'doc-bridge MCP was not listed.' }
            if ($AllowCodexRestartPending -and -not $mcpEnabled -and $cliMode -ne 'LegacyPluginCommands') {
                Add-Check 'Codex MCP visible and enabled' $false ($matchingMcpLine + ' Restart Codex completely and open a new task, then run 2-TEST.cmd.') $false 'WARN'
            } else {
                Add-Check 'Codex MCP visible and enabled' $mcpEnabled $matchingMcpLine $true
            }
        }
    }
}

Write-Host "`nDocBridge doctor results" -ForegroundColor Cyan
foreach ($item in $results) {
    $label = switch ($item.Status) {
        'OK' { '[OK]  ' }
        'FAIL' { '[FAIL]' }
        'WARN' { '[WARN]' }
        'SKIP' { '[SKIP]' }
    }
    $color = switch ($item.Status) {
        'OK' { 'Green' }
        'FAIL' { 'Red' }
        'SKIP' { 'DarkGray' }
        default { 'Yellow' }
    }
    Write-Host "$label $($item.Check): $($item.Detail)" -ForegroundColor $color
}

$report = [ordered]@{
    generatedAt = (Get-Date).ToString('o')
    installRoot = $InstallRoot
    clients = @($selectedClients)
    results = @($results.ToArray())
} | ConvertTo-Json -Depth 10
Write-DocBridgeUtf8NoBom -Path (Join-Path $InstallRoot 'doctor-report.json') -Text $report

$criticalFailures = @($results | Where-Object { $_.Critical -and -not $_.Ok })
if ($criticalFailures.Count -gt 0) { exit 1 }
exit 0
