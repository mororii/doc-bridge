[CmdletBinding()]
param(
    [ValidateSet('All', 'Codex', 'ClaudeCode', 'ClaudeDesktop', 'Kimi', 'Cursor')]
    [string[]]$Clients = @('All'),
    [string]$InstallRoot,
    [string]$UserProfileRoot = $env:USERPROFILE,
    [string]$AppDataRoot = $env:APPDATA,
    [string]$LocalAppDataRoot = $env:LOCALAPPDATA,
    [string]$CodexCliPath,
    [string]$CodexHomeRoot,
    [switch]$SkipHwpSecurity,
    [switch]$SkipClientCommands,
    [switch]$SkipDoctor,
    # Test harness only. Real installations must keep the HWP runtime doctor enabled.
    [switch]$SkipHwpRuntimeDoctor,
    # Advanced recovery only. The installer never stops running DocBridge processes.
    [switch]$AllowRunningDocBridge
)

$ErrorActionPreference = 'Stop'
$modulePath = Join-Path $PSScriptRoot 'support\DocBridge.Deployment.psm1'
if (-not (Test-Path -LiteralPath $modulePath)) {
    $modulePath = Join-Path $PSScriptRoot 'deployment\DocBridge.Deployment.psm1'
}
Import-Module $modulePath -Force

function Get-RunningDocBridgeProcessInfo {
    $result = New-Object System.Collections.Generic.List[object]
    $processes = @(Get-Process -Name @('doc-bridge-mcp', 'doc-bridge-hwp-worker') -ErrorAction SilentlyContinue)
    foreach ($process in $processes) {
        $path = '<unavailable>'
        try {
            if (-not [string]::IsNullOrWhiteSpace([string]$process.Path)) {
                $path = [string]$process.Path
            }
        } catch { }
        [void]$result.Add([pscustomobject]@{
            Name = [string]$process.ProcessName
            Id = [int]$process.Id
            Path = $path
        })
        try { $process.Dispose() } catch { }
    }
    return @($result | Sort-Object Id)
}

$runningDocBridge = @(Get-RunningDocBridgeProcessInfo)
if ($runningDocBridge.Count -gt 0) {
    $runningDetails = @($runningDocBridge | ForEach-Object {
        "PID=$($_.Id) name=$($_.Name) path=$($_.Path)"
    })
    $message = "Running DocBridge automation processes were detected.`n  " +
        ($runningDetails -join "`n  ") +
        "`nClose Codex, Claude, Kimi, and Cursor completely, then run the installer again. " +
        "The installer will not stop or kill these processes."
    if (-not $AllowRunningDocBridge) {
        throw $message
    }
    Write-Warning ($message + "`n-AllowRunningDocBridge was supplied, so installation will continue at the caller's risk.")
}

if ([string]::IsNullOrWhiteSpace($InstallRoot)) {
    $InstallRoot = Get-DocBridgeDefaultRoot -LocalAppDataRoot $LocalAppDataRoot
}
$InstallRoot = [System.IO.Path]::GetFullPath($InstallRoot).TrimEnd('\')
$installParent = Split-Path -Parent $InstallRoot
if ([string]::IsNullOrWhiteSpace($installParent)) { throw "Invalid install path: $InstallRoot" }
$previousInstallation = $null
$previousInstallationPath = Join-Path $InstallRoot 'installation.json'
if (Test-Path -LiteralPath $previousInstallationPath -PathType Leaf) {
    try {
        $previousInstallation = Get-Content -LiteralPath $previousInstallationPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        throw "Cannot safely upgrade an unreadable installation record: $previousInstallationPath`n$($_.Exception.Message)"
    }
}

$payloadRoot = Join-Path $PSScriptRoot 'payload\codex-marketplace'
$payloadPlugin = Join-Path $payloadRoot 'plugins\doc-bridge'
$payloadMarketplace = Join-Path $payloadRoot '.agents\plugins\marketplace.json'
$payloadExe = Join-Path $payloadPlugin 'dist\doc-bridge-mcp.exe'
$payloadHwpWorker = Join-Path $payloadPlugin 'dist\doc-bridge-hwp-worker.exe'
$payloadCoreClr = Join-Path $payloadPlugin 'dist\coreclr.dll'
$payloadHwpSecurity = Join-Path $payloadPlugin 'dist\hwp-security\FilePathCheckerModuleExample.dll'
$pluginManifestPath = Join-Path $payloadPlugin '.codex-plugin\plugin.json'
foreach ($required in @($payloadMarketplace, $payloadExe, $payloadHwpWorker, $payloadCoreClr, $payloadHwpSecurity, $pluginManifestPath)) {
    if (-not (Test-Path -LiteralPath $required -PathType Leaf)) {
        throw "Incomplete deployment payload: $required"
    }
}

$pluginManifest = Get-Content -LiteralPath $pluginManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$marketplaceInfo = Get-Content -LiteralPath $payloadMarketplace -Raw -Encoding UTF8 | ConvertFrom-Json
if ($pluginManifest.name -ne 'doc-bridge') { throw 'Payload plugin name must be doc-bridge.' }
if ([string]::IsNullOrWhiteSpace([string]$marketplaceInfo.name)) { throw 'Marketplace name is missing.' }
$pluginBaseVersion = ([string]$pluginManifest.version).Split('+')[0]

Write-Host "Installing DocBridge $($pluginManifest.version)" -ForegroundColor Cyan
Write-Host "Install root: $InstallRoot"
$targetMarketplace = Join-Path $InstallRoot 'codex-marketplace'
$selectedClients = Get-DocBridgeClientSet -Clients $Clients
$codexMarketplaceOwnership = if ($null -ne $previousInstallation -and
    $null -ne $previousInstallation.PSObject.Properties['codexMarketplaceOwnership']) {
    $previousInstallation.codexMarketplaceOwnership
} else { $null }
$previousManagedClients = @()
if ($null -ne $previousInstallation -and $null -ne $previousInstallation.PSObject.Properties['installedClients']) {
    $previousManagedClients = @($previousInstallation.installedClients)
} elseif ($null -ne $previousInstallation -and $null -ne $previousInstallation.PSObject.Properties['clients']) {
    $previousManagedClients = @($previousInstallation.clients)
}
if ($selectedClients -contains 'Codex' -or $previousManagedClients -contains 'Codex') {
    $codexMarketplaceState = Get-DocBridgeCodexMarketplaceState -MarketplaceName ([string]$marketplaceInfo.name) `
        -UserProfileRoot $UserProfileRoot -CodexHomeRoot $CodexHomeRoot
    if ($null -ne $codexMarketplaceState.Error) {
        throw "Cannot safely verify the existing Codex marketplace entry '$($marketplaceInfo.name)': $($codexMarketplaceState.Error)"
    }
    if ($codexMarketplaceState.Exists) {
        $allowedSources = New-Object System.Collections.Generic.List[string]
        $currentSource = ConvertTo-DocBridgeNormalizedLocalPath -Path $targetMarketplace
        if (-not [string]::IsNullOrWhiteSpace($currentSource)) { [void]$allowedSources.Add($currentSource) }
        if ($null -ne $previousInstallation -and
            $null -ne $previousInstallation.PSObject.Properties['marketplaceRoot'] -and
            -not [string]::IsNullOrWhiteSpace([string]$previousInstallation.marketplaceRoot)) {
            $previousSource = ConvertTo-DocBridgeNormalizedLocalPath -Path ([string]$previousInstallation.marketplaceRoot)
            if (-not [string]::IsNullOrWhiteSpace($previousSource) -and -not $allowedSources.Contains($previousSource)) {
                [void]$allowedSources.Add($previousSource)
            }
        }
        if ($null -ne $codexMarketplaceOwnership -and
            $null -ne $codexMarketplaceOwnership.PSObject.Properties['normalizedSource'] -and
            -not [string]::IsNullOrWhiteSpace([string]$codexMarketplaceOwnership.normalizedSource)) {
            $ownedSource = ConvertTo-DocBridgeNormalizedLocalPath -Path ([string]$codexMarketplaceOwnership.normalizedSource)
            if (-not [string]::IsNullOrWhiteSpace($ownedSource) -and -not $allowedSources.Contains($ownedSource)) {
                [void]$allowedSources.Add($ownedSource)
            }
        }
        $preexistingOwned = [string]$codexMarketplaceState.SourceType -ieq 'local' -and
            @($allowedSources | Where-Object {
                ([string]$_).Equals([string]$codexMarketplaceState.NormalizedSource, [System.StringComparison]::OrdinalIgnoreCase)
            }).Count -gt 0
        if (-not $preexistingOwned) {
            throw "Codex already has a foreign marketplace named '$($marketplaceInfo.name)' and it was preserved. source_type=$($codexMarketplaceState.SourceType); source=$($codexMarketplaceState.Source)"
        }
    }
}
New-Item -ItemType Directory -Path $InstallRoot -Force | Out-Null
$backupRoot = Join-Path $InstallRoot 'backups'
$stagingRoot = Join-Path $InstallRoot ('.staging-' + [guid]::NewGuid().ToString('N'))
$stagedMarketplace = Join-Path $stagingRoot 'codex-marketplace'
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
$previousMarketplace = $null

try {
    New-Item -ItemType Directory -Path $stagedMarketplace -Force | Out-Null
    Copy-Item -Path (Join-Path $payloadRoot '*') -Destination $stagedMarketplace -Recurse -Force
    Copy-Item -LiteralPath (Join-Path $payloadRoot '.agents') -Destination $stagedMarketplace -Recurse -Force

    $stagedExe = Join-Path $stagedMarketplace 'plugins\doc-bridge\dist\doc-bridge-mcp.exe'
    $versionResult = Invoke-DocBridgeNative -FilePath $stagedExe -Arguments @('--version')
    if ($versionResult.ExitCode -ne 0) { throw "Staged executable verification failed: $($versionResult.Output)" }
    if ($versionResult.Output.Trim() -ne $pluginBaseVersion) {
        throw "Payload version mismatch: plugin=$pluginBaseVersion executable=$($versionResult.Output.Trim())"
    }
    Write-Host "Staged MCP version: $($versionResult.Output)"

    if (Test-Path -LiteralPath $targetMarketplace) {
        $safeTarget = Assert-DocBridgeSafePath -Path $targetMarketplace -Root $InstallRoot
        New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
        $previousMarketplace = Join-Path $backupRoot "codex-marketplace-$timestamp"
        Move-Item -LiteralPath $safeTarget -Destination $previousMarketplace
        Write-Host "Previous installation backup: $previousMarketplace"
    }

    try {
        Move-Item -LiteralPath $stagedMarketplace -Destination $targetMarketplace
    } catch {
        if ($null -ne $previousMarketplace -and
            (Test-Path -LiteralPath $previousMarketplace) -and
            -not (Test-Path -LiteralPath $targetMarketplace)) {
            Move-Item -LiteralPath $previousMarketplace -Destination $targetMarketplace
        }
        throw
    }
} finally {
    if (Test-Path -LiteralPath $stagingRoot) {
        $safeStage = Assert-DocBridgeSafePath -Path $stagingRoot -Root $InstallRoot
        Remove-Item -LiteralPath $safeStage -Recurse -Force
    }
}

$installedPlugin = Join-Path $targetMarketplace 'plugins\doc-bridge'
$mcpExe = Join-Path $installedPlugin 'dist\doc-bridge-mcp.exe'
$cliExe = Join-Path $installedPlugin 'dist\doc-bridge-cli.exe'
$installedMarketplaceJson = Join-Path $targetMarketplace '.agents\plugins\marketplace.json'
$installedMarketplace = Get-Content -LiteralPath $installedMarketplaceJson -Raw -Encoding UTF8 | ConvertFrom-Json
$installedClients = New-Object System.Collections.Generic.List[string]
$previousClients = @()
if ($null -ne $previousInstallation -and $null -ne $previousInstallation.PSObject.Properties['installedClients']) {
    $previousClients = @($previousInstallation.installedClients)
} elseif ($null -ne $previousInstallation -and $null -ne $previousInstallation.PSObject.Properties['clients']) {
    $previousClients = @($previousInstallation.clients)
}
if ($previousClients.Count -gt 0) {
    foreach ($client in $previousClients) {
        $clientName = [string]$client
        if (-not [string]::IsNullOrWhiteSpace($clientName) -and -not $installedClients.Contains($clientName)) {
            [void]$installedClients.Add($clientName)
        }
    }
}
foreach ($client in $selectedClients) {
    if (-not $installedClients.Contains([string]$client)) { [void]$installedClients.Add([string]$client) }
}
$clientMcpProvenance = [ordered]@{}
if ($null -ne $previousInstallation -and
    $null -ne $previousInstallation.PSObject.Properties['clientMcpProvenance'] -and
    $null -ne $previousInstallation.clientMcpProvenance) {
    foreach ($property in $previousInstallation.clientMcpProvenance.PSObject.Properties) {
        $clientMcpProvenance[$property.Name] = $property.Value
    }
}
$warnings = New-Object System.Collections.Generic.List[string]
$criticalErrors = New-Object System.Collections.Generic.List[string]
$codexEnvironment = $null
$codexExecutable = $null
$codexCliMode = $null
$claudeCodeOwnership = if ($null -ne $previousInstallation -and
    $null -ne $previousInstallation.PSObject.Properties['claudeCodeOwnership']) {
    $previousInstallation.claudeCodeOwnership
} else { $null }
if (-not [string]::IsNullOrWhiteSpace($CodexHomeRoot)) {
    $CodexHomeRoot = [System.IO.Path]::GetFullPath($CodexHomeRoot).TrimEnd('\')
    New-Item -ItemType Directory -Path $CodexHomeRoot -Force | Out-Null
    $codexEnvironment = @{ CODEX_HOME = $CodexHomeRoot }
}

function Get-ClientMcpConfigPath([string]$ClientName) {
    switch ($ClientName) {
        'ClaudeDesktop' {
            if ([string]::IsNullOrWhiteSpace($AppDataRoot)) { return $null }
            return Join-Path $AppDataRoot 'Claude\claude_desktop_config.json'
        }
        'Kimi' {
            if ([string]::IsNullOrWhiteSpace($UserProfileRoot)) { return $null }
            return Join-Path $UserProfileRoot '.kimi\mcp.json'
        }
        'Cursor' {
            if ([string]::IsNullOrWhiteSpace($UserProfileRoot)) { return $null }
            return Join-Path $UserProfileRoot '.cursor\mcp.json'
        }
        default { return $null }
    }
}

function Initialize-ClientMcpProvenance([string]$ClientName, [string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path) -or $clientMcpProvenance.Contains($ClientName)) { return }
    $state = Get-DocBridgeMcpJsonEntryState -Path $Path
    $legacyOwned = $false
    if ($null -ne $previousInstallation -and
        $null -ne $previousInstallation.PSObject.Properties['mcpExecutable'] -and
        -not [string]::IsNullOrWhiteSpace([string]$previousInstallation.mcpExecutable) -and
        $state.entryExisted) {
        $legacyOwned = Test-DocBridgeMcpJsonEntryMatches -Entry $state.entry `
            -Command ([string]$previousInstallation.mcpExecutable) -Arguments @('--stdio')
    }
    $clientMcpProvenance[$ClientName] = [ordered]@{
        path = $Path
        previousEntryExisted = [bool]($state.entryExisted -and -not $legacyOwned)
        previousEntry = $(if ($state.entryExisted -and -not $legacyOwned) { $state.entry } else { $null })
        installedCommand = $mcpExe
        installedArgs = @('--stdio')
        upgradedFromLegacyRecord = [bool]$legacyOwned
    }
}

foreach ($clientName in @('ClaudeDesktop', 'Kimi', 'Cursor')) {
    if ($installedClients.Contains($clientName)) {
        $clientPath = Get-ClientMcpConfigPath $clientName
        Initialize-ClientMcpProvenance -ClientName $clientName -Path $clientPath
        if ($clientMcpProvenance.Contains($clientName)) {
            $clientMcpProvenance[$clientName].path = $clientPath
            $clientMcpProvenance[$clientName].installedCommand = $mcpExe
            $clientMcpProvenance[$clientName].installedArgs = @('--stdio')
        }
    }
}

if (-not $SkipHwpSecurity) {
    $securityInstaller = Join-Path $installedPlugin 'dist\install-hwp-security.ps1'
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $securityInstaller
        if ($LASTEXITCODE -ne 0) { throw "exit code $LASTEXITCODE" }
    } catch {
        $warnings.Add("HWP security module registration failed: $($_.Exception.Message)")
    }
}

if ($selectedClients -contains 'ClaudeDesktop') {
    if ([string]::IsNullOrWhiteSpace($AppDataRoot)) {
        $warnings.Add('APPDATA is unavailable; Claude Desktop configuration was skipped.')
    } else {
        $claudeDesktopConfig = Get-ClientMcpConfigPath 'ClaudeDesktop'
        [void](Set-DocBridgeMcpJsonEntry -Path $claudeDesktopConfig -Command $mcpExe -Arguments @('--stdio') -BackupRoot $backupRoot)
        Write-Host "Claude Desktop configured: $claudeDesktopConfig"
    }
}

if ($selectedClients -contains 'Kimi') {
    if ([string]::IsNullOrWhiteSpace($UserProfileRoot)) {
        $warnings.Add('User profile is unavailable; Kimi configuration was skipped.')
    } else {
        $kimiConfig = Get-ClientMcpConfigPath 'Kimi'
        [void](Set-DocBridgeMcpJsonEntry -Path $kimiConfig -Command $mcpExe -Arguments @('--stdio') -BackupRoot $backupRoot)
        Write-Host "Kimi configured: $kimiConfig"
    }
}

if ($selectedClients -contains 'Cursor') {
    if ([string]::IsNullOrWhiteSpace($UserProfileRoot)) {
        $warnings.Add('User profile is unavailable; Cursor configuration was skipped.')
    } else {
        # Cursor's global MCP file is user-scoped. Do not modify a project's
        # .cursor\mcp.json because project configuration may intentionally differ.
        $cursorConfig = Get-ClientMcpConfigPath 'Cursor'
        [void](Set-DocBridgeMcpJsonEntry -Path $cursorConfig -Command $mcpExe -Arguments @('--stdio') -BackupRoot $backupRoot)
        Write-Host "Cursor configured: $cursorConfig"
    }
}

$generatedRoot = Join-Path $InstallRoot 'generated-configs'
New-Item -ItemType Directory -Path $generatedRoot -Force | Out-Null
$portableJson = [ordered]@{
    mcpServers = [ordered]@{
        'doc-bridge' = [ordered]@{ command = $mcpExe; args = @('--stdio') }
    }
} | ConvertTo-Json -Depth 8
Write-DocBridgeUtf8NoBom -Path (Join-Path $generatedRoot 'mcp.json') -Text $portableJson
foreach ($clientConfigName in @('claude_desktop_config.json', 'claude-code.mcp.json', 'kimi-mcp.json', 'cursor-mcp.json')) {
    Write-DocBridgeUtf8NoBom -Path (Join-Path $generatedRoot $clientConfigName) -Text $portableJson
}
$mcpExeToml = $mcpExe.Replace('\', '/')
$codexToml = @"
[mcp_servers.doc-bridge]
command = "$mcpExeToml"
args = ["--stdio"]
startup_timeout_sec = 30
tool_timeout_sec = 300

[mcp_servers.doc-bridge.tools.excel_apply_ops]
approval_mode = "approve"

[mcp_servers.doc-bridge.tools.hwp_apply_ops]
approval_mode = "approve"

[mcp_servers.doc-bridge.tools.cad_apply_ops]
approval_mode = "approve"

[mcp_servers.doc-bridge.tools.core_restore_snapshot]
approval_mode = "approve"
"@
Write-DocBridgeUtf8NoBom -Path (Join-Path $generatedRoot 'codex-config.toml') -Text $codexToml

if ($selectedClients -contains 'Cursor') {
    $cursorClientRoot = Join-Path $installedPlugin 'clients\cursor'
    $cursorGeneratedRoot = Join-Path $generatedRoot 'cursor'
    New-Item -ItemType Directory -Path $cursorGeneratedRoot -Force | Out-Null
    foreach ($cursorTemplate in @(
        @{ Source = (Join-Path $cursorClientRoot 'mcp.example.json'); Target = 'mcp.example.json' },
        @{ Source = (Join-Path $cursorClientRoot 'rules\docbridge-safe-automation.mdc'); Target = 'rules\docbridge-safe-automation.mdc' },
        @{ Source = (Join-Path $cursorClientRoot 'docbridge-user-rule.txt'); Target = 'docbridge-user-rule.txt' },
        @{ Source = (Join-Path $cursorClientRoot 'CURSOR_USAGE.md'); Target = 'CURSOR_USAGE.md' }
    )) {
        if (Test-Path -LiteralPath $cursorTemplate.Source -PathType Leaf) {
            $cursorTemplateTarget = Join-Path $cursorGeneratedRoot $cursorTemplate.Target
            New-Item -ItemType Directory -Path (Split-Path -Parent $cursorTemplateTarget) -Force | Out-Null
            Copy-Item -LiteralPath $cursorTemplate.Source -Destination $cursorTemplateTarget -Force
        } else {
            $warnings.Add("Cursor guidance template was not found: $($cursorTemplate.Source)")
        }
    }
}

if ($selectedClients -contains 'ClaudeCode') {
    $claude = Get-Command claude -ErrorAction SilentlyContinue
    if ($SkipClientCommands -or $null -eq $claude) {
        $commandText = "claude mcp add --scope user doc-bridge -- `"$mcpExe`" --stdio`r`n"
        Write-DocBridgeUtf8NoBom -Path (Join-Path $generatedRoot 'install-claude-code.txt') -Text $commandText
        $warnings.Add('Claude Code CLI registration was skipped. See generated-configs\install-claude-code.txt.')
    } else {
        $claudeState = Get-DocBridgeClaudeMcpState -FilePath $claude.Source
        if ($null -ne $claudeState.Error) {
            $criticalErrors.Add("Claude Code MCP ownership could not be verified: $($claudeState.Error) $($claudeState.Output)")
        } else {
            $ownedCommands = New-Object System.Collections.Generic.List[string]
            [void]$ownedCommands.Add($mcpExe)
            if ($null -ne $previousInstallation -and
                $null -ne $previousInstallation.PSObject.Properties['mcpExecutable'] -and
                -not [string]::IsNullOrWhiteSpace([string]$previousInstallation.mcpExecutable) -and
                -not $ownedCommands.Contains([string]$previousInstallation.mcpExecutable)) {
                [void]$ownedCommands.Add([string]$previousInstallation.mcpExecutable)
            }
            $ownedEntry = $false
            if ($claudeState.Exists) {
                foreach ($ownedCommand in $ownedCommands) {
                    if (Test-DocBridgeClaudeMcpStateMatches -State $claudeState -Command $ownedCommand -Arguments @('--stdio')) {
                        $ownedEntry = $true
                        break
                    }
                }
            }
            $foreignEntry = $claudeState.Exists -and -not $ownedEntry
            if ($foreignEntry) {
                $criticalErrors.Add("Claude Code already has a foreign or user-modified user-scope doc-bridge entry and it was preserved: command=$($claudeState.Command); args=$($claudeState.ArgsText); environment=$($claudeState.Environment)")
            } else {
                if ($claudeState.Exists) {
                    $claudeRemove = Invoke-DocBridgeNative -FilePath $claude.Source -Arguments @('mcp', 'remove', '--scope', 'user', 'doc-bridge')
                    if ($claudeRemove.ExitCode -ne 0) {
                        $criticalErrors.Add("Claude Code owned MCP removal failed: $($claudeRemove.Output)")
                    }
                }
                if ($criticalErrors.Count -eq 0) {
                    $claudeAdd = Invoke-DocBridgeNative -FilePath $claude.Source -Arguments @('mcp', 'add', '--scope', 'user', 'doc-bridge', '--', $mcpExe, '--stdio')
                    $claudeInstalledState = if ($claudeAdd.ExitCode -eq 0) {
                        Get-DocBridgeClaudeMcpState -FilePath $claude.Source
                    } else { $null }
                    $claudeConfigured = $claudeAdd.ExitCode -eq 0 -and
                        (Test-DocBridgeClaudeMcpStateMatches -State $claudeInstalledState -Command $mcpExe -Arguments @('--stdio'))
                    if ($claudeConfigured) {
                        $claudeCodeOwnership = [ordered]@{
                            installedCommand = $mcpExe
                            installedArgs = @('--stdio')
                            installedEnvironment = [ordered]@{}
                        }
                        Write-Host 'Claude Code user-scope MCP configured and verified.'
                    } else {
                        $stateOutput = if ($null -ne $claudeInstalledState) { $claudeInstalledState.Output } else { '' }
                        $criticalErrors.Add("Claude Code MCP registration failed or its command/args/environment could not be verified. add=$($claudeAdd.Output); get=$stateOutput")
                    }
                }
            }
        }
    }
}

if ($selectedClients -contains 'Codex') {
    $codexExecutable = Get-DocBridgeCodexCli -UserProfileRoot $UserProfileRoot -LocalAppDataRoot $LocalAppDataRoot -ExplicitPath $CodexCliPath
    $codexDetected = Test-DocBridgeCodexInstalled -UserProfileRoot $UserProfileRoot -LocalAppDataRoot $LocalAppDataRoot
    $marketplaceName = [string]$installedMarketplace.name
    if (-not [string]::IsNullOrWhiteSpace($codexExecutable)) {
        $codexCliMode = Get-DocBridgeCodexPluginCliMode -FilePath $codexExecutable -Environment $codexEnvironment
    }
    $commandPrefix = if ([string]::IsNullOrWhiteSpace($codexExecutable)) { 'codex' } else { '& "' + $codexExecutable + '"' }
    $commandLines = New-Object System.Collections.Generic.List[string]
    [void]$commandLines.Add("$commandPrefix plugin marketplace add `"$targetMarketplace`"")
    if ($codexCliMode -eq 'LegacyPluginCommands') {
        [void]$commandLines.Add("$commandPrefix plugin add doc-bridge@$marketplaceName")
    } else {
        [void]$commandLines.Add('# Codex를 완전히 종료한 뒤 다시 실행하고 새 작업을 여세요.')
        [void]$commandLines.Add('# 필요하면 Codex의 Plugins 화면 또는 /plugins에서 DocBridge를 확인하세요.')
    }
    $commandText = ($commandLines -join "`r`n") + "`r`n"
    Write-DocBridgeUtf8NoBom -Path (Join-Path $generatedRoot 'install-codex.txt') -Text $commandText
    if ($SkipClientCommands) {
        $warnings.Add('Codex CLI registration was skipped. See generated-configs\install-codex.txt.')
    } elseif ([string]::IsNullOrWhiteSpace($codexExecutable)) {
        $message = 'Codex is installed, but its working CLI could not be found. Open Codex once, close it completely, and run 1-INSTALL.cmd again.'
        if ($codexDetected) { $criticalErrors.Add($message) }
        else { $warnings.Add('Codex is not installed on this PC; Codex registration was skipped.') }
    } else {
        Write-Host "Codex CLI found: $codexExecutable ($codexCliMode)"
        $marketplaceResult = Invoke-DocBridgeNative -FilePath $codexExecutable -Arguments @('plugin', 'marketplace', 'add', $targetMarketplace) -Environment $codexEnvironment
        if ($marketplaceResult.ExitCode -ne 0) {
            $criticalErrors.Add("Codex marketplace registration failed: $($marketplaceResult.Output)")
        } else {
            $registeredState = Get-DocBridgeCodexMarketplaceState -MarketplaceName $marketplaceName `
                -UserProfileRoot $UserProfileRoot -CodexHomeRoot $CodexHomeRoot
            $expectedMarketplaceSource = ConvertTo-DocBridgeNormalizedLocalPath -Path $targetMarketplace
            $marketplaceOwned = Test-DocBridgeCodexMarketplaceStateMatches -State $registeredState `
                -SourceType 'local' -NormalizedSource $expectedMarketplaceSource
            if (-not $marketplaceOwned) {
                $criticalErrors.Add("Codex marketplace command succeeded, but source ownership could not be verified: source_type=$($registeredState.SourceType); source=$($registeredState.Source); error=$($registeredState.Error)")
            } else {
                $codexMarketplaceOwnership = [ordered]@{
                    marketplaceName = $marketplaceName
                    sourceType = 'local'
                    normalizedSource = $expectedMarketplaceSource
                }
            }
            if ($marketplaceOwned -and $codexCliMode -eq 'LegacyPluginCommands') {
                $pluginResult = Invoke-DocBridgeNative -FilePath $codexExecutable -Arguments @('plugin', 'add', "doc-bridge@$marketplaceName") -Environment $codexEnvironment
                $listResult = Invoke-DocBridgeNative -FilePath $codexExecutable -Arguments @('plugin', 'list') -Environment $codexEnvironment
                $pluginEnabled = $pluginResult.ExitCode -eq 0 -and $listResult.ExitCode -eq 0 -and
                    (Test-DocBridgeCodexPluginEnabled -Output $listResult.Output -MarketplaceName $marketplaceName)
                $mcpListResult = Invoke-DocBridgeNative -FilePath $codexExecutable -Arguments @('mcp', 'list') -Environment $codexEnvironment
                $mcpEnabled = $mcpListResult.ExitCode -eq 0 -and (Test-DocBridgeCodexMcpEnabled -Output $mcpListResult.Output)
                if ($pluginEnabled -and $mcpEnabled) {
                    Write-Host "Codex plugin and MCP installed and enabled: doc-bridge@$marketplaceName"
                } else {
                    $matchingLine = @($listResult.Output -split "`r?`n" | Where-Object { $_ -match 'doc-bridge@' }) -join ' | '
                    $matchingMcpLine = @($mcpListResult.Output -split "`r?`n" | Where-Object { $_ -match '^\s*doc-bridge\s+' }) -join ' | '
                    $criticalErrors.Add("Codex plugin or MCP is not enabled. add=$($pluginResult.Output); plugin=$matchingLine; mcp=$matchingMcpLine")
                }
            } elseif ($marketplaceOwned) {
                $mcpListResult = Invoke-DocBridgeNative -FilePath $codexExecutable -Arguments @('mcp', 'list') -Environment $codexEnvironment
                $mcpEnabled = $mcpListResult.ExitCode -eq 0 -and (Test-DocBridgeCodexMcpEnabled -Output $mcpListResult.Output)
                if ($mcpEnabled) {
                    Write-Host "Codex marketplace and MCP registered: doc-bridge@$marketplaceName"
                } else {
                    $warnings.Add('Codex marketplace is registered. The currently running Codex process must be fully restarted before DocBridge appears in a new task.')
                }
            }
        }
    }
}

$installation = [ordered]@{
    product = 'DocBridge'
    version = [string]$pluginManifest.version
    installedAt = (Get-Date).ToString('o')
    installRoot = $InstallRoot
    marketplaceRoot = $targetMarketplace
    marketplaceName = [string]$installedMarketplace.name
    mcpExecutable = $mcpExe
    cliExecutable = $cliExe
    clients = @($selectedClients)
    installedClients = @($installedClients.ToArray())
    clientMcpProvenance = $clientMcpProvenance
    claudeCodeOwnership = $claudeCodeOwnership
    codexMarketplaceOwnership = $codexMarketplaceOwnership
    hwpSecurityRequested = -not [bool]$SkipHwpSecurity
    codexCli = $codexExecutable
    codexCliMode = $codexCliMode
    codexHome = $CodexHomeRoot
} | ConvertTo-Json -Depth 8
Write-DocBridgeUtf8NoBom -Path (Join-Path $InstallRoot 'installation.json') -Text $installation

if (-not $SkipDoctor) {
    $doctor = Join-Path $PSScriptRoot 'Test-DocBridge.ps1'
    if (Test-Path -LiteralPath $doctor) {
        $doctorArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $doctor,
            '-InstallRoot', $InstallRoot, '-UserProfileRoot', $UserProfileRoot,
            '-AppDataRoot', $AppDataRoot, '-LocalAppDataRoot', $LocalAppDataRoot,
            '-ClientsCsv', ($selectedClients -join ','))
        if ($SkipHwpSecurity) { $doctorArgs += '-SkipHwpSecurity' }
        if ($SkipHwpRuntimeDoctor) { $doctorArgs += '-SkipHwpRuntimeDoctor' }
        if ($SkipClientCommands) { $doctorArgs += '-SkipClientCommands' }
        if ($selectedClients -contains 'Codex' -and $codexCliMode -ne 'LegacyPluginCommands') { $doctorArgs += '-AllowCodexRestartPending' }
        if (-not [string]::IsNullOrWhiteSpace($codexExecutable)) { $doctorArgs += @('-CodexCliPath', $codexExecutable) }
        if (-not [string]::IsNullOrWhiteSpace($CodexHomeRoot)) { $doctorArgs += @('-CodexHomeRoot', $CodexHomeRoot) }
        & powershell.exe @doctorArgs
        if ($LASTEXITCODE -ne 0) { $criticalErrors.Add('Post-install doctor found one or more critical failures.') }
    }
}

if ($criticalErrors.Count -gt 0) {
    Write-Host "`nDocBridge installation is incomplete." -ForegroundColor Red
    $criticalErrors | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    throw 'DocBridge client registration or verification failed. Fix the errors above and run 1-INSTALL.cmd again.'
}

Write-Host "`nDocBridge installation completed." -ForegroundColor Green
Write-Host "MCP: $mcpExe"
Write-Host 'Restart Cursor, Claude Desktop, and Codex completely before using the plugin.'
if ($warnings.Count -gt 0) {
    Write-Host 'Warnings:' -ForegroundColor Yellow
    $warnings | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
}
