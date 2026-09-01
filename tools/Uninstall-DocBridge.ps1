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
    [switch]$RemoveHwpSecurity,
    [switch]$RemoveData,
    [switch]$SkipClientCommands
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
$installation = $null
$installationPath = Join-Path $InstallRoot 'installation.json'
if (Test-Path -LiteralPath $installationPath -PathType Leaf) {
    try {
        $installation = Get-Content -LiteralPath $installationPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        throw "Cannot safely uninstall with an unreadable installation record: $installationPath`n$($_.Exception.Message)"
    }
}
if (-not $PSBoundParameters.ContainsKey('Clients')) {
    if ($null -ne $installation -and $null -ne $installation.PSObject.Properties['installedClients']) {
        $Clients = @($installation.installedClients | ForEach-Object { [string]$_ })
    } elseif ($null -ne $installation -and $null -ne $installation.PSObject.Properties['clients']) {
        $Clients = @($installation.clients | ForEach-Object { [string]$_ })
    } else {
        $Clients = @('All')
    }
}
$selectedClients = Get-DocBridgeClientSet -Clients $Clients
$managedClients = New-Object System.Collections.Generic.List[string]
$recordedClients = @()
if ($null -ne $installation -and $null -ne $installation.PSObject.Properties['installedClients']) {
    $recordedClients = @($installation.installedClients)
} elseif ($null -ne $installation -and $null -ne $installation.PSObject.Properties['clients']) {
    $recordedClients = @($installation.clients)
} else {
    $recordedClients = @($selectedClients)
}
foreach ($client in $recordedClients) {
    $name = [string]$client
    if (-not [string]::IsNullOrWhiteSpace($name) -and -not $managedClients.Contains($name)) {
        [void]$managedClients.Add($name)
    }
}
$remainingClients = @($managedClients | Where-Object { $selectedClients -notcontains $_ })
$isFinalManagedRemoval = $remainingClients.Count -eq 0
if ($RemoveData -and -not $isFinalManagedRemoval) {
    throw '-RemoveData cannot be used for a partial client uninstall. Remove the last managed client first.'
}
$marketplaceRoot = Join-Path $InstallRoot 'codex-marketplace'
$backupRoot = Join-Path $InstallRoot 'backups'
$generatedRoot = Join-Path $InstallRoot 'generated-configs'
$warnings = New-Object System.Collections.Generic.List[string]
$removed = New-Object System.Collections.Generic.List[string]
$codexEnvironment = $null
if (-not [string]::IsNullOrWhiteSpace($CodexHomeRoot)) {
    $CodexHomeRoot = [System.IO.Path]::GetFullPath($CodexHomeRoot).TrimEnd('\')
    $codexEnvironment = @{ CODEX_HOME = $CodexHomeRoot }
}

New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null

function Get-ClientProvenance([string]$ClientName, [string]$FallbackPath) {
    $record = $null
    if ($null -ne $installation -and
        $null -ne $installation.PSObject.Properties['clientMcpProvenance'] -and
        $null -ne $installation.clientMcpProvenance) {
        $property = $installation.clientMcpProvenance.PSObject.Properties[$ClientName]
        if ($null -ne $property) { $record = $property.Value }
    }
    if ($null -ne $record) { return $record }
    if (-not $managedClients.Contains($ClientName)) {
        return $null
    }
    if ($null -eq $installation -or
        $null -eq $installation.PSObject.Properties['mcpExecutable'] -or
        [string]::IsNullOrWhiteSpace([string]$installation.mcpExecutable)) {
        return $null
    }
    return [pscustomobject][ordered]@{
        path = $FallbackPath
        previousEntryExisted = $false
        previousEntry = $null
        installedCommand = [string]$installation.mcpExecutable
        installedArgs = @('--stdio')
        upgradedFromLegacyRecord = $true
    }
}

function Remove-OwnedClientMcpEntry([string]$ClientName, [string]$FallbackPath, [string]$DisplayName) {
    $provenance = Get-ClientProvenance -ClientName $ClientName -FallbackPath $FallbackPath
    if ($null -eq $provenance) {
        $warnings.Add("$DisplayName configuration was preserved because no installation ownership record exists: $FallbackPath")
        return
    }
    $path = if ($null -ne $provenance.PSObject.Properties['path'] -and
        -not [string]::IsNullOrWhiteSpace([string]$provenance.path)) { [string]$provenance.path } else { $FallbackPath }
    $expectedCommand = if ($null -ne $provenance.PSObject.Properties['installedCommand']) {
        [string]$provenance.installedCommand
    } else { [string]$installation.mcpExecutable }
    if ([string]::IsNullOrWhiteSpace($expectedCommand)) {
        $warnings.Add("$DisplayName configuration was preserved because its installation ownership command is missing: $path")
        return
    }
    $expectedArguments = @('--stdio')
    if ($null -ne $provenance.PSObject.Properties['installedArgs'] -and $null -ne $provenance.installedArgs) {
        $expectedArguments = @($provenance.installedArgs | ForEach-Object { [string]$_ })
    }
    $previousExisted = $null -ne $provenance.PSObject.Properties['previousEntryExisted'] -and
        [bool]$provenance.previousEntryExisted
    $previousEntry = if ($null -ne $provenance.PSObject.Properties['previousEntry']) { $provenance.previousEntry } else { $null }
    $result = Remove-DocBridgeMcpJsonEntry -Path $path -BackupRoot $backupRoot `
        -ExpectedCommand $expectedCommand -ExpectedArguments $expectedArguments `
        -PreviousEntryExisted $previousExisted -PreviousEntry $previousEntry
    if ($result.Changed) {
        [void]$removed.Add("$DisplayName ($($result.Action)): $path")
    } elseif ($result.Action -eq 'preserved-user-change') {
        $warnings.Add("$DisplayName doc-bridge entry was changed after installation and was preserved: $path")
    }
}

if ($selectedClients -contains 'ClaudeDesktop' -and -not [string]::IsNullOrWhiteSpace($AppDataRoot)) {
    $path = Join-Path $AppDataRoot 'Claude\claude_desktop_config.json'
    Remove-OwnedClientMcpEntry -ClientName 'ClaudeDesktop' -FallbackPath $path -DisplayName 'Claude Desktop'
}

if ($selectedClients -contains 'Kimi' -and -not [string]::IsNullOrWhiteSpace($UserProfileRoot)) {
    $path = Join-Path $UserProfileRoot '.kimi\mcp.json'
    Remove-OwnedClientMcpEntry -ClientName 'Kimi' -FallbackPath $path -DisplayName 'Kimi'
}

if ($selectedClients -contains 'Cursor' -and -not [string]::IsNullOrWhiteSpace($UserProfileRoot)) {
    $path = Join-Path $UserProfileRoot '.cursor\mcp.json'
    Remove-OwnedClientMcpEntry -ClientName 'Cursor' -FallbackPath $path -DisplayName 'Cursor'
}

if ($selectedClients -contains 'ClaudeCode') {
    $expectedClaudeCommand = $null
    if ($null -ne $installation -and
        $null -ne $installation.PSObject.Properties['claudeCodeOwnership'] -and
        $null -ne $installation.claudeCodeOwnership -and
        $null -ne $installation.claudeCodeOwnership.PSObject.Properties['installedCommand']) {
        $expectedClaudeCommand = [string]$installation.claudeCodeOwnership.installedCommand
    } elseif ($null -ne $installation -and $null -ne $installation.PSObject.Properties['mcpExecutable']) {
        $expectedClaudeCommand = [string]$installation.mcpExecutable
    }
    $expectedClaudeArgs = @('--stdio')
    if ($null -ne $installation -and
        $null -ne $installation.PSObject.Properties['claudeCodeOwnership'] -and
        $null -ne $installation.claudeCodeOwnership -and
        $null -ne $installation.claudeCodeOwnership.PSObject.Properties['installedArgs']) {
        $expectedClaudeArgs = @($installation.claudeCodeOwnership.installedArgs)
    }
    $manual = "claude mcp get doc-bridge`r`n# Remove only when Command equals: $expectedClaudeCommand`r`n# Args must equal: $($expectedClaudeArgs -join ' ')`r`n# Environment must be empty.`r`nclaude mcp remove --scope user doc-bridge"
    Write-DocBridgeUtf8NoBom -Path (Join-Path $generatedRoot 'uninstall-claude-code.txt') -Text ($manual + "`r`n")
    $claude = Get-Command claude -ErrorAction SilentlyContinue
    if ($SkipClientCommands -or $null -eq $claude) {
        $warnings.Add('Claude Code CLI removal was skipped. See generated-configs\uninstall-claude-code.txt.')
    } elseif ([string]::IsNullOrWhiteSpace($expectedClaudeCommand)) {
        $warnings.Add('Claude Code doc-bridge entry was preserved because no installation ownership record exists.')
    } else {
        $claudeState = Get-DocBridgeClaudeMcpState -FilePath $claude.Source
        if ($null -ne $claudeState.Error) {
            $warnings.Add("Claude Code MCP ownership query failed; the entry was preserved: $($claudeState.Error) $($claudeState.Output)")
        } elseif (-not $claudeState.Exists) {
            [void]$removed.Add('Claude Code user-scope MCP was already absent')
        } elseif (-not (Test-DocBridgeClaudeMcpStateMatches -State $claudeState -Command $expectedClaudeCommand -Arguments $expectedClaudeArgs)) {
            $warnings.Add("Claude Code doc-bridge entry was changed, foreign, or could not be fully ownership-matched and was preserved: command=$($claudeState.Command); args=$($claudeState.ArgsText); environment=$($claudeState.Environment)")
        } else {
            $claudeRemove = Invoke-DocBridgeNative -FilePath $claude.Source -Arguments @('mcp', 'remove', '--scope', 'user', 'doc-bridge')
            if ($claudeRemove.ExitCode -ne 0) {
                $warnings.Add("Claude Code MCP removal failed (exit $($claudeRemove.ExitCode)): $($claudeRemove.Output)")
            } else { [void]$removed.Add('Claude Code user-scope MCP') }
        }
    }
}

if ($selectedClients -contains 'Codex') {
    $marketplaceName = 'docbridge-local'
    $marketplaceOwnership = $null
    if ($null -ne $installation -and
        $null -ne $installation.PSObject.Properties['codexMarketplaceOwnership'] -and
        $null -ne $installation.codexMarketplaceOwnership) {
        $marketplaceOwnership = $installation.codexMarketplaceOwnership
        if ($null -ne $marketplaceOwnership.PSObject.Properties['marketplaceName'] -and
            -not [string]::IsNullOrWhiteSpace([string]$marketplaceOwnership.marketplaceName)) {
            $marketplaceName = [string]$marketplaceOwnership.marketplaceName
        }
    }
    $marketplaceJson = Join-Path $marketplaceRoot '.agents\plugins\marketplace.json'
    if (Test-Path -LiteralPath $marketplaceJson) {
        try {
            $marketplaceName = [string](Get-Content -LiteralPath $marketplaceJson -Raw -Encoding UTF8 | ConvertFrom-Json).name
        } catch { $warnings.Add("Could not read marketplace metadata: $($_.Exception.Message)") }
    }
    $codexExecutable = Get-DocBridgeCodexCli -UserProfileRoot $UserProfileRoot -LocalAppDataRoot $LocalAppDataRoot -ExplicitPath $CodexCliPath
    $codexCliMode = if ([string]::IsNullOrWhiteSpace($codexExecutable)) { 'Unknown' } else {
        Get-DocBridgeCodexPluginCliMode -FilePath $codexExecutable -Environment $codexEnvironment
    }
    $commandPrefix = if ([string]::IsNullOrWhiteSpace($codexExecutable)) { 'codex' } else { '& "' + $codexExecutable + '"' }
    $manualLines = New-Object System.Collections.Generic.List[string]
    $expectedSourceType = if ($null -ne $marketplaceOwnership -and
        $null -ne $marketplaceOwnership.PSObject.Properties['sourceType']) { [string]$marketplaceOwnership.sourceType } else { '' }
    $expectedSource = if ($null -ne $marketplaceOwnership -and
        $null -ne $marketplaceOwnership.PSObject.Properties['normalizedSource']) { [string]$marketplaceOwnership.normalizedSource } else { '' }
    [void]$manualLines.Add("# First inspect [marketplaces.$marketplaceName] in config.toml.")
    [void]$manualLines.Add("# Remove only when source_type='$expectedSourceType' and normalized source='$expectedSource'.")
    if ($codexCliMode -eq 'LegacyPluginCommands') {
        [void]$manualLines.Add("$commandPrefix plugin remove doc-bridge@$marketplaceName")
    }
    [void]$manualLines.Add("$commandPrefix plugin marketplace remove $marketplaceName")
    $manual = ($manualLines -join "`r`n") + "`r`n"
    Write-DocBridgeUtf8NoBom -Path (Join-Path $generatedRoot 'uninstall-codex.txt') -Text $manual
    $marketplaceState = Get-DocBridgeCodexMarketplaceState -MarketplaceName $marketplaceName `
        -UserProfileRoot $UserProfileRoot -CodexHomeRoot $CodexHomeRoot
    $marketplaceRemovalAllowed = $false
    if ($null -ne $marketplaceState.Error) {
        $warnings.Add("Codex marketplace ownership could not be parsed; it was preserved: $($marketplaceState.Error)")
    } elseif (-not $marketplaceState.Exists) {
        [void]$removed.Add("Codex marketplace was already absent: $marketplaceName")
    } elseif ($null -eq $marketplaceOwnership -or
        [string]::IsNullOrWhiteSpace($expectedSourceType) -or
        [string]::IsNullOrWhiteSpace($expectedSource)) {
        $warnings.Add("Codex marketplace '$marketplaceName' was preserved because no installation ownership record exists.")
    } elseif (-not (Test-DocBridgeCodexMarketplaceStateMatches -State $marketplaceState `
        -SourceType $expectedSourceType -NormalizedSource $expectedSource)) {
        $warnings.Add("Codex marketplace '$marketplaceName' source was changed or is foreign and was preserved: source_type=$($marketplaceState.SourceType); source=$($marketplaceState.Source)")
    } else {
        $marketplaceRemovalAllowed = $true
    }
    if (-not $marketplaceRemovalAllowed) {
        # The ownership decision above is authoritative. Never invoke a broad
        # same-name CLI removal when source provenance is missing or changed.
    } elseif ($SkipClientCommands -or [string]::IsNullOrWhiteSpace($codexExecutable)) {
        $warnings.Add('Codex CLI removal was skipped. See generated-configs\uninstall-codex.txt.')
    } else {
        if ($codexCliMode -eq 'LegacyPluginCommands') {
            $removePlugin = Invoke-DocBridgeNative -FilePath $codexExecutable -Arguments @('plugin', 'remove', "doc-bridge@$marketplaceName") -Environment $codexEnvironment
            if ($removePlugin.ExitCode -ne 0) { $warnings.Add('Codex plugin removal failed; use the generated manual command.') }
            else { [void]$removed.Add("Codex plugin: doc-bridge@$marketplaceName") }
        }
        $removeMarketplace = Invoke-DocBridgeNative -FilePath $codexExecutable -Arguments @('plugin', 'marketplace', 'remove', $marketplaceName) -Environment $codexEnvironment
        if ($removeMarketplace.ExitCode -ne 0) { $warnings.Add('Codex marketplace removal failed; use the generated manual command.') }
        else { [void]$removed.Add("Codex marketplace: $marketplaceName") }
    }
}

if ($RemoveHwpSecurity -and -not $isFinalManagedRemoval) {
    $warnings.Add("HWP security registration was retained because managed clients remain: $($remainingClients -join ', ')")
} elseif ($RemoveHwpSecurity) {
    $securityUninstaller = Join-Path $marketplaceRoot 'plugins\doc-bridge\dist\uninstall-hwp-security.ps1'
    if (Test-Path -LiteralPath $securityUninstaller) {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $securityUninstaller
        if ($LASTEXITCODE -ne 0) { $warnings.Add("HWP security module removal failed (exit $LASTEXITCODE).") }
        else { [void]$removed.Add('HWP security module registration') }
    } else {
        $warnings.Add('HWP security uninstaller was not found; registry entry was not changed.')
    }
}

if ($isFinalManagedRemoval -and (Test-Path -LiteralPath $marketplaceRoot)) {
    $safeMarketplace = Assert-DocBridgeSafePath -Path $marketplaceRoot -Root $InstallRoot
    $archive = Join-Path $backupRoot ('uninstalled-codex-marketplace-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    Move-Item -LiteralPath $safeMarketplace -Destination $archive
    [void]$removed.Add("Installed payload archived to: $archive")
} elseif (-not $isFinalManagedRemoval) {
    [void]$removed.Add("Shared payload retained for remaining clients: $($remainingClients -join ', ')")
}

if (-not $isFinalManagedRemoval -and $null -ne $installation) {
    foreach ($propertyName in @('clients', 'installedClients')) {
        $property = $installation.PSObject.Properties[$propertyName]
        if ($null -eq $property) {
            $installation | Add-Member -MemberType NoteProperty -Name $propertyName -Value @($remainingClients)
        } else {
            $property.Value = @($remainingClients)
        }
    }
    if ($null -ne $installation.PSObject.Properties['clientMcpProvenance'] -and
        $null -ne $installation.clientMcpProvenance) {
        foreach ($client in $selectedClients) {
            $installation.clientMcpProvenance.PSObject.Properties.Remove([string]$client)
        }
    }
    if ($selectedClients -contains 'ClaudeCode' -and
        $null -ne $installation.PSObject.Properties['claudeCodeOwnership']) {
        $installation.claudeCodeOwnership = $null
    }
    if ($selectedClients -contains 'Codex' -and
        $null -ne $installation.PSObject.Properties['codexMarketplaceOwnership']) {
        $installation.codexMarketplaceOwnership = $null
    }
    $updatedAt = $installation.PSObject.Properties['updatedAt']
    if ($null -eq $updatedAt) {
        $installation | Add-Member -MemberType NoteProperty -Name 'updatedAt' -Value (Get-Date).ToString('o')
    } else {
        $updatedAt.Value = (Get-Date).ToString('o')
    }
    Write-DocBridgeUtf8NoBom -Path $installationPath -Text ($installation | ConvertTo-Json -Depth 50)
} elseif ($isFinalManagedRemoval -and (Test-Path -LiteralPath $installationPath -PathType Leaf)) {
    $installationArchive = Join-Path $backupRoot ('installation.json.' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff') + '.uninstalled.bak')
    Move-Item -LiteralPath $installationPath -Destination $installationArchive
    [void]$removed.Add("Installation ownership record archived to: $installationArchive")
}

$report = [ordered]@{
    product = 'DocBridge'
    uninstalledAt = (Get-Date).ToString('o')
    installRoot = $InstallRoot
    clients = @($selectedClients)
    remainingClients = @($remainingClients)
    finalRemoval = [bool]$isFinalManagedRemoval
    removed = @($removed.ToArray())
    warnings = @($warnings.ToArray())
    dataRemoved = [bool]$RemoveData
} | ConvertTo-Json -Depth 8
Write-DocBridgeUtf8NoBom -Path (Join-Path $InstallRoot 'uninstall-report.json') -Text $report

if ($RemoveData) {
    if ((Split-Path -Leaf $InstallRoot) -ne 'DocBridge') {
        throw "Refusing -RemoveData because the install folder is not named DocBridge: $InstallRoot"
    }
    $parent = Split-Path -Parent $InstallRoot
    $safeInstallRoot = Assert-DocBridgeSafePath -Path $InstallRoot -Root $parent
    Remove-Item -LiteralPath $safeInstallRoot -Recurse -Force
    Write-Host "Removed installation data: $safeInstallRoot"
} else {
    Write-Host "Backups and reports retained at: $InstallRoot"
}

Write-Host 'DocBridge uninstall completed.' -ForegroundColor Green
if ($warnings.Count -gt 0) {
    Write-Host 'Warnings:' -ForegroundColor Yellow
    $warnings | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
}
