Set-StrictMode -Version 2.0

function Get-DocBridgeDefaultRoot {
    param([string]$LocalAppDataRoot)
    if ([string]::IsNullOrWhiteSpace($LocalAppDataRoot)) {
        $LocalAppDataRoot = $env:LOCALAPPDATA
    }
    if ([string]::IsNullOrWhiteSpace($LocalAppDataRoot)) {
        throw 'LOCALAPPDATA is unavailable. Specify -InstallRoot explicitly.'
    }
    return [System.IO.Path]::GetFullPath((Join-Path $LocalAppDataRoot 'DocBridge'))
}

function Assert-DocBridgeSafePath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Root
    )
    $resolvedPath = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    $resolvedRoot = [System.IO.Path]::GetFullPath($Root).TrimEnd('\')
    if ($resolvedPath.Length -le $resolvedRoot.Length -or
        -not $resolvedPath.StartsWith($resolvedRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing unsafe deployment path: $resolvedPath (allowed root: $resolvedRoot)"
    }
    return $resolvedPath
}

function Backup-DocBridgeFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$BackupRoot
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $leaf = Split-Path -Leaf $Path
    $target = Join-Path $BackupRoot "$leaf.$stamp.bak"
    Copy-Item -LiteralPath $Path -Destination $target -Force
    return $target
}

function Read-DocBridgeJsonObject {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{}
    }
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) { return [pscustomobject]@{} }
    try {
        $value = $raw | ConvertFrom-Json
        if ($null -eq $value -or $value -is [System.Array] -or $value -is [string] -or $value -is [ValueType]) {
            throw 'The JSON root must be an object.'
        }
        return $value
    } catch {
        throw "Cannot safely merge invalid JSON config: $Path`n$($_.Exception.Message)"
    }
}

function Assert-DocBridgeMcpServersObject {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Path
    )
    if ($null -eq $Config.PSObject.Properties['mcpServers']) {
        $Config | Add-Member -MemberType NoteProperty -Name 'mcpServers' -Value ([pscustomobject]@{})
    }
    $servers = $Config.mcpServers
    if ($null -eq $servers -or $servers -is [System.Array] -or $servers -is [string] -or $servers -is [ValueType]) {
        throw "mcpServers must be a JSON object: $Path"
    }
    return $servers
}

function Set-DocBridgeMcpJsonEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Command,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [Parameter(Mandatory = $true)][string]$BackupRoot
    )
    $parent = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
    [void](Backup-DocBridgeFile -Path $Path -BackupRoot $BackupRoot)
    $config = Read-DocBridgeJsonObject -Path $Path
    $servers = Assert-DocBridgeMcpServersObject -Config $config -Path $Path
    $entry = [pscustomobject][ordered]@{
        command = $Command
        args = @($Arguments)
    }
    $existing = $servers.PSObject.Properties['doc-bridge']
    if ($null -eq $existing) {
        $servers | Add-Member -MemberType NoteProperty -Name 'doc-bridge' -Value $entry
    } else {
        $existing.Value = $entry
    }
    $json = $config | ConvertTo-Json -Depth 50
    Write-DocBridgeUtf8NoBom -Path $Path -Text $json
    return $Path
}

function Get-DocBridgeMcpJsonEntryState {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fileExisted = Test-Path -LiteralPath $Path -PathType Leaf
    $config = Read-DocBridgeJsonObject -Path $Path
    if ($null -eq $config.PSObject.Properties['mcpServers']) {
        return [pscustomobject][ordered]@{
            path = $Path
            fileExisted = [bool]$fileExisted
            entryExisted = $false
            entry = $null
        }
    }

    $servers = Assert-DocBridgeMcpServersObject -Config $config -Path $Path
    $property = $servers.PSObject.Properties['doc-bridge']
    return [pscustomobject][ordered]@{
        path = $Path
        fileExisted = [bool]$fileExisted
        entryExisted = $null -ne $property
        entry = if ($null -eq $property) { $null } else { $property.Value }
    }
}

function Test-DocBridgeMcpJsonEntryMatches {
    param(
        [Parameter(Mandatory = $true)]$Entry,
        [Parameter(Mandatory = $true)][string]$Command,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    if ($null -eq $Entry -or $null -eq $Entry.PSObject.Properties['command']) { return $false }
    $actualCommand = [string]$Entry.command
    if (-not $actualCommand.Equals($Command, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
    $actualArguments = @()
    if ($null -ne $Entry.PSObject.Properties['args']) {
        $actualArguments = @($Entry.args | ForEach-Object { [string]$_ })
    }
    if ($actualArguments.Count -ne $Arguments.Count) { return $false }
    for ($index = 0; $index -lt $Arguments.Count; $index++) {
        if (-not $actualArguments[$index].Equals($Arguments[$index], [System.StringComparison]::Ordinal)) {
            return $false
        }
    }
    return $true
}

function Remove-DocBridgeMcpJsonEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$BackupRoot,
        [Parameter(Mandatory = $true)][string]$ExpectedCommand,
        [string[]]$ExpectedArguments = @('--stdio'),
        [bool]$PreviousEntryExisted = $false,
        $PreviousEntry
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{ Changed = $false; Action = 'missing'; Detail = $Path }
    }
    $config = Read-DocBridgeJsonObject -Path $Path
    if ($null -eq $config.PSObject.Properties['mcpServers']) {
        return [pscustomobject]@{ Changed = $false; Action = 'not-found'; Detail = $Path }
    }
    $servers = Assert-DocBridgeMcpServersObject -Config $config -Path $Path
    $current = $servers.PSObject.Properties['doc-bridge']
    if ($null -eq $current) {
        return [pscustomobject]@{ Changed = $false; Action = 'not-found'; Detail = $Path }
    }
    if ([string]::IsNullOrWhiteSpace($ExpectedCommand)) {
        return [pscustomobject]@{
            Changed = $false
            Action = 'preserved-unverified'
            Detail = 'The installation ownership record has no command to match.'
        }
    }
    if (-not (Test-DocBridgeMcpJsonEntryMatches -Entry $current.Value -Command $ExpectedCommand -Arguments $ExpectedArguments)) {
        return [pscustomobject]@{
            Changed = $false
            Action = 'preserved-user-change'
            Detail = 'The current doc-bridge entry no longer matches the entry installed by DocBridge.'
        }
    }
    [void](Backup-DocBridgeFile -Path $Path -BackupRoot $BackupRoot)
    if ($PreviousEntryExisted) {
        if ($null -eq $PreviousEntry) {
            throw "Cannot restore a missing previous doc-bridge entry: $Path"
        }
        $current.Value = $PreviousEntry
        $action = 'restored-previous'
    } else {
        $servers.PSObject.Properties.Remove('doc-bridge')
        $action = 'removed'
    }
    $json = $config | ConvertTo-Json -Depth 50
    Write-DocBridgeUtf8NoBom -Path $Path -Text $json
    return [pscustomobject]@{ Changed = $true; Action = $action; Detail = $Path }
}

function Get-DocBridgeClientSet {
    param([string[]]$Clients)
    $all = @('Codex', 'ClaudeCode', 'ClaudeDesktop', 'Kimi', 'Cursor')
    if ($null -eq $Clients -or $Clients.Count -eq 0 -or $Clients -contains 'All') { return $all }
    return @($all | Where-Object { $Clients -contains $_ })
}

function Test-DocBridgeCodexInstalled {
    param(
        [string]$UserProfileRoot = $env:USERPROFILE,
        [string]$LocalAppDataRoot = $env:LOCALAPPDATA
    )
    if (-not [string]::IsNullOrWhiteSpace($LocalAppDataRoot) -and
        (Test-Path -LiteralPath (Join-Path $LocalAppDataRoot 'OpenAI\Codex'))) {
        return $true
    }
    if (-not [string]::IsNullOrWhiteSpace($UserProfileRoot) -and
        (Test-Path -LiteralPath (Join-Path $UserProfileRoot '.codex\config.toml') -PathType Leaf)) {
        return $true
    }
    try {
        if ($null -ne (Get-Command Get-AppxPackage -ErrorAction SilentlyContinue)) {
            return $null -ne (Get-AppxPackage -Name 'OpenAI.Codex' -ErrorAction SilentlyContinue | Select-Object -First 1)
        }
    } catch { }
    return $null -ne (Get-Command codex -ErrorAction SilentlyContinue)
}

function Get-DocBridgeCodexCli {
    param(
        [string]$UserProfileRoot = $env:USERPROFILE,
        [string]$LocalAppDataRoot = $env:LOCALAPPDATA,
        [string]$ExplicitPath
    )

    $candidates = New-Object System.Collections.Generic.List[string]
    function Add-CodexCandidate([string]$Candidate) {
        if ([string]::IsNullOrWhiteSpace($Candidate)) { return }
        try { $Candidate = [System.IO.Path]::GetFullPath($Candidate) } catch { return }
        if ($candidates -notcontains $Candidate) { [void]$candidates.Add($Candidate) }
    }

    Add-CodexCandidate $ExplicitPath

    if (-not [string]::IsNullOrWhiteSpace($UserProfileRoot)) {
        $configPath = Join-Path $UserProfileRoot '.codex\config.toml'
        if (Test-Path -LiteralPath $configPath -PathType Leaf) {
            try {
                $raw = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8
                $match = [regex]::Match($raw, "(?m)^\s*CODEX_CLI_PATH\s*=\s*'([^']+)'\s*$")
                if (-not $match.Success) {
                    $match = [regex]::Match($raw, '(?m)^\s*CODEX_CLI_PATH\s*=\s*"([^"]+)"\s*$')
                }
                if ($match.Success) { Add-CodexCandidate $match.Groups[1].Value }
            } catch { }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($LocalAppDataRoot)) {
        $binRoot = Join-Path $LocalAppDataRoot 'OpenAI\Codex\bin'
        if (Test-Path -LiteralPath $binRoot -PathType Container) {
            Get-ChildItem -LiteralPath $binRoot -Directory -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending |
                ForEach-Object { Add-CodexCandidate (Join-Path $_.FullName 'codex.exe') }
        }
        Add-CodexCandidate (Join-Path $LocalAppDataRoot 'Microsoft\WindowsApps\codex.exe')
    }

    $command = Get-Command codex -ErrorAction SilentlyContinue
    if ($null -ne $command) {
        if (-not [string]::IsNullOrWhiteSpace([string]$command.Source)) { Add-CodexCandidate $command.Source }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$command.Path)) { Add-CodexCandidate $command.Path }
    }

    try {
        if ($null -ne (Get-Command Get-AppxPackage -ErrorAction SilentlyContinue)) {
            Get-AppxPackage -Name 'OpenAI.Codex' -ErrorAction SilentlyContinue |
                Sort-Object Version -Descending |
                ForEach-Object { Add-CodexCandidate (Join-Path $_.InstallLocation 'app\resources\codex.exe') }
        }
    } catch { }

    foreach ($candidate in $candidates) {
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { continue }
        try {
            $probe = Invoke-DocBridgeNative -FilePath $candidate -Arguments @('--version')
            if ($probe.ExitCode -eq 0) { return $candidate }
        } catch { }
    }
    return $null
}

function Test-DocBridgeCodexPluginEnabled {
    param(
        [Parameter(Mandatory = $true)][string]$Output,
        [Parameter(Mandatory = $true)][string]$MarketplaceName
    )
    $selector = 'doc-bridge@' + [regex]::Escape($MarketplaceName)
    return [regex]::IsMatch($Output, '(?im)^\s*' + $selector + '\s+installed,\s*enabled(?:\s|$)')
}

function Get-DocBridgeCodexPluginCliMode {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [hashtable]$Environment
    )
    try {
        $help = Invoke-DocBridgeNative -FilePath $FilePath -Arguments @('plugin', '--help') -Environment $Environment
        if ($help.ExitCode -ne 0) { return 'Unknown' }
        if ($help.Output -match '(?m)^\s+add\s+' -and $help.Output -match '(?m)^\s+list\s+') {
            return 'LegacyPluginCommands'
        }
        if ($help.Output -match '(?m)^\s+marketplace\s+') {
            return 'MarketplaceOnly'
        }
    } catch { }
    return 'Unknown'
}

function Get-DocBridgeCodexConfigPath {
    param(
        [string]$UserProfileRoot = $env:USERPROFILE,
        [string]$CodexHomeRoot
    )
    if (-not [string]::IsNullOrWhiteSpace($CodexHomeRoot)) {
        return Join-Path $CodexHomeRoot 'config.toml'
    }
    if ([string]::IsNullOrWhiteSpace($UserProfileRoot)) { return $null }
    return Join-Path $UserProfileRoot '.codex\config.toml'
}

function ConvertTo-DocBridgeNormalizedLocalPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $candidate = $Path.Trim().Replace('/', '\')
    if ($candidate.StartsWith('\\?\UNC\', [System.StringComparison]::OrdinalIgnoreCase)) {
        $candidate = '\\' + $candidate.Substring(8)
    } elseif ($candidate.StartsWith('\\?\', [System.StringComparison]::OrdinalIgnoreCase)) {
        $candidate = $candidate.Substring(4)
    }
    if (-not [System.IO.Path]::IsPathRooted($candidate)) { return $null }
    try { return [System.IO.Path]::GetFullPath($candidate).TrimEnd('\') } catch { return $null }
}

function ConvertFrom-DocBridgeTomlString {
    param([Parameter(Mandatory = $true)][string]$Value)
    $trimmed = $Value.Trim()
    if ($trimmed -match "^'([^']*)'$") { return $matches[1] }
    if ($trimmed -match '^"(?:\\.|[^"\\])*"$') {
        try { return [string]($trimmed | ConvertFrom-Json) } catch { return $null }
    }
    return $null
}

function Get-DocBridgeCodexMarketplaceState {
    param(
        [Parameter(Mandatory = $true)][string]$MarketplaceName,
        [string]$UserProfileRoot = $env:USERPROFILE,
        [string]$CodexHomeRoot
    )
    $configPath = Get-DocBridgeCodexConfigPath -UserProfileRoot $UserProfileRoot -CodexHomeRoot $CodexHomeRoot
    if ([string]::IsNullOrWhiteSpace($configPath) -or
        -not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
        return [pscustomobject]@{
            Exists = $false; ConfigPath = $configPath; SourceType = $null; Source = $null
            NormalizedSource = $null; Error = $null
        }
    }
    try { $lines = @(Get-Content -LiteralPath $configPath -Encoding UTF8) }
    catch {
        return [pscustomobject]@{
            Exists = $null; ConfigPath = $configPath; SourceType = $null; Source = $null
            NormalizedSource = $null; Error = $_.Exception.Message
        }
    }
    $escapedName = [regex]::Escape($MarketplaceName)
    $sectionPattern = '^\s*\[marketplaces\.(?:' + $escapedName + '|"' + $escapedName + '"|''' + $escapedName + ''')\]\s*(?:#.*)?$'
    $sectionIndexes = New-Object System.Collections.Generic.List[int]
    for ($index = 0; $index -lt $lines.Count; $index++) {
        if ($lines[$index] -match $sectionPattern) { [void]$sectionIndexes.Add($index) }
    }
    if ($sectionIndexes.Count -eq 0) {
        return [pscustomobject]@{
            Exists = $false; ConfigPath = $configPath; SourceType = $null; Source = $null
            NormalizedSource = $null; Error = $null
        }
    }
    if ($sectionIndexes.Count -ne 1) {
        return [pscustomobject]@{
            Exists = $null; ConfigPath = $configPath; SourceType = $null; Source = $null
            NormalizedSource = $null; Error = "Duplicate marketplace sections were found for $MarketplaceName."
        }
    }
    $values = @{}
    for ($index = $sectionIndexes[0] + 1; $index -lt $lines.Count; $index++) {
        $line = [string]$lines[$index]
        if ($line -match '^\s*\[') { break }
        if ($line -match '^\s*(source_type|source)\s*=\s*(?<value>''[^'']*''|"(?:\\.|[^"\\])*")\s*(?:#.*)?$') {
            $values[$matches[1]] = ConvertFrom-DocBridgeTomlString -Value $matches['value']
        }
    }
    if (-not $values.ContainsKey('source_type') -or -not $values.ContainsKey('source') -or
        [string]::IsNullOrWhiteSpace([string]$values.source_type) -or
        [string]::IsNullOrWhiteSpace([string]$values.source)) {
        return [pscustomobject]@{
            Exists = $null; ConfigPath = $configPath; SourceType = [string]$values.source_type; Source = [string]$values.source
            NormalizedSource = $null; Error = "Marketplace source/source_type could not be parsed for $MarketplaceName."
        }
    }
    $normalized = ConvertTo-DocBridgeNormalizedLocalPath -Path ([string]$values.source)
    return [pscustomobject]@{
        Exists = $true
        ConfigPath = $configPath
        SourceType = [string]$values.source_type
        Source = [string]$values.source
        NormalizedSource = $normalized
        Error = $(if ([string]$values.source_type -ieq 'local' -and $null -eq $normalized) {
            "Local marketplace source is not an absolute path: $($values.source)"
        } else { $null })
    }
}

function Test-DocBridgeCodexMarketplaceStateMatches {
    param(
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][string]$SourceType,
        [Parameter(Mandatory = $true)][string]$NormalizedSource
    )
    if ($null -eq $State -or $State.Exists -ne $true -or $null -ne $State.Error) { return $false }
    if (-not ([string]$State.SourceType).Equals($SourceType, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
    $expected = ConvertTo-DocBridgeNormalizedLocalPath -Path $NormalizedSource
    if ([string]::IsNullOrWhiteSpace($expected) -or [string]::IsNullOrWhiteSpace([string]$State.NormalizedSource)) { return $false }
    return ([string]$State.NormalizedSource).Equals($expected, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-DocBridgeCodexMarketplaceRegistered {
    param(
        [Parameter(Mandatory = $true)][string]$MarketplaceName,
        [string]$UserProfileRoot = $env:USERPROFILE,
        [string]$CodexHomeRoot
    )
    $configPath = Get-DocBridgeCodexConfigPath -UserProfileRoot $UserProfileRoot -CodexHomeRoot $CodexHomeRoot
    if ([string]::IsNullOrWhiteSpace($configPath) -or
        -not (Test-Path -LiteralPath $configPath -PathType Leaf)) { return $false }
    try {
        $raw = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8
        $escaped = [regex]::Escape($MarketplaceName)
        return $raw -match ('(?m)^\s*\[marketplaces\.(?:' + $escaped + '|"' + $escaped + '")\]\s*$')
    } catch { return $false }
}

function Test-DocBridgeCodexMcpEnabled {
    param([Parameter(Mandatory = $true)][string]$Output)
    $matchingLine = @($Output -split "`r?`n" | Where-Object { $_ -match '^\s*doc-bridge\s+' } | Select-Object -First 1)
    return $matchingLine.Count -eq 1 -and $matchingLine[0] -match '(?i)\benabled\b'
}

function ConvertFrom-DocBridgeClaudeMcpOutput {
    param(
        [Parameter(Mandatory = $true)][int]$ExitCode,
        [AllowEmptyString()][string]$Output,
        [string]$Name = 'doc-bridge'
    )
    if ($ExitCode -ne 0) {
        if ($Output -match '(?i)no(?:\s+user-scoped)?\s+MCP server found') {
            return [pscustomobject]@{
                Exists = $false
                Command = $null
                Arguments = $null
                ArgsText = $null
                Environment = $null
                Output = $Output
                Error = $null
            }
        }
        return [pscustomobject]@{
            Exists = $null
            Command = $null
            Arguments = $null
            ArgsText = $null
            Environment = $null
            Output = $Output
            Error = "Claude MCP query failed (exit $ExitCode)."
        }
    }

    $fields = @{}
    foreach ($line in @($Output -split "\r?\n")) {
        if ($line -match '^\s*(Command|Args|Environment):[ \t]*(.*)$') {
            $fields[$matches[1]] = $matches[2].Trim()
        }
    }
    $missing = @(@('Command', 'Args', 'Environment') | Where-Object { -not $fields.ContainsKey($_) })
    if ($missing.Count -gt 0 -or [string]::IsNullOrWhiteSpace([string]$fields.Command)) {
        return [pscustomobject]@{
            Exists = $null
            Command = $null
            Arguments = $null
            ArgsText = $null
            Environment = $null
            Output = $Output
            Error = "Claude MCP query succeeded but required ownership fields could not be parsed: $($missing -join ', ')."
        }
    }
    $argsText = [string]$fields.Args
    $arguments = if ([string]::IsNullOrWhiteSpace($argsText)) { @() } else { @($argsText -split '\s+') }
    return [pscustomobject]@{
        Exists = $true
        Command = ([string]$fields.Command).Trim().Trim('"')
        Arguments = @($arguments)
        ArgsText = $argsText
        Environment = [string]$fields.Environment
        Output = $Output
        Error = $null
    }
}

function Test-DocBridgeClaudeMcpStateMatches {
    param(
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][string]$Command,
        [string[]]$Arguments = @('--stdio')
    )
    if ($null -eq $State -or $null -ne $State.Error -or $State.Exists -ne $true) { return $false }
    if (-not ([string]$State.Command).Equals($Command, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
    if ($null -eq $State.PSObject.Properties['Arguments'] -or $null -eq $State.Arguments) { return $false }
    if ($null -eq $State.PSObject.Properties['Environment']) { return $false }
    $actualArguments = @($State.Arguments)
    if ($actualArguments.Count -ne $Arguments.Count) { return $false }
    for ($index = 0; $index -lt $Arguments.Count; $index++) {
        if ([string]$actualArguments[$index] -cne [string]$Arguments[$index]) { return $false }
    }
    return [string]::IsNullOrWhiteSpace([string]$State.Environment)
}

function Get-DocBridgeClaudeMcpState {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string]$Name = 'doc-bridge'
    )
    $result = Invoke-DocBridgeNative -FilePath $FilePath -Arguments @('mcp', 'get', $Name)
    return ConvertFrom-DocBridgeClaudeMcpOutput -ExitCode $result.ExitCode -Output $result.Output -Name $Name
}

function Write-DocBridgeUtf8NoBom {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text
    )
    $parent = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, $Text, (New-Object System.Text.UTF8Encoding($false)))
}

function Invoke-DocBridgeNative {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$Arguments = @(),
        [string]$WorkingDirectory,
        [hashtable]$Environment
    )
    $quotedArguments = @($Arguments | ForEach-Object {
        if ($_ -match '[\s"]') { '"' + ($_ -replace '"', '\"') + '"' } else { $_ }
    }) -join ' '
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $FilePath
    $startInfo.Arguments = $quotedArguments
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) {
        $startInfo.WorkingDirectory = $WorkingDirectory
    }
    # Windows PowerShell 5.1 can enumerate an empty ProcessStartInfo.EnvironmentVariables
    # dictionary as $null, so indexing that property is not reliable. Temporarily set only
    # the requested process-level variables; the child inherits them and finally restores them.
    $previousEnvironment = @{}
    if ($null -ne $Environment) {
        foreach ($entry in $Environment.GetEnumerator()) {
            if ($null -eq $entry.Value) { continue }
            $key = [string]$entry.Key
            $previousEnvironment[$key] = [Environment]::GetEnvironmentVariable($key, 'Process')
            [Environment]::SetEnvironmentVariable($key, [string]$entry.Value, 'Process')
        }
    }
    try {
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        if (-not $process.Start()) { throw "Failed to start process: $FilePath" }
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        $combined = ($stdout + $stderr).Trim()
        return [pscustomobject]@{
            ExitCode = $process.ExitCode
            StdOut = $stdout.Trim()
            StdErr = $stderr.Trim()
            Output = $combined
        }
    } finally {
        foreach ($entry in $previousEnvironment.GetEnumerator()) {
            [Environment]::SetEnvironmentVariable([string]$entry.Key, $entry.Value, 'Process')
        }
    }
}

Export-ModuleMember -Function @(
    'Get-DocBridgeDefaultRoot',
    'Assert-DocBridgeSafePath',
    'Backup-DocBridgeFile',
    'Read-DocBridgeJsonObject',
    'Get-DocBridgeMcpJsonEntryState',
    'Test-DocBridgeMcpJsonEntryMatches',
    'Set-DocBridgeMcpJsonEntry',
    'Remove-DocBridgeMcpJsonEntry',
    'Get-DocBridgeClientSet',
    'Test-DocBridgeCodexInstalled',
    'Get-DocBridgeCodexCli',
    'Get-DocBridgeCodexPluginCliMode',
    'Get-DocBridgeCodexConfigPath',
    'Test-DocBridgeCodexMarketplaceRegistered',
    'Test-DocBridgeCodexPluginEnabled',
    'Test-DocBridgeCodexMcpEnabled',
    'ConvertTo-DocBridgeNormalizedLocalPath',
    'Get-DocBridgeCodexMarketplaceState',
    'Test-DocBridgeCodexMarketplaceStateMatches',
    'ConvertFrom-DocBridgeClaudeMcpOutput',
    'Test-DocBridgeClaudeMcpStateMatches',
    'Get-DocBridgeClaudeMcpState',
    'Write-DocBridgeUtf8NoBom',
    'Invoke-DocBridgeNative'
)
