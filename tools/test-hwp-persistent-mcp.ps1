[CmdletBinding()]
param([int]$TimeoutSec = 180)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$exe = Join-Path $repo 'dist\doc-bridge-mcp.exe'
$source = Join-Path (Split-Path -Parent $repo) 'demo-hwp.hwp'
if (-not (Test-Path -LiteralPath $exe)) { throw "MCP executable not found: $exe" }
if (-not (Test-Path -LiteralPath $source)) { throw "HWP sample not found: $source" }

$scratch = Join-Path $env:TEMP ('docbridge-hwp-persistent-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
New-Item -ItemType Directory -Path $scratch -Force | Out-Null
$work = Join-Path $scratch 'hwp-test.hwp'
Copy-Item -LiteralPath $source -Destination $work

$utf8 = New-Object System.Text.UTF8Encoding($false)
$psi = [System.Diagnostics.ProcessStartInfo]::new()
$psi.FileName = $exe
$psi.Arguments = '--stdio'
$psi.UseShellExecute = $false
$psi.CreateNoWindow = $true
$psi.RedirectStandardInput = $true
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$process = [System.Diagnostics.Process]::Start($psi)
$stderrTask = $process.StandardError.ReadToEndAsync()
$stream = $process.StandardInput.BaseStream
$nextId = 0

function Send-Line([object]$value) {
    $line = ($value | ConvertTo-Json -Compress -Depth 30) + "`n"
    $bytes = $utf8.GetBytes($line)
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Flush()
}

function Invoke-McpTool([string]$name, [hashtable]$arguments) {
    $script:nextId++
    Send-Line @{ jsonrpc = '2.0'; id = $script:nextId; method = 'tools/call'; params = @{ name = $name; arguments = $arguments } }
    $readTask = $process.StandardOutput.ReadLineAsync()
    if (-not $readTask.Wait($TimeoutSec * 1000)) { throw "$name timed out after ${TimeoutSec}s" }
    $response = $readTask.Result | ConvertFrom-Json
    if ($response.error) { throw "$name JSON-RPC error: $($response.error.message)" }
    if ($response.result.structuredContent) { return $response.result.structuredContent }
    if ($response.result.content) { return $response.result.content[0].text | ConvertFrom-Json }
    throw "$name returned no payload"
}

try {
    $nextId++
    Send-Line @{ jsonrpc = '2.0'; id = $nextId; method = 'initialize'; params = @{ protocolVersion = '2025-06-18'; capabilities = @{}; clientInfo = @{ name = 'hwp-persistent-e2e'; version = '1.0' } } }
    $initTask = $process.StandardOutput.ReadLineAsync()
    if (-not $initTask.Wait($TimeoutSec * 1000)) { throw 'initialize timed out' }
    $init = $initTask.Result | ConvertFrom-Json
    Send-Line @{ jsonrpc = '2.0'; method = 'notifications/initialized' }

    $read = Invoke-McpTool 'hwp_read_text' @{ file = $work; scope = 'document' }
    if (-not $read.ok) { throw "initial read failed: $($read.errors -join '; ')" }
    $original = [string]$read.text
    $word = $original -split '[\s\r\n\.,]+' | Where-Object { $_.Length -ge 2 } | Select-Object -First 1
    if (-not $word) { throw 'no replacement word found' }
    # The replacement deliberately must not contain the search term.  The HWP
    # adapter treats that case specially because a successful replacement can
    # still leave the original term inside the replacement text.
    $replacement = 'DocBridgeHwpE2E_' + [Guid]::NewGuid().ToString('N')
    $ops = @(@{ op = 'find_replace'; file = $work; find = $word; replace = $replacement })

    $dry = Invoke-McpTool 'hwp_apply_ops' @{ dryRun = $true; ops = $ops }
    if (-not $dry.ok -or -not $dry.confirmToken) { throw "dry-run failed: $($dry.errors -join '; ')" }
    $apply = Invoke-McpTool 'hwp_apply_ops' @{ dryRun = $false; confirmToken = $dry.confirmToken; ops = $ops }
    if (-not $apply.ok -or -not $apply.readback.verified) { throw "apply failed: $($apply.errors -join '; ')" }
    $after = Invoke-McpTool 'hwp_read_text' @{ file = $work; scope = 'document' }
    if ([string]$after.text -notlike "*$replacement*") {
        throw "changed text was not found; apply=$($apply | ConvertTo-Json -Compress -Depth 20); afterPrefix=$(([string]$after.text).Substring(0, [Math]::Min(200, ([string]$after.text).Length)))"
    }

    $restoreDry = Invoke-McpTool 'core_restore_snapshot' @{ snapshotId = $dry.snapshotId }
    $restore = Invoke-McpTool 'core_restore_snapshot' @{ snapshotId = $dry.snapshotId; confirmToken = $restoreDry.confirmToken }
    if (-not $restore.ok) { throw "restore failed: $($restore.errors -join '; ')" }
    $restored = Invoke-McpTool 'hwp_read_text' @{ file = $work; scope = 'document' }
    if ([string]$restored.text -ne $original) { throw 'restored text does not match the original' }

    [pscustomobject]@{
        ok = $true
        version = [string]$init.result.serverInfo.version
        file = $work
        replacement = $replacement
        interaction = $apply.interaction
        snapshotId = $dry.snapshotId
    } | ConvertTo-Json -Depth 10
}
finally {
    try { $stream.Close() } catch { }
    if (-not $process.WaitForExit(10000)) { try { $process.Kill() } catch { } }
    try {
        $stderr = $stderrTask.GetAwaiter().GetResult()
        if ($stderr) { Write-Verbose $stderr }
    } catch { }
}
