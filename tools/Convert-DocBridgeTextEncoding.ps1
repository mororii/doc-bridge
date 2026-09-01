<#
.SYNOPSIS
  Normalizes DocBridge PowerShell files and converts legacy VBA .bas files safely.

.DESCRIPTION
  PowerShell scripts are written as UTF-8 with BOM so Windows PowerShell 5.1 can
  parse Korean text reliably. VBA .bas files are written as CP949 without BOM
  and with CRLF line endings for legacy Korean Office/VBE import compatibility.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('NormalizePowerShell', 'CheckPowerShell', 'ConvertBasToCp949', 'CheckBas')]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$Path,

    [string]$OutputPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Test-BytePrefix {
    param([byte[]]$Bytes, [byte[]]$Prefix)
    if ($Bytes.Length -lt $Prefix.Length) { return $false }
    for ($index = 0; $index -lt $Prefix.Length; $index++) {
        if ($Bytes[$index] -ne $Prefix[$index]) { return $false }
    }
    return $true
}

function Get-Cp949Encoding {
    $encoderFallback = New-Object System.Text.EncoderExceptionFallback
    $decoderFallback = New-Object System.Text.DecoderExceptionFallback
    return [System.Text.Encoding]::GetEncoding(949, $encoderFallback, $decoderFallback)
}

function Read-TextWithEncodingDetection {
    param([Parameter(Mandatory = $true)][string]$FilePath)

    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    if (Test-BytePrefix $bytes ([byte[]](0xEF, 0xBB, 0xBF))) {
        $encoding = New-Object System.Text.UTF8Encoding -ArgumentList $true, $true
        return $encoding.GetString($bytes, 3, $bytes.Length - 3)
    }
    if (Test-BytePrefix $bytes ([byte[]](0xFF, 0xFE))) {
        $encoding = New-Object System.Text.UnicodeEncoding -ArgumentList $false, $true, $true
        return $encoding.GetString($bytes, 2, $bytes.Length - 2)
    }
    if (Test-BytePrefix $bytes ([byte[]](0xFE, 0xFF))) {
        $encoding = New-Object System.Text.UnicodeEncoding -ArgumentList $true, $true, $true
        return $encoding.GetString($bytes, 2, $bytes.Length - 2)
    }

    try {
        $encoding = New-Object System.Text.UTF8Encoding -ArgumentList $false, $true
        return $encoding.GetString($bytes)
    } catch [System.Text.DecoderFallbackException] {
        return (Get-Cp949Encoding).GetString($bytes)
    }
}

function Get-PowerShellFiles {
    param([Parameter(Mandatory = $true)][string]$InputPath)

    $resolved = [System.IO.Path]::GetFullPath($InputPath)
    if (Test-Path -LiteralPath $resolved -PathType Leaf) {
        $item = Get-Item -LiteralPath $resolved
        if ($item.Extension -notin @('.ps1', '.psm1', '.psd1')) {
            throw "Not a PowerShell source file: $resolved"
        }
        return @($item)
    }
    if (-not (Test-Path -LiteralPath $resolved -PathType Container)) {
        throw "Path was not found: $resolved"
    }
    return @(Get-ChildItem -LiteralPath $resolved -Recurse -File | Where-Object {
        $_.Extension -in @('.ps1', '.psm1', '.psd1')
    })
}

function Assert-Utf8Bom {
    param([Parameter(Mandatory = $true)][System.IO.FileInfo[]]$Files)

    $failures = New-Object System.Collections.Generic.List[string]
    foreach ($file in $Files) {
        $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
        if (-not (Test-BytePrefix $bytes ([byte[]](0xEF, 0xBB, 0xBF)))) {
            [void]$failures.Add($file.FullName)
        }
    }
    if ($failures.Count -gt 0) {
        throw "PowerShell files without UTF-8 BOM:`n$($failures -join "`n")"
    }
}

function Assert-BasCp949 {
    param([Parameter(Mandatory = $true)][string]$FilePath)

    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    if (Test-BytePrefix $bytes ([byte[]](0xEF, 0xBB, 0xBF))) {
        throw "BAS must not contain a UTF-8 BOM: $FilePath"
    }
    $cp949 = Get-Cp949Encoding
    $text = $cp949.GetString($bytes)
    $roundTrip = $cp949.GetBytes($text)
    if ($roundTrip.Length -ne $bytes.Length) {
        throw "BAS is not a lossless CP949 stream: $FilePath"
    }
    for ($index = 0; $index -lt $bytes.Length; $index++) {
        if ($bytes[$index] -ne $roundTrip[$index]) {
            throw "BAS is not a lossless CP949 stream: $FilePath"
        }
    }
    if ($text -match '(?<!\r)\n|\r(?!\n)') {
        throw "BAS line endings must be CRLF: $FilePath"
    }
}

$resolvedPath = [System.IO.Path]::GetFullPath($Path)
switch ($Mode) {
    'NormalizePowerShell' {
        $files = @(Get-PowerShellFiles $resolvedPath)
        $utf8Bom = New-Object System.Text.UTF8Encoding -ArgumentList $true
        foreach ($file in $files) {
            $text = Read-TextWithEncodingDetection $file.FullName
            [System.IO.File]::WriteAllText($file.FullName, $text, $utf8Bom)
        }
        Assert-Utf8Bom $files
        Write-Host "Normalized PowerShell files: $($files.Count)" -ForegroundColor Green
    }
    'CheckPowerShell' {
        $files = @(Get-PowerShellFiles $resolvedPath)
        Assert-Utf8Bom $files
        Write-Host "PowerShell UTF-8 BOM check passed: $($files.Count)" -ForegroundColor Green
    }
    'ConvertBasToCp949' {
        if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
            throw "BAS source was not found: $resolvedPath"
        }
        if ([System.IO.Path]::GetExtension($resolvedPath) -ne '.bas') {
            throw "Expected a .bas source file: $resolvedPath"
        }
        if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = $resolvedPath }
        $resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)
        if ([System.IO.Path]::GetExtension($resolvedOutput) -ne '.bas') {
            throw "Expected a .bas output file: $resolvedOutput"
        }
        $parent = Split-Path -Parent $resolvedOutput
        if (-not [string]::IsNullOrWhiteSpace($parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        $text = Read-TextWithEncodingDetection $resolvedPath
        $text = $text -replace '\r\n|\r|\n', "`r`n"
        [System.IO.File]::WriteAllText($resolvedOutput, $text, (Get-Cp949Encoding))
        Assert-BasCp949 $resolvedOutput
        Write-Host "Converted BAS to CP949/CRLF: $resolvedOutput" -ForegroundColor Green
    }
    'CheckBas' {
        if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
            throw "BAS file was not found: $resolvedPath"
        }
        if ([System.IO.Path]::GetExtension($resolvedPath) -ne '.bas') {
            throw "Expected a .bas file: $resolvedPath"
        }
        Assert-BasCp949 $resolvedPath
        Write-Host "BAS CP949/CRLF check passed: $resolvedPath" -ForegroundColor Green
    }
}
