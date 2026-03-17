# ─────────────────────────────────────────────────────────────────────────────
# Helpers.ps1 — Shared utility functions
# ─────────────────────────────────────────────────────────────────────────────

$script:Sep    = "=" * 90
$script:SubSep = "-" * 70

function Run-Command {
    param([string]$Exe, [string[]]$Arguments, [int]$TimeoutSeconds = 300, [string]$StatusLabel = "")
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $Exe
        $psi.Arguments = ($Arguments -join ' ')
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow  = $true

        $proc = [System.Diagnostics.Process]::Start($psi)

        # Async reads to prevent deadlock on large stderr/stdout
        $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
        $stderrTask = $proc.StandardError.ReadToEndAsync()

        # Wait with visible elapsed counter
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $label = if ($StatusLabel) { $StatusLabel } else { (Split-Path $Exe -Leaf) }
        while (-not $proc.WaitForExit(2000)) {
            $elapsed = $sw.Elapsed
            Write-Host "`r      Running $label... $($elapsed.ToString('mm\:ss'))" -NoNewline -ForegroundColor DarkYellow
            if ($sw.Elapsed.TotalSeconds -gt $TimeoutSeconds) {
                Write-Host "`r      $label TIMED OUT after ${TimeoutSeconds}s - skipping    " -ForegroundColor Red
                try { $proc.Kill() } catch {}
                return @{ ExitCode = -1; StdOut = ""; StdErr = "TIMEOUT after ${TimeoutSeconds}s" }
            }
        }
        if ($sw.Elapsed.TotalSeconds -gt 3) {
            Write-Host "`r      $label done ($([math]::Round($sw.Elapsed.TotalSeconds,1))s)        " -ForegroundColor DarkGray
        } else {
            Write-Host "`r                                                        `r" -NoNewline
        }

        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()

        return @{ ExitCode = $proc.ExitCode; StdOut = $stdout; StdErr = $stderr }
    }
    catch {
        return @{ ExitCode = -1; StdOut = ""; StdErr = $_.Exception.Message }
    }
}

function Format-Bitrate {
    param([double]$Kbps)
    if ($Kbps -ge 1000) { return "{0:N2} Mbps" -f ($Kbps / 1000) }
    return "{0:N0} kbps" -f $Kbps
}

function Format-Size {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes bytes"
}

function Write-Section {
    param([System.Text.StringBuilder]$sb, [string]$Title)
    $sb.AppendLine("") | Out-Null
    $sb.AppendLine($script:SubSep) | Out-Null
    $sb.AppendLine("  $Title") | Out-Null
    $sb.AppendLine($script:SubSep) | Out-Null
}

function Write-Field {
    param(
        [System.Text.StringBuilder]$sb,
        [string]$Label,
        [string]$Value,
        [int]$Indent = 4,
        [int]$Width  = 40
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return }
    $pad = ' ' * $Indent
    $sb.AppendLine("${pad}$($Label.PadRight($Width)): $Value") | Out-Null
}

function BoolStr {
    param($val)
    if ($null -eq $val) { return $null }
    if ($val) { "Enabled" } else { "Disabled" }
}

function Parse-EncoderSettings {
    param([string]$SettingsString)
    if ([string]::IsNullOrWhiteSpace($SettingsString)) { return $null }
    $parsed = [ordered]@{}
    foreach ($part in ($SettingsString -split '\s*/\s*')) {
        $t = $part.Trim()
        if ($t -match '^([^=]+)=(.+)$') { $parsed[$Matches[1].Trim()] = $Matches[2].Trim() }
        elseif ($t -match '^no-(.+)$')   { $parsed[$Matches[1].Trim()] = 'disabled' }
        elseif ($t.Length -gt 0)          { $parsed[$t] = 'enabled' }
    }
    return $parsed
}

function Find-Tool {
    param([string]$Name, [string[]]$Alternates = @())

    # First check the tools folder next to the script
    $toolsDir = Join-Path $PSScriptRoot "..\tools"
    if (-not (Test-Path $toolsDir)) { $toolsDir = Join-Path $PSScriptRoot "tools" }

    foreach ($n in @($Name) + $Alternates) {
        # Priority 1: tools folder (flat)
        $toolPath = Join-Path $toolsDir $n
        if (Test-Path $toolPath -PathType Leaf) { return (Resolve-Path $toolPath).Path }

        # Priority 2: tools subfolder (e.g. tools/NVEncC64/NVEncC64.exe)
        if (Test-Path $toolsDir) {
            $subMatch = Get-ChildItem -Path $toolsDir -Filter $n -Recurse -Depth 2 -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($subMatch) { return $subMatch.FullName }
        }

        # Priority 3: system PATH
        $cmd = Get-Command $n -ErrorAction SilentlyContinue
        if ($cmd -and (Test-Path $cmd.Source -PathType Leaf)) { return $cmd.Source }

        # Priority 4: common locations
        foreach ($p in @(
            "$env:ProgramFiles\$n", "${env:ProgramFiles(x86)}\$n",
            "$env:LOCALAPPDATA\$n", "$env:USERPROFILE\$n", ".\$n"
        )) {
            if (Test-Path $p -PathType Leaf) { return $p }
        }
    }
    return $null
}
