<#
.SYNOPSIS
    Comprehensive media compression analyzer v2.0
.DESCRIPTION
    Runs FFprobe, FFmpeg, MediaInfo, and CheckBitrate to extract every detail
    about how a file was encoded, including bitstream-level SPS/PPS/NAL params,
    QP statistics, GOP structure, and generates encoder settings reconstruction.
.EXAMPLE
    .\Analyze-Media.ps1 -Path "D:\Movie.mkv"
    .\Analyze-Media.ps1 -Path "D:\Movies" -Recurse -ExportJson
    .\Analyze-Media.ps1 -Path "D:\Movie.mkv" -MaxAnalysisFrames 5000
    .\Analyze-Media.ps1 -Path "D:\Movie.mkv" -CheckGPU -Verify
    .\Analyze-Media.ps1 -Path "D:\Movie.mkv" -Verify -VerifyDuration 120
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, Position=0)][string]$Path,
    [switch]$Recurse,
    [string]$OutputDir,
    [double]$CheckBitrateInterval = 0,
    [switch]$SkipCheckBitrate,
    [switch]$SkipQP,
    [int]$MaxAnalysisFrames = 1000,
    [switch]$FullFrameScan,
    [switch]$ExportJson,
    [switch]$CheckGPU,
    [switch]$Verify,
    [int]$VerifyDuration = 0,
    [switch]$DebugMode
)

# ── Load library files ──
$scriptDir = $PSScriptRoot
. "$scriptDir\lib\Helpers.ps1"
. "$scriptDir\lib\Collectors.ps1"
. "$scriptDir\lib\HMAnalyser.ps1"
. "$scriptDir\lib\JMAnalyser.ps1"
. "$scriptDir\lib\ReportWriter.ps1"
. "$scriptDir\lib\Reconstruction.ps1"
. "$scriptDir\lib\GPUCapabilities.ps1"
. "$scriptDir\lib\Verification.ps1"

$script:SkipCB = $SkipCheckBitrate
$script:DebugMode = $DebugMode
$MediaExts = '.mkv','.mp4','.m4v','.avi','.mov','.wmv','.flv','.webm','.ts','.m2ts','.mts','.mpg','.mpeg','.vob','.3gp','.ogv'

# ── Discover tools ──
$Tools = @{
    FFprobe      = Find-Tool 'ffprobe' @('ffprobe.exe')
    FFmpeg       = Find-Tool 'ffmpeg' @('ffmpeg.exe')
    MediaInfo    = Find-Tool 'mediainfo' @('mediainfo.exe','MediaInfo.exe')
    CheckBitrate = Find-Tool 'CheckBitrate' @('CheckBitrate.exe','checkbitrate.exe')
    TAppDecoder  = Find-Tool 'TAppDecoderAnalyser' @('TAppDecoderAnalyser.exe','TAppDecoder.exe','TAppDecoder')
    JMDecoder    = Find-Tool 'ldecod' @('ldecod.exe','ldecod')
    NVEncC       = Find-Tool 'NVEncC64' @('NVEncC64.exe','NVEncC.exe')
    QSVEncC      = Find-Tool 'QSVEncC64' @('QSVEncC64.exe','QSVEncC.exe')
    DoviTool     = Find-Tool 'dovi_tool' @('dovi_tool.exe','dovi_tool')
}

Write-Host "`n$($script:Sep)" -ForegroundColor Cyan
Write-Host "  MEDIA COMPRESSION ANALYZER v2.0" -ForegroundColor Cyan
Write-Host $script:Sep -ForegroundColor Cyan

# ── Get tool versions (quick, parallel-safe) ──
$ToolVersions = @{}
foreach ($t in $Tools.GetEnumerator()) {
    if (-not $t.Value) { continue }
    $ver = $null
    try {
        switch ($t.Key) {
            'FFprobe' {
                $r = Run-Command $t.Value @('-version') -TimeoutSeconds 5
                if ($r.StdOut -match 'ffprobe version\s+(\S+)') { $ver = $Matches[1] }
            }
            'FFmpeg' {
                $r = Run-Command $t.Value @('-version') -TimeoutSeconds 5
                if ($r.StdOut -match 'ffmpeg version\s+(\S+)') { $ver = $Matches[1] }
            }
            'MediaInfo' {
                $r = Run-Command $t.Value @('--version') -TimeoutSeconds 5
                $out = "$($r.StdOut) $($r.StdErr)"
                if ($out -match 'v([\d.]+)') { $ver = "v$($Matches[1])" }
            }
            'DoviTool' {
                $r = Run-Command $t.Value @('--version') -TimeoutSeconds 5
                if ("$($r.StdOut) $($r.StdErr)" -match 'dovi_tool\s+([\d.]+)') { $ver = "v$($Matches[1])" }
            }
            'TAppDecoder' {
                $r = Run-Command $t.Value @('--help') -TimeoutSeconds 5
                if ("$($r.StdOut) $($r.StdErr)" -match 'Version\s*\[([\d.]+)\]') { $ver = "v$($Matches[1])" }
            }
            'JMDecoder' {
                $r = Run-Command $t.Value @('-version') -TimeoutSeconds 5
                if ("$($r.StdOut) $($r.StdErr)" -match 'JM\s+([\d.]+)') { $ver = "JM $($Matches[1])" }
            }
            'CheckBitrate' {
                $r = Run-Command $t.Value @() -TimeoutSeconds 5
                if ("$($r.StdOut) $($r.StdErr)" -match 'CheckBitrate\s+([\d.]+)') { $ver = "v$($Matches[1])" }
            }
            'NVEncC' {
                $r = Run-Command $t.Value @('--version') -TimeoutSeconds 5
                if ("$($r.StdOut) $($r.StdErr)" -match 'NVEncC\s+\(x64\)\s+([\d.]+)\s+\((r\d+)\)') {
                    $ver = "v$($Matches[1]) ($($Matches[2]))"
                }
            }
            'QSVEncC' {
                $r = Run-Command $t.Value @('--version') -TimeoutSeconds 5
                if ("$($r.StdOut) $($r.StdErr)" -match 'QSVEncC\s+\(x64\)\s+([\d.]+)\s+\((r\d+)\)') {
                    $ver = "v$($Matches[1]) ($($Matches[2]))"
                }
            }
        }
    } catch {}
    if ($ver) { $ToolVersions[$t.Key] = $ver }
}

Write-Host "`n  Tool Status:" -ForegroundColor Yellow
foreach ($t in ($Tools.GetEnumerator() | Sort-Object Key)) {
    $color = if ($t.Value) { 'Green' } else { 'Red' }
    $verStr = if ($ToolVersions[$t.Key]) { " [$($ToolVersions[$t.Key])]" } else { '' }
    $status = if ($t.Value) { "FOUND$verStr -> $($t.Value)" } else { "NOT FOUND" }
    Write-Host "    $($t.Key.PadRight(16)) : $status" -ForegroundColor $color
}
if (-not $Tools.FFprobe -and -not $Tools.MediaInfo) { Write-Error "Need at least FFprobe or MediaInfo."; exit 1 }

# ── Collect files ──
$Files = @()
if (Test-Path $Path -PathType Leaf) { $Files = @(Get-Item $Path) }
elseif (Test-Path $Path -PathType Container) {
    $gp = @{ Path = $Path; File = $true }; if ($Recurse) { $gp['Recurse'] = $true }
    $Files = Get-ChildItem @gp | Where-Object { $MediaExts -contains $_.Extension.ToLower() }
} else { Write-Error "Path not found: $Path"; exit 1 }

if ($Files.Count -eq 0) { Write-Warning "No media files found."; exit 0 }
$verifyLabel = if ($Verify) { " | Verify: ON$(if($VerifyDuration -gt 0){" (${VerifyDuration}s)"}else{' (quick)'})" } else { '' }
Write-Host "`n  Files: $($Files.Count)  |  Frame budget: $(if($FullFrameScan){'FULL SCAN'}else{"$MaxAnalysisFrames (distributed)"})$verifyLabel`n" -ForegroundColor Yellow

if (-not $OutputDir) {
    $sourceDir = if (Test-Path $Path -PathType Leaf) {
        Split-Path $Path -Parent
    } else { $Path }

    $defaultOut = Join-Path $sourceDir "MediaAnalysis"
    $canWrite = $false

    # Test if source location is writable (handles optical drives, ISO mounts, read-only network shares)
    try {
        $drive = [System.IO.Path]::GetPathRoot($sourceDir)
        if ($drive) {
            $driveInfo = [System.IO.DriveInfo]::new($drive)
            # CD-ROM, DVD, Blu-ray drives are never writable
            if ($driveInfo.DriveType -eq 'CDRom') {
                $canWrite = $false
            } elseif ($driveInfo.DriveType -eq 'Network') {
                # Network drives — test by trying to create directory
                try {
                    New-Item -ItemType Directory -Path $defaultOut -Force | Out-Null
                    $canWrite = $true
                } catch { $canWrite = $false }
            } else {
                # Local/removable — check if drive is ready and not read-only
                if ($driveInfo.IsReady) {
                    try {
                        New-Item -ItemType Directory -Path $defaultOut -Force | Out-Null
                        $canWrite = $true
                    } catch { $canWrite = $false }
                }
            }
        }
    } catch {
        # DriveInfo failed (UNC paths, unusual mounts) — try direct write test
        try {
            New-Item -ItemType Directory -Path $defaultOut -Force | Out-Null
            $canWrite = $true
        } catch { $canWrite = $false }
    }

    if ($canWrite) {
        $OutputDir = $defaultOut
    } else {
        # Fall back to user's Documents folder or script directory
        $fallback = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "MediaAnalysis"
        if (-not (Test-Path $fallback)) { New-Item -ItemType Directory -Path $fallback -Force | Out-Null }
        $OutputDir = $fallback
        Write-Host "  Source location is read-only (optical/ISO/network). Output: $OutputDir" -ForegroundColor Yellow
    }
} else {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# ═══════════════════════════════════════════════════════════════════════════
# GPU CAPABILITY DETECTION (once, shared across all files)
# ═══════════════════════════════════════════════════════════════════════════

$script:GPUCaps = $null
$script:GPUValidation = $null

# -Verify implies -CheckGPU (need GPU caps to build verification commands)
# -VerifyDuration > 0 implies -Verify
if ($VerifyDuration -gt 0 -and -not $Verify) { $Verify = [switch]::Present }
if ($Verify -and -not $CheckGPU) { $CheckGPU = [switch]::Present }

if ($CheckGPU) {
    Write-Host "`n  Detecting GPU capabilities..." -ForegroundColor Yellow
    $nvPath  = $Tools.NVEncC
    $qsvPath = $Tools.QSVEncC
    $script:GPUCaps = Get-GPUCapabilities -NVEncPath $nvPath -QSVEncPath $qsvPath

    # Enumerate GPU devices
    if ($nvPath) {
        $devR = Run-Command $nvPath @('--check-device') -TimeoutSeconds 10
        $devOut = "$($devR.StdOut)`n$($devR.StdErr)"
        $nvDevices = @()
        foreach ($line in ($devOut -split "`n")) {
            if ($line -match 'DeviceId\s*#(\d+):\s*(.+)') {
                $nvDevices += @{ Id = [int]$Matches[1]; Name = $Matches[2].Trim() }
            }
        }
        if ($nvDevices.Count -gt 0) {
            foreach ($d in $nvDevices) {
                Write-Host "    NVIDIA Device #$($d.Id): $($d.Name)" -ForegroundColor Green
            }
            if ($nvDevices.Count -gt 1) {
                Write-Host "      Multi-GPU: Use --device <N> to select" -ForegroundColor DarkGray
            }
        } elseif ($script:GPUCaps.NVEnc) {
            Write-Host "    NVIDIA: $($script:GPUCaps.NVEnc.GPU)" -ForegroundColor Green
        }
    }
    if ($qsvPath) {
        $devR = Run-Command $qsvPath @('--check-device') -TimeoutSeconds 10
        $devOut = "$($devR.StdOut)`n$($devR.StdErr)"
        $qsvDevices = @()
        foreach ($line in ($devOut -split "`n")) {
            if ($line -match 'Device\s*#(\d+):\s*(.+)') {
                $qsvDevices += @{ Id = [int]$Matches[1]; Name = $Matches[2].Trim() }
            }
        }
        if ($qsvDevices.Count -gt 0) {
            foreach ($d in $qsvDevices) {
                Write-Host "    Intel  Device #$($d.Id): $($d.Name)" -ForegroundColor Green
            }
            if ($qsvDevices.Count -gt 1) {
                Write-Host "      Multi-GPU: Use -d <N> to select" -ForegroundColor DarkGray
            }
        } elseif ($script:GPUCaps.QSVEnc) {
            Write-Host "    Intel:  $($script:GPUCaps.QSVEnc.GPU)" -ForegroundColor Green
        }
    }

    foreach ($e in $script:GPUCaps.Errors) { Write-Host "    $e" -ForegroundColor DarkYellow }
} elseif ($Tools.NVEncC -or $Tools.QSVEncC) {
    Write-Host "  Tip: Use -CheckGPU to validate commands against your GPU hardware" -ForegroundColor DarkGray
}

# ═══════════════════════════════════════════════════════════════════════════
# ANALYZE EACH FILE
# ═══════════════════════════════════════════════════════════════════════════

$allReports = @()
$fileIdx = 0

foreach ($file in $Files) {
    $fileIdx++
    $fp = $file.FullName
    Write-Host "`n  [$fileIdx/$($Files.Count)] $($file.Name)" -ForegroundColor Green
    Write-Host "  $(Format-Size $file.Length)" -ForegroundColor DarkGray
    $fileTimer = [System.Diagnostics.Stopwatch]::StartNew()

    $rpt = [System.Text.StringBuilder]::new()
    $rc  = [ordered]@{}   # reconstruction collector
    $rc['FileName'] = $file.Name
    $jd  = @{ File = $file.Name; Size = $file.Length; Path = $fp }

    # Report header
    $rpt.AppendLine($script:Sep) | Out-Null
    $rpt.AppendLine("  MEDIA COMPRESSION ANALYSIS REPORT v2.0") | Out-Null
    $rpt.AppendLine("  Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')") | Out-Null
    $rpt.AppendLine($script:Sep) | Out-Null
    $rpt.AppendLine("  File: $($file.Name)") | Out-Null
    $rpt.AppendLine("  Path: $fp") | Out-Null
    $rpt.AppendLine("  Size: $(Format-Size $file.Length) ($($file.Length) bytes)") | Out-Null

    # ── 1. FFprobe streams ──
    Write-Progress -Activity "Analyzing $($file.Name)" -Status "[1/9] FFprobe streams..." -PercentComplete 5
    Write-Host "    [1/9] FFprobe streams..." -ForegroundColor DarkGray
    $probe = Get-ProbeJson $fp
    if ($probe) { $jd.Probe = $probe; Write-ContainerReport $rpt $probe $rc }

    # Get duration early (needed for distributed frame sampling)
    $duration = if ($probe -and $probe.format -and $probe.format.duration) { [double]$probe.format.duration } else { 0 }
    $totalFramesEst = if ($duration -gt 0) { [math]::Round($duration * 24) } else { 0 }

    # ── 2. Frame/GOP analysis (distributed across file) ──
    $samplingMode = if ($FullFrameScan) { 'FULL' } elseif ($duration -gt 120) { 'distributed' } else { 'sequential' }
    Write-Progress -Activity "Analyzing $($file.Name)" -Status "[2/9] Frame/GOP analysis ($MaxAnalysisFrames frames, $samplingMode)..." -PercentComplete 15
    Write-Host "    [2/9] Frame/GOP analysis ($samplingMode)..." -ForegroundColor DarkGray
    $frames = Get-FrameData $fp -MaxFrames $MaxAnalysisFrames -AllFrames:$FullFrameScan -Duration $duration
    if ($frames) { $jd.Frames = $frames; Write-FrameReport $rpt $frames $probe $rc }

    # ── 3. Multi-point sampling (additional consistency check) ──
    # Now only runs if frame analysis was sequential (short files) and file is long enough
    if ($duration -gt 300 -and $frames -and $frames._sampling -eq 'sequential') {
        Write-Progress -Activity "Analyzing $($file.Name)" -Status "[3/9] Multi-point sampling (5 positions)..." -PercentComplete 28
        Write-Host "    [3/9] Multi-point sampling..." -ForegroundColor DarkGray
        $multi = Get-MultiPointFrames $fp $duration
        if ($multi -and $multi.Count -gt 0) {
            Write-Section $rpt "MULTI-POINT FRAME SAMPLING"
            $rpt.AppendLine("    (Sampled from 5 positions to verify consistency)") | Out-Null
            foreach ($s in $multi) {
                $si = @($s.Frames | Where-Object { $_.pict_type -eq 'I' }).Count
                $sp = @($s.Frames | Where-Object { $_.pict_type -eq 'P' }).Count
                $sb = @($s.Frames | Where-Object { $_.pict_type -eq 'B' }).Count
                $sizes = $s.Frames | Where-Object { $_.pkt_size } | ForEach-Object { [int]$_.pkt_size }
                $avgSz = if ($sizes.Count -gt 0) { Format-Size ([long]($sizes | Measure-Object -Average).Average) } else { "N/A" }
                $rpt.AppendLine("    $($s.Label) (@$($s.Start)s): $($s.Frames.Count)fr I:$si P:$sp B:$sb avg:$avgSz") | Out-Null
            }
        }
    } else { Write-Host "    [3/9] Multi-point... (covered by distributed sampling)" -ForegroundColor DarkGray }

    # ── 4. QP / HM Analyser ──
    if ($SkipQP) {
        Write-Host "    [4/9] QP analysis... (skipped via -SkipQP)" -ForegroundColor DarkYellow
    } else {
        Write-Progress -Activity "Analyzing $($file.Name)" -Status "[4/9] QP / HM Analyser..." -PercentComplete 38
        Write-Host "    [4/9] QP analysis..." -ForegroundColor DarkGray

        # Try HM Analyser first (HEVC only, gives QP + CABAC + header params)
        $hmData = $null
        $qp = $null
        $codec = $rc['Codec']
        if ($DebugMode) {
            Write-Host "      [DEBUG] rc['Codec'] = '$codec'" -ForegroundColor Magenta
            Write-Host "      [DEBUG] codec -match 'hevc|h265' = $($codec -match 'hevc|h265')" -ForegroundColor Magenta
            Write-Host "      [DEBUG] Tools.TAppDecoder = '$($Tools.TAppDecoder)'" -ForegroundColor Magenta
            Write-Host "      [DEBUG] Tools.FFmpeg = '$($Tools.FFmpeg)'" -ForegroundColor Magenta
        }
        if ($codec -match 'hevc|h265' -and $Tools.TAppDecoder) {
            if ($DebugMode) { Write-Host "      [DEBUG] Calling Get-HMAnalysis..." -ForegroundColor Magenta }
            $hmData = Get-HMAnalysis $fp -MaxFrames $MaxAnalysisFrames
            if ($DebugMode) {
                Write-Host "      [DEBUG] hmData = $(if($hmData){'object with keys: '+($hmData.Keys -join ', ')}else{'NULL'})" -ForegroundColor Magenta
                if ($hmData -and $hmData.QP) { Write-Host "      [DEBUG] hmData.QP.FrameCount = $($hmData.QP.FrameCount)" -ForegroundColor Magenta }
            }
            if ($hmData -and $hmData.QP) {
                $qp = $hmData.QP
                Write-Host "      (QP via HM Reference Decoder Analyser)" -ForegroundColor Green
            }
        } elseif ($codec -match 'h264|avc' -and $Tools.JMDecoder) {
            # H.264: use JM reference decoder for QP stats, B-pyramid confirmation, accurate SPS ref count
            if ($DebugMode) { Write-Host "      [DEBUG] Calling Get-JMAnalysis..." -ForegroundColor Magenta }
            $jmData = Get-JMAnalysis $fp -MaxFrames $MaxAnalysisFrames
            if ($DebugMode) {
                Write-Host "      [DEBUG] jmData = $(if($jmData){'object with keys: '+($jmData.Keys -join ', ')}else{'NULL'})" -ForegroundColor Magenta
                if ($jmData -and $jmData.QP) { Write-Host "      [DEBUG] jmData.QP.FrameCount = $($jmData.QP.FrameCount)" -ForegroundColor Magenta }
            }
            if ($jmData -and $jmData.QP) {
                $qp = $jmData.QP
                Write-Host "      (QP via JM H.264 Reference Decoder)" -ForegroundColor Green
            }
        }

        # Fallback to ffmpeg for codecs without a reference decoder (VP9, AV1, etc.)
        if (-not $qp) {
            $qp = Get-QPData $fp -MaxFrames ([math]::Min($MaxAnalysisFrames, 200))
        }

        if ($qp) { $jd.QP = $qp.Stats; Write-QPReport $rpt $qp $rc }
        else { Write-Host "      (QP data unavailable for this codec)" -ForegroundColor DarkYellow }

        # Write HM Analyser detailed report if available (HEVC)
        if ($hmData) {
            $jd.HM = $hmData
            Write-HMAnalyserReport $rpt $hmData $rc

            # Override FFprobe ref count with actual HM-measured max references
            if ($hmData.QP -and $hmData.QP.RefStructure -and $hmData.QP.RefStructure.MaxL0Refs) {
                $hmMaxRef = [int]$hmData.QP.RefStructure.MaxL0Refs
                if ($hmMaxRef -gt [int]$rc['Refs']) {
                    $rc['Refs'] = $hmMaxRef
                    if ($DebugMode) { Write-Host "      [DEBUG] Refs overridden: FFprobe=$($rc['Refs']) -> HM MaxL0=$hmMaxRef" -ForegroundColor Magenta }
                }
            }
            # Store DPB size and MaxReorder from SPS for encoder command generation
            # hmData.SPS is a hashtable — use ['key'] not .Property
            # Note: rc['DPBSize'] is also set by Write-HMAnalyserReport (ReportWriter); this is a belt-and-suspenders backup
            if ($hmData.SPS) {
                if ($hmData.SPS['dpb_size'])         { $rc['DPBSize']   = $hmData.SPS['dpb_size'] }
                if ($hmData.SPS['max_reorder_pics']) { $rc['MaxReorder'] = $hmData.SPS['max_reorder_pics'] }
            }
        }

        # Write JM Analyser detailed report if available (H.264)
        if ($jmData) {
            $jd.JM = $jmData
            Write-JMAnalyserReport $rpt $jmData $rc

            # Override FFprobe ref count with JM-probed SPS value (fixes the refs=1 container bug)
            if ($jmData.SPS -and $jmData.SPS['refs'] -and [int]$jmData.SPS['refs'] -gt [int]$rc['Refs']) {
                if ($DebugMode) { Write-Host "      [DEBUG] Refs overridden: FFprobe=$($rc['Refs']) -> JM SPS refs=$($jmData.SPS['refs'])" -ForegroundColor Magenta }
                $rc['Refs'] = $jmData.SPS['refs']
            }
            # Store B-pyramid in rc for encoder command generation
            if ($jmData.QP.BPyramid) { $rc['BPyramid'] = $true; $rc['BPyramid_Source'] = 'JM-confirmed' }
        }

        # ── Post-hoc: patch B-Pyramid line in GOP section if decoder confirmed it ──
        # Write-FrameReport runs at step 2 (before HM/JM), so it writes the heuristic.
        # Now that we have a definitive answer, replace that line in the report buffer.
        if ($rc['BPyramid_Source'] -match 'HM-confirmed|JM-confirmed') {
            $confirmedStr = if ($rc['BPyramid'] -eq $true) {
                if ($rc['BPyramid_Source'] -eq 'HM-confirmed') {
                    "CONFIRMED (from HM reference list analysis)"
                } else {
                    "CONFIRMED (from JM decoded output)"
                }
            } else {
                if ($rc['BPyramid_Source'] -eq 'HM-confirmed') {
                    "NOT DETECTED (HM reference list analysis)"
                } else {
                    "NOT DETECTED (JM decoded output)"
                }
            }
            # Replace any heuristic line — matches "Likely ENABLED..." or "Likely DISABLED..."
            $rptStr = $rpt.ToString()
            $rptStr = [regex]::Replace(
                $rptStr,
                '(    B-Pyramid\s+: )Likely (?:ENABLED|DISABLED)[^\r\n]*',
                "`${1}$confirmedStr"
            )
            $rpt.Clear() | Out-Null
            $rpt.Append($rptStr) | Out-Null
            if ($DebugMode) { Write-Host "      [DEBUG] B-Pyramid GOP line patched: $confirmedStr" -ForegroundColor Magenta }
        }
    }

    # ── Intermediate save (in case later steps hang) ──
    $reportPath = Join-Path $OutputDir "$($file.BaseName)_analysis.txt"
    $rpt.ToString() | Out-File -FilePath $reportPath -Encoding UTF8

    # ── 5. NAL/bitstream analysis (skip if HM Analyser already provided this) ──
    if ($hmData -and $hmData.HeaderParams -and $hmData.HeaderParams.Count -gt 10) {
        Write-Host "    [5/9] NAL/bitstream... (covered by HM Analyser)" -ForegroundColor DarkGray
    } elseif ($jmData -and $jmData.SPS -and $jmData.SPS['refs']) {
        Write-Host "    [5/9] NAL/bitstream... (covered by JM Analyser)" -ForegroundColor DarkGray
    } else {
        Write-Progress -Activity "Analyzing $($file.Name)" -Status "[5/9] NAL/bitstream params..." -PercentComplete 50
        Write-Host "    [5/9] NAL/bitstream params..." -ForegroundColor DarkGray
        $nal = Get-NALData $fp
        if ($nal) { $jd.NAL = $nal; Write-NALReport $rpt $nal $rc }
    }

    # ── 6. Encoder detection ──
    Write-Progress -Activity "Analyzing $($file.Name)" -Status "[6/9] Encoder detection..." -PercentComplete 60
    Write-Host "    [6/9] Encoder detection..." -ForegroundColor DarkGray
    $enc = Get-EncoderInfo $fp
    if ($enc) { Write-EncoderReport $rpt $enc $rc }

    # ── 7. MediaInfo ──
    Write-Progress -Activity "Analyzing $($file.Name)" -Status "[7/9] MediaInfo..." -PercentComplete 70
    Write-Host "    [7/9] MediaInfo..." -ForegroundColor DarkGray
    $mi = Get-MediaInfoData $fp
    if ($mi) { $jd.MI = $mi.Json; Write-MediaInfoReport $rpt $mi $rc }

    # ── Intermediate save ──
    $rpt.ToString() | Out-File -FilePath $reportPath -Encoding UTF8

    # ── 8. Scene cuts ──
    Write-Progress -Activity "Analyzing $($file.Name)" -Status "[8/9] Scene cut detection..." -PercentComplete 80
    Write-Host "    [8/9] Scene cuts..." -ForegroundColor DarkGray
    $scenes = Get-SceneCuts $fp -MaxSeconds ([math]::Min([int]$duration, 300))
    if ($scenes) { Write-SceneCutReport $rpt $scenes $rc }

    # ── 9. CheckBitrate ──
    Write-Progress -Activity "Analyzing $($file.Name)" -Status "[9/9] CheckBitrate analysis..." -PercentComplete 88
    Write-Host "    [9/9] CheckBitrate..." -ForegroundColor DarkGray
    $cb = Get-CheckBitrateData $fp -Interval $CheckBitrateInterval
    if ($cb) { $jd.CB = $cb; Write-CheckBitrateReport $rpt $cb $OutputDir $file.BaseName $rc }

    # ── 10. GPU Capability validation (if --CheckGPU, runs BEFORE reconstruction) ──
    if ($script:GPUCaps) {
        $script:GPUValidation = Validate-EncoderCommands -rc $rc -caps $script:GPUCaps
        # Apply GPU-validated adjustments to $rc so reconstruction uses corrected values
        if ($script:GPUValidation.AdjustedRC) {
            foreach ($key in $script:GPUValidation.AdjustedRC.Keys) {
                $rc[$key] = $script:GPUValidation.AdjustedRC[$key]
            }
        }
    }

    # ── 11. RECONSTRUCTION (the big payoff) ──
    Write-Progress -Activity "Analyzing $($file.Name)" -Status "Generating reconstruction report..." -PercentComplete 95
    Write-ReconstructionReport $rpt $rc

    # ── 12. GPU Capability report (write after reconstruction) ──
    if ($script:GPUCaps) {
        Write-GPUCapabilityReport $rpt $script:GPUCaps $script:GPUValidation
    }

    # ── 13. Verification encode (if -Verify) ──
    if ($Verify -and $script:GPUCaps -and $frames -and $frames._segments) {
        Write-Progress -Activity "Analyzing $($file.Name)" -Status "Verification encode..." -PercentComplete 97
        $vdLabel = if ($VerifyDuration -gt 0) { "extended (${VerifyDuration}s)" } else { "quick" }
        Write-Host "    [Verify] Running $vdLabel verification encodes..." -ForegroundColor Cyan

        $verifySeg = Get-VerificationSegments -SegmentInfo $frames._segments -FileDuration $duration `
            -FPS $(if($rc['FPS']){$rc['FPS']}else{24}) -MaxSegments 3 `
            -ExtendedDuration $VerifyDuration

        if ($verifySeg.Count -gt 0) {
            $verifyResults = Invoke-VerificationEncode `
                -SourcePath $fp -rc $rc -Segments $verifySeg `
                -WorkDir $OutputDir -Tools $Tools -GPUCaps $script:GPUCaps

            Write-VerificationReport $rpt $verifyResults $verifySeg

            # Clean up verification files
            $verifyDir = Join-Path $OutputDir "verify_encode"
            if (Test-Path $verifyDir) {
                Remove-Item $verifyDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        } else {
            Write-Host "      No segments available for verification" -ForegroundColor Yellow
        }
    }

    # ── Save report ──
    $rpt.AppendLine("") | Out-Null
    $rpt.AppendLine($script:Sep) | Out-Null
    $rpt.AppendLine("  END OF REPORT") | Out-Null
    $rpt.AppendLine($script:Sep) | Out-Null

    $reportPath = Join-Path $OutputDir "$($file.BaseName)_analysis.txt"
    $rpt.ToString() | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Progress -Activity "Analyzing $($file.Name)" -Completed
    $fileTimer.Stop()
    Write-Host "    Report: $reportPath" -ForegroundColor Green
    Write-Host "    Completed in $([math]::Round($fileTimer.Elapsed.TotalSeconds,1))s" -ForegroundColor DarkGray
    $allReports += $reportPath

    if ($ExportJson) {
        $jsonPath = Join-Path $OutputDir "$($file.BaseName)_analysis.json"
        $jd | ConvertTo-Json -Depth 20 | Out-File -FilePath $jsonPath -Encoding UTF8
        Write-Host "    JSON:   $jsonPath" -ForegroundColor Green
    }
}

# ── Summary ──
Write-Host "`n$($script:Sep)" -ForegroundColor Cyan
Write-Host "  COMPLETE: $($allReports.Count)/$($Files.Count) files analyzed" -ForegroundColor Cyan
Write-Host "  Reports: $OutputDir" -ForegroundColor Yellow
Write-Host $script:Sep -ForegroundColor Cyan
Write-Host ""