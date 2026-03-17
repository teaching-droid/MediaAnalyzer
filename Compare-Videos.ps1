<#
.SYNOPSIS
    Compare one original video against one or more re-encoded files and show a side-by-side
    quality/parameter table with winners highlighted.
.DESCRIPTION
    First argument is always the reference (original). All remaining arguments are encoded
    files to compare against it. Runs VMAF/PSNR/SSIM and bitstream analysis (HM/JM) for
    each encoded file, then prints a summary comparison table.
.EXAMPLE
    # Two-file compare (original vs one encode)
    .\Compare-Videos.ps1 original.mkv encode_nvenc.mkv

    # Multi-file compare (original vs three encodes)
    .\Compare-Videos.ps1 original.mkv enc_nvenc.mkv enc_qsvenc.mkv enc_x265.mkv

    # Drag-and-drop: drag reference first, then encodes, onto this script

    # Options
    .\Compare-Videos.ps1 original.mkv enc1.mkv enc2.mkv -SampleDuration 120
    .\Compare-Videos.ps1 original.mkv enc1.mkv enc2.mkv -FullFile
    .\Compare-Videos.ps1 original.mkv enc1.mkv enc2.mkv -SkipMetrics
#>
[CmdletBinding()]
param(
    [Parameter(Position=0)][string]$Reference,
    [Parameter(Position=1, ValueFromRemainingArguments=$true)][string[]]$EncodedFiles,
    [string]$OutputDir,
    [int]$SampleDuration = 60,
    [switch]$FullFile,
    [switch]$DebugMode,
    [switch]$SkipMetrics
)

# ── Clean up any quoted paths (drag-and-drop on Windows can add quotes) ──
if ($Reference)     { $Reference = $Reference.Trim('"').Trim("'") }
if ($EncodedFiles)  { $EncodedFiles = $EncodedFiles | ForEach-Object { $_.Trim('"').Trim("'") } | Where-Object { $_ -ne '' } }

if (-not $Reference -or -not $EncodedFiles -or $EncodedFiles.Count -eq 0) {
    Write-Host ""
    Write-Host "  Usage: .\Compare-Videos.ps1 <reference> <encoded1> [encoded2] [encoded3] ..." -ForegroundColor Yellow
    Write-Host "  Or:    Drag and drop reference first, then encoded files, onto this script." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Options:" -ForegroundColor DarkGray
    Write-Host "    -SampleDuration <sec>   VMAF sample length (default: 60s)" -ForegroundColor DarkGray
    Write-Host "    -FullFile               Run metrics on full file (slow)" -ForegroundColor DarkGray
    Write-Host "    -SkipMetrics            Parameter comparison only, no VMAF" -ForegroundColor DarkGray
    Write-Host "    -DebugMode              Verbose output" -ForegroundColor DarkGray
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

if (-not (Test-Path $Reference)) {
    Write-Error "Reference file not found: $Reference"
    Read-Host "Press Enter"; exit 1
}
foreach ($ef in $EncodedFiles) {
    if (-not (Test-Path $ef)) {
        Write-Error "Encoded file not found: $ef"
        Read-Host "Press Enter"; exit 1
    }
}

# ── Load lib files ──
$scriptDir = $PSScriptRoot
. "$scriptDir\lib\Helpers.ps1"
. "$scriptDir\lib\Collectors.ps1"
. "$scriptDir\lib\HMAnalyser.ps1"
. "$scriptDir\lib\JMAnalyser.ps1"

$script:DebugMode = $DebugMode
$script:SkipCB    = $true

# ── Discover tools ──
$Tools = @{
    FFprobe     = Find-Tool 'ffprobe'             @('ffprobe.exe')
    FFmpeg      = Find-Tool 'ffmpeg'              @('ffmpeg.exe')
    MediaInfo   = Find-Tool 'mediainfo'           @('mediainfo.exe','MediaInfo.exe')
    TAppDecoder = Find-Tool 'TAppDecoderAnalyser' @('TAppDecoderAnalyser.exe','TAppDecoder.exe')
    JMDecoder   = Find-Tool 'ldecod'              @('ldecod.exe','ldecod')
    DoviTool    = Find-Tool 'dovi_tool'           @('dovi_tool.exe','dovi_tool')
}

$Sep  = "=" * 90
$Sep2 = "-" * 90
$n    = $EncodedFiles.Count

Write-Host ""
Write-Host $Sep -ForegroundColor Cyan
Write-Host "  VIDEO COMPARISON TOOL v2.0  ($n file$(if($n -ne 1){'s'}) vs reference)" -ForegroundColor Cyan
Write-Host $Sep -ForegroundColor Cyan
Write-Host ""
Write-Host "  Reference : $(Split-Path $Reference -Leaf)" -ForegroundColor White
for ($i = 0; $i -lt $n; $i++) {
    Write-Host "  Encode #$($i+1)  : $(Split-Path $EncodedFiles[$i] -Leaf)" -ForegroundColor White
}
Write-Host ""

# ═══════════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

function BoolStr2 {
    param($val)
    if ($null -eq $val) { return 'N/A' }
    if ($val) { return 'Enabled' } else { return 'Disabled' }
}

function Fmt-Duration {
    param($sec)
    if (-not $sec -or [double]$sec -le 0) { return '?' }
    $s  = [int][double]$sec
    $h  = [math]::Floor($s / 3600)
    $m  = [math]::Floor(($s % 3600) / 60)
    $ss = $s % 60
    if ($h -gt 0)     { return "${h}h ${m}m ${ss}s" }
    elseif ($m -gt 0) { return "${m}m ${ss}s" }
    else              { return "${ss}s" }
}

# FFprobe a file into a rc hashtable
function Get-FileRC {
    param([string]$FilePath)
    $rc = @{}
    $probe = Get-ProbeJson $FilePath
    if (-not $probe) { return $rc }
    foreach ($st in $probe.streams) {
        if ($st.codec_type -eq 'video' -and $st.codec_name -ne 'png' -and -not $rc['Codec']) {
            $rc['Codec']         = $st.codec_name
            $rc['Profile']       = $st.profile
            $rc['Level']         = $st.level
            $rc['Resolution']    = "$($st.width)x$($st.height)"
            $rc['Width']         = [int]$st.width
            $rc['Height']        = [int]$st.height
            $rc['PixFmt']        = $st.pix_fmt
            $rc['Refs']          = $st.refs
            $rc['ColorSpace']    = $st.color_space
            $rc['ColorTransfer'] = $st.color_transfer
            $rc['ColorPrimaries']= $st.color_primaries
            $rc['ColorRange']    = $st.color_range
            $rc['BitDepth'] = if ($st.bits_per_raw_sample -and [int]$st.bits_per_raw_sample -gt 0) {
                                  $st.bits_per_raw_sample
                              } elseif ($st.pix_fmt -match '10') { 10 } else { 8 }
            if ($st.bit_rate) { $rc['VideoBitrateKbps'] = [math]::Round([double]$st.bit_rate / 1000) }
            if ($st.r_frame_rate) {
                $p = $st.r_frame_rate -split '/'
                if ($p.Count -eq 2 -and [double]$p[1] -gt 0) {
                    $rc['FPS'] = [math]::Round([double]$p[0] / [double]$p[1], 3)
                }
            }
        }
    }
    if ($probe.format) {
        $rc['Duration']             = if ($probe.format.duration) { [math]::Round([double]$probe.format.duration, 1) } else { 0 }
        $rc['FileSizeMB']           = [math]::Round([double]$probe.format.size / 1MB, 2)
        $rc['FileSizeBytes']        = [long]$probe.format.size
        $rc['ContainerBitrateKbps'] = if ($probe.format.bit_rate) { [math]::Round([double]$probe.format.bit_rate / 1000) } else { 0 }
    }
    return $rc
}

# Extract keyint and max B-frames from frame data
function Get-GOPData {
    param($frameData)
    $result = @{ Keyint = $null; BFrames = $null }
    if (-not $frameData -or -not $frameData.frames) { return $result }
    $gopLengths = @(); $gopLen = 0; $maxB = 0; $curB = 0
    foreach ($f in $frameData.frames) {
        $gopLen++
        if ($f.pict_type -eq 'I' -and $gopLen -gt 1) { $gopLengths += $gopLen; $gopLen = 0 }
        if ($f.pict_type -eq 'B') { $curB++; if ($curB -gt $maxB) { $maxB = $curB } }
        else { $curB = 0 }
    }
    if ($gopLengths.Count -gt 0) { $result['Keyint']  = ($gopLengths | Measure-Object -Maximum).Maximum }
    if ($maxB -gt 0)              { $result['BFrames'] = $maxB }
    return $result
}

# Apply HM analysis into rc hashtable
function Apply-HMtoRC {
    param($hm, $rc)
    if (-not $hm) { return }
    if ($hm.QP) {
        if ($hm.QP.Stats) {
            $rc['QP_Avg'] = $hm.QP.Stats.Avg
            $rc['QP_Min'] = $hm.QP.Stats.Min
            $rc['QP_Max'] = $hm.QP.Stats.Max
        }
        if ($null -ne $hm.QP.BPyramid) {
            $rc['BPyramid']        = $hm.QP.BPyramid
            $rc['BPyramid_Source'] = 'HM-confirmed'
        }
        if ($hm.QP.RefStructure) {
            $rc['MaxL0Refs'] = $hm.QP.RefStructure.MaxL0Refs
            $rc['MaxL1Refs'] = $hm.QP.RefStructure.MaxL1Refs
        }
    }
    if ($hm.SPS) {
        if ($null -ne $hm.SPS['sao_enabled'])            { $rc['SAO']         = $hm.SPS['sao_enabled'] }
        if ($null -ne $hm.SPS['amp_enabled'])            { $rc['AMP']         = $hm.SPS['amp_enabled'] }
        if ($null -ne $hm.SPS['strong_intra_smoothing']) { $rc['StrongIntra'] = $hm.SPS['strong_intra_smoothing'] }
        if ($null -ne $hm.SPS['temporal_mvp'])           { $rc['TemporalMVP'] = $hm.SPS['temporal_mvp'] }
        if ($hm.SPS['log2_max_cu'])                      { $rc['CTU']         = [math]::Pow(2, $hm.SPS['log2_max_cu']) }
        if ($hm.SPS['dpb_size'])                         { $rc['DPBSize']     = $hm.SPS['dpb_size'] }
        if ($hm.SPS['max_reorder_pics'])                 { $rc['MaxReorder']  = $hm.SPS['max_reorder_pics'] }
        if ($hm.SPS['bit_depth_luma'])                   { $rc['BitDepth']    = $hm.SPS['bit_depth_luma'] }
    }
    if ($hm.PPS) {
        if ($null -ne $hm.PPS['weighted_pred'])   { $rc['WeightedPred']   = $hm.PPS['weighted_pred'] }
        if ($null -ne $hm.PPS['weighted_bipred']) { $rc['WeightedBipred'] = $hm.PPS['weighted_bipred'] }
        if ($null -ne $hm.PPS['cu_qp_delta'])     { $rc['CUQPDelta']      = $hm.PPS['cu_qp_delta'] }
    }
    if ($hm.VUI) { $rc['_HM_VUI'] = $hm.VUI }
}

# Apply JM analysis into rc hashtable
function Apply-JMtoRC {
    param($jm, $rc)
    if (-not $jm) { return }
    if ($jm.QP) {
        if ($jm.QP.Stats) {
            $rc['QP_Avg'] = $jm.QP.Stats.Avg
            $rc['QP_Min'] = $jm.QP.Stats.Min
            $rc['QP_Max'] = $jm.QP.Stats.Max
        }
        if ($jm.QP.BPyramid) {
            $rc['BPyramid']        = $true
            $rc['BPyramid_Source'] = 'JM-confirmed'
        }
    }
    if ($jm.SPS -and $jm.SPS['refs']) { $rc['Refs'] = $jm.SPS['refs'] }
}

# Run VMAF/PSNR/SSIM using distributed multi-point sampling (same as HM analyser):
# beginning / 25% / middle / 75% / end — runs each metric on all segments, averages results.
function Get-QualityMetrics {
    param(
        [string]$RefFile,
        [string]$EncFile,
        [double]$Duration,
        [int]$SegSec,
        [bool]$IsFullFile,
        [int]$RefW, [int]$RefH,
        [int]$EncW, [int]$EncH,
        [string]$TempDir,
        [int]$EncIdx
    )
    $result = @{ VMAF = $null; PSNR = $null; SSIM = $null; Segments = 0 }
    if (-not $Tools.FFmpeg) { return $result }

    $scaleFilter = ""
    if ($RefW -gt 0 -and $EncW -gt 0 -and ($RefW -ne $EncW -or $RefH -ne $EncH)) {
        Write-Host "      Resolution mismatch (${RefW}x${RefH} vs ${EncW}x${EncH}) — scaling enc to ref..." -ForegroundColor DarkYellow
        $scaleFilter = "scale=${RefW}:${RefH}:flags=lanczos,"
    }

    $vmafCheck = Run-Command $Tools.FFmpeg @('-filters') -TimeoutSeconds 10
    $vmafAvail = ($vmafCheck.StdOut -match 'libvmaf' -or $vmafCheck.StdErr -match 'libvmaf')

    # ── Build segment list (mirrors HM analyser logic) ──
    if ($IsFullFile -or $Duration -le 0) {
        $segments = @( @{ Start = 0; Label = "full" } )
        $segDur = 0   # 0 = no trimming, use full file
    } else {
        $numSegs = if ($Duration -gt 1200) { 5 }       # >20 min: 5 points
                   elseif ($Duration -gt 300)  { 3 }   # >5 min:  3 points
                   else                        { 1 }   # short:   single segment from start
        $segDur = [math]::Min($SegSec, [math]::Floor($Duration / ($numSegs + 1)))
        $segDur = [math]::Max($segDur, 10)  # at least 10s per segment

        if ($numSegs -ge 5) {
            $segments = @(
                @{ Start = 0;                                                                                Label = "beginning" }
                @{ Start = [math]::Floor($Duration * 0.25) - [math]::Floor($segDur / 2); Label = "25%"       }
                @{ Start = [math]::Floor($Duration * 0.50) - [math]::Floor($segDur / 2); Label = "middle"    }
                @{ Start = [math]::Floor($Duration * 0.75) - [math]::Floor($segDur / 2); Label = "75%"       }
                @{ Start = [math]::Max(0, [math]::Floor($Duration) - $segDur - 5);       Label = "end"       }
            )
        } elseif ($numSegs -ge 3) {
            $segments = @(
                @{ Start = 0;                                                                                Label = "beginning" }
                @{ Start = [math]::Floor($Duration * 0.50) - [math]::Floor($segDur / 2); Label = "middle"    }
                @{ Start = [math]::Max(0, [math]::Floor($Duration) - $segDur - 5);       Label = "end"       }
            )
        } else {
            $segments = @( @{ Start = 0; Label = "start" } )
        }
        # Clamp all starts to valid range
        $segments = $segments | ForEach-Object {
            $_.Start = [math]::Max(0, [math]::Min($_.Start, [math]::Floor($Duration) - $segDur - 1))
            $_
        }
    }

    $nSeg = $segments.Count
    Write-Host "      Multi-point sampling: $nSeg x ${segDur}s segments" -ForegroundColor DarkGray

    $vmafVals = @(); $psnrVals = @(); $ssimVals = @()

    for ($si = 0; $si -lt $nSeg; $si++) {
        $seg   = $segments[$si]
        $label = $seg.Label
        $start = $seg.Start
        $sIdx  = $si + 1

        Write-Host "      [$sIdx/$nSeg] $label (@${start}s, ${segDur}s)..." -ForegroundColor DarkGray

        # Extract matching ref + enc segments
        $rSeg = Join-Path $TempDir "ref_e${EncIdx}_s${si}.mkv"
        $eSeg = Join-Path $TempDir "enc_e${EncIdx}_s${si}.mkv"

        if ($IsFullFile -or $segDur -le 0) {
            $rSeg = $RefFile
            $eSeg = $EncFile
        } else {
            $exR = @('-hide_banner','-y','-ss',"$start",'-t',"$segDur",'-i',"`"$RefFile`"",'-c:v','copy','-an',"`"$rSeg`"")
            $exE = @('-hide_banner','-y','-ss',"$start",'-t',"$segDur",'-i',"`"$EncFile`"",'-c:v','copy','-an',"`"$eSeg`"")
            Run-Command $Tools.FFmpeg $exR -TimeoutSeconds 120 | Out-Null
            Run-Command $Tools.FFmpeg $exE -TimeoutSeconds 120 | Out-Null
        }
        if (-not (Test-Path $rSeg) -or -not (Test-Path $eSeg)) {
            Write-Host "      [$sIdx/$nSeg] Segment extraction failed, skipping." -ForegroundColor DarkYellow
            continue
        }

        # VMAF
        if ($vmafAvail) {
            $vf = "[0:v]setpts=PTS-STARTPTS[ref];[1:v]${scaleFilter}setpts=PTS-STARTPTS[enc];[ref][enc]libvmaf=shortest=1:n_threads=4"
            $vr = Run-Command $Tools.FFmpeg @('-hide_banner','-an','-sn','-i',"`"$rSeg`"",'-i',"`"$eSeg`"",'-lavfi',"`"$vf`"",'-f','null','-') -TimeoutSeconds 1800 -StatusLabel "VMAF ($label)"
            $vOut = "$($vr.StdOut)`n$($vr.StdErr)"
            if ($script:DebugMode) { Write-Host "      [DEBUG] VMAF($label) exit=$($vr.ExitCode) $(($vOut -split '\n' | Select-String 'VMAF score'))" -ForegroundColor Magenta }
            if ($vr.ExitCode -eq 0 -and $vOut -match 'VMAF score[:\s=]+([\d.]+)') { $vmafVals += [double]$Matches[1] }
        }

        # PSNR
        $pf = "[0:v]setpts=PTS-STARTPTS[ref];[1:v]${scaleFilter}setpts=PTS-STARTPTS[enc];[ref][enc]psnr=shortest=1"
        $pr = Run-Command $Tools.FFmpeg @('-hide_banner','-an','-sn','-i',"`"$rSeg`"",'-i',"`"$eSeg`"",'-lavfi',"`"$pf`"",'-f','null','-') -TimeoutSeconds 1800 -StatusLabel "PSNR ($label)"
        $pOut = "$($pr.StdOut)`n$($pr.StdErr)"
        if ($script:DebugMode) { Write-Host "      [DEBUG] PSNR($label) exit=$($pr.ExitCode) $(($pOut -split '\n' | Select-String 'average:'))" -ForegroundColor Magenta }
        if ($pr.ExitCode -eq 0) {
            if    ($pOut -match '\bPSNR\b.*\baverage:([\d.]+)') { $psnrVals += [double]$Matches[1] }
            elseif ($pOut -match 'average:([\d.]+)')             { $psnrVals += [double]$Matches[1] }
        }

        # SSIM
        $sf = "[0:v]setpts=PTS-STARTPTS[ref];[1:v]${scaleFilter}setpts=PTS-STARTPTS[enc];[ref][enc]ssim=shortest=1"
        $sr = Run-Command $Tools.FFmpeg @('-hide_banner','-an','-sn','-i',"`"$rSeg`"",'-i',"`"$eSeg`"",'-lavfi',"`"$sf`"",'-f','null','-') -TimeoutSeconds 1800 -StatusLabel "SSIM ($label)"
        $sOut = "$($sr.StdOut)`n$($sr.StdErr)"
        if ($script:DebugMode) { Write-Host "      [DEBUG] SSIM($label) exit=$($sr.ExitCode) $(($sOut -split '\n' | Select-String 'All:'))" -ForegroundColor Magenta }
        if ($sr.ExitCode -eq 0) {
            if    ($sOut -match '\bSSIM\b.*\bAll:([\d.]+)') { $ssimVals += [double]$Matches[1] }
            elseif ($sOut -match '\bAll:([\d.]+)')           { $ssimVals += [double]$Matches[1] }
        }
    }

    # Average across all segments
    if ($vmafVals.Count -gt 0) { $result['VMAF'] = [math]::Round(($vmafVals | Measure-Object -Average).Average, 2) }
    if ($psnrVals.Count -gt 0) { $result['PSNR'] = [math]::Round(($psnrVals | Measure-Object -Average).Average, 2) }
    if ($ssimVals.Count -gt 0) { $result['SSIM'] = [math]::Round(($ssimVals | Measure-Object -Average).Average, 4) }
    $result['Segments'] = $nSeg

    if ($script:DebugMode) {
        Write-Host "      [DEBUG] Averaged over $nSeg segments: VMAF=$($result['VMAF']) PSNR=$($result['PSNR']) SSIM=$($result['SSIM'])" -ForegroundColor Magenta
    }

    return $result
}

# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 2 — Create temp dir (segments extracted per-encode inside Get-QualityMetrics)
# ═══════════════════════════════════════════════════════════════════════════════
$tempDir = Join-Path $env:TEMP "MediaAnalyzer_CMP_$(Get-Random)"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 1 — Analyze reference file
# ═══════════════════════════════════════════════════════════════════════════════
Write-Host "  Analyzing reference file..." -ForegroundColor DarkGray
$rcRef   = Get-FileRC $Reference
$refItem = Get-Item $Reference
$codecRef = "$($rcRef['Codec'])".ToLower()

if ($Tools.FFprobe) {
    $fdRef  = Get-FrameData $Reference -MaxFrames 500 -Duration ([double]$rcRef['Duration'])
    $gopRef = Get-GOPData $fdRef
    if ($null -ne $gopRef['Keyint'])  { $rcRef['Keyint']  = $gopRef['Keyint'] }
    if ($null -ne $gopRef['BFrames']) { $rcRef['BFrames'] = $gopRef['BFrames'] }
}

$hmRef = $null; $jmRef = $null
if ($codecRef -match 'hevc|h265' -and $Tools.TAppDecoder) {
    Write-Host "    HM bitstream analysis..." -ForegroundColor DarkGray
    $hmRef = Get-HMAnalysis $Reference -MaxFrames 500
    Apply-HMtoRC $hmRef $rcRef
} elseif ($codecRef -match 'h264|avc' -and $Tools.JMDecoder) {
    Write-Host "    JM bitstream analysis..." -ForegroundColor DarkGray
    $jmRef = Get-JMAnalysis $Reference -MaxFrames 500
    Apply-JMtoRC $jmRef $rcRef
}

# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 3 — Analyze each encoded file
# ═══════════════════════════════════════════════════════════════════════════════
# Each entry: rc hashtable + quality metrics
$encResults = @()

for ($i = 0; $i -lt $n; $i++) {
    $encFile  = $EncodedFiles[$i]
    $encLabel = "Encode #$($i+1) — $(Split-Path $encFile -Leaf)"
    Write-Host ""
    Write-Host "  [$($i+1)/$n] $encLabel" -ForegroundColor Cyan

    $rcE     = Get-FileRC $encFile
    $codecE  = "$($rcE['Codec'])".ToLower()

    # Frame/GOP
    if ($Tools.FFprobe) {
        $fdE  = Get-FrameData $encFile -MaxFrames 500 -Duration ([double]$rcE['Duration'])
        $gopE = Get-GOPData $fdE
        if ($null -ne $gopE['Keyint'])  { $rcE['Keyint']  = $gopE['Keyint'] }
        if ($null -ne $gopE['BFrames']) { $rcE['BFrames'] = $gopE['BFrames'] }
    }

    # HM/JM analysis
    $hmE = $null; $jmE = $null
    if ($codecE -match 'hevc|h265' -and $Tools.TAppDecoder) {
        Write-Host "    HM bitstream analysis..." -ForegroundColor DarkGray
        $hmE = Get-HMAnalysis $encFile -MaxFrames 500
        Apply-HMtoRC $hmE $rcE
    } elseif ($codecE -match 'h264|avc' -and $Tools.JMDecoder) {
        Write-Host "    JM bitstream analysis..." -ForegroundColor DarkGray
        $jmE = Get-JMAnalysis $encFile -MaxFrames 500
        Apply-JMtoRC $jmE $rcE
    } else {
        Write-Host "    No reference decoder for '$codecE'" -ForegroundColor DarkYellow
    }

    # Quality metrics
    $metrics = @{ VMAF = $null; PSNR = $null; SSIM = $null }
    if (-not $SkipMetrics -and $Tools.FFmpeg) {
        Write-Host "    Quality metrics (VMAF/PSNR/SSIM)..." -ForegroundColor DarkGray
        $metrics = Get-QualityMetrics `
            -RefFile   $Reference `
            -EncFile   $encFile `
            -Duration  ([double]$rcRef['Duration']) `
            -SegSec    $SampleDuration `
            -IsFullFile ([bool]$FullFile) `
            -RefW      ([int]$rcRef['Width']) `
            -RefH      ([int]$rcRef['Height']) `
            -EncW      ([int]$rcE['Width']) `
            -EncH      ([int]$rcE['Height']) `
            -TempDir   $tempDir `
            -EncIdx    $i
    }


    # Encoder name guess from filename / MediaInfo
    $encName = [System.IO.Path]::GetFileNameWithoutExtension($encFile)

    $encResults += @{
        File    = $encFile
        Item    = Get-Item $encFile
        Label   = $encName
        RC      = $rcE
        Metrics = $metrics
        HM      = $hmE
        JM      = $jmE
    }

    # Quick per-file console summary
    $vmafStr = if ($null -ne $metrics['VMAF']) { "VMAF=$($metrics['VMAF'])" } else { "VMAF=N/A" }
    $psnrStr = if ($null -ne $metrics['PSNR']) { "PSNR=$($metrics['PSNR'])dB" } else { "" }
    $ssimStr = if ($null -ne $metrics['SSIM']) { "SSIM=$($metrics['SSIM'])" } else { "" }
    $brStr   = if ($rcE['ContainerBitrateKbps'] -gt 0) { "Bitrate=$($rcE['ContainerBitrateKbps'])kbps" } else { "" }
    $summary = @($vmafStr,$psnrStr,$ssimStr,$brStr) | Where-Object { $_ -ne '' }
    Write-Host "    → $($summary -join '  ')" -ForegroundColor White
}

# Cleanup temp
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }

# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 4 — Build report
# ═══════════════════════════════════════════════════════════════════════════════
$rpt = [System.Text.StringBuilder]::new()
function S { param([string]$line) $rpt.AppendLine($line) | Out-Null; Write-Host $line }

# ── Find winners (highest is best for VMAF/PSNR/SSIM, lowest is best for size/bitrate) ──
$bestVMAF   = $null; $bestPSNR = $null; $bestSSIM = $null
$bestSizeI  = -1; $bestVMAFI = -1; $bestPSNRI = -1; $bestSSIMI = -1
$bestBRI    = -1

for ($i = 0; $i -lt $n; $i++) {
    $m = $encResults[$i].Metrics
    $rc = $encResults[$i].RC
    if ($null -ne $m['VMAF'] -and ($null -eq $bestVMAF -or $m['VMAF'] -gt $bestVMAF)) { $bestVMAF = $m['VMAF']; $bestVMAFI = $i }
    if ($null -ne $m['PSNR'] -and ($null -eq $bestPSNR -or $m['PSNR'] -gt $bestPSNR)) { $bestPSNR = $m['PSNR']; $bestPSNRI = $i }
    if ($null -ne $m['SSIM'] -and ($null -eq $bestSSIM -or $m['SSIM'] -gt $bestSSIM)) { $bestSSIM = $m['SSIM']; $bestSSIMI = $i }
    if ($bestSizeI -eq -1 -or [long]$rc['FileSizeBytes'] -lt [long]$encResults[$bestSizeI].RC['FileSizeBytes']) { $bestSizeI = $i }
    if ($bestBRI -eq -1 -or [double]$rc['ContainerBitrateKbps'] -lt [double]$encResults[$bestBRI].RC['ContainerBitrateKbps']) { $bestBRI = $i }
}

# ── Table layout helpers ──
# Column widths: label col + one col per encode
$labelW = 28
$colW   = [math]::Max(18, (($encResults | ForEach-Object { $_.Label.Length }) | Measure-Object -Maximum).Maximum + 4)

function Pad  { param([string]$s, [int]$w) $s.PadRight($w) }
function LPad { param([string]$s, [int]$w) $s.PadLeft($w) }

# Build a table row. Optional -RefVal adds a Reference column before encode columns.
function Table-Row {
    param(
        [string]$Label,
        [string[]]$Values,
        [int[]]$WinnerIdx = @(),
        [string]$WinColor = 'Green',
        [string]$BaseColor = 'White',
        [string]$LabelColor = 'DarkGray',
        [bool]$IsHeader = $false,
        [string]$RefVal = $null        # if set, shown as first column in DarkGray/Cyan
    )
    $hasRef = ($null -ne $RefVal -and $RefVal -ne '')
    $line = "  $(Pad $Label $labelW)"
    if ($hasRef) { $line += "$(Pad $RefVal $colW)" }
    foreach ($v in $Values) { $line += "$(Pad $v $colW)" }
    $rpt.AppendLine($line) | Out-Null

    Write-Host -NoNewline "  "
    if ($IsHeader) {
        Write-Host -NoNewline (Pad $Label $labelW) -ForegroundColor Cyan
    } else {
        Write-Host -NoNewline (Pad $Label $labelW) -ForegroundColor $LabelColor
    }
    if ($hasRef) {
        $rc = if ($IsHeader) { 'Cyan' } else { 'DarkGray' }
        Write-Host -NoNewline (Pad $RefVal $colW) -ForegroundColor $rc
    }
    for ($vi = 0; $vi -lt $Values.Count; $vi++) {
        $col = if ($IsHeader) { 'Cyan' } elseif ($WinnerIdx -contains $vi) { $WinColor } else { $BaseColor }
        Write-Host -NoNewline (Pad $Values[$vi] $colW) -ForegroundColor $col
    }
    Write-Host ""
}

function Table-Sep {
    param([string]$char = '-', [bool]$WithRef = $false)
    $extra = if ($WithRef) { $colW } else { 0 }
    $line = "  " + ($char * ($labelW + $colW * $n + $extra))
    $rpt.AppendLine($line) | Out-Null; Write-Host $line -ForegroundColor DarkGray
}

# ── Report header ──
Write-Host ""
S ""
S $Sep
S "  VIDEO COMPARISON REPORT"
S $Sep
S "  Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
S ""
S "  Reference  : $($refItem.FullName)"
for ($i = 0; $i -lt $n; $i++) {
    $rpt.AppendLine("  Encode #$($i+1)   : $($encResults[$i].Item.FullName)") | Out-Null
    Write-Host "  Encode #$($i+1)   : $($encResults[$i].Item.FullName)"
}
S ""

# ═══════════════════════════════════════════════════════════════════════════════
# QUALITY METRICS TABLE
# ═══════════════════════════════════════════════════════════════════════════════
S $Sep2
S "  QUALITY METRICS"
S $Sep2
    $nSegs = ($encResults | ForEach-Object { $_.Metrics['Segments'] } | Measure-Object -Maximum).Maximum
    $metricSampleLabel = if ($FullFile) { '(full file)' } elseif ($SkipMetrics) { '(skipped)' } elseif ($nSegs -gt 1) { "($nSegs-point distributed sample, ${SampleDuration}s each)" } else { "(${SampleDuration}s sample)" }
S "  vs Reference $metricSampleLabel"
S ""

# Header row
$headerVals = $encResults | ForEach-Object { "Enc: $($_.Label)" }
Table-Row "Metric" $headerVals -IsHeader $true

Table-Sep

# VMAF row
$vmafVals = $encResults | ForEach-Object {
    if ($null -ne $_.Metrics['VMAF']) { "$($_.Metrics['VMAF'])" } else { 'N/A' }
}
Table-Row "VMAF" $vmafVals -WinnerIdx @($bestVMAFI) -WinColor 'Green' -BaseColor 'White'

# PSNR row
$psnrVals = $encResults | ForEach-Object {
    if ($null -ne $_.Metrics['PSNR']) { "$($_.Metrics['PSNR']) dB" } else { 'N/A' }
}
Table-Row "PSNR (Y)" $psnrVals -WinnerIdx @($bestPSNRI) -WinColor 'Green' -BaseColor 'White'

# SSIM row
$ssimVals = $encResults | ForEach-Object {
    if ($null -ne $_.Metrics['SSIM']) {
        $db = [math]::Round(-10 * [math]::Log10([math]::Max(1 - $_.Metrics['SSIM'], 1e-10)), 2)
        "$($_.Metrics['SSIM']) ($db dB)"
    } else { 'N/A' }
}
Table-Row "SSIM" $ssimVals -WinnerIdx @($bestSSIMI) -WinColor 'Green' -BaseColor 'White'

# VMAF verdict row
$verdictVals = $encResults | ForEach-Object {
    $v = $_.Metrics['VMAF']
    if ($null -eq $v) { 'N/A' }
    elseif ($v -ge 97) { 'Near-transparent' }
    elseif ($v -ge 93) { 'Excellent' }
    elseif ($v -ge 85) { 'Good' }
    elseif ($v -ge 75) { 'Acceptable' }
    elseif ($v -ge 60) { 'Noticeable loss' }
    else               { 'Significant loss' }
}
Table-Row "Quality verdict" $verdictVals -WinnerIdx @($bestVMAFI) -WinColor 'Green' -BaseColor 'DarkGray'

S ""

# ═══════════════════════════════════════════════════════════════════════════════
# FILE / BITRATE TABLE
# ═══════════════════════════════════════════════════════════════════════════════
S $Sep2
S "  FILE & BITRATE"
S $Sep2
S ""

$refBR = [double]$rcRef['ContainerBitrateKbps']

# Reference row
$refVals = @("$($rcRef['Resolution'])", "$(Format-Size $refItem.Length)", "$([math]::Round($refBR))kbps")
$rLine = "  $('REFERENCE'.PadRight($labelW))$('Resolution'.PadRight($colW))$('File Size'.PadRight($colW))$('Bitrate'.PadRight($colW))"
$rpt.AppendLine($rLine) | Out-Null; Write-Host $rLine -ForegroundColor Cyan
Table-Sep
$refDataLine = "  $(''.PadRight($labelW))$($rcRef['Resolution'].PadRight($colW))$((Format-Size $refItem.Length).PadRight($colW))$([math]::Round($refBR))kbps"
$rpt.AppendLine($refDataLine) | Out-Null; Write-Host $refDataLine -ForegroundColor DarkGray
S ""

# Per-encode rows
$hdrVals = $encResults | ForEach-Object { "Enc: $($_.Label)" }
Table-Row "Metric" $hdrVals -IsHeader $true
Table-Sep

$resVals = $encResults | ForEach-Object { "$($_.RC['Resolution'])" }
Table-Row "Resolution" $resVals -WinnerIdx @() -BaseColor 'White'

$sizeVals = $encResults | ForEach-Object { Format-Size $_.Item.Length }
Table-Row "File Size" $sizeVals -WinnerIdx @($bestSizeI) -WinColor 'Cyan' -BaseColor 'White'

$brVals = $encResults | ForEach-Object { "$([math]::Round([double]$_.RC['ContainerBitrateKbps']))kbps" }
Table-Row "Bitrate" $brVals -WinnerIdx @($bestBRI) -WinColor 'Cyan' -BaseColor 'White'

# Bitrate vs reference ratio
$ratioVals = $encResults | ForEach-Object {
    $eb = [double]$_.RC['ContainerBitrateKbps']
    if ($refBR -gt 0 -and $eb -gt 0) {
        $pct = [math]::Round(($eb - $refBR) / $refBR * 100, 1)
        $sign = if ($pct -ge 0) { "+$pct%" } else { "$pct%" }
        "$([math]::Round($eb/$refBR,3))x  ($sign)"
    } else { 'N/A' }
}
Table-Row "vs Ref bitrate" $ratioVals -WinnerIdx @($bestBRI) -WinColor 'Cyan' -BaseColor 'White'

# Bit depth
$bdVals = $encResults | ForEach-Object { "$($_.RC['BitDepth'])-bit" }
Table-Row "Bit Depth" $bdVals -WinnerIdx @() -BaseColor 'White'

S ""

# ═══════════════════════════════════════════════════════════════════════════════
# ENCODING PARAMETERS TABLE
# ═══════════════════════════════════════════════════════════════════════════════
S $Sep2
S "  ENCODING PARAMETERS  (bitstream-level)"
S $Sep2
S ""

Table-Row "Parameter" ($encResults | ForEach-Object { "Enc: $($_.Label)" }) -IsHeader $true -RefVal "Reference"
Table-Sep -WithRef $true

# Helper to highlight encode values that match the reference
function Cmp-Vals {
    param([string]$refVal, [string[]]$encVals)
    $winners = @()
    for ($vi = 0; $vi -lt $encVals.Count; $vi++) {
        if ($encVals[$vi] -eq $refVal) { $winners += $vi }
    }
    return $winners
}

$refCodec = "$($rcRef['Codec'])"

$codecVals  = $encResults | ForEach-Object { "$($_.RC['Codec'])" }
$matchCodec = Cmp-Vals $refCodec $codecVals
Table-Row "Codec" $codecVals -WinnerIdx $matchCodec -WinColor 'Green' -BaseColor 'Yellow' -RefVal $refCodec

$refProfStr = "$($rcRef['Profile'])"
$profVals  = $encResults | ForEach-Object { "$($_.RC['Profile'])" }
$matchProf = Cmp-Vals $refProfStr $profVals
Table-Row "Profile" $profVals -WinnerIdx $matchProf -WinColor 'Green' -BaseColor 'White' -RefVal $refProfStr

# HDR/SDR detection
function Get-HDRLabel {
    param($rc)
    $ct = "$($rc['ColorTransfer'])"
    $cp = "$($rc['ColorPrimaries'])"
    if ($ct -match 'smpte2084')      { return "HDR10 (PQ)" }
    elseif ($ct -match 'arib-std-b67') { return "HLG" }
    elseif ($cp -match 'bt2020')       { return "WCG/SDR" }
    else                               { return "SDR" }
}
$refHDRStr = Get-HDRLabel $rcRef
$hdrVals   = $encResults | ForEach-Object { Get-HDRLabel $_.RC }
$matchHDR  = Cmp-Vals $refHDRStr $hdrVals
Table-Row "HDR/SDR" $hdrVals -WinnerIdx $matchHDR -WinColor 'Green' -BaseColor 'White' -RefVal $refHDRStr

$refKeyStr = if ($rcRef['Keyint']) { "$($rcRef['Keyint'])" } else { 'N/A' }
$keyVals  = $encResults | ForEach-Object { if ($_.RC['Keyint']) { "$($_.RC['Keyint'])" } else { 'N/A' } }
$matchKey = Cmp-Vals $refKeyStr $keyVals
Table-Row "Keyint" $keyVals -WinnerIdx $matchKey -WinColor 'Green' -BaseColor 'White' -RefVal $refKeyStr

$refBfStr = if ($rcRef['BFrames']) { "$($rcRef['BFrames'])" } else { 'N/A' }
$bfVals  = $encResults | ForEach-Object { if ($_.RC['BFrames']) { "$($_.RC['BFrames'])" } else { 'N/A' } }
$matchBf = Cmp-Vals $refBfStr $bfVals
Table-Row "B-Frames (max)" $bfVals -WinnerIdx $matchBf -WinColor 'Green' -BaseColor 'White' -RefVal $refBfStr

$refBpStr = if ($rcRef['BPyramid'] -eq $true) { 'Yes' } elseif ($rcRef['BPyramid'] -eq $false) { 'No' } else { 'N/A' }
$bpVals = $encResults | ForEach-Object {
    if ($_.RC['BPyramid'] -eq $true) { 'Yes' } elseif ($_.RC['BPyramid'] -eq $false) { 'No' } else { 'N/A' }
}
$matchBp = Cmp-Vals $refBpStr $bpVals
Table-Row "B-Pyramid" $bpVals -WinnerIdx $matchBp -WinColor 'Green' -BaseColor 'White' -RefVal $refBpStr

$refRefStr = if ($rcRef['Refs']) { "$($rcRef['Refs'])" } else { 'N/A' }
$refVals2 = $encResults | ForEach-Object { if ($_.RC['Refs']) { "$($_.RC['Refs'])" } else { 'N/A' } }
$matchRef2 = Cmp-Vals $refRefStr $refVals2
Table-Row "Ref Frames" $refVals2 -WinnerIdx $matchRef2 -WinColor 'Green' -BaseColor 'White' -RefVal $refRefStr

# HEVC-specific rows (only if at least one file is HEVC)
$anyHevc = ($codecRef -match 'hevc|h265') -or ($encResults | Where-Object { $_.RC['Codec'] -match 'hevc|h265' }).Count -gt 0
if ($anyHevc) {
    $refCTUStr = if ($rcRef['CTU']) { "$($rcRef['CTU'])" } else { 'N/A' }
    $ctuVals   = $encResults | ForEach-Object { if ($_.RC['CTU']) { "$($_.RC['CTU'])" } else { 'N/A' } }
    $matchCTU  = Cmp-Vals $refCTUStr $ctuVals
    Table-Row "CTU Size" $ctuVals -WinnerIdx $matchCTU -WinColor 'Green' -BaseColor 'White' -RefVal $refCTUStr

    $refSAOStr = BoolStr2 $rcRef['SAO']
    $saoVals  = $encResults | ForEach-Object { BoolStr2 $_.RC['SAO'] }
    $matchSAO = Cmp-Vals $refSAOStr $saoVals
    Table-Row "SAO" $saoVals -WinnerIdx $matchSAO -WinColor 'Green' -BaseColor 'White' -RefVal $refSAOStr

    $refWPStr = BoolStr2 $rcRef['WeightedPred']
    $wpVals  = $encResults | ForEach-Object { BoolStr2 $_.RC['WeightedPred'] }
    $matchWP = Cmp-Vals $refWPStr $wpVals
    Table-Row "Weighted-P" $wpVals -WinnerIdx $matchWP -WinColor 'Green' -BaseColor 'White' -RefVal $refWPStr

    $refWBStr = BoolStr2 $rcRef['WeightedBipred']
    $wbVals  = $encResults | ForEach-Object { BoolStr2 $_.RC['WeightedBipred'] }
    $matchWB = Cmp-Vals $refWBStr $wbVals
    Table-Row "Weighted-B" $wbVals -WinnerIdx $matchWB -WinColor 'Green' -BaseColor 'White' -RefVal $refWBStr
}

# QP range rows (only if available)
$hasQP = $encResults | Where-Object { $null -ne $_.RC['QP_Avg'] }
if ($hasQP -and $null -ne $rcRef['QP_Avg']) {
    S ""
    S "  (QP from reference decoder — higher QP = more compression = lower quality)"
    $refQP    = [double]$rcRef['QP_Avg']
    $refQPStr = "$([math]::Round($refQP,1))"
    $qpAvgVals = $encResults | ForEach-Object {
        if ($null -ne $_.RC['QP_Avg']) { "$([math]::Round([double]$_.RC['QP_Avg'],1))" } else { 'N/A' }
    }
    # Best = closest to reference QP
    $bestQPDiff = $null; $bestQPI = -1
    for ($i = 0; $i -lt $encResults.Count; $i++) {
        if ($null -ne $encResults[$i].RC['QP_Avg']) {
            $diff = [math]::Abs([double]$encResults[$i].RC['QP_Avg'] - $refQP)
            if ($null -eq $bestQPDiff -or $diff -lt $bestQPDiff) { $bestQPDiff = $diff; $bestQPI = $i }
        }
    }
    Table-Row "QP Avg" $qpAvgVals -WinnerIdx @($bestQPI) -WinColor 'Green' -BaseColor 'White' -RefVal $refQPStr

    $refQPMinStr = if ($null -ne $rcRef['QP_Min']) { "$($rcRef['QP_Min'])" } else { 'N/A' }
    $refQPMaxStr = if ($null -ne $rcRef['QP_Max']) { "$($rcRef['QP_Max'])" } else { 'N/A' }
    $qpMinVals = $encResults | ForEach-Object { if ($null -ne $_.RC['QP_Min']) { "$($_.RC['QP_Min'])" } else { 'N/A' } }
    $qpMaxVals = $encResults | ForEach-Object { if ($null -ne $_.RC['QP_Max']) { "$($_.RC['QP_Max'])" } else { 'N/A' } }
    Table-Row "QP Min" $qpMinVals -WinnerIdx @() -BaseColor 'White' -RefVal $refQPMinStr
    Table-Row "QP Max" $qpMaxVals -WinnerIdx @() -BaseColor 'White' -RefVal $refQPMaxStr
}

S ""

# ═══════════════════════════════════════════════════════════════════════════════
# PROPERTY DIFF vs REFERENCE
# ═══════════════════════════════════════════════════════════════════════════════
S $Sep2
S "  DIFFERENCES vs REFERENCE"
S $Sep2
S ""

# Render an RC value using the same display logic as the parameter table
function Render-RCVal {
    param($rc, [string]$key)
    switch ($key) {
        'BPyramid'      { if ($rc[$key] -eq $true -or [int]"$($rc[$key])" -eq 1) { 'Yes' } elseif ($null -ne $rc[$key]) { 'No' } else { '' } }
        'SAO'           { BoolStr2 $rc[$key] }
        'WeightedPred'  { BoolStr2 $rc[$key] }
        'WeightedBipred'{ BoolStr2 $rc[$key] }
        'HDR'           { Get-HDRLabel $rc }
        'CTU'           { if ($rc[$key]) { "$([math]::Round([double]$rc[$key]))" } else { '' } }
        'Keyint'        { if ($rc[$key]) { "$($rc[$key])" } else { '' } }
        'BFrames'       { if ($null -ne $rc[$key]) { "$($rc[$key])" } else { '' } }
        'Refs'          { if ($rc[$key]) { "$($rc[$key])" } else { '' } }
        default         { "$($rc[$key])" }
    }
}

$checkKeys = @(
    @{ Key='Codec';          Label='Codec' }
    @{ Key='Resolution';     Label='Resolution' }
    @{ Key='BitDepth';       Label='Bit Depth' }
    @{ Key='Profile';        Label='Profile' }
    @{ Key='HDR';            Label='HDR/SDR' }
    @{ Key='ColorPrimaries'; Label='Color Primaries' }
    @{ Key='ColorTransfer';  Label='Color Transfer' }
    @{ Key='Keyint';         Label='Keyint' }
    @{ Key='BFrames';        Label='B-Frames' }
    @{ Key='BPyramid';       Label='B-Pyramid' }
    @{ Key='Refs';           Label='Ref Frames' }
    @{ Key='CTU';            Label='CTU Size' }
    @{ Key='SAO';            Label='SAO' }
    @{ Key='WeightedPred';   Label='Weighted-P' }
    @{ Key='WeightedBipred'; Label='Weighted-B' }
)

for ($i = 0; $i -lt $n; $i++) {
    $rcE   = $encResults[$i].RC
    $diffs = @()
    foreach ($ck in $checkKeys) {
        $rv = Render-RCVal $rcRef $ck.Key
        $ev = Render-RCVal $rcE  $ck.Key
        if ($rv -ne '' -and $ev -ne '' -and $rv -ne $ev) {
            $diffs += "$($ck.Label): $rv → $ev"
        }
    }
    $diffLabel = "Encode #$($i+1) $(Split-Path $encResults[$i].File -Leaf)"
    if ($diffs.Count -eq 0) {
        $okLine = "  [==] $($diffLabel.PadRight(50)) No property differences"
        $rpt.AppendLine($okLine) | Out-Null; Write-Host $okLine -ForegroundColor Green
    } else {
        foreach ($d in $diffs) {
            $diffLine = "  [<>] $($diffLabel.PadRight(50)) $d"
            $rpt.AppendLine($diffLine) | Out-Null; Write-Host $diffLine -ForegroundColor Yellow
        }
    }
}

S ""

# ═══════════════════════════════════════════════════════════════════════════════
# OVERALL WINNER SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════
S $Sep
S "  OVERALL WINNER SUMMARY"
S $Sep
S ""

$categories = [ordered]@{
    'Best quality (VMAF)'    = if ($bestVMAFI -ge 0) { "Encode #$($bestVMAFI+1) — $($encResults[$bestVMAFI].Label) ($bestVMAF)" } else { 'N/A (metrics not run)' }
    'Best PSNR'              = if ($bestPSNRI -ge 0) { "Encode #$($bestPSNRI+1) — $($encResults[$bestPSNRI].Label) ($bestPSNR dB)" } else { 'N/A' }
    'Best SSIM'              = if ($bestSSIMI -ge 0) { "Encode #$($bestSSIMI+1) — $($encResults[$bestSSIMI].Label) ($bestSSIM)" } else { 'N/A' }
    'Smallest file size'     = if ($bestSizeI -ge 0) { "Encode #$($bestSizeI+1) — $($encResults[$bestSizeI].Label) ($(Format-Size $encResults[$bestSizeI].Item.Length))" } else { 'N/A' }
    'Lowest bitrate'         = if ($bestBRI -ge 0)   { "Encode #$($bestBRI+1) — $($encResults[$bestBRI].Label) ($([math]::Round([double]$encResults[$bestBRI].RC['ContainerBitrateKbps']))kbps)" } else { 'N/A' }
}

foreach ($cat in $categories.GetEnumerator()) {
    $wLine = "  $($cat.Key.PadRight(30)): $($cat.Value)"
    $rpt.AppendLine($wLine) | Out-Null
    $col = if ($cat.Value -ne 'N/A' -and $cat.Value -ne 'N/A (metrics not run)') { 'Green' } else { 'DarkGray' }
    Write-Host $wLine -ForegroundColor $col
}

# Best quality-per-bit (VMAF per kbps) if we have both
if ($bestVMAFI -ge 0) {
    $bestEffI = -1; $bestEff = $null
    for ($i = 0; $i -lt $n; $i++) {
        $v = $encResults[$i].Metrics['VMAF']
        $b = [double]$encResults[$i].RC['ContainerBitrateKbps']
        if ($null -ne $v -and $b -gt 0) {
            $eff = [math]::Round($v / $b * 100, 4)
            if ($null -eq $bestEff -or $eff -gt $bestEff) { $bestEff = $eff; $bestEffI = $i }
        }
    }
    if ($bestEffI -ge 0) {
        $effLine = "  $('Best quality/bitrate'.PadRight(30)): Encode #$($bestEffI+1) — $($encResults[$bestEffI].Label) (VMAF $($encResults[$bestEffI].Metrics['VMAF']) @ $([math]::Round([double]$encResults[$bestEffI].RC['ContainerBitrateKbps']))kbps)"
        $rpt.AppendLine($effLine) | Out-Null; Write-Host $effLine -ForegroundColor Cyan
    }
}

S ""

# ── Legend ──
S "  Legend: [==] = matches reference  [<>] = differs from reference  [--] = unknown"
S "  Highlighted in GREEN = best quality metric across all encodes"
S "  Highlighted in CYAN  = most compact (smallest file / lowest bitrate)"
S ""

# ═══════════════════════════════════════════════════════════════════════════════
# WRITE REPORT FILE
# ═══════════════════════════════════════════════════════════════════════════════
$targetDir = if ($OutputDir) {
    $OutputDir
} else {
    $sub = Join-Path (Split-Path $Reference -Parent) "MediaAnalysis"
    if (-not (Test-Path $sub)) { New-Item -ItemType Directory -Path $sub -Force | Out-Null }
    $sub
}

$refBase  = [System.IO.Path]::GetFileNameWithoutExtension($Reference)
$encNames = ($encResults | ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.File) }) -join '_vs_'
$rptPath  = Join-Path $targetDir "compare_${refBase}_vs_${encNames}.txt"
$rpt.ToString() | Out-File -FilePath $rptPath -Encoding UTF8

Write-Host ""
Write-Host "  Report saved: $rptPath" -ForegroundColor Cyan
Write-Host ""

# Keep window open when launched via drag-and-drop
if ($Host.Name -eq 'ConsoleHost') {
    try {
        $proc    = Get-CimInstance Win32_Process -Filter "ProcessId=$PID" -ErrorAction SilentlyContinue
        $cmdLine = if ($proc) { $proc.CommandLine } else { '' }
        if ($cmdLine -match '\.(mkv|mp4|m4v|avi|mov|ts|m2ts|wmv|flv)\b') {
            Read-Host "  Press Enter to close"
        }
    } catch {}
}