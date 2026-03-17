# ─────────────────────────────────────────────────────────────────────────────
# Collectors.ps1 — Data collection functions for each tool
# Requires: $Tools hashtable, Run-Command from Helpers.ps1
# ─────────────────────────────────────────────────────────────────────────────

# ── FFprobe: Full stream/format/chapter JSON ──
function Get-ProbeJson {
    param([string]$FilePath)
    if (-not $Tools.FFprobe) { return $null }
    $r = Run-Command $Tools.FFprobe @(
        '-v','quiet','-print_format','json',
        '-show_format','-show_streams','-show_chapters',
        "`"$FilePath`""
    )
    if ($r.ExitCode -eq 0 -and $r.StdOut) {
        $json = $r.StdOut | ConvertFrom-Json
        if ($script:DebugMode) {
            Write-Host "      [DEBUG] FFprobe JSON: $($r.StdOut.Length) chars, streams=$($json.streams.Count)" -ForegroundColor Magenta
            if ($json.streams) {
                foreach ($s in $json.streams) {
                    Write-Host "      [DEBUG]   idx=$($s.index) type=$($s.codec_type) name=$($s.codec_name)" -ForegroundColor Magenta
                }
            }
        }
        return $json
    }
    if ($script:DebugMode) { Write-Host "      [DEBUG] FFprobe FAILED: exit=$($r.ExitCode) stderr=$($r.StdErr)" -ForegroundColor Red }
    return $null
}

# ── FFprobe: Frame-level analysis (GOP, frame types, sizes) ──
# Distributes frame budget across the file for representative sampling
function Get-FrameData {
    param(
        [string]$FilePath,
        [int]$MaxFrames = 1000,
        [switch]$AllFrames,
        [double]$Duration = 0
    )
    if (-not $Tools.FFprobe) { return $null }

    # Full scan or short file: read everything from start
    if ($AllFrames -or $Duration -le 0 -or $MaxFrames -ge ($Duration * 24)) {
        $interval = if (-not $AllFrames) { @('-read_intervals', "0%+#$MaxFrames") } else { @() }
        $r = Run-Command $Tools.FFprobe (@(
            '-v','quiet','-print_format','json','-select_streams','v:0','-show_frames',
            '-show_entries','frame=pict_type,key_frame,pkt_size,pkt_pts_time,pkt_duration_time,interlaced_frame,top_field_first,coded_picture_number,display_picture_number,side_data_list'
        ) + $interval + @("`"$FilePath`"")) -TimeoutSeconds 600 -StatusLabel "ffprobe frames"
        if ($r.ExitCode -eq 0 -and $r.StdOut) {
            $parsed = $r.StdOut | ConvertFrom-Json
            if ($parsed.frames) {
                $parsed | Add-Member -NotePropertyName '_sampling' -NotePropertyValue 'sequential' -Force
                $parsed | Add-Member -NotePropertyName '_segments' -NotePropertyValue @(@{Label="Start";Start=0;FrameCount=$parsed.frames.Count}) -Force
            }
            return $parsed
        }
        return $null
    }

    # Distributed sampling: spread frames across entire file
    # Use more segments for longer files
    $numSegments = if ($Duration -gt 3600) { 10 }        # >1hr: 10 segments
                   elseif ($Duration -gt 1200) { 8 }     # >20min: 8 segments
                   elseif ($Duration -gt 300) { 6 }      # >5min: 6 segments
                   else { 4 }                              # short: 4 segments

    $framesPerSeg = [math]::Floor($MaxFrames / $numSegments)
    # Ensure each segment gets at least 100 frames for meaningful GOP analysis
    $framesPerSeg = [math]::Max($framesPerSeg, 100)

    # Build segment positions spread evenly across duration
    # Include start and near-end, with equal spacing in between
    $segStarts = @()
    for ($i = 0; $i -lt $numSegments; $i++) {
        $pos = [math]::Floor(($Duration - 30) * $i / [math]::Max(1, $numSegments - 1))
        $pos = [math]::Max(0, $pos)
        $segStarts += $pos
    }
    # Make sure last segment doesn't start too late
    if ($segStarts[-1] -gt ($Duration - 15)) {
        $segStarts[-1] = [math]::Max(0, [math]::Floor($Duration - 30))
    }

    $segLabels = @()
    foreach ($s in $segStarts) {
        $pct = [math]::Round($s / $Duration * 100)
        $ts  = [System.TimeSpan]::FromSeconds($s).ToString('hh\:mm\:ss')
        $segLabels += "${pct}% ($ts)"
    }

    $allFrames_list = [System.Collections.ArrayList]::new()
    $segInfo = @()

    for ($si = 0; $si -lt $segStarts.Count; $si++) {
        $segStart = $segStarts[$si]
        $label = $segLabels[$si]

        $r = Run-Command $Tools.FFprobe @(
            '-v','quiet','-print_format','json','-select_streams','v:0','-show_frames',
            '-read_intervals',"$segStart%+#$framesPerSeg",
            '-show_entries','frame=pict_type,key_frame,pkt_size,pkt_pts_time,pkt_duration_time,interlaced_frame,top_field_first,coded_picture_number,display_picture_number,side_data_list',
            "`"$FilePath`""
        ) -TimeoutSeconds 120

        if ($r.ExitCode -eq 0 -and $r.StdOut) {
            try {
                $parsed = $r.StdOut | ConvertFrom-Json
                if ($parsed.frames -and $parsed.frames.Count -gt 0) {
                    $segInfo += @{ Label = $label; Start = $segStart; FrameCount = $parsed.frames.Count }
                    foreach ($f in $parsed.frames) {
                        $allFrames_list.Add($f) | Out-Null
                    }
                }
            } catch {}
        }

        # Don't show inline progress since Run-Command has its own timer
    }
    Write-Host "      Distributed sampling complete: $($segStarts.Count) segments, $($allFrames_list.Count) frames total" -ForegroundColor DarkGray

    if ($script:DebugMode) {
        $iCount = ($allFrames_list | Where-Object { $_.pict_type -eq 'I' }).Count
        $pCount = ($allFrames_list | Where-Object { $_.pict_type -eq 'P' }).Count
        $bCount = ($allFrames_list | Where-Object { $_.pict_type -eq 'B' }).Count
        Write-Host "      [DEBUG] Frame types: I=$iCount P=$pCount B=$bCount" -ForegroundColor Magenta
    }

    if ($allFrames_list.Count -eq 0) { return $null }

    # Build a result object compatible with the sequential format
    $result = [PSCustomObject]@{
        frames = $allFrames_list.ToArray()
    }
    $result | Add-Member -NotePropertyName '_sampling' -NotePropertyValue 'distributed' -Force
    $result | Add-Member -NotePropertyName '_segments' -NotePropertyValue $segInfo -Force

    return $result
}

# ── FFprobe: Multi-point sampling (beginning, 25%, 50%, 75%, end) ──
function Get-MultiPointFrames {
    param([string]$FilePath, [double]$Duration, [int]$SamplesPerSeg = 200)
    if (-not $Tools.FFprobe -or $Duration -lt 60) { return $null }
    $segments = @(
        @{ Start = 0; Label = "Beginning" },
        @{ Start = [math]::Floor($Duration * 0.25); Label = "25%" },
        @{ Start = [math]::Floor($Duration * 0.50); Label = "50%" },
        @{ Start = [math]::Floor($Duration * 0.75); Label = "75%" },
        @{ Start = [math]::Max(0, [math]::Floor($Duration) - 60); Label = "End" }
    )
    $results = @()
    foreach ($seg in $segments) {
        $r = Run-Command $Tools.FFprobe @(
            '-v','quiet','-print_format','json','-select_streams','v:0','-show_frames',
            '-read_intervals',"$($seg.Start)%+#$SamplesPerSeg",
            '-show_entries','frame=pict_type,key_frame,pkt_size,pkt_pts_time',
            "`"$FilePath`""
        ) -TimeoutSeconds 120
        if ($r.ExitCode -eq 0 -and $r.StdOut) {
            try {
                $parsed = $r.StdOut | ConvertFrom-Json
                if ($parsed.frames) { $results += @{ Label = $seg.Label; Start = $seg.Start; Frames = $parsed.frames } }
            } catch {}
        }
    }
    return $results
}

# QP extraction via FFmpeg (H.264 only fallback - HEVC handled by HMAnalyser.ps1)
function Get-QPData {
    param([string]$FilePath, [int]$MaxFrames = 200)
    if (-not $Tools.FFmpeg) { return $null }

    $result = Get-QPViaFFmpeg $FilePath $MaxFrames
    if ($result) { Write-Host "      (QP via ffmpeg debug)" -ForegroundColor Green }
    return $result
}

function Get-QPViaFFmpeg {
    param([string]$FilePath, [int]$MaxFrames = 200)
    if (-not $Tools.FFmpeg) { return $null }

    $r = Run-Command $Tools.FFmpeg @(
        '-debug','qp','-threads','1','-i',"`"$FilePath`"",
        '-vframes',$MaxFrames.ToString(),'-f','null','-an','-'
    ) -TimeoutSeconds 90 -StatusLabel "ffmpeg QP decode"
    if (-not $r.StdErr) { return $null }

    $allQPs   = [System.Collections.ArrayList]::new()
    $frameQPs = [System.Collections.ArrayList]::new()
    $curType  = $null

    foreach ($line in ($r.StdErr -split "`n")) {
        if ($line -match '\[h26[45].*\]\s*(I|P|B)\s') { $curType = $Matches[1] }
        if ($line -match '(?:QP|qp)\s*[:=]\s*(\d+)') {
            $q = [int]$Matches[1]
            [void]$allQPs.Add($q)
            [void]$frameQPs.Add(@{ QP = $q; Type = $curType })
        }
    }
    if ($allQPs.Count -eq 0) { return $null }
    return Build-QPStats $allQPs $frameQPs
}

function Build-QPStats {
    param($allQPs, $frameQPs)
    $vals = $allQPs | ForEach-Object { [double]$_ }
    $mean = ($vals | Measure-Object -Average).Average
    $ssq  = ($vals | ForEach-Object { [math]::Pow($_ - $mean, 2) } | Measure-Object -Sum).Sum

    $stats = @{
        Count  = $vals.Count
        Min    = ($vals | Measure-Object -Minimum).Minimum
        Max    = ($vals | Measure-Object -Maximum).Maximum
        Avg    = [math]::Round($mean, 2)
        Median = ($vals | Sort-Object)[([math]::Floor($vals.Count / 2))]
        StdDev = [math]::Round([math]::Sqrt($ssq / $vals.Count), 2)
    }

    $byType = @{}
    foreach ($t in 'I','P','B') {
        $tv = ($frameQPs | Where-Object { $_.Type -eq $t }).QP
        if ($tv -and $tv.Count -gt 0) {
            $td = $tv | ForEach-Object { [double]$_ }
            $byType[$t] = @{
                Count = $td.Count
                Min   = ($td | Measure-Object -Minimum).Minimum
                Max   = ($td | Measure-Object -Maximum).Maximum
                Avg   = [math]::Round(($td | Measure-Object -Average).Average, 2)
            }
        }
    }
    return @{ Stats = $stats; ByType = $byType }
}

# ── FFprobe: NAL unit / SPS / PPS bitstream parameter extraction ──
function Get-NALData {
    param([string]$FilePath)
    if (-not $Tools.FFprobe) { return $null }
    $r = Run-Command $Tools.FFprobe @(
        '-v','trace','-read_intervals','0%+#50',
        '-select_streams','v:0','-show_packets',
        "`"$FilePath`""
    ) -TimeoutSeconds 120 -StatusLabel "ffprobe NAL trace"
    if (-not $r.StdErr) { return $null }

    $d = @{ NALTypes = @{}; SEI = [System.Collections.ArrayList]::new() }

    foreach ($line in ($r.StdErr -split "`n")) {
        # NAL unit types
        if ($line -match 'nal_unit_type\s*[:=]\s*(\d+)\s*\(([^)]+)\)') {
            $k = $Matches[2].Trim(); $d.NALTypes[$k] = ($d.NALTypes[$k] + 0) + 1
        } elseif ($line -match 'nal_unit_type\s*[:=]\s*(\d+)') {
            $k = "NAL_$($Matches[1])"; $d.NALTypes[$k] = ($d.NALTypes[$k] + 0) + 1
        }
        # SEI
        if ($line -match 'SEI\s+type\s*[:=]?\s*(\d+)') { [void]$d.SEI.Add([int]$Matches[1]) }

        # ── SPS fields ──
        if ($line -match 'general_profile_idc\s*[:=]\s*(\d+)')                  { $d.ProfileIDC = [int]$Matches[1] }
        if ($line -match 'general_level_idc\s*[:=]\s*(\d+)')                    { $d.LevelIDC = [int]$Matches[1] }
        if ($line -match 'general_tier_flag\s*[:=]\s*(\d+)')                    { $d.Tier = [int]$Matches[1] }
        if ($line -match 'chroma_format_idc\s*[:=]\s*(\d+)')                    { $d.Chroma = [int]$Matches[1] }
        if ($line -match 'bit_depth_luma_minus8\s*[:=]\s*(\d+)')                { $d.BitDepthLuma = [int]$Matches[1] + 8 }
        if ($line -match 'bit_depth_chroma_minus8\s*[:=]\s*(\d+)')              { $d.BitDepthChroma = [int]$Matches[1] + 8 }
        if ($line -match 'log2_min_luma_coding_block_size_minus3\s*[:=]\s*(\d+)') { $d.MinCU = [math]::Pow(2, [int]$Matches[1] + 3) }
        if ($line -match 'log2_diff_max_min_luma_coding_block_size\s*[:=]\s*(\d+)') { $d.CUDepth = [int]$Matches[1] }
        if ($line -match 'log2_min_luma_transform_block_size_minus2\s*[:=]\s*(\d+)') { $d.MinTU = [math]::Pow(2, [int]$Matches[1] + 2) }
        if ($line -match 'log2_diff_max_min_luma_transform_block_size\s*[:=]\s*(\d+)') { $d.TUDepth = [int]$Matches[1] }
        if ($line -match 'max_transform_hierarchy_depth_inter\s*[:=]\s*(\d+)')  { $d.TUDepthInter = [int]$Matches[1] }
        if ($line -match 'max_transform_hierarchy_depth_intra\s*[:=]\s*(\d+)')  { $d.TUDepthIntra = [int]$Matches[1] }
        if ($line -match 'num_short_term_ref_pic_sets\s*[:=]\s*(\d+)')          { $d.STRPS = [int]$Matches[1] }
        if ($line -match 'long_term_ref_pics_present_flag\s*[:=]\s*(\d+)')      { $d.LongTermRef = [int]$Matches[1] }
        if ($line -match 'sps_temporal_mvp_enabled_flag\s*[:=]\s*(\d+)')        { $d.TemporalMVP = [int]$Matches[1] }
        if ($line -match 'strong_intra_smoothing_enabled_flag\s*[:=]\s*(\d+)')  { $d.StrongIntra = [int]$Matches[1] }
        if ($line -match 'sample_adaptive_offset_enabled_flag\s*[:=]\s*(\d+)')  { $d.SAO = [int]$Matches[1] }
        if ($line -match 'amp_enabled_flag\s*[:=]\s*(\d+)')                     { $d.AMP = [int]$Matches[1] }
        if ($line -match 'pcm_enabled_flag\s*[:=]\s*(\d+)')                     { $d.PCM = [int]$Matches[1] }
        if ($line -match 'scaling_list_enabled_flag\s*[:=]\s*(\d+)')            { $d.ScalingList = [int]$Matches[1] }
        if ($line -match 'sps_max_dec_pic_buffering_minus1\s*[:=]\s*(\d+)')     { $d.DPBSize = [int]$Matches[1] + 1 }

        # ── PPS fields ──
        if ($line -match 'transform_skip_enabled_flag\s*[:=]\s*(\d+)')          { $d.TransformSkip = [int]$Matches[1] }
        if ($line -match 'cu_qp_delta_enabled_flag\s*[:=]\s*(\d+)')             { $d.CUQPDelta = [int]$Matches[1] }
        if ($line -match 'diff_cu_qp_delta_depth\s*[:=]\s*(\d+)')               { $d.CUQPDeltaDepth = [int]$Matches[1] }
        if ($line -match 'weighted_pred_flag\s*[:=]\s*(\d+)')                   { $d.WeightedPred = [int]$Matches[1] }
        if ($line -match 'weighted_bipred_flag\s*[:=]\s*(\d+)')                 { $d.WeightedBipred = [int]$Matches[1] }
        if ($line -match 'constrained_intra_pred_flag\s*[:=]\s*(\d+)')          { $d.ConstrainedIntra = [int]$Matches[1] }
        if ($line -match 'sign_data_hiding_enabled_flag\s*[:=]\s*(\d+)')        { $d.SignDataHiding = [int]$Matches[1] }
        if ($line -match 'entropy_coding_sync_enabled_flag\s*[:=]\s*(\d+)')     { $d.WPP = [int]$Matches[1] }
        if ($line -match 'tiles_enabled_flag\s*[:=]\s*(\d+)')                   { $d.Tiles = [int]$Matches[1] }
        if ($line -match 'num_tile_columns_minus1\s*[:=]\s*(\d+)')              { $d.TileCols = [int]$Matches[1] + 1 }
        if ($line -match 'num_tile_rows_minus1\s*[:=]\s*(\d+)')                 { $d.TileRows = [int]$Matches[1] + 1 }
        if ($line -match 'deblocking_filter_disabled_flag\s*[:=]\s*(\d+)')      { $d.DeblockDisabled = [int]$Matches[1] }
        if ($line -match 'pps_beta_offset_div2\s*[:=]\s*(-?\d+)')               { $d.DeblockBeta = [int]$Matches[1] * 2 }
        if ($line -match 'pps_tc_offset_div2\s*[:=]\s*(-?\d+)')                 { $d.DeblockTC = [int]$Matches[1] * 2 }
        if ($line -match 'cabac_init_present_flag\s*[:=]\s*(\d+)')              { $d.CABACInit = [int]$Matches[1] }
        if ($line -match 'num_ref_idx_l0_default_active_minus1\s*[:=]\s*(\d+)') { $d.RefL0 = [int]$Matches[1] + 1 }
        if ($line -match 'num_ref_idx_l1_default_active_minus1\s*[:=]\s*(\d+)') { $d.RefL1 = [int]$Matches[1] + 1 }

        # ── H.264 specific ──
        if ($line -match 'num_ref_frames\s*[:=]\s*(\d+)')                       { $d.H264RefFrames = [int]$Matches[1] }
        if ($line -match 'direct_spatial_mv_pred_flag\s*[:=]\s*(\d+)')          { $d.H264DirectSpatial = [int]$Matches[1] }
    }
    return $d
}

# ── FFmpeg: Encoder string detection ──
function Get-EncoderInfo {
    param([string]$FilePath)
    if (-not $Tools.FFmpeg) { return $null }
    $r = Run-Command $Tools.FFmpeg @('-i', "`"$FilePath`"", '-f','null','-t','0','-') -TimeoutSeconds 30
    if ($script:DebugMode) {
        $encoderLines = ($r.StdErr -split "`n") | Where-Object { $_ -match 'encoder\s*:' }
        Write-Host "      [DEBUG] Encoder tags found: $($encoderLines.Count)" -ForegroundColor Magenta
        foreach ($el in $encoderLines) { Write-Host "      [DEBUG]   $($el.Trim())" -ForegroundColor Magenta }
    }
    return $r.StdErr
}

# ── MediaInfo: Full JSON + text ──
function Get-MediaInfoData {
    param([string]$FilePath)
    if (-not $Tools.MediaInfo) { return $null }
    $rText = Run-Command $Tools.MediaInfo @('-f', "`"$FilePath`"")
    $rJson = Run-Command $Tools.MediaInfo @('--Output=JSON', "`"$FilePath`"")
    $json = $null
    if ($rJson.ExitCode -eq 0 -and $rJson.StdOut) {
        try { $json = $rJson.StdOut | ConvertFrom-Json } catch {}
    }
    if ($script:DebugMode) {
        $tracks = if ($json -and $json.media -and $json.media.track) { $json.media.track } else { @() }
        Write-Host "      [DEBUG] MediaInfo JSON tracks: $($tracks.Count)" -ForegroundColor Magenta
        foreach ($tk in $tracks) {
            $hdr = if ($tk.HDR_Format) { " HDR=$($tk.HDR_Format)" } else { '' }
            Write-Host "      [DEBUG]   @type=$($tk.'@type') Format=$($tk.Format)$hdr" -ForegroundColor Magenta
        }
        Write-Host "      [DEBUG] MediaInfo text: $($rText.StdOut.Length) chars" -ForegroundColor Magenta
    }
    return @{
        Text = if ($rText.ExitCode -eq 0) { $rText.StdOut } else { $null }
        Json = $json
    }
}

# ── CheckBitrate: Bitrate distribution with statistics ──
# CheckBitrate.exe writes CSV files to same dir as input: <input>.track<N>.bitrate.csv
function Get-CheckBitrateData {
    param([string]$FilePath, [double]$Interval = 0)
    if (-not $Tools.CheckBitrate -or $script:SkipCB) {
        if ($script:DebugMode) { Write-Host "      [DEBUG] CheckBitrate skipped: tool=$($Tools.CheckBitrate) skipCB=$($script:SkipCB)" -ForegroundColor Magenta }
        return $null
    }
    $cbArgs = @()
    if ($Interval -gt 0) { $cbArgs += @('-i', $Interval.ToString()) }
    $cbArgs += "`"$FilePath`""

    $r = Run-Command $Tools.CheckBitrate $cbArgs -TimeoutSeconds 600 -StatusLabel "CheckBitrate"
    if ($script:DebugMode) { Write-Host "      [DEBUG] CheckBitrate exit=$($r.ExitCode) stderr='$($r.StdErr.Substring(0, [math]::Min(200, $r.StdErr.Length)))'" -ForegroundColor Magenta }

    # CheckBitrate outputs CSV files: <filepath>.track<N>.bitrate.csv
    # Find the video track CSV (track1 for most files)
    $csvPattern = "$FilePath.track*.bitrate.csv"
    $csvFiles = Get-ChildItem -Path (Split-Path $FilePath -Parent) -Filter "$(Split-Path $FilePath -Leaf).track*.bitrate.csv" -ErrorAction SilentlyContinue
    if ($script:DebugMode) { Write-Host "      [DEBUG] CheckBitrate CSV files found: $($csvFiles.Count) ($($csvFiles.Name -join ', '))" -ForegroundColor Magenta }

    if (-not $csvFiles -or $csvFiles.Count -eq 0) { return $null }

    # Pick the track1 CSV (main video), or the largest CSV if track1 not found
    $csvFile = $csvFiles | Where-Object { $_.Name -match 'track1' } | Select-Object -First 1
    if (-not $csvFile) { $csvFile = $csvFiles | Sort-Object Length -Descending | Select-Object -First 1 }

    $csvContent = Get-Content $csvFile.FullName -ErrorAction SilentlyContinue
    if (-not $csvContent -or $csvContent.Count -lt 2) { return $null }

    $data = @()
    foreach ($line in $csvContent) {
        $trimmed = $line.Trim()
        # Skip header line (starts with comma or contains 'kbps')
        if ($trimmed -match '^,|kbps') { continue }
        # Data lines: "     0.000,3050.88,3050.88"
        $parts = $trimmed -split ','
        if ($parts.Count -ge 3) {
            $time = $parts[0].Trim()
            $kbps = $parts[1].Trim()
            $avg  = $parts[2].Trim()
            # Skip inf values (cover art track)
            if ($kbps -eq 'inf' -or $avg -eq 'inf') { continue }
            try {
                $data += [PSCustomObject]@{
                    TimeSec     = [double]$time
                    BitrateKbps = [double]$kbps
                    AvgKbps     = [double]$avg
                }
            } catch { continue }
        }
    }
    if ($data.Count -eq 0) { return $null }

    # Clean up CSV files
    foreach ($f in $csvFiles) { Remove-Item $f.FullName -Force -ErrorAction SilentlyContinue }

    $bv   = $data | ForEach-Object { $_.BitrateKbps }
    $mean = ($bv | Measure-Object -Average).Average
    $ssq  = ($bv | ForEach-Object { [math]::Pow($_ - $mean, 2) } | Measure-Object -Sum).Sum
    $sorted = $bv | Sort-Object

    $stats = @{
        DataPoints = $data.Count
        Duration   = $data[-1].TimeSec
        FinalAvg   = $data[-1].AvgKbps
        Peak       = ($bv | Measure-Object -Maximum).Maximum
        Min        = ($bv | Measure-Object -Minimum).Minimum
        Median     = $sorted[([math]::Floor($sorted.Count / 2))]
        StdDev     = [math]::Sqrt($ssq / $bv.Count)
        P05        = $sorted[([math]::Floor($sorted.Count * 0.05))]
        P95        = $sorted[([math]::Floor($sorted.Count * 0.95))]
        P99        = $sorted[([math]::Floor($sorted.Count * 0.99))]
        Data       = $data
    }

    # Peak capping detection (VBV/maxrate in effect?)
    $top1pct = $sorted | Select-Object -Last ([math]::Max(1, [math]::Floor($sorted.Count * 0.01)))
    $topMean = ($top1pct | Measure-Object -Average).Average
    $topStd  = 0
    if ($top1pct.Count -gt 1) {
        $ts = ($top1pct | ForEach-Object { [math]::Pow($_ - $topMean, 2) } | Measure-Object -Sum).Sum
        $topStd = [math]::Sqrt($ts / $top1pct.Count)
    }
    $stats.PeaksCapped     = ($topStd / [math]::Max($topMean, 1) * 100) -lt 5
    $stats.PeakClusterKbps = $topMean

    return $stats
}

# ── FFmpeg: Scene cut detection ──
function Get-SceneCuts {
    param([string]$FilePath, [int]$MaxSeconds = 300)
    if (-not $Tools.FFmpeg) { return $null }
    $r = Run-Command $Tools.FFmpeg @(
        '-i', "`"$FilePath`"", '-t', $MaxSeconds.ToString(),
        '-vf', "select='gt(scene,0.3)',metadata=print:file=-",
        '-an', '-f', 'null', '-'
    ) -TimeoutSeconds 300 -StatusLabel "ffmpeg scene detect"

    $cuts = @()
    if ($r.StdOut) {
        foreach ($line in ($r.StdOut -split "`n")) {
            if ($line -match 'pts_time:(\d+\.?\d*)') { $cuts += [double]$Matches[1] }
        }
    }
    if ($script:DebugMode) {
        Write-Host "      [DEBUG] Scene detect: exit=$($r.ExitCode) threshold=0.3 duration=${MaxSeconds}s cuts=$($cuts.Count)" -ForegroundColor Magenta
        if ($cuts.Count -gt 0 -and $cuts.Count -le 10) {
            Write-Host "      [DEBUG] Cut times: $(($cuts | ForEach-Object { '{0:F1}s' -f $_ }) -join ', ')" -ForegroundColor Magenta
        }
    }
    return @{
        Cuts        = $cuts
        Count       = $cuts.Count
        Duration    = $MaxSeconds
        AvgInterval = if ($cuts.Count -gt 1) { [math]::Round($MaxSeconds / $cuts.Count, 2) } else { $MaxSeconds }
    }
}
