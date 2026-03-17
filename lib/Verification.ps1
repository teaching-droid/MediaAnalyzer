# ─────────────────────────────────────────────────────────────────────────────
# Verification.ps1 — Quick verification encodes using detected GPU encoders
# Encodes the same segments that were analyzed, then compares output vs source.
# ─────────────────────────────────────────────────────────────────────────────

function Get-VerificationSegments {
    <#
    .SYNOPSIS
        Build segment extraction parameters for verification encodes.
        Quick mode: short clips from beginning/middle/end (command validation).
        Extended mode: single continuous segment for bitrate convergence testing.
    #>
    param(
        [array]$SegmentInfo,        # from frameData._segments
        [double]$FileDuration,      # total file duration in seconds
        [double]$FPS = 24,          # frames per second
        [int]$MaxSegments = 3,      # limit segments for quick mode
        [int]$ExtendedDuration = 0  # 0=quick mode, >0=continuous segment of N seconds
    )

    if (-not $SegmentInfo -or $SegmentInfo.Count -eq 0) { return @() }

    # ── Extended mode: single continuous segment from middle of file ──
    if ($ExtendedDuration -gt 0) {
        $dur = [math]::Min($ExtendedDuration, [math]::Floor($FileDuration * 0.9))
        # Start from ~25% into the file (past opening credits, representative content)
        $startSec = [math]::Floor($FileDuration * 0.25)
        # Make sure we don't overshoot
        if (($startSec + $dur) -gt ($FileDuration - 5)) {
            $startSec = [math]::Max(0, [math]::Floor($FileDuration - $dur - 5))
        }
        $estFrames = [math]::Round($dur * $FPS)
        $pctPos = [math]::Round($startSec / $FileDuration * 100)
        $ts = [TimeSpan]::FromSeconds($startSec).ToString('hh\:mm\:ss')
        return @(@{
            Index    = 0
            Label    = "Extended ${dur}s from ${pctPos}% ($ts)"
            StartSec = $startSec
            Duration = $dur
            Frames   = $estFrames
        })
    }

    # ── Quick mode: short clips from beginning, middle, end ──
    $indices = @()
    if ($SegmentInfo.Count -le $MaxSegments) {
        $indices = 0..($SegmentInfo.Count - 1)
    } else {
        $indices += 0                                                        # first
        $indices += [math]::Floor($SegmentInfo.Count / 2)                    # middle
        $indices += ($SegmentInfo.Count - 1)                                 # last
    }

    $segments = @()
    foreach ($i in $indices) {
        $seg = $SegmentInfo[$i]
        $startSec = [double]$seg.Start
        # Duration = frames / fps + small buffer for GOP completion
        $durSec = [math]::Ceiling([int]$seg.FrameCount / $FPS) + 2
        # Clamp to file bounds
        if (($startSec + $durSec) -gt $FileDuration) {
            $durSec = [math]::Max(2, [math]::Floor($FileDuration - $startSec))
        }
        $segments += @{
            Index    = $i
            Label    = $seg.Label
            StartSec = $startSec
            Duration = $durSec
            Frames   = $seg.FrameCount
        }
    }
    return $segments
}

function Invoke-VerificationEncode {
    <#
    .SYNOPSIS
        Run verification encodes with available GPU encoders on extracted segments.
    .DESCRIPTION
        1. Extracts short segments from source via FFmpeg
        2. Encodes each segment with NVEncC and/or QSVEncC
        3. Analyzes encoded output with FFprobe
        4. Compares against source parameters
    .RETURNS
        Hashtable with per-encoder results: @{NVEnc=@{...}; QSVEnc=@{...}}
    #>
    param(
        [string]$SourcePath,
        [hashtable]$rc,
        [array]$Segments,           # from Get-VerificationSegments
        [string]$WorkDir,           # temp directory for encode outputs
        [hashtable]$Tools,          # tool paths
        [hashtable]$GPUCaps         # GPU capabilities
    )

    $results = @{}

    if (-not $Segments -or $Segments.Count -eq 0) {
        Write-Host "      No segments available for verification" -ForegroundColor Yellow
        return $results
    }

    # Create work directory
    $verifyDir = Join-Path $WorkDir "verify_encode"
    if (-not (Test-Path $verifyDir)) { New-Item -ItemType Directory -Path $verifyDir -Force | Out-Null }

    # ── Step 1: Extract segments from source ──
    Write-Host "      Extracting $($Segments.Count) segments for verification..." -ForegroundColor DarkGray
    $extractedSegments = @()

    foreach ($seg in $Segments) {
        $segFile = Join-Path $verifyDir "seg_$($seg.Index).mkv"
        # Use stream copy for speed — no re-encoding, just extract
        $ffArgs = @(
            '-y', '-hide_banner', '-loglevel', 'error',
            '-ss', "$($seg.StartSec)",
            '-i', "`"$SourcePath`"",
            '-t', "$($seg.Duration)",
            '-map', '0:v:0',
            '-c:v', 'copy',
            '-an', '-sn',
            "`"$segFile`""
        )
        $r = Run-Command $Tools.FFmpeg $ffArgs -TimeoutSeconds ([math]::Max(120, [int]($seg.Duration / 2)))
        if ($r.ExitCode -eq 0 -and (Test-Path $segFile)) {
            $finfo = Get-Item $segFile
            if ($finfo.Length -gt 1000) {
                $extractedSegments += @{ Path = $segFile; Seg = $seg }
            } else {
                Write-Host "        Segment $($seg.Index) too small ($($finfo.Length) bytes), skipping" -ForegroundColor Yellow
            }
        } else {
            Write-Host "        Failed to extract segment $($seg.Index): $($r.StdErr)" -ForegroundColor Red
        }
    }

    if ($extractedSegments.Count -eq 0) {
        Write-Host "      No segments extracted successfully" -ForegroundColor Red
        return $results
    }

    # Use first extracted segment for verification (shortest encode time)
    # If it passes, the commands are valid. Use middle segment for most representative content.
    $useIdx = if ($extractedSegments.Count -ge 2) { 1 } else { 0 }  # prefer middle
    $testSeg = $extractedSegments[$useIdx]
    $testInput = $testSeg.Path

    # Scale encode timeout based on segment duration (4K HEVC ~2-10fps on GPU)
    $segDurSec = $testSeg.Seg.Duration
    $encTimeout = [math]::Max(300, [int]($segDurSec * 15))  # ~15x realtime worst case

    # ── Step 2: Build encoder commands ──
    $fps = if ($rc['FPS']) { $rc['FPS'] } else { 24 }
    $isHEVC = $rc['Codec'] -match 'hevc|h265'
    $isHDR = $rc['ColorTransfer'] -match 'smpte2084|arib-std-b67'
    $is10b = $rc['BitDepth'] -eq '10' -or $rc['BitDepth'] -eq 10 -or $rc['PixFmt'] -match '10'

    # Common verification parameters (use lower bitrate for speed, same structure)
    $maxrateK = if ($rc['MaxrateLikely']) { $rc['MaxrateLikely'] }
                elseif ($rc['SuggestedMaxrate']) { [math]::Round($rc['SuggestedMaxrate']) }
                else { $null }
    $avgBitK = if ($rc['BitrateAvgKbps']) { [math]::Round($rc['BitrateAvgKbps']) } else { $null }

    # NVEnc bframes (GPU-adjusted)
    $nvBFrames = if ($rc['NVEnc_MaxBFrames']) { [math]::Min([int]$rc['BFrames'], [int]$rc['NVEnc_MaxBFrames']) }
                 elseif ($rc['BFrames']) { [int]$rc['BFrames'] }
                 else { 0 }

    # ── NVEncC verification ──
    if ($Tools.NVEncC -and $GPUCaps.NVEnc) {
        Write-Host "      Encoding verification segment with NVEncC..." -ForegroundColor DarkGray
        $nvOut = Join-Path $verifyDir "verify_nvenc.mkv"
        $nvArgs = @(
            '--avhw',
            '-i', "`"$testInput`"",
            '-o', "`"$nvOut`"",
            '--codec', 'hevc',
            '--output-depth', $(if($is10b){'10'}else{'8'})
        )
        if ($rc['Profile']) {
            $hwProf = switch -Regex ("$($rc['Profile'])".Trim()) {
                '(?i)Main\s*10\s*444' {'main444_10'; break} '(?i)Main\s*444' {'main444'; break}
                '(?i)Main\s*10' {'main10'; break} '(?i)^Main$' {'main'; break}
                default {$rc['Profile'].ToLower() -replace '\s+','' -replace '"',''}
            }
            $nvArgs += '--profile'; $nvArgs += $hwProf
        }
        $nvArgs += '--preset'; $nvArgs += 'P6'
        $nvArgs += '--tune'; $nvArgs += 'uhq'
        if ($rc['Keyint']) { $nvArgs += '--gop-len'; $nvArgs += "$($rc['Keyint'])" }
        if ($nvBFrames -gt 0) {
            $nvArgs += '--bframes'; $nvArgs += "$nvBFrames"
            if ($nvBFrames -ge 3) { $nvArgs += '--bref-mode'; $nvArgs += 'middle' }
        }
        # Ref: use HM-derived count, cap to NVEnc max (6), minimum 2 with B-pyramid
        $nvRef = if ($rc['Refs']) { [int]$rc['Refs'] } else { 1 }
        if ($nvBFrames -ge 3 -and $nvRef -lt 2) { $nvRef = 2 }
        $nvRef = [math]::Min($nvRef, 6)
        $nvArgs += '--ref'; $nvArgs += "$nvRef"
        # NVEnc: weightp unsupported with B-frames
        if ($rc['WeightedPred'] -and $nvBFrames -eq 0) { $nvArgs += '--weightp' }

        # Rate control
        if ($avgBitK) {
            $nvArgs += '--vbr'; $nvArgs += "$avgBitK"
            $nvArgs += '--multipass'; $nvArgs += '2pass-full'
            if ($maxrateK) { $nvArgs += '--max-bitrate'; $nvArgs += "$maxrateK"; $nvArgs += '--vbv-bufsize'; $nvArgs += "$($maxrateK * 2)" }
        } else {
            $nvArgs += '--qvbr'; $nvArgs += '0'; $nvArgs += '--vbr-quality'; $nvArgs += '20'
        }
        $nvArgs += '--lookahead'; $nvArgs += '32'

        # QP bounds: omitted — let VBR allocate freely on verification segments

        # AQ
        if ($rc['CUQPDelta']) { $nvArgs += '--aq'; $nvArgs += '--aq-temporal'; $nvArgs += '--aq-strength'; $nvArgs += '0' }
        # Temporal filter
        if ($nvBFrames -ge 4) { $nvArgs += '--tf-level'; $nvArgs += '4' }

        # Color
        $nvArgs += '--colormatrix'; $nvArgs += 'auto'
        $nvArgs += '--transfer'; $nvArgs += 'auto'
        $nvArgs += '--colorprim'; $nvArgs += 'auto'
        $nvArgs += '--colorrange'; $nvArgs += 'auto'

        # Compliance
        $nvArgs += '--aud'; $nvArgs += '--repeat-headers'; $nvArgs += '--pic-struct'

        # HDR metadata
        if ($isHDR) {
            $nvArgs += '--max-cll'; $nvArgs += 'copy'
            $nvArgs += '--master-display'; $nvArgs += 'copy'
            $nvArgs += '--dhdr10-info'; $nvArgs += 'copy'
        }
        # DV
        if ($rc['DV_Profile']) {
            $dvP = [int]$rc['DV_Profile']; $dvC = $rc['DV_Compat']; $dvFull = "$dvP.$dvC"
            $supportedDV = @('5.0','8.1','8.2','8.4','10.0','10.1','10.2','10.4')
            $dvStr = if ($dvFull -in $supportedDV) { $dvFull } else { 'copy' }
            $nvArgs += '--dolby-vision-rpu'; $nvArgs += 'copy'
            $nvArgs += '--dolby-vision-profile'; $nvArgs += $dvStr
        }

        $encTimer = [System.Diagnostics.Stopwatch]::StartNew()
        $r = Run-Command $Tools.NVEncC $nvArgs -TimeoutSeconds $encTimeout
        $encTimer.Stop()

        $nvResult = @{
            Success    = ($r.ExitCode -eq 0 -and (Test-Path $nvOut))
            ExitCode   = $r.ExitCode
            Duration   = $encTimer.Elapsed.TotalSeconds
            OutputFile = $nvOut
            StdErr     = $r.StdErr
            Checks     = @()
        }

        if ($nvResult.Success) {
            $nvResult.Checks = Compare-EncodedOutput -OutputFile $nvOut -rc $rc -Tools $Tools -EncoderName 'NVEncC' -SourceSegment $testInput
            $nvResult.OutputSize = (Get-Item $nvOut).Length
        } else {
            $nvResult.Checks += @{ Name = 'Encode'; Status = 'FAIL'; Detail = "Exit code $($r.ExitCode): $($r.StdErr)" }
        }
        $results['NVEnc'] = $nvResult
    }

    # ── QSVEncC verification ──
    if ($Tools.QSVEncC -and $GPUCaps.QSVEnc) {
        Write-Host "      Encoding verification segment with QSVEncC..." -ForegroundColor DarkGray
        $qsvOut = Join-Path $verifyDir "verify_qsvenc.mkv"
        $qsvArgs = @(
            '--avhw',
            '-i', "`"$testInput`"",
            '-o', "`"$qsvOut`"",
            '--codec', 'hevc',
            '--output-depth', $(if($is10b){'10'}else{'8'})
        )
        if ($rc['Profile']) {
            $hwProf = switch -Regex ("$($rc['Profile'])".Trim()) {
                '(?i)Main\s*10\s*444' {'main444_10'; break} '(?i)Main\s*444' {'main444'; break}
                '(?i)Main\s*10' {'main10'; break} '(?i)^Main$' {'main'; break}
                default {$rc['Profile'].ToLower() -replace '\s+','' -replace '"',''}
            }
            $qsvArgs += '--profile'; $qsvArgs += $hwProf
        }
        $qsvArgs += '--quality'; $qsvArgs += 'best'
        if ($rc['Keyint']) { $qsvArgs += '--gop-len'; $qsvArgs += "$($rc['Keyint'])" }
        if ($rc['BFrames']) {
            $qsvArgs += '--bframes'; $qsvArgs += "$($rc['BFrames'])"
            $qsvArgs += '--b-pyramid'; $qsvArgs += '--weightb'
        }
        # Ref: use HM-derived count, cap to QSVEnc max (16), minimum 2 with B-pyramid
        $qsvRef = if ($rc['Refs']) { [int]$rc['Refs'] } else { 1 }
        if ($rc['BFrames'] -and [int]$rc['BFrames'] -ge 3 -and $qsvRef -lt 2) { $qsvRef = 2 }
        $qsvRef = [math]::Min($qsvRef, 16)
        $qsvArgs += '--ref'; $qsvArgs += "$qsvRef"
        if ($rc['WeightedPred']) { $qsvArgs += '--weightp' }

        # Rate control (use GPU-validated fallback)
        if ($avgBitK) {
            if ($rc['QSVEnc_LAFallback']) {
                # Parse fallback string (e.g. "--vbr 66150")
                $qsvArgs += ($rc['QSVEnc_LAFallback'] -split '\s+')
            } else {
                $qsvArgs += '--la'; $qsvArgs += "$avgBitK"; $qsvArgs += '--la-depth'; $qsvArgs += '40'
            }
            if ($maxrateK) { $qsvArgs += '--max-bitrate'; $qsvArgs += "$maxrateK"; $qsvArgs += '--vbv-bufsize'; $qsvArgs += "$($maxrateK * 2)" }
        } else {
            if ($rc['QSVEnc_RCFallback']) {
                $qsvArgs += ($rc['QSVEnc_RCFallback'] -split '\s+')
            } else {
                $qsvArgs += '--la-icq'; $qsvArgs += '23'; $qsvArgs += '--la-depth'; $qsvArgs += '40'
            }
        }

        # SAO
        if ($rc['SAO']) { $qsvArgs += '--sao'; $qsvArgs += 'all' } else { $qsvArgs += '--sao'; $qsvArgs += 'none' }

        # QP bounds: omitted — let VBR allocate freely on verification segments

        # Advanced features (conditionally exclude unsupported)
        $qsvRm = if ($rc['QSVEnc_Remove']) { $rc['QSVEnc_Remove'] } else { @() }
        if ('extbrc'      -notin $qsvRm) { $qsvArgs += '--extbrc' }
        if ('mbbrc'       -notin $qsvRm) { $qsvArgs += '--mbbrc' }
        if ('adapt-ref'   -notin $qsvRm) { $qsvArgs += '--adapt-ref' }
        if ('adapt-ltr'   -notin $qsvRm) { $qsvArgs += '--adapt-ltr' }
        if ('adapt-cqm'   -notin $qsvRm) { $qsvArgs += '--adapt-cqm' }
        if ('fade-detect' -notin $qsvRm) { $qsvArgs += '--fade-detect' }

        # Color
        $qsvArgs += '--colormatrix'; $qsvArgs += 'auto'
        $qsvArgs += '--transfer'; $qsvArgs += 'auto'
        $qsvArgs += '--colorprim'; $qsvArgs += 'auto'
        $qsvArgs += '--colorrange'; $qsvArgs += 'auto'

        # Compliance
        $qsvArgs += '--aud'; $qsvArgs += '--repeat-headers'; $qsvArgs += '--pic-struct'; $qsvArgs += '--buf-period'

        # HDR metadata
        if ($isHDR) {
            $qsvArgs += '--max-cll'; $qsvArgs += 'copy'
            $qsvArgs += '--master-display'; $qsvArgs += 'copy'
            $qsvArgs += '--dhdr10-info'; $qsvArgs += 'copy'
        }
        # DV
        if ($rc['DV_Profile']) {
            $dvP = [int]$rc['DV_Profile']; $dvC = $rc['DV_Compat']; $dvFull = "$dvP.$dvC"
            $supportedDV = @('5.0','8.1','8.2','8.4','10.0','10.1','10.2','10.4')
            $dvStr = if ($dvFull -in $supportedDV) { $dvFull } else { 'copy' }
            $qsvArgs += '--dolby-vision-rpu'; $qsvArgs += 'copy'
            $qsvArgs += '--dolby-vision-profile'; $qsvArgs += $dvStr
        }

        $encTimer = [System.Diagnostics.Stopwatch]::StartNew()
        $r = Run-Command $Tools.QSVEncC $qsvArgs -TimeoutSeconds $encTimeout
        $encTimer.Stop()

        $qsvResult = @{
            Success    = ($r.ExitCode -eq 0 -and (Test-Path $qsvOut))
            ExitCode   = $r.ExitCode
            Duration   = $encTimer.Elapsed.TotalSeconds
            OutputFile = $qsvOut
            StdErr     = $r.StdErr
            Checks     = @()
        }

        if ($qsvResult.Success) {
            $qsvResult.Checks = Compare-EncodedOutput -OutputFile $qsvOut -rc $rc -Tools $Tools -EncoderName 'QSVEncC' -SourceSegment $testInput
            $qsvResult.OutputSize = (Get-Item $qsvOut).Length
        } else {
            $qsvResult.Checks += @{ Name = 'Encode'; Status = 'FAIL'; Detail = "Exit code $($r.ExitCode): $($r.StdErr)" }
        }
        $results['QSVEnc'] = $qsvResult
    }

    # ── Cleanup extracted segments (keep encoded outputs for inspection) ──
    foreach ($es in $extractedSegments) {
        if (Test-Path $es.Path) { Remove-Item $es.Path -Force -ErrorAction SilentlyContinue }
    }

    return $results
}


function Compare-EncodedOutput {
    <#
    .SYNOPSIS
        Run full analysis pipeline on encoded output and compare against source.
        Runs the same checks as the main analysis: FFprobe, Frame/GOP, HM, MediaInfo, CheckBitrate, Quality.
    .RETURNS
        Array of @{Name; Status; Detail} check results (PASS/WARN/FAIL/INFO).
    #>
    param(
        [string]$OutputFile,
        [hashtable]$rc,
        [hashtable]$Tools,
        [string]$EncoderName,
        [string]$SourceSegment = ""   # Path to source segment MKV for quality comparison
    )

    $checks = @()

    # ══════════════════════════════════════════════════════════════
    # Step 1: FFprobe stream analysis
    # ══════════════════════════════════════════════════════════════
    Write-Host "        [1/6] FFprobe stream analysis..." -ForegroundColor DarkGray
    if (-not $Tools.FFprobe) {
        $checks += @{ Name = 'FFprobe'; Status = 'SKIP'; Detail = 'FFprobe not available' }
        return $checks
    }

    $probe = Get-ProbeJson $OutputFile
    if (-not $probe) {
        if ($script:DebugMode) { Write-Host "      [DEBUG] Verify FFprobe: probe returned null for $OutputFile" -ForegroundColor Magenta }
        $checks += @{ Name = 'Probe'; Status = 'FAIL'; Detail = "Cannot probe encoded file" }
        return $checks
    }
    $vs = $probe.streams | Where-Object { $_.codec_type -eq 'video' } | Select-Object -First 1
    if (-not $vs) {
        $checks += @{ Name = 'Video Stream'; Status = 'FAIL'; Detail = 'No video stream' }
        return $checks
    }
    $checks += @{ Name = 'Encode'; Status = 'PASS'; Detail = "Encoded successfully ($EncoderName)" }

    # Basic stream checks
    $codecOK = $vs.codec_name -match 'hevc|h265'
    $checks += @{ Name = 'Codec'; Status = if($codecOK){'PASS'}else{'FAIL'}; Detail = "Output: $($vs.codec_name)" }

    $resOK = "$($vs.width)x$($vs.height)" -eq $rc['Resolution']
    $checks += @{ Name = 'Resolution'; Status = if($resOK){'PASS'}else{'FAIL'}
        Detail = "Output: $($vs.width)x$($vs.height)$(if(-not $resOK){" (expected $($rc['Resolution']))"})" }

    $outDepth = if($vs.pix_fmt -match '10'){'10'}elseif($vs.bits_per_raw_sample){"$($vs.bits_per_raw_sample)"}else{'8'}
    $srcDepth = if($rc['BitDepth'] -and "$($rc['BitDepth'])" -ne ''){"$($rc['BitDepth'])"}elseif($rc['PixFmt'] -match '10'){'10'}else{'8'}
    $checks += @{ Name = 'Bit Depth'; Status = if($outDepth -eq $srcDepth){'PASS'}else{'FAIL'}
        Detail = "Output: ${outDepth}-bit$(if($outDepth -ne $srcDepth){" (expected ${srcDepth}-bit)"})" }

    # Color metadata
    foreach ($prop in @(
        @{N='Color Primaries'; S='ColorPrimaries'; F='color_primaries'},
        @{N='Transfer (HDR)'; S='ColorTransfer'; F='color_transfer'},
        @{N='Color Matrix'; S='ColorSpace'; F='color_space'}
    )) {
        if ($rc[$prop.S]) {
            $outVal = $vs.($prop.F)
            $match = $outVal -eq $rc[$prop.S]
            $checks += @{ Name = $prop.N; Status = if($match){'PASS'}elseif($outVal){'WARN'}else{'WARN'}
                Detail = "Output: $outVal$(if(-not $match){" (expected $($rc[$prop.S]))"})" }
        }
    }

    # Profile
    if ($rc['Profile'] -and $vs.profile) {
        $checks += @{ Name = 'Profile'; Status = if($vs.profile -match [regex]::Escape($rc['Profile'])){'PASS'}else{'WARN'}
            Detail = "Output: $($vs.profile)" }
    }

    # DV RPU
    if ($rc['DV_Profile'] -and $vs.side_data_list) {
        $hasDV = $vs.side_data_list | Where-Object { $_.side_data_type -match 'DOVI' }
        $checks += @{ Name = 'Dolby Vision RPU'; Status = if($hasDV){'PASS'}else{'WARN'}
            Detail = if($hasDV){"DV metadata present"}else{"DV RPU not detected (may need dovi_tool)"} }
    }

    # HDR mastering display (FFprobe often misses SEI-level metadata; MediaInfo step confirms later)
    if ($rc['MasterMaxLum']) {
        $hasMD = $vs.side_data_list | Where-Object { $_.side_data_type -match 'Mastering display' }
        $hasCLL = $vs.side_data_list | Where-Object { $_.side_data_type -match 'Content light level' }
        $checks += @{ Name = 'HDR Mastering Display'; Status = if($hasMD){'PASS'}else{'INFO'}
            Detail = if($hasMD){"Mastering display preserved"}else{"Not in FFprobe side_data (checked via MediaInfo below)"} }
        if ($hasCLL) {
            $checks += @{ Name = 'HDR MaxCLL/FALL'; Status = 'PASS'; Detail = "Content light level present" }
        }
    }

    # ══════════════════════════════════════════════════════════════
    # Step 2: Frame / GOP analysis
    # ══════════════════════════════════════════════════════════════
    Write-Host "        [2/6] Frame/GOP analysis..." -ForegroundColor DarkGray
    $outDurSec = 0
    try { if ($probe.format.duration) { $outDurSec = [double]$probe.format.duration } } catch {}
    $fps = if ($rc['FPS']) { $rc['FPS'] } else { 24 }
    $probeFrameCount = [math]::Min(500, [math]::Max(100, [math]::Round($outDurSec * $fps * 0.3)))

    $fr = Run-Command $Tools.FFprobe @(
        '-v','quiet','-print_format','json','-select_streams','v:0','-show_frames',
        '-read_intervals',"0%+#$probeFrameCount",
        '-show_entries','frame=pict_type,key_frame,pkt_size',
        "`"$OutputFile`""
    ) -TimeoutSeconds ([math]::Max(60, $probeFrameCount / 2))

    if ($fr.ExitCode -eq 0 -and $fr.StdOut) {
        try {
            $frames = ($fr.StdOut | ConvertFrom-Json).frames
            if ($script:DebugMode) {
                Write-Host "      [DEBUG] Verify frames: $($frames.Count) parsed, requested=$probeFrameCount" -ForegroundColor Magenta
            }
            if ($frames -and $frames.Count -gt 0) {
                $iC = @($frames | Where-Object { $_.pict_type -eq 'I' }).Count
                $pC = @($frames | Where-Object { $_.pict_type -eq 'P' }).Count
                $bC = @($frames | Where-Object { $_.pict_type -eq 'B' }).Count
                $checks += @{ Name = 'Frame Distribution'; Status = 'INFO'
                    Detail = "Encoded: I=$iC P=$pC B=$bC ($($frames.Count)f) | Source: I=9% P=16% B=75%" }

                # GOP lengths
                $gopLens = @(); $cur = 0
                foreach ($f in $frames) { $cur++; if ($f.key_frame -eq 1 -and $cur -gt 1) { $gopLens += $cur - 1; $cur = 1 } }
                if ($cur -gt 0) { $gopLens += $cur }
                if ($gopLens.Count -gt 1) {
                    $avgGOP = [math]::Round(($gopLens | Measure-Object -Average).Average, 1)
                    $maxGOP = ($gopLens | Measure-Object -Max).Maximum
                    $checks += @{ Name = 'GOP Structure'; Status = 'INFO'
                        Detail = "Avg GOP=$avgGOP, Max=$maxGOP (keyint=$($rc['Keyint']))" }
                }

                # B-frame consecutive runs
                $maxCB = 0; $cb = 0
                foreach ($f in $frames) { if ($f.pict_type -eq 'B'){$cb++} else { if($cb -gt $maxCB){$maxCB=$cb}; $cb=0 } }
                if ($cb -gt $maxCB) { $maxCB = $cb }
                $expB = if ($EncoderName -eq 'NVEncC' -and $rc['NVEnc_MaxBFrames']) { [int]$rc['NVEnc_MaxBFrames'] } else { [int]$rc['BFrames'] }
                $bSt = if ($maxCB -le $expB) {'PASS'} elseif ($maxCB -le ($expB*3)) {'INFO'} else {'WARN'}
                $bN = if ($maxCB -gt $expB -and $pC -eq 0) {" (GPB mode)"} elseif ($maxCB -gt $expB) {" (hierarchical B)"} else {''}
                $checks += @{ Name = 'Max Consecutive B'; Status = $bSt
                    Detail = "Max $maxCB (setting: $expB, source: $($rc['BFrames']))$bN" }

                # Frame-average bitrate
                $avgSz = ($frames | Where-Object {$_.pkt_size} | ForEach-Object {[int]$_.pkt_size} | Measure-Object -Average).Average
                if ($avgSz -and $fps) {
                    $estMbps = [math]::Round($avgSz * 8 * $fps / 1000000, 2)
                    $tgtMbps = if($rc['BitrateAvgKbps']){[math]::Round($rc['BitrateAvgKbps']/1000,2)}else{0}
                    $checks += @{ Name = 'Bitrate (frame avg)'; Status = 'INFO'; Detail = "~${estMbps} Mbps (target: ${tgtMbps} Mbps)" }
                }
            }
        } catch {}
    }

    # File-level bitrate (always, quick sanity check)
    try {
        $outSize = (Get-Item $OutputFile).Length
        if ($outDurSec -gt 0) {
            $actMbps = [math]::Round($outSize * 8 / $outDurSec / 1000000, 2)
            $tgtMbps = if($rc['BitrateAvgKbps']){[math]::Round($rc['BitrateAvgKbps']/1000,2)}else{0}
            $brD = "File: ${actMbps} Mbps"
            if ($tgtMbps -gt 0) {
                $rat = [math]::Round($actMbps / $tgtMbps, 2)
                $brD += " (target: ${tgtMbps}, ratio: ${rat}x)"
                $brS = if($rat -ge 0.7 -and $rat -le 1.5){'PASS'}elseif($rat -ge 0.3){'WARN'}else{'WARN'}
            } else { $brS = 'INFO' }
            $checks += @{ Name = 'Bitrate (file)'; Status = $brS; Detail = $brD }
        }
    } catch {}

    # ══════════════════════════════════════════════════════════════
    # Step 3: MediaInfo (HDR format, encoder ID)
    # ══════════════════════════════════════════════════════════════
    Write-Host "        [3/6] MediaInfo verification..." -ForegroundColor DarkGray
    if ($Tools.MediaInfo) {
        $miR = Run-Command $Tools.MediaInfo @('--Output=JSON',"`"$OutputFile`"") -TimeoutSeconds 60
        if ($miR.ExitCode -eq 0 -and $miR.StdOut) {
            try {
                $miJson = $miR.StdOut | ConvertFrom-Json
                $miV = $miJson.media.track | Where-Object { $_.'@type' -eq 'Video' } | Select-Object -First 1
                if ($script:DebugMode) {
                    Write-Host "      [DEBUG] Verify MI: tracks=$($miJson.media.track.Count) HDR=$($miV.HDR_Format) MaxCLL=$($miV.MaxCLL)" -ForegroundColor Magenta
                }
                if ($miV) {
                    # HDR format
                    if ($miV.HDR_Format) {
                        $hasDVMI = $miV.HDR_Format -match 'Dolby Vision'
                        $hasHDR10 = $miV.HDR_Format -match 'HDR10|SMPTE ST 2086'
                        $checks += @{ Name = 'MediaInfo HDR'; Status = if($hasDVMI -or $hasHDR10){'PASS'}else{'WARN'}
                            Detail = "$($miV.HDR_Format)" }
                    }
                    # Mastering display (reliable check — catches what FFprobe misses)
                    $mdLum = $miV.MasteringDisplay_Luminance
                    if ($rc['MasterMaxLum']) {
                        if ($mdLum) {
                            $checks += @{ Name = 'MI: Mastering Display'; Status = 'PASS'
                                Detail = "$mdLum" }
                        } else {
                            $checks += @{ Name = 'MI: Mastering Display'; Status = 'WARN'
                                Detail = "Not found in MediaInfo output" }
                        }
                    }
                    # MaxCLL/MaxFALL
                    $maxCLL = $miV.MaxCLL
                    if ($maxCLL) {
                        $checks += @{ Name = 'MI: MaxCLL'; Status = 'PASS'; Detail = "MaxCLL=$maxCLL MaxFALL=$($miV.MaxFALL)" }
                    }
                    # Encoder identification
                    $encLib = $miV.Encoded_Library
                    $encSet = if($miV.Encoded_Library_Settings){$miV.Encoded_Library_Settings.Substring(0,[math]::Min(120,$miV.Encoded_Library_Settings.Length))+'...'}else{$null}
                    $encName2 = if ($encLib -match 'nvenc') {'NVEncC'} elseif ($encLib -match 'qsv|intel') {'QSVEncC'} elseif ($encLib -match 'x265') {'x265'} else { $EncoderName }
                    $checks += @{ Name = 'Encoder (MediaInfo)'; Status = 'INFO'
                        Detail = "$(if($encLib){"$encLib ($encName2)"}else{"No encoder tag ($encName2)"})$(if($encSet){" | $encSet"})" }
                }
            } catch {}
        }
    } else {
        Write-Host "        [3/6] MediaInfo: not available" -ForegroundColor Yellow
    }

    # ══════════════════════════════════════════════════════════════
    # Step 5: HM Reference Decoder (SPS/PPS/QP from bitstream)
    # ══════════════════════════════════════════════════════════════
    if ($Tools.TAppDecoder) {
        Write-Host "        [4/6] HM Reference Decoder analysis..." -ForegroundColor DarkGray
        # Scale frame count: more frames for larger files, but cap for speed
        $hmFrames = [math]::Min(300, [math]::Max(100, [math]::Round($outDurSec * 2)))

        # Pass MKV directly — Get-HMAnalysis handles extraction + DV RPU strip
        $hmResult = Get-HMAnalysis -FilePath $OutputFile -MaxFrames $hmFrames
            if ($hmResult) {
                    # QP comparison
                    if ($hmResult.QP -and $hmResult.QP.Stats) {
                        $hqp = $hmResult.QP.Stats
                        $srcQP = "$($rc['QP_Min'])-$($rc['QP_Max']) (avg $($rc['QP_Avg']))"
                        $encQP = "$($hqp.Min)-$($hqp.Max) (avg $([math]::Round($hqp.Avg,1)))"
                        $qpOK = [int]$hqp.Min -ge ([int]$rc['QP_Min'] - 3) -and [int]$hqp.Max -le ([int]$rc['QP_Max'] + 5)
                        $checks += @{ Name = 'HM: QP Range'; Status = if($qpOK){'PASS'}else{'INFO'}
                            Detail = "Encoded: $encQP | Source: $srcQP" }
                        # Per-type QP
                        if ($hmResult.QP.ByType) {
                            foreach ($t in 'I','P','B') {
                                $ht = $hmResult.QP.ByType[$t]
                                $srcKey = "QP_${t}_Min"
                                if ($ht -and $rc[$srcKey]) {
                                    $checks += @{ Name = "HM: QP $t-Frame"; Status = 'INFO'
                                        Detail = "Enc: avg=$([math]::Round($ht.Avg,1)) min=$($ht.Min) max=$($ht.Max) | Src: $($rc[$srcKey])-$($rc["QP_${t}_Max"]) (avg $([math]::Round([double]$rc["QP_${t}_Avg"],1)))" }
                                }
                            }
                        }
                        # Ref structure
                        if ($hmResult.QP.RefStructure) {
                            $rs = $hmResult.QP.RefStructure
                            $checks += @{ Name = 'HM: Ref Structure'; Status = 'INFO'
                                Detail = "Enc: L0=$($rs.MaxL0Refs) L1=$($rs.MaxL1Refs) TLayers=$($rs.TemporalLayers) | Src: L0=2 L1=1" }
                        }
                    }
                    # SPS comparison
                    if ($hmResult.SPS) {
                        $sps = $hmResult.SPS
                        if ($null -ne $sps['dpb_size']) {
                            $checks += @{ Name = 'HM: DPB/Refs'; Status = 'INFO'
                                Detail = "Enc: DPB=$($sps['dpb_size']) MaxReorder=$($sps['max_reorder_pics']) | Src: DPB=5 MaxReorder=3" }
                        }
                        if ($sps['log2_max_cu']) {
                            $encCTU = [math]::Pow(2, $sps['log2_max_cu'])
                            $checks += @{ Name = 'HM: CTU Size'; Status = if($encCTU -eq $rc['CTU']){'PASS'}else{'INFO'}
                                Detail = "Enc: $encCTU | Src: $($rc['CTU'])" }
                        }
                        if ($null -ne $sps['sao_enabled']) {
                            $checks += @{ Name = 'HM: SAO'; Status = if([bool]$sps['sao_enabled'] -eq [bool]$rc['SAO']){'PASS'}else{'WARN'}
                                Detail = "Enc: $(if($sps['sao_enabled']){'On'}else{'Off'}) | Src: $(if($rc['SAO']){'On'}else{'Off'})" }
                        }
                        if ($sps['bit_depth_luma']) {
                            $checks += @{ Name = 'HM: Bit Depth'; Status = if($sps['bit_depth_luma'] -eq 10){'PASS'}else{'WARN'}
                                Detail = "Enc: $($sps['bit_depth_luma'])/$($sps['bit_depth_chroma']) | Src: 10/10" }
                        }
                        if ($null -ne $sps['amp_enabled']) {
                            $checks += @{ Name = 'HM: AMP'; Status = 'INFO'
                                Detail = "Enc: $(if($sps['amp_enabled']){'On'}else{'Off'}) | Src: $(if($rc['AMP']){'On'}else{'Off'})" }
                        }
                        if ($null -ne $sps['strong_intra_smoothing']) {
                            $checks += @{ Name = 'HM: Strong Intra'; Status = 'INFO'
                                Detail = "Enc: $(if($sps['strong_intra_smoothing']){'On'}else{'Off'}) | Src: $(if($rc['StrongIntra']){'On'}else{'Off'})" }
                        }
                    }
                    # PPS
                    if ($hmResult.PPS) {
                        $pps = $hmResult.PPS
                        if ($null -ne $pps['cu_qp_delta']) {
                            $checks += @{ Name = 'HM: CU QP Delta'; Status = if([bool]$pps['cu_qp_delta'] -eq [bool]$rc['CUQPDelta']){'PASS'}else{'INFO'}
                                Detail = "Enc: $(if($pps['cu_qp_delta']){'On'}else{'Off'}) | Src: $(if($rc['CUQPDelta']){'On'}else{'Off'})" }
                        }
                        if ($null -ne $pps['weighted_pred']) {
                            $checks += @{ Name = 'HM: Weighted Pred'; Status = 'INFO'
                                Detail = "Enc: WP=$(if($pps['weighted_pred']){'On'}else{'Off'}) WB=$(if($pps['weighted_bipred']){'On'}else{'Off'}) | Src: WP=$(if($rc['WeightedPred']){'On'}else{'Off'}) WB=$(if($rc['WeightedBipred']){'On'}else{'Off'})" }
                        }
                    }
                    # VUI color
                    if ($hmResult.VUI) {
                        $vui = $hmResult.VUI
                        if ($vui['colour_primaries']) {
                            $checks += @{ Name = 'HM: VUI Color'; Status = if($vui['colour_primaries'] -eq 9){'PASS'}else{'WARN'}
                                Detail = "Enc: prim=$($vui['colour_primaries']) xfer=$($vui['transfer']) mx=$($vui['matrix_coeffs']) | Src: 9/16/9 (BT.2020/PQ)" }
                        }
                    }
                } else {
                    $checks += @{ Name = 'HM Analysis'; Status = 'WARN'
                        Detail = "HM returned no results (Get-HMAnalysis handles DV strip via dovi_tool)" }
                }
    } else {
        Write-Host "        [4/6] HM Analyser: not available" -ForegroundColor Yellow
    }

    # ══════════════════════════════════════════════════════════════
    # Step 6: CheckBitrate distribution (primary bitrate verifier)
    # ══════════════════════════════════════════════════════════════
    if ($Tools.CheckBitrate -and $outDurSec -gt 10) {
        Write-Host "        [5/6] CheckBitrate analysis..." -ForegroundColor DarkGray
        $cbData = Get-CheckBitrateData $OutputFile
        if ($cbData) {
            $encAvgMbps = [math]::Round($cbData.FinalAvg / 1000, 2)
            $srcAvgMbps = if($rc['BitrateAvgKbps']){[math]::Round($rc['BitrateAvgKbps']/1000,2)}else{0}

            # Avg bitrate with PASS/WARN status
            if ($srcAvgMbps -gt 0) {
                $rat = [math]::Round($encAvgMbps / $srcAvgMbps, 2)
                $brS = if ($rat -ge 0.7 -and $rat -le 1.5) {'PASS'}
                       elseif ($rat -ge 0.3 -and $rat -le 2.5) {'WARN'}
                       else {'WARN'}
                $checks += @{ Name = 'CB: Avg Bitrate'; Status = $brS
                    Detail = "Enc: $encAvgMbps Mbps | Src: $srcAvgMbps Mbps (ratio: ${rat}x)" }
            } else {
                $checks += @{ Name = 'CB: Avg Bitrate'; Status = 'INFO'
                    Detail = "Enc: $encAvgMbps Mbps" }
            }

            # Peak bitrate
            if ($cbData.Peak) {
                $encPeakMbps = [math]::Round($cbData.Peak / 1000, 2)
                $srcPeakMbps = if($rc['PeakKbps']){[math]::Round($rc['PeakKbps']/1000,2)}else{0}
                $srcMaxrateMbps = if($rc['MaxrateLikely']){[math]::Round($rc['MaxrateLikely']/1000,2)}else{0}
                # Peak should not wildly exceed source maxrate
                $peakS = if ($srcMaxrateMbps -gt 0 -and $encPeakMbps -le ($srcMaxrateMbps * 1.1)) {'PASS'}
                         elseif ($srcMaxrateMbps -gt 0 -and $encPeakMbps -le ($srcMaxrateMbps * 1.5)) {'INFO'}
                         else {'INFO'}
                $checks += @{ Name = 'CB: Peak Bitrate'; Status = $peakS
                    Detail = "Enc: $encPeakMbps Mbps | Src peak: $srcPeakMbps Mbps (maxrate: $srcMaxrateMbps)" }
            }

            # VBR pattern (CoV comparison)
            if ($cbData.StdDev -and $cbData.FinalAvg -gt 0) {
                $encCoV = [math]::Round($cbData.StdDev / $cbData.FinalAvg * 100, 1)
                $srcCoV = if($rc['CoefVar']){[math]::Round($rc['CoefVar'],1)}else{29.3}
                $checks += @{ Name = 'CB: VBR Pattern'; Status = 'INFO'
                    Detail = "Enc: CoV=${encCoV}% | Src: ${srcCoV}%" }
            }

            # VBV capping detection
            if ($cbData.PeaksCapped) {
                $effMaxMbps = [math]::Round($cbData.PeakClusterKbps / 1000, 2)
                $srcMaxMbps = if($rc['MaxrateLikely']){[math]::Round($rc['MaxrateLikely']/1000,2)}else{0}
                $checks += @{ Name = 'CB: VBV Ceiling'; Status = 'INFO'
                    Detail = "Enc VBV caps at ~$effMaxMbps Mbps (src: ~$srcMaxMbps Mbps)" }
            }
        }
    } elseif ($Tools.CheckBitrate) {
        Write-Host "        [5/6] CheckBitrate: segment too short ($([math]::Round($outDurSec))s)" -ForegroundColor Yellow
        # Fallback to file-level bitrate if CheckBitrate can't run
        try {
            $outSize = (Get-Item $OutputFile).Length
            if ($outDurSec -gt 0) {
                $actMbps = [math]::Round($outSize * 8 / $outDurSec / 1000000, 2)
                $tgtMbps = if($rc['BitrateAvgKbps']){[math]::Round($rc['BitrateAvgKbps']/1000,2)}else{0}
                $brD = "File: ${actMbps} Mbps"
                if ($tgtMbps -gt 0) {
                    $rat = [math]::Round($actMbps / $tgtMbps, 2)
                    $brD += " (target: ${tgtMbps}, ratio: ${rat}x)"
                    $brS = if($rat -ge 0.7 -and $rat -le 1.5){'PASS'}elseif($rat -ge 0.3){'WARN'}else{'WARN'}
                } else { $brS = 'INFO' }
                $checks += @{ Name = 'Bitrate (file)'; Status = $brS; Detail = $brD }
            }
        } catch {}
    } else {
        Write-Host "        [5/6] CheckBitrate: not available" -ForegroundColor Yellow
    }

    # ══════════════════════════════════════════════════════════════
    # Step 6: Quality metrics (PSNR, SSIM, VMAF) vs source segment
    # ══════════════════════════════════════════════════════════════
    if ($SourceSegment -and (Test-Path $SourceSegment) -and $Tools.FFmpeg) {
        Write-Host "        [6/6] Quality metrics (PSNR/SSIM/VMAF)..." -ForegroundColor DarkGray
        try {
            # Check if FFmpeg has libvmaf support
            $vmafAvail = $false
            $vmafCheck = Run-Command $Tools.FFmpeg @('-filters') -TimeoutSeconds 10
            if ($vmafCheck.StdOut -match 'libvmaf' -or $vmafCheck.StdErr -match 'libvmaf') { $vmafAvail = $true }

            # Timeout: scale with resolution — 4K needs much more time than 1080p
            # Use file size as a proxy: >5GB file segment → assume 4K
            $segSize = if (Test-Path $SourceSegment) { (Get-Item $SourceSegment).Length } else { 0 }
            $metricTimeout = if ($segSize -gt 500MB) { 3600 } elseif ($segSize -gt 100MB) { 1800 } else { 600 }

            # Shared lavfi arg builder — avoids inline quote nesting issues
            function Build-LavfiArgs {
                param([string]$Src, [string]$Enc, [string]$Filter)
                return @(
                    '-hide_banner', '-an', '-sn',
                    '-i', "`"$Src`"",
                    '-i', "`"$Enc`"",
                    '-lavfi', $Filter,
                    '-f', 'null', '-'
                )
            }

            # ── VMAF: separate pass, parse "VMAF score: XX.XX" from stderr ──
            if ($vmafAvail) {
                $vmafFilter = "[0:v]setpts=PTS-STARTPTS[ref];[1:v]setpts=PTS-STARTPTS[enc];[ref][enc]libvmaf=shortest=1:n_threads=4"
                $vArgs = Build-LavfiArgs $SourceSegment $OutputFile $vmafFilter
                $vr = Run-Command $Tools.FFmpeg $vArgs -TimeoutSeconds $metricTimeout -StatusLabel "VMAF"
                if ($script:DebugMode) {
                    Write-Host "        [DEBUG] VMAF exit=$($vr.ExitCode)" -ForegroundColor Magenta
                    $tail = if ($vr.StdErr.Length -gt 400) { $vr.StdErr.Substring($vr.StdErr.Length - 400) } else { $vr.StdErr }
                    Write-Host "        [DEBUG] VMAF stderr tail: $tail" -ForegroundColor Magenta
                }
                $allOut = "$($vr.StdOut)`n$($vr.StdErr)"
                if ($vr.ExitCode -eq 0 -and $allOut -match 'VMAF score[:\s=]+([\d.]+)') {
                    $vmafScore = [math]::Round([double]$Matches[1], 2)
                    $vmafS = if ($vmafScore -ge 93) { 'PASS' } elseif ($vmafScore -ge 80) { 'WARN' } else { 'FAIL' }
                    $checks += @{ Name = 'VMAF Score'; Status = $vmafS; Detail = "Mean: $vmafScore" }
                } else {
                    Write-Host "        VMAF failed (exit=$($vr.ExitCode)), skipping VMAF check" -ForegroundColor Yellow
                }
            }

            # ── PSNR: separate pass ──
            $psnrFilter = "[0:v]setpts=PTS-STARTPTS[ref];[1:v]setpts=PTS-STARTPTS[enc];[ref][enc]psnr=shortest=1"
            $pArgs = Build-LavfiArgs $SourceSegment $OutputFile $psnrFilter
            $pr = Run-Command $Tools.FFmpeg $pArgs -TimeoutSeconds $metricTimeout -StatusLabel "PSNR"
            $pOut = "$($pr.StdOut)`n$($pr.StdErr)"
            if ($pr.ExitCode -eq 0 -and $pOut -match 'PSNR y:([\d.]+)') {
                $psnrY = [math]::Round([double]$Matches[1], 2)
                $psnrS = if ($psnrY -ge 40) { 'PASS' } elseif ($psnrY -ge 35) { 'INFO' } else { 'WARN' }
                $checks += @{ Name = 'PSNR (Y)'; Status = $psnrS; Detail = "Mean: ${psnrY} dB" }
            }

            # ── SSIM: separate pass ──
            $ssimFilter = "[0:v]setpts=PTS-STARTPTS[ref];[1:v]setpts=PTS-STARTPTS[enc];[ref][enc]ssim=shortest=1"
            $sArgs = Build-LavfiArgs $SourceSegment $OutputFile $ssimFilter
            $sr = Run-Command $Tools.FFmpeg $sArgs -TimeoutSeconds $metricTimeout -StatusLabel "SSIM"
            $sOut = "$($sr.StdOut)`n$($sr.StdErr)"
            if ($sr.ExitCode -eq 0 -and $sOut -match 'All:([\d.]+)') {
                $ssimVal = [math]::Round([double]$Matches[1], 4)
                $ssimS = if ($ssimVal -ge 0.95) { 'PASS' } elseif ($ssimVal -ge 0.90) { 'INFO' } else { 'WARN' }
                $ssimDb = if ($ssimVal -lt 1.0) { [math]::Round(-10 * [math]::Log10(1 - $ssimVal), 2) } else { 'Inf' }
                $checks += @{ Name = 'SSIM'; Status = $ssimS; Detail = "Mean: $ssimVal ($ssimDb dB)" }
            }

        } catch {
            Write-Host "        Quality metrics failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    } elseif (-not $SourceSegment -or -not (Test-Path $SourceSegment)) {
        Write-Host "        [6/6] Quality metrics: source segment not available" -ForegroundColor Yellow
    }

    # ── Print all results to console ──
    Write-Host "" -ForegroundColor DarkGray
    foreach ($c in $checks) {
        $icon = switch ($c.Status) { 'PASS' {'[OK]'} 'WARN' {'[!!]'} 'FAIL' {'[XX]'} 'INFO' {'[--]'} 'SKIP' {'[..]'} }
        $color = switch ($c.Status) { 'PASS' {'Green'} 'WARN' {'Yellow'} 'FAIL' {'Red'} default {'DarkGray'} }
        Write-Host "        $icon $($c.Name.PadRight(24)) $($c.Detail)" -ForegroundColor $color
    }

    return $checks
}
function Write-VerificationReport {
    <#
    .SYNOPSIS
        Write verification encode results to the report.
    #>
    param(
        [System.Text.StringBuilder]$rpt,
        [hashtable]$VerifyResults,
        [array]$Segments
    )

    Write-Section $rpt "VERIFICATION ENCODE RESULTS"

    if (-not $VerifyResults -or $VerifyResults.Count -eq 0) {
        $rpt.AppendLine("    No GPU encoders available for verification.") | Out-Null
        return
    }

    # Show segment info
    $totalDur = ($Segments | Measure-Object -Property Duration -Sum).Sum
    $modeLabel = if ($Segments.Count -eq 1 -and $totalDur -gt 30) { "extended (${totalDur}s continuous)" } else { "quick ($($Segments.Count) segments)" }
    $rpt.AppendLine("    Verification mode: $modeLabel") | Out-Null
    foreach ($seg in $Segments) {
        $rpt.AppendLine("      $($seg.Label): $($seg.Duration)s (~$($seg.Frames) frames)") | Out-Null
    }
    $rpt.AppendLine("") | Out-Null

    foreach ($encName in @('NVEnc','QSVEnc')) {
        if (-not $VerifyResults.ContainsKey($encName)) { continue }
        $er = $VerifyResults[$encName]
        $displayName = switch ($encName) { 'NVEnc' { 'NVEncC (NVIDIA)' } 'QSVEnc' { 'QSVEncC (Intel)' } }

        $rpt.AppendLine("    [$displayName]") | Out-Null

        if ($er.Success) {
            $sizeKB = [math]::Round($er.OutputSize / 1024, 1)
            $rpt.AppendLine("        Encode Time: $([math]::Round($er.Duration, 1))s | Output: ${sizeKB} KB") | Out-Null
        }

        # Count results
        $pass = @($er.Checks | Where-Object { $_.Status -eq 'PASS' }).Count
        $warn = @($er.Checks | Where-Object { $_.Status -eq 'WARN' }).Count
        $fail = @($er.Checks | Where-Object { $_.Status -eq 'FAIL' }).Count
        $info = @($er.Checks | Where-Object { $_.Status -eq 'INFO' }).Count
        $rpt.AppendLine("        Results: $pass PASS, $warn WARN, $fail FAIL, $info INFO") | Out-Null
        $rpt.AppendLine("") | Out-Null

        foreach ($c in $er.Checks) {
            $icon = switch ($c.Status) {
                'PASS' { [char]0x2713 }  # ✓
                'WARN' { [char]0x26A0 }  # ⚠
                'FAIL' { [char]0x2717 }  # ✗
                'INFO' { [char]0x2139 }  # ℹ
                'SKIP' { '-' }
            }
            $pad = $c.Name.PadRight(24)
            $rpt.AppendLine("        $icon $pad $($c.Detail)") | Out-Null
        }
        $rpt.AppendLine("") | Out-Null

        # Show stderr excerpt if encode failed
        if (-not $er.Success -and $er.StdErr) {
            $errLines = ($er.StdErr -split "`n" | Select-Object -Last 5) -join "`n          "
            $rpt.AppendLine("        Error output:") | Out-Null
            $rpt.AppendLine("          $errLines") | Out-Null
            $rpt.AppendLine("") | Out-Null
        }
    }

    # Summary
    $allPass = $true
    foreach ($er in $VerifyResults.Values) {
        if (@($er.Checks | Where-Object { $_.Status -eq 'FAIL' }).Count -gt 0) { $allPass = $false; break }
    }
    if ($allPass) {
        $rpt.AppendLine("    Overall: ALL CHECKS PASSED - generated commands produce valid output") | Out-Null
    } else {
        $rpt.AppendLine("    Overall: ISSUES DETECTED - review FAIL items above and adjust commands") | Out-Null
    }
}