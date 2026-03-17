# ─────────────────────────────────────────────────────────────────────────────
# ReportWriter.ps1 — Report generation functions (one per analysis section)
# Requires: Helpers.ps1 functions (Write-Section, Write-Field, etc.)
# ─────────────────────────────────────────────────────────────────────────────

function Write-ContainerReport {
    param([System.Text.StringBuilder]$rpt, $probe, $rc)
    if (-not $probe) { return }
    Write-Section $rpt "CONTAINER / FORMAT"
    $fmt = $probe.format
    if ($fmt) {
        Write-Field $rpt "Format"           $fmt.format_name
        Write-Field $rpt "Format Long"      $fmt.format_long_name
        Write-Field $rpt "Duration"         ("{0} ({1:hh\:mm\:ss\.fff})" -f $fmt.duration, [TimeSpan]::FromSeconds([double]$fmt.duration))
        Write-Field $rpt "Overall Bitrate"  (Format-Bitrate ([double]$fmt.bit_rate / 1000))
        Write-Field $rpt "Streams"          $fmt.nb_streams
        Write-Field $rpt "Start Time"       $fmt.start_time
        Write-Field $rpt "Probe Score"      $fmt.probe_score
        if ($fmt.tags) {
            $rpt.AppendLine("`n    Container Tags:") | Out-Null
            foreach ($t in $fmt.tags.PSObject.Properties) { Write-Field $rpt $t.Name $t.Value -Indent 8 -Width 32 }
        }
    }
    # Streams
    foreach ($st in $probe.streams) {
        $cType = if ($st.codec_type) { $st.codec_type } else { "unknown" }
        $sType = $cType.ToUpper()
        if ($script:DebugMode) { Write-Host "      [DEBUG] Stream #$($st.index): codec_type='$cType' codec_name='$($st.codec_name)'" -ForegroundColor Magenta }
        Write-Section $rpt "$sType STREAM #$($st.index)"
        Write-Field $rpt "Codec"            "$($st.codec_name) ($($st.codec_long_name))"
        Write-Field $rpt "Codec Tag"        $st.codec_tag_string
        Write-Field $rpt "Profile"          $st.profile
        Write-Field $rpt "Level"            $st.level
        Write-Field $rpt "Pixel Format"     $st.pix_fmt
        Write-Field $rpt "Bits/Raw Sample"  $st.bits_per_raw_sample
        Write-Field $rpt "Bits/Coded Sample" $st.bits_per_coded_sample

        if ($st.codec_type -eq 'video') {
            Write-Field $rpt "Resolution"       "$($st.width)x$($st.height)"
            Write-Field $rpt "Coded Resolution"  "$($st.coded_width)x$($st.coded_height)"
            Write-Field $rpt "SAR / DAR"         "$($st.sample_aspect_ratio) / $($st.display_aspect_ratio)"
            Write-Field $rpt "Frame Rate (avg)"  $st.avg_frame_rate
            Write-Field $rpt "Frame Rate (r)"    $st.r_frame_rate
            Write-Field $rpt "Time Base"         $st.time_base
            Write-Field $rpt "Frame Count"       $st.nb_frames
            Write-Field $rpt "Field Order"       $st.field_order
            Write-Field $rpt "Has B-Frames"      $st.has_b_frames
            Write-Field $rpt "Reference Frames"  $st.refs
            Write-Field $rpt "Color Space"       $st.color_space
            Write-Field $rpt "Color Transfer"    $st.color_transfer
            Write-Field $rpt "Color Primaries"   $st.color_primaries
            Write-Field $rpt "Color Range"       $st.color_range
            Write-Field $rpt "Chroma Location"   $st.chroma_location
            Write-Field $rpt "NAL Length Size"   $st.nal_length_size

            # Only set rc from FIRST non-attached-pic video stream
            $isAttached = $st.disposition -and $st.disposition.attached_pic -eq 1
            if (-not $rc['Codec'] -and -not $isAttached) {
                if ($st.bit_rate) {
                    Write-Field $rpt "Video Bitrate" (Format-Bitrate ([double]$st.bit_rate / 1000))
                    $rc['VideoBitrateKbps'] = [double]$st.bit_rate / 1000
                }
                $rc['Codec']=$st.codec_name; $rc['Profile']=$st.profile; $rc['Level']=$st.level
                $rc['Resolution']="$($st.width)x$($st.height)"; $rc['PixFmt']=$st.pix_fmt
                $rc['BitDepth']=$st.bits_per_raw_sample; $rc['BFrameFlag']=$st.has_b_frames
                $rc['Refs']=$st.refs; $rc['ColorSpace']=$st.color_space
                $rc['ColorTransfer']=$st.color_transfer; $rc['ColorPrimaries']=$st.color_primaries; $rc['ColorRange']=$st.color_range
            }

            if ($st.side_data_list) {
                $rpt.AppendLine("`n    Side Data (HDR/DV):") | Out-Null
                foreach ($sd in $st.side_data_list) {
                    $rpt.AppendLine("      - $($sd.side_data_type)") | Out-Null
                    foreach ($p in $sd.PSObject.Properties) {
                        if ($p.Name -ne 'side_data_type') { Write-Field $rpt $p.Name "$($p.Value)" -Indent 10 -Width 34 }
                    }
                    if ($sd.side_data_type -match 'DOVI') {
                        $rc['DV_Profile']=$sd.dv_profile; $rc['DV_Level']=$sd.dv_level
                        $rc['DV_Compat']= if($sd.dv_bl_signal_compatibility_id) {$sd.dv_bl_signal_compatibility_id}
                                          elseif($sd.dv_bl_signal_compatibility) {$sd.dv_bl_signal_compatibility}
                                          else {''}
                    }
                }
            }
        }
        if ($st.codec_type -eq 'audio') {
            Write-Field $rpt "Sample Rate"     "$($st.sample_rate) Hz"
            Write-Field $rpt "Channels"        $st.channels
            Write-Field $rpt "Channel Layout"  $st.channel_layout
            Write-Field $rpt "Sample Format"   $st.sample_fmt
            Write-Field $rpt "Bits/Sample"     $st.bits_per_sample
            if ($st.bit_rate) { Write-Field $rpt "Audio Bitrate" (Format-Bitrate ([double]$st.bit_rate / 1000)) }
        }
        if ($st.codec_type -eq 'subtitle') { Write-Field $rpt "Codec" $st.codec_name }
        if ($st.tags) {
            $rpt.AppendLine("`n    Stream Tags:") | Out-Null
            foreach ($t in $st.tags.PSObject.Properties) { Write-Field $rpt $t.Name "$($t.Value)" -Indent 8 -Width 32 }
        }
        if ($st.disposition) {
            $active = $st.disposition.PSObject.Properties | Where-Object { $_.Value -eq 1 } | ForEach-Object { $_.Name }
            if ($active) { Write-Field $rpt "Disposition" ($active -join ', ') }
        }
    }
    # Chapters
    if ($probe.chapters -and $probe.chapters.Count -gt 0) {
        Write-Section $rpt "CHAPTERS ($($probe.chapters.Count))"
        foreach ($ch in $probe.chapters) {
            $cs = [TimeSpan]::FromSeconds([double]$ch.start_time); $ce = [TimeSpan]::FromSeconds([double]$ch.end_time)
            $ct = if ($ch.tags -and $ch.tags.title) { $ch.tags.title } else { "(untitled)" }
            $rpt.AppendLine("    $($cs.ToString('hh\:mm\:ss\.fff')) - $($ce.ToString('hh\:mm\:ss\.fff'))  $ct") | Out-Null
        }
    }
}

function Write-FrameReport {
    param([System.Text.StringBuilder]$rpt, $frameData, $probe, $rc)
    if (-not $frameData -or -not $frameData.frames) { return }
    $fr = $frameData.frames
    $sampling = if ($frameData._sampling) { $frameData._sampling } else { 'sequential' }
    Write-Section $rpt "FRAME / GOP ANALYSIS ($($fr.Count) frames, $sampling sampling)"

    # Show segment breakdown for distributed sampling
    if ($sampling -eq 'distributed' -and $frameData._segments) {
        $rpt.AppendLine("    Sampled from $($frameData._segments.Count) segments across file:") | Out-Null
        foreach ($seg in $frameData._segments) {
            $rpt.AppendLine("      $($seg.Label): $($seg.FrameCount) frames") | Out-Null
        }
        $rpt.AppendLine("") | Out-Null
    }

    $iF = @($fr | Where-Object { $_.pict_type -eq 'I' })
    $pF = @($fr | Where-Object { $_.pict_type -eq 'P' })
    $bF = @($fr | Where-Object { $_.pict_type -eq 'B' })

    Write-Field $rpt "Frames Analyzed" $fr.Count
    Write-Field $rpt "I-Frames" "$($iF.Count) ($([math]::Round($iF.Count/$fr.Count*100,1))%)"
    Write-Field $rpt "P-Frames" "$($pF.Count) ($([math]::Round($pF.Count/$fr.Count*100,1))%)"
    Write-Field $rpt "B-Frames" "$($bF.Count) ($([math]::Round($bF.Count/$fr.Count*100,1))%)"

    # GOP lengths — for distributed sampling, analyze within each segment separately
    # to avoid artificial short GOPs at segment boundaries
    $gl = @()
    if ($sampling -eq 'distributed' -and $frameData._segments -and $frameData._segments.Count -gt 1) {
        # Process each segment's frames independently
        $offset = 0
        foreach ($seg in $frameData._segments) {
            $segEnd = $offset + $seg.FrameCount
            $segKI = @()
            for ($i = $offset; $i -lt $segEnd -and $i -lt $fr.Count; $i++) {
                if ($fr[$i].key_frame -eq 1) { $segKI += ($i - $offset) }
            }
            # Only count GOPs fully within this segment (skip first/last partial GOPs)
            if ($segKI.Count -ge 2) {
                for ($i = 1; $i -lt $segKI.Count; $i++) {
                    $gopLen = $segKI[$i] - $segKI[$i-1]
                    if ($gopLen -gt 0) { $gl += $gopLen }
                }
            }
            $offset = $segEnd
        }
    } else {
        # Sequential sampling: original logic
        $ki = @(); for ($i=0; $i -lt $fr.Count; $i++) { if ($fr[$i].key_frame -eq 1) { $ki += $i } }
        if ($ki.Count -ge 2) {
            for ($i=1; $i -lt $ki.Count; $i++) { $gl += ($ki[$i] - $ki[$i-1]) }
        }
    }

    if ($gl.Count -ge 2) {
        $avg = ($gl | Measure-Object -Average).Average
        $gMin = ($gl | Measure-Object -Minimum).Minimum; $gMax = ($gl | Measure-Object -Maximum).Maximum
        $gMode = ($gl | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name

        # For distributed sampling, also compute the "max mode" — the most common large GOP
        # which better represents the actual keyint setting
        $largeModeThresh = [math]::Max(8, [int]($gMax * 0.5))
        $largeGOPs = @($gl | Where-Object { $_ -ge $largeModeThresh })
        $effectiveKeyint = $gMode
        if ($largeGOPs.Count -ge 3) {
            $largeMode = ($largeGOPs | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name
            if ([int]$largeMode -gt [int]$gMode) {
                $effectiveKeyint = $largeMode
            }
        }

        Write-Field $rpt "GOP Count" $gl.Count
        Write-Field $rpt "Avg GOP Length" ("{0:N1} frames" -f $avg)
        Write-Field $rpt "Min / Max GOP" "$gMin / $gMax frames"
        Write-Field $rpt "Mode GOP Length" "$gMode frames"
        if ([int]$effectiveKeyint -ne [int]$gMode) {
            Write-Field $rpt "Likely Keyint (max)" "$effectiveKeyint frames (most common full-length GOP)"
        }

        $rc['Keyint'] = $effectiveKeyint; $rc['MinKeyint'] = $gMin
        $rc['GOPType'] = if (($gMax - $gMin) -lt 2) { "Fixed" } else { "Variable" }

        # Keyframe interval in seconds
        $vs = $probe.streams | Where-Object { $_.codec_type -eq 'video' } | Select-Object -First 1
        if ($vs -and $vs.avg_frame_rate -match '(\d+)/(\d+)') {
            $fps = [double]$Matches[1] / [double]$Matches[2]
            if ($fps -gt 0) {
                Write-Field $rpt "Keyframe Interval" "$([math]::Round([int]$gMode/$fps,2))s @ $([math]::Round($fps,3)) fps"
                $rc['FPS'] = $fps
            }
        }

        # Distribution
        $rpt.AppendLine("`n    GOP Length Distribution:") | Out-Null
        foreach ($g in ($gl | Group-Object | Sort-Object { [int]$_.Name })) {
            Write-Field $rpt "  $($g.Name) frames" "$($g.Count) GOPs ($([math]::Round($g.Count/$gl.Count*100,1))%)" -Indent 8 -Width 20
        }
    }

    # Frame sizes
    $iS = $iF | Where-Object { $_.pkt_size } | ForEach-Object { [int]$_.pkt_size }
    $pS = $pF | Where-Object { $_.pkt_size } | ForEach-Object { [int]$_.pkt_size }
    $bS = $bF | Where-Object { $_.pkt_size } | ForEach-Object { [int]$_.pkt_size }
    if ($iS.Count -gt 0) {
        $rpt.AppendLine("`n    Frame Sizes:") | Out-Null
        Write-Field $rpt "I avg/max" "$(Format-Size ([long]($iS|Measure-Object -Average).Average)) / $(Format-Size ([long]($iS|Measure-Object -Maximum).Maximum))" -Indent 8 -Width 20
        if ($pS.Count -gt 0) { Write-Field $rpt "P avg/max" "$(Format-Size ([long]($pS|Measure-Object -Average).Average)) / $(Format-Size ([long]($pS|Measure-Object -Maximum).Maximum))" -Indent 8 -Width 20 }
        if ($bS.Count -gt 0) { Write-Field $rpt "B avg/max" "$(Format-Size ([long]($bS|Measure-Object -Average).Average)) / $(Format-Size ([long]($bS|Measure-Object -Maximum).Maximum))" -Indent 8 -Width 20 }
        if ($iS.Count -gt 0 -and ($pS.Count -gt 0 -or $bS.Count -gt 0)) {
            $base = if ($bS.Count -gt 0) { ($bS|Measure-Object -Average).Average } else { ($pS|Measure-Object -Average).Average }
            if ($base -gt 0) { Write-Field $rpt "I:P:B ratio" "$([math]::Round(($iS|Measure-Object -Average).Average/$base,1)):$([math]::Round(($pS|Measure-Object -Average).Average/$base,1)):$(if($bS.Count -gt 0){'1.0'}else{'N/A'})" -Indent 8 -Width 20 }
        }
    }

    # GOP patterns (first 4)
    $allPat = ($fr | ForEach-Object { $_.pict_type }) -join ''
    $rpt.AppendLine("`n    First GOP Patterns:") | Out-Null
    $gs = 0; $gn = 0
    for ($i = 0; $i -lt $fr.Count -and $gn -lt 4; $i++) {
        if ($fr[$i].key_frame -eq 1 -and $i -gt 0) {
            $rpt.AppendLine("      GOP#$($gn+1): $(($fr[$gs..($i-1)]|ForEach-Object{$_.pict_type}) -join '')") | Out-Null
            $gs = $i; $gn++
        }
    }

    # B-frame structure
    $bM = [regex]::Matches($allPat, 'B+')
    if ($bM.Count -gt 0) {
        $bLens = $bM | ForEach-Object { $_.Length }
        $maxB  = ($bLens | Measure-Object -Maximum).Maximum
        $modeB = ($bLens | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name

        Write-Field $rpt "Max Consecutive B-Frames" "$maxB (bframes=$maxB)"
        Write-Field $rpt "Typical B-run Length" "$modeB frames"
        $rc['BFrames'] = $maxB; $rc['BRunMode'] = $modeB

        # B-pyramid detection
        # Prefer HM-confirmed result (set by HMAnalyser POC cross-reference check) over heuristic.
        # HM-confirmed: rc['BPyramid'] is boolean, rc['BPyramid_Source'] = 'HM-confirmed'
        # Heuristic fallback: frame size bimodal distribution test (less reliable)
        if ($rc['BPyramid_Source'] -eq 'HM-confirmed') {
            # Already set by Write-HMAnalyserReport — just reflect it in the GOP section
            $bpyrStr = if ($rc['BPyramid']) {
                "CONFIRMED (B-SLICE references B-SLICE in decoded output)"
            } else {
                "NOT DETECTED (no B-SLICE references another B-SLICE)"
            }
            Write-Field $rpt "B-Pyramid" $bpyrStr
        } elseif ($bS -and $bS.Count -gt 10 -and $maxB -ge 3) {
            $bSorted = $bS | Sort-Object
            $bMed    = $bSorted[([math]::Floor($bSorted.Count / 2))]
            $bLarge  = @($bS | Where-Object { $_ -gt ($bMed * 1.5) }).Count
            $bRat    = $bLarge / $bS.Count
            $bpyr = if ($bRat -gt 0.15 -and $bRat -lt 0.6) {
                "Likely ENABLED ($([math]::Round($bRat*100))% large B-frames)"
            } else { "Likely DISABLED (uniform B sizes)" }
            Write-Field $rpt "B-Pyramid" $bpyr
            $rc['BPyramid'] = $bpyr
        }

        # B-run distribution
        $rpt.AppendLine("`n    B-Run Length Distribution:") | Out-Null
        foreach ($g in ($bLens | Group-Object | Sort-Object { [int]$_.Name })) {
            Write-Field $rpt "  $($g.Name)B runs" "$($g.Count) ($([math]::Round($g.Count/$bM.Count*100,1))%)" -Indent 8 -Width 16
        }
    }

    # Interlacing
    $intF = @($fr | Where-Object { $_.interlaced_frame -eq 1 })
    if ($intF.Count -gt 0) {
        $tff = @($intF | Where-Object { $_.top_field_first -eq 1 }).Count
        Write-Field $rpt "Interlaced" "$($intF.Count)/$($fr.Count) frames (TFF:$tff BFF:$($intF.Count-$tff))"
        $rc['Interlaced'] = $true; $rc['FieldOrder'] = if ($tff -gt ($intF.Count-$tff)) { "TFF" } else { "BFF" }
    }
}

function Write-QPReport {
    param([System.Text.StringBuilder]$rpt, $qpData, $rc)
    if (-not $qpData -or -not $qpData.Stats) { return }
    Write-Section $rpt "QUANTIZATION PARAMETER (QP) ANALYSIS"

    # Show data source
    if ($qpData.Source) { Write-Field $rpt "Data Source" $qpData.Source }
    if ($qpData.FrameCount) { Write-Field $rpt "Frames Decoded" $qpData.FrameCount }

    $s = $qpData.Stats
    Write-Field $rpt "QP Samples"   $s.Count
    Write-Field $rpt "QP Range"     "$($s.Min) - $($s.Max)"
    Write-Field $rpt "QP Average"   $s.Avg
    Write-Field $rpt "QP Median"    $s.Median
    Write-Field $rpt "QP Std Dev"   $s.StdDev

    $rc['QP_Avg'] = $s.Avg; $rc['QP_Min'] = $s.Min; $rc['QP_Max'] = $s.Max; $rc['QP_SD'] = $s.StdDev

    foreach ($t in 'I','P','B') {
        if ($qpData.ByType[$t]) {
            $v = $qpData.ByType[$t]
            Write-Field $rpt "$t-Frame QP" "avg=$($v.Avg) min=$($v.Min) max=$($v.Max) (n=$($v.Count))"
            $rc["QP_$t"] = $v.Avg
            $rc["QP_${t}_Min"] = $v.Min
            $rc["QP_${t}_Max"] = $v.Max
        }
    }

    # Reference structure (from HM decoder)
    if ($qpData.RefStructure) {
        $ref = $qpData.RefStructure
        $rpt.AppendLine("`n    Reference Structure (from HM decoder):") | Out-Null
        Write-Field $rpt "Max L0 References" $ref.MaxL0Refs -Indent 8 -Width 28
        Write-Field $rpt "Max L1 References" $ref.MaxL1Refs -Indent 8 -Width 28
        Write-Field $rpt "Avg L0 References" $ref.AvgL0Refs -Indent 8 -Width 28
        Write-Field $rpt "Avg L1 References" $ref.AvgL1Refs -Indent 8 -Width 28
        Write-Field $rpt "Temporal Layers" $ref.TemporalLayers -Indent 8 -Width 28
        Write-Field $rpt "Max Temporal ID" $ref.MaxTId -Indent 8 -Width 28
        $rc['MaxL0Refs'] = $ref.MaxL0Refs
        $rc['MaxL1Refs'] = $ref.MaxL1Refs
        $rc['TemporalLayers'] = $ref.TemporalLayers
    }

    # Rate control estimation
    $rpt.AppendLine("`n    Rate Control Estimation:") | Out-Null
    $qR  = $s.Max - $s.Min
    $qCV = if ($s.Avg -gt 0) { $s.StdDev / $s.Avg * 100 } else { 0 }

    # Special case: very low QP (disc authoring, near-lossless)
    $rcMode = if ($s.Max -le 5 -and $s.Avg -le 3)           { "CQP/Fixed QP (near-lossless, QP ~$($s.Median))" }
              elseif ($qR -le 2 -and $qCV -lt 5)             { "CQP (Constant QP = $($s.Median))" }
              elseif ($qR -le 8 -and $qCV -lt 15)            { "CRF (estimated ~$($s.Avg))" }
              elseif ($qCV -gt 25)                            { "VBR/ABR (wide QP range)" }
              else                                            { "Constrained VBR or CRF + VBV" }
    Write-Field $rpt "Likely RC Mode" $rcMode -Indent 8 -Width 28
    $rc['RateControl'] = $rcMode

    $aqEst = if ($qCV -gt 10) { "AQ likely enabled" } elseif ($qCV -gt 5) { "Moderate AQ" } else { "AQ minimal/disabled" }
    Write-Field $rpt "AQ Estimate" "$([math]::Round($qCV,1))% CV - $aqEst" -Indent 8 -Width 28
}

function Write-HMAnalyserReport {
    param([System.Text.StringBuilder]$rpt, $hmData, $rc)
    if (-not $hmData) { return }

    Write-Section $rpt "HM ANALYSER - DETAILED BITSTREAM ANALYSIS"
    $rpt.AppendLine("    Source: $($hmData.Source)") | Out-Null

    # SPS parameters
    if ($hmData.SPS) {
        $s = $hmData.SPS
        $rpt.AppendLine("`n    [SPS - Sequence Parameters (exact values)]") | Out-Null
        Write-Field $rpt "Profile IDC"               $s.general_profile_idc -Indent 8 -Width 36
        Write-Field $rpt "Tier"                       $(if($null -ne $s.general_tier_flag){if($s.general_tier_flag){"High"}else{"Main"}}else{$null}) -Indent 8 -Width 36
        Write-Field $rpt "Level IDC"                  $(if($s.general_level_idc){"$($s.general_level_idc) (Level $([math]::Round($s.general_level_idc/30,1)))"}else{$null}) -Indent 8 -Width 36
        Write-Field $rpt "Resolution"                 "$($s.pic_width)x$($s.pic_height)" -Indent 8 -Width 36
        Write-Field $rpt "Chroma Format"              $(switch($s.chroma_format_idc){0{"Mono"}1{"4:2:0"}2{"4:2:2"}3{"4:4:4"}default{$s.chroma_format_idc}}) -Indent 8 -Width 36
        Write-Field $rpt "Bit Depth (Luma/Chroma)"    $(if($s.bit_depth_luma){"$($s.bit_depth_luma)/$($s.bit_depth_chroma)"}else{$null}) -Indent 8 -Width 36
        if ($s.log2_max_cu) {
            $minCU = if ($s.log2_min_cu) { [math]::Pow(2, $s.log2_min_cu) } else { "?" }
            $maxCTU = [math]::Pow(2, $s.log2_max_cu)
            Write-Field $rpt "CU Size Range"          "${minCU}x${minCU} to ${maxCTU}x${maxCTU}" -Indent 8 -Width 36
            $rc['CTU'] = $maxCTU
        }
        if ($s.log2_max_tu) {
            $minTU = if ($s.log2_min_tu) { [math]::Pow(2, $s.log2_min_tu) } else { "?" }
            $maxTU = [math]::Pow(2, $s.log2_max_tu)
            Write-Field $rpt "TU Size Range"          "${minTU}x${minTU} to ${maxTU}x${maxTU}" -Indent 8 -Width 36
        }
        Write-Field $rpt "TU Depth Inter/Intra"       $(if($null -ne $s.tu_depth_inter){"$($s.tu_depth_inter)/$($s.tu_depth_intra)"}else{$null}) -Indent 8 -Width 36
        Write-Field $rpt "AMP"                         (BoolStr $s.amp_enabled) -Indent 8 -Width 36
        Write-Field $rpt "SAO"                         (BoolStr $s.sao_enabled) -Indent 8 -Width 36
        Write-Field $rpt "PCM"                         (BoolStr $s.pcm_enabled) -Indent 8 -Width 36
        Write-Field $rpt "Strong Intra Smoothing"      (BoolStr $s.strong_intra_smoothing) -Indent 8 -Width 36
        Write-Field $rpt "Temporal MVP"                (BoolStr $s.temporal_mvp) -Indent 8 -Width 36
        Write-Field $rpt "Scaling List"                (BoolStr $s.scaling_list) -Indent 8 -Width 36
        Write-Field $rpt "DPB Size"                    $s.dpb_size -Indent 8 -Width 36
        Write-Field $rpt "Max Reorder Pics"            $s.max_reorder_pics -Indent 8 -Width 36
        # Collect for reconstruction from SPS
        if ($null -ne $s.sao_enabled) { $rc['SAO'] = $s.sao_enabled }
        if ($null -ne $s.amp_enabled) { $rc['AMP'] = $s.amp_enabled }
        if ($null -ne $s.strong_intra_smoothing) { $rc['StrongIntra'] = $s.strong_intra_smoothing }
        if ($null -ne $s.temporal_mvp) { $rc['TemporalMVP'] = $s.temporal_mvp }
        if ($null -ne $s.scaling_list) { $rc['ScalingList'] = $s.scaling_list }
        if ($s.log2_max_cu) { $rc['CTU'] = [math]::Pow(2, $s.log2_max_cu) }
        if ($null -ne $s.tu_depth_inter) { $rc['TUDepthInter'] = $s.tu_depth_inter; $rc['TUDepthIntra'] = $s.tu_depth_intra }
        if ($s.dpb_size) { $rc['DPBSize'] = $s.dpb_size }
        if ($s.max_reorder_pics) { $rc['ReorderPics'] = $s.max_reorder_pics }
    }

    # PPS parameters
    if ($hmData.PPS) {
        $p = $hmData.PPS
        $rpt.AppendLine("`n    [PPS - Picture Parameters (exact values)]") | Out-Null
        Write-Field $rpt "CU QP Delta"                $(if($p.cu_qp_delta){"Enabled (depth=$($p.cu_qp_delta_depth))"}else{(BoolStr $p.cu_qp_delta)}) -Indent 8 -Width 36
        Write-Field $rpt "Cb QP Offset"                $p.cb_qp_offset -Indent 8 -Width 36
        Write-Field $rpt "Cr QP Offset"                $p.cr_qp_offset -Indent 8 -Width 36
        Write-Field $rpt "Init QP"                     $p.init_qp -Indent 8 -Width 36
        Write-Field $rpt "Sign Data Hiding"            (BoolStr $p.sign_data_hiding) -Indent 8 -Width 36
        Write-Field $rpt "Transform Skip"              (BoolStr $p.transform_skip) -Indent 8 -Width 36
        Write-Field $rpt "Constrained Intra"           (BoolStr $p.constrained_intra) -Indent 8 -Width 36
        Write-Field $rpt "Tiles"                       (BoolStr $p.tiles) -Indent 8 -Width 36
        Write-Field $rpt "WPP (Entropy Coding Sync)"   (BoolStr $p.wpp) -Indent 8 -Width 36
        Write-Field $rpt "Weighted Prediction"         (BoolStr $p.weighted_pred) -Indent 8 -Width 36
        Write-Field $rpt "Weighted Biprediction"       (BoolStr $p.weighted_bipred) -Indent 8 -Width 36
        Write-Field $rpt "Transquant Bypass"           (BoolStr $p.transquant_bypass) -Indent 8 -Width 36
        Write-Field $rpt "Loop Filter Across Slices"   (BoolStr $p.loop_filter_across) -Indent 8 -Width 36
        Write-Field $rpt "Deblocking Control Present"  (BoolStr $p.deblock_control) -Indent 8 -Width 36

        # Collect for reconstruction from PPS
        if ($null -ne $p.cb_qp_offset) { $rc['CbQPOffset'] = $p.cb_qp_offset; $rc['CrQPOffset'] = $p.cr_qp_offset }
        if ($null -ne $p.weighted_pred) { $rc['WeightedPred'] = $p.weighted_pred }
        if ($null -ne $p.weighted_bipred) { $rc['WeightedBipred'] = $p.weighted_bipred }
        if ($null -ne $p.sign_data_hiding) { $rc['SignDataHiding'] = $p.sign_data_hiding }
        if ($null -ne $p.wpp) { $rc['WPP'] = $p.wpp }
        if ($null -ne $p.tiles) { $rc['Tiles'] = $p.tiles }
        if ($null -ne $p.cu_qp_delta) { $rc['CUQPDelta'] = $p.cu_qp_delta }
        if ($null -ne $p.transform_skip) { $rc['TransformSkip'] = $p.transform_skip }
        if ($null -ne $p.constrained_intra) { $rc['ConstrainedIntra'] = $p.constrained_intra }
        if ($null -ne $p.deblock_control) { $rc['DeblockControl'] = $p.deblock_control }
    }

    # Slice-level
    if ($hmData.Slice) {
        $sl = $hmData.Slice
        $rpt.AppendLine("`n    [Slice-Level Parameters]") | Out-Null
        Write-Field $rpt "Slices Analyzed"             $sl.count -Indent 8 -Width 36
        Write-Field $rpt "Max Merge Candidates"        $sl.max_merge -Indent 8 -Width 36
        Write-Field $rpt "Temporal MVP per Slice"      (BoolStr $sl.temporal_mvp) -Indent 8 -Width 36
        Write-Field $rpt "SAO Luma per Slice"          (BoolStr $sl.sao_luma) -Indent 8 -Width 36
        Write-Field $rpt "SAO Chroma per Slice"        (BoolStr $sl.sao_chroma) -Indent 8 -Width 36
        Write-Field $rpt "WPP Entry Points/Slice"      $sl.entry_points -Indent 8 -Width 36
        if ($sl.max_merge) { $rc['MaxMerge'] = $sl.max_merge }
    }

    # B-pyramid — write to rc from HM QP data so GOP section and Reconstruction can use it
    if ($hmData.QP -and $null -ne $hmData.QP.BPyramid) {
        $rc['BPyramid']        = $hmData.QP.BPyramid
        $rc['BPyramid_Source'] = $hmData.QP.BPyramid_Source   # 'HM-confirmed'
        $rc['BPyramid_BRefs']  = $hmData.QP.BPyramid_BRefs    # count of cross-B references

        # Write a dedicated [B-Pyramid Detection] sub-section in the HM report
        $rpt.AppendLine("`n    [B-Pyramid Detection (from decoded reference lists)]") | Out-Null
        if ($hmData.QP.BPyramid) {
            Write-Field $rpt "B-Pyramid" "CONFIRMED — $($hmData.QP.BPyramid_BRefs) B-SLICE(s) reference another B-SLICE in decoded output" -Indent 8 -Width 36
        } else {
            Write-Field $rpt "B-Pyramid" "NOT DETECTED — no B-SLICE references another B-SLICE in decoded output" -Indent 8 -Width 36
        }
    }
    if ($hmData.VUI) {
        $v = $hmData.VUI
        $cpNames = @{1='BT.709';4='BT.470M';5='BT.601';9='BT.2020';10='BT.2020-10';11='BT.2020-12'}
        $tcNames = @{1='BT.709';4='BT.470M';6='BT.601';14='BT.2020-10';15='BT.2020-12';16='SMPTE ST 2084 (PQ)';18='ARIB STD-B67 (HLG)'}
        $rpt.AppendLine("`n    [VUI - Color / Timing (exact values from bitstream)]") | Out-Null
        Write-Field $rpt "Colour Primaries"            $(if($v.colour_primaries){"$($v.colour_primaries) ($($cpNames[[int]$v.colour_primaries]))"}else{$null}) -Indent 8 -Width 36
        Write-Field $rpt "Transfer Characteristics"    $(if($v.transfer){"$($v.transfer) ($($tcNames[[int]$v.transfer]))"}else{$null}) -Indent 8 -Width 36
        Write-Field $rpt "Matrix Coefficients"         $(if($v.matrix_coeffs){"$($v.matrix_coeffs) ($($cpNames[[int]$v.matrix_coeffs]))"}else{$null}) -Indent 8 -Width 36
        Write-Field $rpt "Video Full Range"            (BoolStr $v.full_range) -Indent 8 -Width 36
        Write-Field $rpt "Frame Rate (from VUI)"       $(if($v.fps){"$($v.fps) fps (tick=$($v.tick) scale=$($v.timescale))"}else{$null}) -Indent 8 -Width 36
    }

    # HDR metadata
    if ($hmData.HDR) {
        $h = $hmData.HDR
        $rpt.AppendLine("`n    [HDR Metadata (exact values from SEI)]") | Out-Null
        Write-Field $rpt "MaxCLL"                      $(if($h.MaxCLL){"$($h.MaxCLL) nits"}else{$null}) -Indent 8 -Width 36
        Write-Field $rpt "MaxFALL"                     $(if($h.MaxFALL){"$($h.MaxFALL) nits"}else{$null}) -Indent 8 -Width 36
        Write-Field $rpt "Mastering Max Luminance"     $(if($h.max_lum){"$($h.max_lum) cd/m2"}else{$null}) -Indent 8 -Width 36
        Write-Field $rpt "Mastering Min Luminance"     $(if($h.min_lum){"$($h.min_lum) cd/m2"}else{$null}) -Indent 8 -Width 36
        if ($h['px0']) {
            # Convert from SEI units (1/50000) to CIE chromaticity coordinates
            # HEVC SEI order: [0]=Green, [1]=Blue, [2]=Red
            $fmtP = { param($v) if ($v) { [math]::Round($v / 50000, 4) } else { "?" } }
            $rpt.AppendLine("        Display Primaries (x,y): G($(& $fmtP $h.px0),$(& $fmtP $h.py0)) B($(& $fmtP $h.px1),$(& $fmtP $h.py1)) R($(& $fmtP $h.px2),$(& $fmtP $h.py2))") | Out-Null
            $rpt.AppendLine("        White Point (x,y): ($(& $fmtP $h.wpx),$(& $fmtP $h.wpy))") | Out-Null
        }
        # Store HM SEI HDR data in rc (don't overwrite existing values with null)
        if ($h.MaxCLL)     { $rc['MaxCLL'] = $h.MaxCLL; $rc['MaxCLL_Source'] = 'SEI' }
        if ($h.MaxFALL)    { $rc['MaxFALL'] = $h.MaxFALL; $rc['MaxFALL_Source'] = 'SEI' }
        if ($h.max_lum)    { $rc['MasterMaxLum'] = $h.max_lum; $rc['MasterLum_Source'] = 'SEI' }
        if ($h.min_lum)    { $rc['MasterMinLum'] = $h.min_lum }
        # Store display primaries in rc (raw SEI units /50000 for CIE)
        if ($h['px0']) {
            for ($pi = 0; $pi -lt 3; $pi++) {
                $rc["dp_x$pi"] = $h["px$pi"]; $rc["dp_y$pi"] = $h["py$pi"]
            }
            $rc['wp_x'] = $h.wpx; $rc['wp_y'] = $h.wpy
            $rc['DisplayPrimaries_Source'] = 'SEI'
        }
    }

    # CABAC encoding analysis
    if ($hmData.Analysis -and $hmData.Analysis.TotalBits) {
        $a = $hmData.Analysis
        $rpt.AppendLine("`n    [Encoding Tool Analysis (from CABAC statistics)]") | Out-Null
        Write-Field $rpt "Total Coded Bits"            "$($a.TotalBits) ($([math]::Round($a.TotalBits/8/1024,1)) KB)" -Indent 8 -Width 36
        if ($a.SkipRatio)    { Write-Field $rpt "Skip Mode Usage"      "$($a.SkipRatio)% of CUs" -Indent 8 -Width 36 }
        if ($a.MergeCount)   { Write-Field $rpt "Merge Mode Count"     "$($a.MergeCount) CUs" -Indent 8 -Width 36 }
        if ($a.SAOPercent)   { Write-Field $rpt "SAO Overhead"         "$($a.SAOPercent)% of total bits" -Indent 8 -Width 36 }
        if ($a.IntraPercent) { Write-Field $rpt "Intra Direction"      "$($a.IntraPercent)% of total bits" -Indent 8 -Width 36 }
        if ($a.MVDPercent)   { Write-Field $rpt "Motion Vectors"       "$($a.MVDPercent)% of total bits" -Indent 8 -Width 36 }
    }

    # SEI messages
    if ($hmData.SEIMessages) {
        $rpt.AppendLine("`n    [Decoded SEI Messages]") | Out-Null
        foreach ($line in ($hmData.SEIMessages -split "`n" | Select-Object -First 30)) {
            if ($line.Trim()) { $rpt.AppendLine("        $($line.Trim())") | Out-Null }
        }
    }
}

function Write-JMAnalyserReport {
    param([System.Text.StringBuilder]$rpt, $jmData, $rc)
    if (-not $jmData) { return }

    Write-Section $rpt "JM ANALYSER - H.264 BITSTREAM ANALYSIS"
    $rpt.AppendLine("    Source: $($jmData.Source)") | Out-Null

    # SPS fields from FFprobe on raw Annex B stream
    if ($jmData.SPS) {
        $s = $jmData.SPS
        $rpt.AppendLine("`n    [SPS - Sequence Parameters (from Annex B bitstream)]") | Out-Null
        Write-Field $rpt "Profile"               $s['profile']      -Indent 8 -Width 36
        $lvl = $s['level']
        Write-Field $rpt "Level"                 $(if($lvl){"$lvl (Level $([math]::Round($lvl/10,1)))"}else{$null}) -Indent 8 -Width 36
        Write-Field $rpt "Max Ref Frames"        $s['refs']         -Indent 8 -Width 36
        Write-Field $rpt "Pixel Format"          $s['pix_fmt']      -Indent 8 -Width 36
        Write-Field $rpt "Resolution"            $(if($s['width']){"$($s['width'])x$($s['height'])"}else{$null}) -Indent 8 -Width 36
        Write-Field $rpt "Has B-Frames (SPS)"    $s['has_b_frames'] -Indent 8 -Width 36
        # Write MediaInfo-confirmed parameters
        $rpt.AppendLine("`n    [PPS / Format Settings (from MediaInfo)]") | Out-Null
        # These are set in rc by Write-MediaInfoReport when it runs after JM analysis
        $miRefs = if ($rc['Refs_Source'] -eq 'MediaInfo') { $rc['Refs'] } else { $null }
        $miCabac = $rc['CABAC']
        $miCabacSrc = $rc['CABAC_Source']
        if ($null -ne $miRefs -or $null -ne $miCabac) {
            Write-Field $rpt "CABAC"              $(if($null -ne $miCabac){if($miCabac){'Yes'}else{'No'}}else{'(pending MediaInfo)'}) -Indent 8 -Width 36
            Write-Field $rpt "Ref Frames (MI)"   $(if($miRefs){"$miRefs frames"}else{'(pending MediaInfo)'}) -Indent 8 -Width 36
        } else {
            $rpt.AppendLine("        (MediaInfo runs after JM — see MediaInfo section below for CABAC / ref frame confirmation)") | Out-Null
        }
    }

    # QP analysis
    if ($jmData.QP) {
        $q = $jmData.QP
        $s = $q.Stats
        $rpt.AppendLine("`n    [QP Analysis ($($q.FrameCount) frames decoded)]") | Out-Null
        Write-Field $rpt "QP Range"              "$($s.Min) - $($s.Max)"   -Indent 8 -Width 36
        Write-Field $rpt "QP Average"            $s.Avg                    -Indent 8 -Width 36
        Write-Field $rpt "QP Median"             $s.Median                 -Indent 8 -Width 36
        Write-Field $rpt "QP Std Dev"            $s.StdDev                 -Indent 8 -Width 36

        $rpt.AppendLine("`n    [QP by Frame Type]") | Out-Null
        # Keys match JMAnalyser.ps1: Bref=uppercase B (B-pyramid ref), Bleaf=lowercase b (leaf B)
        $typeDisplay = [ordered]@{
            IDR   = 'IDR (keyframe)'
            I     = 'I-frame'
            P     = 'P-frame'
            Bref  = 'B-frame (pyramid ref, uppercase B)'
            Bleaf = 'b-frame (leaf B, lowercase b)'
        }
        foreach ($t in $typeDisplay.Keys) {
            if ($q.ByType[$t] -and $q.ByType[$t].Count -gt 0) {
                $v = $q.ByType[$t]
                Write-Field $rpt $typeDisplay[$t] "avg=$($v.Avg) min=$($v.Min) max=$($v.Max) (n=$($v.Count))" -Indent 8 -Width 40
            }
        }

        $rpt.AppendLine("`n    [B-Pyramid Detection]") | Out-Null
        if ($q.BPyramid) {
            Write-Field $rpt "B-Pyramid"  "CONFIRMED — uppercase B-ref frames observed in decode output" -Indent 8 -Width 36
        } else {
            Write-Field $rpt "B-Pyramid"  "NOT DETECTED — no B-reference frames in decoded output"       -Indent 8 -Width 36
        }
        # Store definitive result in rc (replaces heuristic)
        $rc['BPyramid']        = $q.BPyramid
        $rc['BPyramid_Source'] = 'JM-confirmed'
        $rc['QP_Avg']  = $s.Avg; $rc['QP_Min'] = $s.Min; $rc['QP_Max'] = $s.Max; $rc['QP_SD'] = $s.StdDev
        if ($q.ByType['IDR'])   { $rc['QP_I'] = $q.ByType['IDR'].Avg }
        elseif ($q.ByType['I']) { $rc['QP_I'] = $q.ByType['I'].Avg }
        if ($q.ByType['P'])     { $rc['QP_P'] = $q.ByType['P'].Avg }
        if ($q.ByType['Bref'])  { $rc['QP_B'] = $q.ByType['Bref'].Avg }
    }
}

function Write-NALReport {
    param([System.Text.StringBuilder]$rpt, $nalData, $rc)
    if (-not $nalData) { return }
    Write-Section $rpt "BITSTREAM PARAMETERS (SPS / PPS / NAL)"

    if ($nalData.NALTypes.Count -gt 0) {
        $rpt.AppendLine("    NAL Unit Types:") | Out-Null
        foreach ($n in ($nalData.NALTypes.GetEnumerator() | Sort-Object Value -Descending)) {
            Write-Field $rpt $n.Key "$($n.Value) units" -Indent 8 -Width 40
        }
    }

    $rpt.AppendLine("`n    SPS (Sequence Parameters):") | Out-Null
    Write-Field $rpt "Profile IDC"            $nalData.ProfileIDC                                                              -Indent 8 -Width 36
    Write-Field $rpt "Level IDC"              $(if($nalData.LevelIDC){"$($nalData.LevelIDC) (=$([math]::Round($nalData.LevelIDC/30,1)))"}else{$null}) -Indent 8 -Width 36
    Write-Field $rpt "Tier"                   $(if($null -ne $nalData.Tier){if($nalData.Tier){"High"}else{"Main"}}else{$null}) -Indent 8 -Width 36
    Write-Field $rpt "Chroma Format"          $(switch($nalData.Chroma){0{"Mono"}1{"4:2:0"}2{"4:2:2"}3{"4:4:4"}default{$nalData.Chroma}}) -Indent 8 -Width 36
    Write-Field $rpt "Bit Depth (L/C)"        $(if($nalData.BitDepthLuma){"$($nalData.BitDepthLuma)/$($nalData.BitDepthChroma)"}else{$null}) -Indent 8 -Width 36
    Write-Field $rpt "Min CU"                 $(if($nalData.MinCU){"$($nalData.MinCU)x$($nalData.MinCU)"}else{$null})         -Indent 8 -Width 36
    Write-Field $rpt "CU Depth"               $nalData.CUDepth                                                                 -Indent 8 -Width 36
    if ($nalData.MinCU -and $nalData.CUDepth) {
        $maxCTU = $nalData.MinCU * [math]::Pow(2, $nalData.CUDepth)
        Write-Field $rpt "Max CTU (derived)"  "${maxCTU}x${maxCTU}"                                                            -Indent 8 -Width 36
        $rc['CTU'] = $maxCTU
    }
    Write-Field $rpt "Min TU"                 $(if($nalData.MinTU){"$($nalData.MinTU)x$($nalData.MinTU)"}else{$null})         -Indent 8 -Width 36
    Write-Field $rpt "TU Depth (Inter/Intra)" $(if($nalData.TUDepthInter){"$($nalData.TUDepthInter)/$($nalData.TUDepthIntra)"}else{$null}) -Indent 8 -Width 36
    Write-Field $rpt "Short-Term Ref Sets"    $nalData.STRPS                -Indent 8 -Width 36
    Write-Field $rpt "Long-Term Ref"          (BoolStr $nalData.LongTermRef) -Indent 8 -Width 36
    Write-Field $rpt "Temporal MVP"           (BoolStr $nalData.TemporalMVP) -Indent 8 -Width 36
    Write-Field $rpt "SAO"                    (BoolStr $nalData.SAO)         -Indent 8 -Width 36
    Write-Field $rpt "AMP"                    (BoolStr $nalData.AMP)         -Indent 8 -Width 36
    Write-Field $rpt "PCM"                    (BoolStr $nalData.PCM)         -Indent 8 -Width 36
    Write-Field $rpt "Strong Intra Smoothing" (BoolStr $nalData.StrongIntra) -Indent 8 -Width 36
    Write-Field $rpt "Scaling List"           (BoolStr $nalData.ScalingList) -Indent 8 -Width 36
    Write-Field $rpt "DPB Size"               $nalData.DPBSize              -Indent 8 -Width 36

    $rpt.AppendLine("`n    PPS (Picture Parameters):") | Out-Null
    Write-Field $rpt "Transform Skip"         (BoolStr $nalData.TransformSkip)    -Indent 8 -Width 36
    Write-Field $rpt "CU QP Delta"            $(if($null -ne $nalData.CUQPDelta){if($nalData.CUQPDelta){"Enabled (depth=$($nalData.CUQPDeltaDepth))"}else{"Disabled"}}else{$null}) -Indent 8 -Width 36
    Write-Field $rpt "Weighted Prediction"    (BoolStr $nalData.WeightedPred)     -Indent 8 -Width 36
    Write-Field $rpt "Weighted Biprediction"  (BoolStr $nalData.WeightedBipred)   -Indent 8 -Width 36
    Write-Field $rpt "Constrained Intra"      (BoolStr $nalData.ConstrainedIntra) -Indent 8 -Width 36
    Write-Field $rpt "Sign Data Hiding"       (BoolStr $nalData.SignDataHiding)   -Indent 8 -Width 36
    Write-Field $rpt "WPP (Wavefront)"        (BoolStr $nalData.WPP)              -Indent 8 -Width 36
    Write-Field $rpt "Tiles"                  $(if($null -ne $nalData.Tiles){if($nalData.Tiles){"$($nalData.TileCols)x$($nalData.TileRows)"}else{"Disabled"}}else{$null}) -Indent 8 -Width 36
    Write-Field $rpt "Deblocking"             $(if($null -ne $nalData.DeblockDisabled){if($nalData.DeblockDisabled){"Disabled"}else{"Enabled (beta=$($nalData.DeblockBeta) tc=$($nalData.DeblockTC))"}}else{$null}) -Indent 8 -Width 36
    Write-Field $rpt "Ref Idx L0/L1"          $(if($nalData.RefL0){"$($nalData.RefL0)/$($nalData.RefL1)"}else{$null}) -Indent 8 -Width 36

    # Collect for reconstruction
    foreach ($k in @('SAO','AMP','TemporalMVP','StrongIntra','WeightedPred','WeightedBipred','TransformSkip','SignDataHiding','WPP','Tiles','CUQPDelta')) {
        if ($null -ne $nalData.$k) { $rc[$k] = $nalData.$k }
    }
    if ($null -ne $nalData.DeblockDisabled) { $rc['Deblock'] = (-not $nalData.DeblockDisabled); $rc['DeblockBeta'] = $nalData.DeblockBeta; $rc['DeblockTC'] = $nalData.DeblockTC }
    if ($nalData.TUDepthInter) { $rc['TUDepthInter'] = $nalData.TUDepthInter; $rc['TUDepthIntra'] = $nalData.TUDepthIntra }

    # SEI messages
    if ($nalData.SEI.Count -gt 0) {
        $rpt.AppendLine("`n    SEI Messages:") | Out-Null
        $seiNames = @{ 0='Buffering Period';1='Picture Timing';6='Recovery Point';129='Active Params';130='Decoding Unit';132='Decoded Pic Hash';133='Scalable Nesting';137='Mastering Display';144='Content Light Level';176='3D Ref Displays';177='Dolby Vision RPU';178='DV EL' }
        foreach ($g in ($nalData.SEI | Group-Object | Sort-Object { [int]$_.Name })) {
            $nm = if ($seiNames[[int]$g.Name]) { $seiNames[[int]$g.Name] } else { "Unknown" }
            Write-Field $rpt "Type $($g.Name) ($nm)" "$($g.Count)x" -Indent 8 -Width 40
        }
    }
}

function Write-EncoderReport {
    param([System.Text.StringBuilder]$rpt, [string]$encInfo, $rc)
    if (-not $encInfo) { return }
    Write-Section $rpt "ENCODER DETECTION"

    # Only use actual metadata lines, not ffmpeg capability/library listing
    $metaLines = $encInfo -split "`n" | Where-Object {
        $_ -match '(encoder|creation|writing|muxing|handler|vendor|major_brand)\s*[:=]' -and
        $_ -notmatch 'configuration:|libav|built with|^\s*$|^\[|^Input|^Output|^Stream|^Duration|^video:|^audio:|^frame='
    }

    foreach ($l in $metaLines) { $t = $l.Trim(); if ($t) { $rpt.AppendLine("    $t") | Out-Null } }

    # Only match encoder patterns against metadata lines, not full ffmpeg output
    $metaText = $metaLines -join "`n"
    $hints = @()
    $patterns = [ordered]@{
        'x265|libx265'='x265 (libx265)'; 'x264|libx264'='x264 (libx264)'; 'nvenc|NVENC|NVEncC|nvencc'='NVIDIA NVENC';
        'amf|AMF'='AMD AMF'; 'qsv|QSV'='Intel QSV'; 'videotoolbox|VideoToolbox'='Apple VideoToolbox';
        'HandBrake|handbrake'='HandBrake'; 'DaVinci|Resolve'='DaVinci Resolve';
        'Adobe|Premiere|Media Encoder'='Adobe'; 'MainConcept'='MainConcept';
        'SVT-HEVC|SVT-AV1'='SVT'; 'av1an|Av1an'='av1an';
        'rav1e'='rav1e'; 'libaom'='libaom (AV1)'; 'libvpx'='libvpx (VP9)';
        'Ateme|TITAN'='Ateme TITAN'; 'Elemental'='AWS Elemental'; 'Harmonic'='Harmonic'
    }
    foreach ($p in $patterns.GetEnumerator()) { if ($metaText -match $p.Key) { $hints += $p.Value } }

    if ($hints.Count -gt 0) {
        Write-Field $rpt "Detected Encoder(s)" ($hints -join ', ')
        $rc['Encoder'] = $hints[0]
    } else {
        $rpt.AppendLine("    No specific encoder identified from metadata") | Out-Null
    }
}

function Write-MediaInfoReport {
    param([System.Text.StringBuilder]$rpt, $miData, $rc)
    if (-not $miData -or -not $miData.Json) { return }
    Write-Section $rpt "MEDIAINFO DETAILS"

    $tracks = if ($miData.Json.media) { $miData.Json.media.track } elseif ($miData.Json.track) { $miData.Json.track } else { $null }
    if (-not $tracks) { return }

    foreach ($tk in $tracks) {
        $tt = $tk.'@type'
        if ($tt -eq 'General') {
            $rpt.AppendLine("    [General]") | Out-Null
            foreach ($f in 'Format','Format_Version','OverallBitRate','OverallBitRate_Mode','Encoded_Application','Encoded_Library','Encoded_Library_Settings') {
                Write-Field $rpt $f $tk.$f -Indent 8 -Width 34
            }
            if ($tk.extra) { Write-Field $rpt "Writing_Application" $tk.extra.Writing_Application -Indent 8 -Width 34 }
        }
        if ($tt -eq 'Video') {
            $rpt.AppendLine("`n    [Video]") | Out-Null
            foreach ($f in 'Format','Format_Profile','Format_Level','Format_Tier','Format_Settings',
                'HDR_Format','HDR_Format_Version','HDR_Format_Profile','HDR_Format_Level','HDR_Format_Compatibility',
                'CodecID','BitRate','BitRate_Mode','BitRate_Nominal','BitRate_Maximum',
                'Width','Height','PixelAspectRatio','DisplayAspectRatio','FrameRate_Mode','FrameRate','FrameCount',
                'ColorSpace','ChromaSubsampling','BitDepth','ScanType','ScanOrder','Compression_Mode',
                'Encoded_Library','Encoded_Library_Settings',
                'colour_range','colour_primaries','transfer_characteristics','matrix_coefficients',
                'MasteringDisplay_ColorPrimaries','MasteringDisplay_Luminance','MaxCLL','MaxFALL') {
                Write-Field $rpt $f $tk.$f -Indent 8 -Width 34
            }
            if ($tk.HDR_Format -match 'Dolby Vision') { $rpt.AppendLine("`n        *** DOLBY VISION ***") | Out-Null }

            # Cache entire video track object so JMAnalyserReport (and others) can read it
            $rc['MI_VideoTrack'] = $tk

            # H.264: MediaInfo reads actual reference list depth from bitstream slice headers,
            # which is the correct value for encoder reproduction. FFprobe/JM return
            # max_num_ref_frames from the SPS which is often set to 1 as a DPB hint.
            # Override rc['Refs'] with the MediaInfo value when available and higher.
            if ($tk.Format_Settings_RefFrames) {
                # Field is "4" or "4 frames" — extract the integer
                if ($tk.Format_Settings_RefFrames -match '(\d+)') {
                    $miRefs = [int]$Matches[1]
                    if ($miRefs -gt [int]$rc['Refs']) {
                        if ($script:DebugMode) { Write-Host "      [DEBUG] Refs overridden: $($rc['Refs']) -> $miRefs (MediaInfo Format_Settings_RefFrames)" -ForegroundColor Magenta }
                        $rc['Refs'] = $miRefs
                        $rc['Refs_Source'] = 'MediaInfo'
                    }
                }
            }

            # CABAC: store in rc for encoder command reconstruction
            if ($tk.Format_Settings_CABAC) {
                $rc['CABAC'] = ($tk.Format_Settings_CABAC -match 'Yes|1|true')
                $rc['CABAC_Source'] = 'MediaInfo'
            }

            # Store MediaInfo HDR metadata in rc as fallback
            if ($tk.MaxCLL -and -not $rc['MaxCLL']) { $rc['MaxCLL'] = $tk.MaxCLL; $rc['MaxCLL_Source'] = 'MediaInfo' }
            if ($tk.MaxFALL -and -not $rc['MaxFALL']) { $rc['MaxFALL'] = $tk.MaxFALL; $rc['MaxFALL_Source'] = 'MediaInfo' }
            if ($tk.MasteringDisplay_Luminance -and -not $rc['MasterMaxLum']) {
                if ($tk.MasteringDisplay_Luminance -match 'max:\s*([\d.]+)') { $rc['MasterMaxLum'] = [double]$Matches[1] }
                if ($tk.MasteringDisplay_Luminance -match 'min:\s*([\d.]+)') { $rc['MasterMinLum'] = [double]$Matches[1] }
            }
            # Parse encoder settings
            $settingsStr = if ($tk.Encoded_Library_Settings) { $tk.Encoded_Library_Settings }
                           elseif ($tracks | Where-Object { $_.'@type' -eq 'General' -and $_.Encoded_Library_Settings }) { ($tracks | Where-Object { $_.'@type' -eq 'General' }).Encoded_Library_Settings }
                           else { $null }
            if ($settingsStr) {
                $parsed = Parse-EncoderSettings $settingsStr
                if ($parsed) {
                    $rpt.AppendLine("`n        Parsed Encoding Settings ($($parsed.Count) params):") | Out-Null
                    foreach ($p in $parsed.GetEnumerator()) {
                        $rpt.AppendLine("            $($p.Key.PadRight(30))= $($p.Value)") | Out-Null
                    }
                    $rc['ParsedSettings'] = $parsed
                }
            }
        }
        if ($tt -eq 'Audio') {
            $rpt.AppendLine("`n    [Audio]") | Out-Null
            foreach ($f in 'Format','Format_Profile','Format_Info','Format_Settings','CodecID','BitRate','BitRate_Mode','Channels','ChannelLayout','SamplingRate','BitDepth','Compression_Mode','Language','Title') {
                Write-Field $rpt $f $tk.$f -Indent 8 -Width 34
            }
            if ($tk.Format -match 'TrueHD' -and $tk.Format_Profile -match 'Atmos') { $rpt.AppendLine("        *** DOLBY ATMOS ***") | Out-Null }
            if ($tk.Format -match 'E-AC-3' -and $tk.Format_Profile -match 'Atmos') { $rpt.AppendLine("        *** DOLBY ATMOS (DD+) ***") | Out-Null }
        }
        if ($tt -eq 'Text') {
            $rpt.AppendLine("`n    [Subtitle]") | Out-Null
            foreach ($f in 'Format','CodecID','Language','Title','Forced') { Write-Field $rpt $f $tk.$f -Indent 8 -Width 34 }
        }
    }
    # Full text
    if ($miData.Text) { Write-Section $rpt "MEDIAINFO FULL TEXT"; $rpt.AppendLine($miData.Text) | Out-Null }
}

function Write-CheckBitrateReport {
    param([System.Text.StringBuilder]$rpt, $cbData, [string]$OutputDir, [string]$BaseName, $rc)
    if (-not $cbData) { return }
    Write-Section $rpt "BITRATE DISTRIBUTION (CheckBitrate)"

    Write-Field $rpt "Data Points"     $cbData.DataPoints
    Write-Field $rpt "Duration"        ("{0:N1}s ({1:hh\:mm\:ss})" -f $cbData.Duration, [TimeSpan]::FromSeconds($cbData.Duration))
    Write-Field $rpt "Avg Bitrate"     (Format-Bitrate $cbData.FinalAvg)
    Write-Field $rpt "Peak"            (Format-Bitrate $cbData.Peak)
    Write-Field $rpt "Min"             (Format-Bitrate $cbData.Min)
    Write-Field $rpt "Median"          (Format-Bitrate $cbData.Median)
    Write-Field $rpt "Std Dev"         ("{0:N0} kbps" -f $cbData.StdDev)
    Write-Field $rpt "P05 / P95"       "$(Format-Bitrate $cbData.P05) / $(Format-Bitrate $cbData.P95)"
    Write-Field $rpt "P99"             (Format-Bitrate $cbData.P99)
    Write-Field $rpt "Peak:Avg Ratio"  ("{0:N2}x" -f ($cbData.Peak / $cbData.FinalAvg))

    $coefV = $cbData.StdDev / $cbData.FinalAvg * 100
    $brMode = if ($coefV -lt 15) { "Likely CBR/Constrained VBR" } elseif ($coefV -lt 50) { "Moderate VBR" } else { "High VBR (scene-adaptive)" }
    Write-Field $rpt "Coef. of Variation" ("{0:N1}% - $brMode" -f $coefV)
    Write-Field $rpt "Peaks Capped (VBV)" $(if ($cbData.PeaksCapped) { "YES - ceiling ~$(Format-Bitrate $cbData.PeakClusterKbps)" } else { "NO - natural distribution" })

    $rc['BitrateAvgKbps'] = $cbData.FinalAvg; $rc['PeakKbps'] = $cbData.Peak; $rc['CoefVar'] = [math]::Round($coefV, 1)
    if ($cbData.PeaksCapped) { $rc['MaxrateLikely'] = [math]::Round($cbData.PeakClusterKbps) }

    # VBV / Maxrate analysis
    $rpt.AppendLine("`n    VBV / Maxrate Analysis:") | Out-Null
    $peakAvgRatio = $cbData.Peak / [math]::Max($cbData.FinalAvg, 1)
    if ($cbData.PeaksCapped) {
        $rc['VBV'] = "Constrained"
        Write-Field $rpt "VBV Buffer" "DETECTED - peaks cluster at ~$(Format-Bitrate $cbData.PeakClusterKbps)" -Indent 8 -Width 32
        Write-Field $rpt "Effective Maxrate" "~$(Format-Bitrate $cbData.PeakClusterKbps)" -Indent 8 -Width 32
        Write-Field $rpt "VBV Ratio" ("{0:N2}x avg bitrate" -f ($cbData.PeakClusterKbps / $cbData.FinalAvg)) -Indent 8 -Width 32
    } else {
        $rc['VBV'] = "Unconstrained"
        Write-Field $rpt "VBV Buffer" "NOT DETECTED - natural peak distribution" -Indent 8 -Width 32
        Write-Field $rpt "Peak/Avg Ratio" ("{0:N2}x" -f $peakAvgRatio) -Indent 8 -Width 32
        # Estimate what maxrate would be needed to reproduce this bitrate profile
        $suggestedMax = [math]::Ceiling($cbData.P99 / 100) * 100
        Write-Field $rpt "Suggested Maxrate" "~$(Format-Bitrate $suggestedMax) (covers P99)" -Indent 8 -Width 32
        Write-Field $rpt "Absolute Peak" (Format-Bitrate $cbData.Peak) -Indent 8 -Width 32
        $rc['SuggestedMaxrate'] = $suggestedMax
    }
    # Sustained rate analysis (10s window)
    if ($cbData.Data.Count -ge 3) {
        $windowSize = [math]::Min(3, [math]::Floor($cbData.Data.Count / 3))  # ~12s window at 4s intervals
        $maxSustained = 0
        for ($i = 0; $i -le ($cbData.Data.Count - $windowSize); $i++) {
            $windowAvg = ($cbData.Data[$i..($i + $windowSize - 1)] | ForEach-Object { $_.BitrateKbps } | Measure-Object -Average).Average
            if ($windowAvg -gt $maxSustained) { $maxSustained = $windowAvg }
        }
        Write-Field $rpt "Max Sustained (~12s)" (Format-Bitrate $maxSustained) -Indent 8 -Width 32
        $rc['MaxSustained12s'] = $maxSustained
    }

    # Histogram
    $rpt.AppendLine("`n    Histogram:") | Out-Null
    foreach ($rng in @(
        @{L="  0-1M";A=0;B=1000},@{L=" 1-3M";A=1000;B=3000},@{L=" 3-5M";A=3000;B=5000},
        @{L=" 5-10M";A=5000;B=10000},@{L="10-20M";A=10000;B=20000},@{L="  20M+";A=20000;B=[double]::MaxValue}
    )) {
        $cnt = @($cbData.Data | Where-Object { $_.BitrateKbps -ge $rng.A -and $_.BitrateKbps -lt $rng.B }).Count
        $pct = [math]::Round($cnt / $cbData.DataPoints * 100, 1)
        $rpt.AppendLine("        $($rng.L) : $('#' * [math]::Floor($pct/2)) $pct% ($cnt)") | Out-Null
    }

    # Peaks & lows
    $rpt.AppendLine("`n    Top 5 Peaks:") | Out-Null
    foreach ($p in ($cbData.Data | Sort-Object BitrateKbps -Descending | Select-Object -First 5)) {
        $rpt.AppendLine("        $([TimeSpan]::FromSeconds($p.TimeSec).ToString('hh\:mm\:ss')) - $(Format-Bitrate $p.BitrateKbps)") | Out-Null
    }
    $rpt.AppendLine("`n    Top 5 Lows:") | Out-Null
    foreach ($p in ($cbData.Data | Sort-Object BitrateKbps | Select-Object -First 5)) {
        $rpt.AppendLine("        $([TimeSpan]::FromSeconds($p.TimeSec).ToString('hh\:mm\:ss')) - $(Format-Bitrate $p.BitrateKbps)") | Out-Null
    }

    $csv = Join-Path $OutputDir "${BaseName}_bitrate.csv"
    $cbData.Data | Export-Csv -Path $csv -NoTypeInformation
    $rpt.AppendLine("`n    CSV exported: $csv") | Out-Null
}

function Write-SceneCutReport {
    param([System.Text.StringBuilder]$rpt, $sceneData, $rc)
    if (-not $sceneData) { return }
    Write-Section $rpt "SCENE CUT ANALYSIS (first $($sceneData.Duration)s)"
    Write-Field $rpt "Scene Cuts Detected"  $sceneData.Count
    Write-Field $rpt "Avg Cut Interval"     "$($sceneData.AvgInterval)s"
    $rc['SceneCutInterval'] = $sceneData.AvgInterval
}
