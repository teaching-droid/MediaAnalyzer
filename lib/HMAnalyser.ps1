# HMAnalyser.ps1 - HEVC Reference Decoder (HM TAppDecoderAnalyser) integration

# Header param lookup helpers (pass $headerParams explicitly to avoid scope issues)
function _HVal  { param($hp,$n) if ($hp[$n]) { $hp[$n].Sum } else { $null } }
function _HCnt  { param($hp,$n) if ($hp[$n]) { $hp[$n].Count } else { 0 } }
function _HAvg  { param($hp,$n) $c = _HCnt $hp $n; if ($c -gt 0) { (_HVal $hp $n) / $c } else { $null } }
function _HBool { param($hp,$n) $a = _HAvg $hp $n; if ($null -ne $a) { [int][math]::Round($a) } else { $null } }

function Get-HMAnalysis {
    param([string]$FilePath, [int]$MaxFrames = 1000)
    if (-not $Tools.TAppDecoder -or -not $Tools.FFmpeg) { return $null }

    $duration = 0
    if ($Tools.FFprobe) {
        $dr = Run-Command $Tools.FFprobe @('-v','quiet','-show_entries','format=duration','-of','csv=p=0',"`"$FilePath`"") -TimeoutSeconds 15
        if ($dr.ExitCode -eq 0 -and $dr.StdOut.Trim() -match '[\d.]+') { $duration = [double]$dr.StdOut.Trim() }
    }

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "MediaAnalyzer_HM_$([System.IO.Path]::GetRandomFileName())"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    try {
        $fps = 24
        # Scale number of segments based on file duration
        $numSegs = if ($duration -gt 1200) { 5 }       # >20 min: 5 segments
                   elseif ($duration -gt 300) { 3 }     # >5 min: 3 segments
                   else { 1 }                            # short: single segment
        $framesPerSeg = [math]::Ceiling($MaxFrames / $numSegs)
        $segDuration = [math]::Ceiling($framesPerSeg / $fps) + 1

        $segments = @()
        if ($numSegs -ge 5 -and $duration -gt ($segDuration * 6)) {
            $segments = @(
                @{ Start = 0; Label = "beginning" },
                @{ Start = [math]::Floor($duration * 0.25) - [math]::Floor($segDuration / 2); Label = "25%" },
                @{ Start = [math]::Floor($duration * 0.50) - [math]::Floor($segDuration / 2); Label = "middle" },
                @{ Start = [math]::Floor($duration * 0.75) - [math]::Floor($segDuration / 2); Label = "75%" },
                @{ Start = [math]::Max(0, [math]::Floor($duration) - $segDuration - 5); Label = "end" }
            )
            Write-Host "      Multi-point sampling: $numSegs x ${segDuration}s segments" -ForegroundColor DarkGray
        } elseif ($numSegs -ge 3 -and $duration -gt ($segDuration * 4)) {
            $segments = @(
                @{ Start = 0; Label = "beginning" },
                @{ Start = [math]::Floor($duration / 2) - [math]::Floor($segDuration / 2); Label = "middle" },
                @{ Start = [math]::Max(0, [math]::Floor($duration) - $segDuration - 5); Label = "end" }
            )
            Write-Host "      Multi-point sampling: 3 x ${segDuration}s segments" -ForegroundColor DarkGray
        } else {
            $segDuration = [math]::Ceiling($MaxFrames / $fps) + 2
            $segments = @( @{ Start = 0; Label = "full" } )
        }

        $allOutput = [System.Text.StringBuilder]::new()
        $headerSegOutput = ""   # Best segment for CAVLC headers (longest = most frames decoded)
        $segIdx = 0
        foreach ($seg in $segments) {
            $segIdx++
            $rawStream = Join-Path $tempDir "seg_${segIdx}.265"

            Write-Host "      [$segIdx/$($segments.Count)] Extracting $($seg.Label) (@$($seg.Start)s, ${segDuration}s)..." -ForegroundColor DarkGray
            $r = Run-Command $Tools.FFmpeg @(
                '-ss', $seg.Start.ToString(), '-i', "`"$FilePath`"",
                '-t', $segDuration.ToString(), '-c:v', 'copy', '-bsf:v', 'hevc_mp4toannexb',
                '-an', '-sn', '-y', "`"$rawStream`""
            ) -TimeoutSeconds 60 -StatusLabel "ffmpeg extract $($seg.Label)"

            if ($r.ExitCode -ne 0 -or -not (Test-Path $rawStream)) { continue }

            # If DV content, strip RPU NALs before HM (HM can't parse DV NAL units)
            if ($Tools.DoviTool) {
                $cleanStream = $rawStream -replace '\.265$','_clean.265'
                $dvR = Run-Command $Tools.DoviTool @('remove','-i',"`"$rawStream`"",'-o',"`"$cleanStream`"") -TimeoutSeconds 60
                if ($dvR.ExitCode -eq 0 -and (Test-Path $cleanStream)) {
                    $origSize  = (Get-Item $rawStream).Length
                    $cleanSize = (Get-Item $cleanStream).Length
                    $rpuKB = [math]::Round([math]::Abs($origSize - $cleanSize) / 1KB)
                    Write-Host "      dovi_tool: stripped ${rpuKB}KB RPU data" -ForegroundColor DarkGray
                    # Point directly at clean file — avoids Remove+Rename race condition
                    # where dovi_tool briefly holds a file lock on the original stream
                    $rawStream = $cleanStream
                } else {
                    Write-Host "      dovi_tool: no RPU data found (exit=$($dvR.ExitCode))" -ForegroundColor DarkGray
                }
            }

            Write-Host "      [$segIdx/$($segments.Count)] Decoding $($seg.Label)..." -ForegroundColor DarkGray
            $seiArg = @()
            if ($segIdx -eq 1) {
                $seiFile = Join-Path $tempDir "sei.txt"
                # HM uses --key=value syntax; path without extra quotes
                $seiArg = @("--OutputDecodedSEIMessagesFilename=$seiFile")
            }
            if ($script:DebugMode) { Write-Host "      [DEBUG] HM input: $rawStream ($(if(Test-Path $rawStream){(Get-Item $rawStream).Length}else{'MISSING'})B)" -ForegroundColor Magenta }
            $r = Run-Command $Tools.TAppDecoder (@('-b', "`"$rawStream`"") + $seiArg) -TimeoutSeconds 300 -StatusLabel "HM Decoder ($($seg.Label))"
            if ($r.StdOut.Length -lt 200 -and $r.StdErr.Length -lt 200) {
                Write-Host "      HM: minimal output ($($r.StdOut.Length + $r.StdErr.Length) chars) — may not have parsed this stream" -ForegroundColor Yellow
                if ($script:DebugMode) {
                    Write-Host "      [DEBUG] HM stdout: $($r.StdOut.Substring(0,[math]::Min(300,$r.StdOut.Length)))" -ForegroundColor Magenta
                    Write-Host "      [DEBUG] HM stderr: $($r.StdErr.Substring(0,[math]::Min(300,$r.StdErr.Length)))" -ForegroundColor Magenta
                }
            }

            $segOut = "$($r.StdOut)`n$($r.StdErr)"
            [void]$allOutput.AppendLine($segOut)

            # Pick segment with most output for CAVLC headers (avoids empty short segments)
            if ($segOut.Length -gt $headerSegOutput.Length) {
                $headerSegOutput = $segOut
                if ($script:DebugMode) { Write-Host "      [DEBUG] Best header segment: $($seg.Label) ($($segOut.Length) chars)" -ForegroundColor Magenta }
            }
        }

        $output = $allOutput.ToString()
        if ([string]::IsNullOrWhiteSpace($output)) { return $null }

        $result = @{ Source = "HM Reference Decoder Analyser v18.0" }

        # 1. POC/QP/Reference parsing
        $allQPs   = [System.Collections.ArrayList]::new()
        $frameQPs = [System.Collections.ArrayList]::new()
        $refInfo  = [System.Collections.ArrayList]::new()

        foreach ($line in ($output -split "`n")) {
            if ($line -match 'POC\s+(\d+)\s+TId:\s*(\d+)\s*\(\s*(I|P|B)-SLICE,\s*QP\s+(\d+)\s*\)') {
                $poc=[int]$Matches[1]; $tid=[int]$Matches[2]; $type=$Matches[3]; $qp=[int]$Matches[4]
                [void]$allQPs.Add($qp)
                [void]$frameQPs.Add(@{QP=$qp;Type=$type;POC=$poc;TId=$tid})
                $l0=0; $l1=0
                if ($line -match '\[L0\s*([^\]]*)\]') { $l0=@($Matches[1].Trim() -split '\s+'|Where-Object{$_ -match '^\d+$'}).Count }
                if ($line -match '\[L1\s*([^\]]*)\]') { $l1=@($Matches[1].Trim() -split '\s+'|Where-Object{$_ -match '^\d+$'}).Count }
                [void]$refInfo.Add(@{POC=$poc;Type=$type;L0=$l0;L1=$l1;TId=$tid})
                if ($frameQPs.Count -ge $MaxFrames) { break }
            }
        }

        if ($allQPs.Count -gt 0) {
            $result.QP = Build-QPStats $allQPs $frameQPs
            $result.QP.Source = $result.Source
            $result.QP.FrameCount = $frameQPs.Count
            if ($refInfo.Count -gt 0) {
                $result.QP.RefStructure = @{
                    MaxL0Refs = ($refInfo|ForEach-Object{$_.L0}|Measure-Object -Maximum).Maximum
                    MaxL1Refs = ($refInfo|ForEach-Object{$_.L1}|Measure-Object -Maximum).Maximum
                    AvgL0Refs = [math]::Round(($refInfo|ForEach-Object{$_.L0}|Measure-Object -Average).Average,1)
                    AvgL1Refs = [math]::Round(($refInfo|ForEach-Object{$_.L1}|Measure-Object -Average).Average,1)
                    TemporalLayers = ($refInfo|ForEach-Object{$_.TId}|Sort-Object -Unique).Count
                    MaxTId = ($refInfo|ForEach-Object{$_.TId}|Measure-Object -Maximum).Maximum
                }
            }
        }

        # 2. CAVLC Header fields - parse from BEST SEGMENT (most frames decoded)
        # (Header params are constants; using one segment avoids value multiplication)
        $headerParams = [ordered]@{}
        $inHeader = $false
        foreach ($line in ($headerSegOutput -split "`n")) {
            if ($line -match 'CAVLC HEADER BITS') { $inHeader = $true; continue }
            if ($inHeader -and $line -match '^\[CAVLC') { $inHeader = $false; continue }
            if ($inHeader -and $line -match '^-') { continue }
            if ($inHeader -and $line -match '^\s+([\w\[\]\.]+)\s+:') {
                $name = $Matches[1]
                $afterColon = ($line -split ':\s*', 2)[1]
                $nums = [regex]::Matches($afterColon, '-?\d+')
                # CAVLC lines: the numbers after "- -" are: EPCount EPSum EPBits TotalBits (bytes)
                # We want EPCount (=count) and EPSum (=value)
                if ($nums.Count -ge 4) {
                    $headerParams[$name] = @{
                        Count = [int]$nums[0].Value
                        Sum   = [long]$nums[1].Value
                    }
                }
            }
        }
        $result.HeaderParams = $headerParams
        if ($script:DebugMode) {
            Write-Host "      [DEBUG] headerParams count: $($headerParams.Count)" -ForegroundColor Magenta
            if ($headerParams.Count -gt 0) {
                $first3 = ($headerParams.Keys | Select-Object -First 3) -join ', '
                Write-Host "      [DEBUG] First 3 keys: $first3" -ForegroundColor Magenta
            }
        }

        # Extract structured parameters
        $sps = [ordered]@{}; $pps = [ordered]@{}; $slice = [ordered]@{}; $vui = [ordered]@{}; $hdr = [ordered]@{}

        # SPS
        $sps['general_profile_idc']    = _HAvg $headerParams 'general_profile_idc'
        $sps['general_tier_flag']      = _HBool $headerParams 'general_tier_flag'
        $sps['general_level_idc']      = _HAvg $headerParams 'general_level_idc'
        $sps['chroma_format_idc']      = _HAvg $headerParams 'chroma_format_idc'
        $v = _HAvg $headerParams 'bit_depth_luma_minus8'; $sps['bit_depth_luma'] = if ($null -ne $v) { $v + 8 } else { $null }
        $v = _HAvg $headerParams 'bit_depth_chroma_minus8'; $sps['bit_depth_chroma'] = if ($null -ne $v) { $v + 8 } else { $null }
        $sps['pic_width']  = _HAvg $headerParams 'pic_width_in_luma_samples'
        $sps['pic_height'] = _HAvg $headerParams 'pic_height_in_luma_samples'
        $minCU = _HAvg $headerParams 'log2_min_luma_coding_block_size_minus3'
        $diffCU = _HAvg $headerParams 'log2_diff_max_min_luma_coding_block_size'
        if ($null -ne $minCU -and $null -ne $diffCU) {
            $sps['log2_min_cu'] = $minCU + 3; $sps['log2_max_cu'] = $minCU + 3 + $diffCU
        }
        $minTU = _HAvg $headerParams 'log2_min_luma_transform_block_size_minus2'
        $diffTU = _HAvg $headerParams 'log2_diff_max_min_luma_transform_block_size'
        if ($null -ne $minTU -and $null -ne $diffTU) {
            $sps['log2_min_tu'] = $minTU + 2; $sps['log2_max_tu'] = $minTU + 2 + $diffTU
        }
        $sps['tu_depth_inter']         = _HAvg $headerParams 'max_transform_hierarchy_depth_inter'
        $sps['tu_depth_intra']         = _HAvg $headerParams 'max_transform_hierarchy_depth_intra'
        $sps['amp_enabled']            = _HBool $headerParams 'amp_enabled_flag'
        $sps['sao_enabled']            = _HBool $headerParams 'sample_adaptive_offset_enabled_flag'
        $sps['pcm_enabled']            = _HBool $headerParams 'pcm_enabled_flag'
        $sps['strong_intra_smoothing'] = _HBool $headerParams 'strong_intra_smoothing_enable_flag'
        $sps['temporal_mvp']           = _HBool $headerParams 'sps_temporal_mvp_enabled_flag'
        $sps['scaling_list']           = _HBool $headerParams 'scaling_list_enabled_flag'
        $v = _HAvg $headerParams 'sps_max_dec_pic_buffering_minus1[i]'; $sps['dpb_size'] = if ($null -ne $v) { $v + 1 } else { $null }
        $sps['max_reorder_pics']       = _HAvg $headerParams 'sps_max_num_reorder_pics[i]'

        # PPS
        $pps['cu_qp_delta']            = _HBool $headerParams 'cu_qp_delta_enabled_flag'
        $pps['cu_qp_delta_depth']      = _HAvg $headerParams 'diff_cu_qp_delta_depth'
        $pps['cb_qp_offset']           = _HAvg $headerParams 'pps_cb_qp_offset'
        $pps['cr_qp_offset']           = _HAvg $headerParams 'pps_cr_qp_offset'
        $pps['sign_data_hiding']       = _HBool $headerParams 'sign_data_hiding_enabled_flag'
        $pps['transform_skip']         = _HBool $headerParams 'transform_skip_enabled_flag'
        $pps['constrained_intra']      = _HBool $headerParams 'constrained_intra_pred_flag'
        $pps['tiles']                  = _HBool $headerParams 'tiles_enabled_flag'
        $pps['wpp']                    = _HBool $headerParams 'entropy_coding_sync_enabled_flag'
        $pps['weighted_pred']          = _HBool $headerParams 'weighted_pred_flag'
        $pps['weighted_bipred']        = _HBool $headerParams 'weighted_bipred_flag'
        $pps['transquant_bypass']      = _HBool $headerParams 'transquant_bypass_enabled_flag'
        $pps['loop_filter_across']     = _HBool $headerParams 'pps_loop_filter_across_slices_enabled_flag'
        $pps['deblock_control']        = _HBool $headerParams 'deblocking_filter_control_present_flag'
        $v = _HAvg $headerParams 'init_qp_minus26'; $pps['init_qp'] = if ($null -ne $v) { $v + 26 } else { $null }

        # Slice
        $sliceCount = _HCnt $headerParams 'slice_type'
        $slice['count'] = $sliceCount
        if ($sliceCount -gt 0) {
            $fmm = _HVal $headerParams 'five_minus_max_num_merge_cand'
            if ($null -ne $fmm) { $slice['max_merge'] = 5 - [math]::Round($fmm / $sliceCount) }
            $slice['temporal_mvp']  = _HBool $headerParams 'slice_temporal_mvp_enabled_flag'
            $slice['sao_luma']      = _HBool $headerParams 'slice_sao_luma_flag'
            $slice['sao_chroma']    = _HBool $headerParams 'slice_sao_chroma_flag'
            $ep = _HVal $headerParams 'num_entry_point_offsets'
            if ($null -ne $ep) { $slice['entry_points'] = [math]::Round($ep / $sliceCount) }
        }

        # VUI
        $vui['colour_primaries']    = _HAvg $headerParams 'colour_primaries'
        $vui['transfer']            = _HAvg $headerParams 'transfer_characteristics'
        $vui['matrix_coeffs']       = _HAvg $headerParams 'matrix_coeffs'
        $vui['full_range']          = _HBool $headerParams 'video_full_range_flag'
        $vui['tick']                = _HAvg $headerParams 'vui_num_units_in_tick'
        $vui['timescale']           = _HAvg $headerParams 'vui_time_scale'
        if ($vui['tick'] -and $vui['timescale']) {
            # HEVC spec: fps = time_scale / (2 * num_units_in_tick) for frame-based
            # But some encoders set tick as half-period already, so fps = time_scale / num_units_in_tick
            $fps1 = $vui['timescale'] / (2 * $vui['tick'])   # spec formula
            $fps2 = $vui['timescale'] / $vui['tick']          # direct
            # Use whichever is closer to a standard rate (23.976, 24, 25, 29.97, 30, 50, 59.94, 60)
            $stdRates = @(23.976, 24, 25, 29.97, 30, 50, 59.94, 60)
            $best1 = ($stdRates | ForEach-Object { [math]::Abs($_ - $fps1) } | Measure-Object -Minimum).Minimum
            $best2 = ($stdRates | ForEach-Object { [math]::Abs($_ - $fps2) } | Measure-Object -Minimum).Minimum
            $vui['fps'] = if ($best1 -le $best2) { [math]::Round($fps1, 3) } else { [math]::Round($fps2, 3) }
        }

        # HDR
        $hdr['MaxCLL']  = _HAvg $headerParams 'max_content_light_level'
        $hdr['MaxFALL'] = _HAvg $headerParams 'max_pic_average_light_level'
        $v = _HAvg $headerParams 'max_display_mastering_luminance'; $hdr['max_lum'] = if ($null -ne $v) { $v / 10000 } else { $null }
        $v = _HAvg $headerParams 'min_display_mastering_luminance'; $hdr['min_lum'] = if ($null -ne $v) { $v / 10000 } else { $null }
        foreach ($i in 0..2) {
            $hdr["px$i"] = _HAvg $headerParams "display_primaries_x[$i]"
            $hdr["py$i"] = _HAvg $headerParams "display_primaries_y[$i]"
        }
        $hdr['wpx'] = _HAvg $headerParams 'white_point_x'
        $hdr['wpy'] = _HAvg $headerParams 'white_point_y'

        $result.SPS = $sps; $result.PPS = $pps; $result.Slice = $slice; $result.VUI = $vui; $result.HDR = $hdr

        # 3. CABAC statistics
        $cabac = [ordered]@{}
        foreach ($line in ($output -split "`n")) {
            if ($line -match '^\s+CABAC_BITS__(\w+)\s+:\s+(-|\d+)\s+(-|\w+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)') {
                $cabac["$($Matches[1])_$($Matches[2])_$($Matches[3])"] = @{ CABACCount=[long]$Matches[4]; CABACSum=[long]$Matches[5]; TotalBits=[long]$Matches[10] }
            }
            if ($line -match '^\[CABAC_BITS__(\w+)\s+~\s+~~ST~~\s+~~ST~~\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)') {
                $cabac["$($Matches[1])_TOTAL"] = @{ CABACCount=[long]$Matches[2]; CABACSum=[long]$Matches[3]; TotalBits=[long]$Matches[8] }
            }
        }
        $result.CABAC = $cabac

        # Analysis summary
        $analysis = [ordered]@{}
        foreach ($line in ($output -split "`n")) {
            if ($line -match '\[TOTAL\s+~\s+~~GT~~\s+~~GT~~\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+(\d+)') {
                $analysis['TotalBits'] = [long]$Matches[1]; break
            }
        }
        $tb = $analysis['TotalBits']
        if ($tb -and $tb -gt 0) {
            if ($cabac['SKIP_FLAG_-_-']) { $analysis['SkipRatio'] = [math]::Round($cabac['SKIP_FLAG_-_-'].CABACSum / $cabac['SKIP_FLAG_-_-'].CABACCount * 100, 1) }
            if ($cabac['MERGE_FLAG_-_-']) { $analysis['MergeCount'] = $cabac['MERGE_FLAG_-_-'].CABACSum }
            if ($cabac['SAO_-_-']) { $analysis['SAOPercent'] = [math]::Round($cabac['SAO_-_-'].TotalBits / $tb * 100, 2) }
            if ($cabac['INTRA_DIR_ANG_TOTAL']) { $analysis['IntraPercent'] = [math]::Round($cabac['INTRA_DIR_ANG_TOTAL'].TotalBits / $tb * 100, 2) }
            $mvd = 0
            if ($cabac['MVD_-_-']) { $mvd += $cabac['MVD_-_-'].TotalBits }
            if ($cabac['MVD_EP_-_-']) { $mvd += $cabac['MVD_EP_-_-'].TotalBits }
            if ($mvd -gt 0) { $analysis['MVDPercent'] = [math]::Round($mvd / $tb * 100, 2) }
        }
        $result.Analysis = $analysis

        # 4. SEI
        $seiPath = Join-Path $tempDir "sei.txt"
        if (Test-Path $seiPath) {
            $result.SEIMessages = Get-Content $seiPath -Raw -ErrorAction SilentlyContinue
        }

        Write-Host "      HM: $($headerParams.Count) params, $($frameQPs.Count) QP frames, $($cabac.Count) CABAC" -ForegroundColor Green
        return $result
    }
    catch {
        Write-Host "      HM Analyser error: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
    finally {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}