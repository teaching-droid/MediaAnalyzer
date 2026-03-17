# ─────────────────────────────────────────────────────────────────────────────
# Reconstruction.ps1 — Encoder settings reconstruction
# ─────────────────────────────────────────────────────────────────────────────

function Write-ReconstructionReport {
    param([System.Text.StringBuilder]$rpt, $rc)

    Write-Section $rpt "ENCODER SETTINGS RECONSTRUCTION"
    $rpt.AppendLine("    Confidence: HIGH=metadata, MED=bitstream, LOW=heuristic") | Out-Null

    # ── If we have embedded encoder settings, show them organized ──
    if ($rc['ParsedSettings']) {
        $rpt.AppendLine("`n    *** EMBEDDED ENCODER SETTINGS FOUND (HIGH confidence) ***") | Out-Null
        $ps = $rc['ParsedSettings']
        $groups = [ordered]@{
            "Rate Control"    = 'rc','crf','bitrate','qp','vbv-maxrate','vbv-bufsize','aq-mode','aq-strength','qcomp','qpmin','qpmax','qpstep','ipratio','pbratio','crf-max','crf-min','strict-cbr','lossless','cutree'
            "Frame Structure" = 'keyint','min-keyint','scenecut','open-gop','bframes','b-adapt','b-pyramid','bframe-bias','ref','gop-lookahead','rc-lookahead','lookahead-slices'
            "Analysis"        = 'me','subme','merange','rect','amp','limit-modes','max-merge','early-skip','fast-intra','b-intra','tu-inter-depth','tu-intra-depth','limit-tu','rdoq','rdoq-level','rd','rskip','psy-rd','psy-rdoq','tskip','cu-lossless','max-tu-size','min-cu-size'
            "Filtering"       = 'deblock','sao','sao-non-deblock','selective-sao','strong-intra-smoothing','constrained-intra'
            "Coding Tools"    = 'weightp','weightb','wpp','pmode','pme','tiles','slices','signhide','cabac-init','transform-skip'
            "Performance"     = 'preset','tune','profile','level-idc','pools','frame-threads','numa-pools','asm'
            "HDR/Color"       = 'hdr10','hdr10-opt','dhdr10-info','master-display','max-cll','repeat-headers','aud','hrd','info','hash','colorprim','transfer','colormatrix','chromaloc','range','sar'
        }
        foreach ($grp in $groups.GetEnumerator()) {
            $found = @()
            foreach ($key in $grp.Value) { if ($ps.Contains($key)) { $found += @{K=$key;V=$ps[$key]} } }
            if ($found.Count -gt 0) {
                $rpt.AppendLine("`n    [$($grp.Key)]") | Out-Null
                foreach ($f in $found) { Write-Field $rpt $f.K $f.V -Indent 8 -Width 30 }
            }
        }
        $allGrpKeys = $groups.Values | ForEach-Object { $_ }
        $remaining = $ps.GetEnumerator() | Where-Object { $allGrpKeys -notcontains $_.Key }
        if ($remaining) {
            $rpt.AppendLine("`n    [Other]") | Out-Null
            foreach ($r in $remaining) { Write-Field $rpt $r.Key $r.Value -Indent 8 -Width 30 }
        }
    }

    # ── Bitstream-derived parameters table ──
    Write-Section $rpt "BITSTREAM-DERIVED PARAMETERS"
    $params = [ordered]@{
        "Codec"="$($rc['Codec'])|HIGH";"Profile"="$($rc['Profile'])|HIGH";"Level"="$($rc['Level'])|HIGH"
        "Resolution"="$($rc['Resolution'])|HIGH";"Pixel Format"="$($rc['PixFmt'])|HIGH";"Bit Depth"="$($rc['BitDepth'])|HIGH"
        "Color Space"="$($rc['ColorSpace'])|HIGH";"Color Transfer"="$($rc['ColorTransfer'])|HIGH"
        "Color Primaries"="$($rc['ColorPrimaries'])|HIGH";"Color Range"="$($rc['ColorRange'])|HIGH"
        "CTU Size"="$($rc['CTU'])|HIGH"
        "Ref Frames"="$($rc['Refs'])|$(if($rc['Refs_Source'] -eq 'MediaInfo'){'HIGH'}else{'MED'})"
        "CABAC"="$(if($null -ne $rc['CABAC']){if($rc['CABAC']){'Yes'}else{'No'}}else{$null})|$(if($rc['CABAC_Source']){'HIGH'}else{'MED'})"
        "B-Frames (max consecutive)"="$($rc['BFrames'])|MED"
        "B-Pyramid"="$(if($rc['BPyramid'] -is [bool]){if($rc['BPyramid']){'True'}else{'False'}}else{$rc['BPyramid']})|$(if($rc['BPyramid_Source'] -match 'HM-confirmed|JM-confirmed'){'HIGH'}else{'MED'})"
        "Keyint (mode GOP)"="$($rc['Keyint'])|MED";"Min Keyint"="$($rc['MinKeyint'])|MED";"GOP Type"="$($rc['GOPType'])|MED"
        "SAO"="$(BoolStr $rc['SAO'])|HIGH";"AMP"="$(BoolStr $rc['AMP'])|HIGH"
        "Temporal MVP"="$(BoolStr $rc['TemporalMVP'])|HIGH"
        "Strong Intra Smoothing"="$(BoolStr $rc['StrongIntra'])|HIGH"
        "Weighted Pred"="$(BoolStr $rc['WeightedPred'])|HIGH";"Weighted Bipred"="$(BoolStr $rc['WeightedBipred'])|HIGH"
        "Transform Skip"="$(BoolStr $rc['TransformSkip'])|HIGH";"Sign Data Hiding"="$(BoolStr $rc['SignDataHiding'])|HIGH"
        "WPP"="$(BoolStr $rc['WPP'])|HIGH";"Tiles"="$(BoolStr $rc['Tiles'])|HIGH"
        "Deblocking"="$(if($null -ne $rc['Deblock']){if($rc['Deblock']){"On (b=$($rc['DeblockBeta']),tc=$($rc['DeblockTC']))"}else{"Off"}}else{$null})|HIGH"
        "CU QP Delta"="$(BoolStr $rc['CUQPDelta'])|HIGH"
        "TU Depth Inter/Intra"="$(if($rc['TUDepthInter']){"$($rc['TUDepthInter'])/$($rc['TUDepthIntra'])"}else{$null})|HIGH"
        "Rate Control"="$($rc['RateControl'])|LOW";"QP Average"="$($rc['QP_Avg'])|MED"
        "QP I-Frame"="$(if($rc['QP_I_Min']){"$($rc['QP_I_Min'])-$($rc['QP_I_Max']) (avg $($rc['QP_I']))"}else{$null})|MED"
        "QP P-Frame"="$(if($rc['QP_P_Min']){"$($rc['QP_P_Min'])-$($rc['QP_P_Max']) (avg $($rc['QP_P']))"}else{$null})|MED"
        "QP B-Frame"="$(if($rc['QP_B_Min']){"$($rc['QP_B_Min'])-$($rc['QP_B_Max']) (avg $($rc['QP_B']))"}else{$null})|MED"
        "Avg Bitrate"="$(if($rc['BitrateAvgKbps']){Format-Bitrate $rc['BitrateAvgKbps']}else{$null})|HIGH"
        "Peak Bitrate"="$(if($rc['PeakKbps']){Format-Bitrate $rc['PeakKbps']}else{$null})|HIGH"
        "Maxrate (if capped)"="$(if($rc['MaxrateLikely']){Format-Bitrate $rc['MaxrateLikely']}else{$null})|LOW"
        "VBV Buffer"="$($rc['VBV'])|MED"
        "Suggested Maxrate"="$(if($rc['SuggestedMaxrate']){Format-Bitrate $rc['SuggestedMaxrate']}else{$null})|LOW"
        "Max Sustained (12s)"="$(if($rc['MaxSustained12s']){Format-Bitrate $rc['MaxSustained12s']}else{$null})|MED"
        "Encoder"="$($rc['Encoder'])|MED";"DV Profile"="$($rc['DV_Profile'])|HIGH";"DV Level"="$($rc['DV_Level'])|HIGH"
    }
    foreach ($p in $params.GetEnumerator()) {
        $parts = "$($p.Value)" -split '\|'; $val = $parts[0]; $conf = $parts[1]
        if (-not [string]::IsNullOrWhiteSpace($val) -and $val -ne '') {
            Write-Field $rpt "$($p.Key) [$conf]" $val -Indent 4 -Width 44
        }
    }

    # ── Suggested command lines ──
    Write-CommandLines $rpt $rc

    # ── What cannot be determined ──
    Write-Section $rpt "PARAMETERS NOT DETERMINABLE FROM BITSTREAM"
    foreach ($u in @(
        "Preset (speed/quality tradeoff) - not stored in output"
        "Tune setting (film/animation/grain) - not stored"
        "Lookahead depth (rc-lookahead) - encoder-internal only"
        "Motion estimation method (me) - not signaled"
        "Subpixel ME quality (subme) - not signaled"
        "ME range (merange) - not signaled"
        "Psychovisual RD (psy-rd/psy-rdoq) - affects decisions not flags"
        "RD level - not signaled"
        "B-adapt mode - only result visible, not method"
        "AQ mode/strength - only CU QP delta flag visible"
        "Number of passes - not stored"
        "Source filtering applied before encoding"
    )) { $rpt.AppendLine("    - $u") | Out-Null }
    if (-not $rc['ParsedSettings']) { $rpt.AppendLine("`n    *** No Encoded_Library_Settings found - many settings unknown ***") | Out-Null }

    # ── Quality estimate ──
    Write-Section $rpt "COMPRESSION QUALITY ESTIMATE"
    if ($rc['Resolution'] -match '(\d+)x(\d+)') {
        $w=[int]$Matches[1];$h=[int]$Matches[2];$px=$w*$h
        $fps=if($rc['FPS']){$rc['FPS']}else{23.976}
        $vbr=if($rc['VideoBitrateKbps']){$rc['VideoBitrateKbps']}elseif($rc['BitrateAvgKbps']){$rc['BitrateAvgKbps']*0.9}else{0}
        if($fps -gt 0 -and $px -gt 0 -and $vbr -gt 0){
            $bppf=($vbr*1000)/($px*$fps)
            Write-Field $rpt "Bits/Pixel/Frame" ("{0:N4}" -f $bppf)
            $q=if($bppf -gt 0.15){"High/transparent"}elseif($bppf -gt 0.08){"Good"}elseif($bppf -gt 0.04){"Moderate (streaming)"}elseif($bppf -gt 0.02){"Low"}else{"Very low"}
            Write-Field $rpt "Quality Estimate" $q
            $bd=if($rc['BitDepth']){[int]$rc['BitDepth']}else{8};$raw=($px*$fps*$bd*1.5)/1000
            Write-Field $rpt "Compression Ratio" ("{0:N0}:1 vs raw" -f ($raw/$vbr))
        }
        if($rc['ColorTransfer'] -match 'smpte2084|arib-std-b67'){
            $rpt.AppendLine("`n    Note: HDR 10-bit needs ~20-30% more bitrate than SDR for equal quality.")|Out-Null
        }
    }
}

function Write-CommandLines {
    param([System.Text.StringBuilder]$rpt, $rc)

    $codec = $rc['Codec']
    if ($codec -notmatch 'hevc|h265|x265|h264|avc|x264') { return }

    $isHEVC = $codec -match 'hevc|h265|x265'
    $isHDR  = $rc['ColorTransfer'] -match 'smpte2084|arib-std-b67'
    $isDV   = $null -ne $rc['DV_Profile']
    $is10b  = $rc['BitDepth'] -eq '10' -or $rc['BitDepth'] -eq 10 -or $rc['PixFmt'] -match '10'

    # ── Smart keyint: align to ~1 second at actual FPS ──
    # Standard frame rates and their natural keyint (1-second GOP):
    #   23.976 → 24,  24 → 24,  25 → 25 (PAL),  29.97 → 30,  30 → 30
    #   48 → 48,  50 → 50 (PAL HFR),  59.94 → 60,  60 → 60
    $fps = if ($rc['FPS']) { $rc['FPS'] } else { 24 }
    $naturalKeyint = [math]::Round($fps)  # 1 second
    $detectedKeyint = $rc['Keyint']

    if ($detectedKeyint -and $fps -gt 0) {
        $keyintSec = $detectedKeyint / $fps
        # If detected keyint is NOT close to a whole-second boundary (0.8-1.2s), or
        # if it doesn't match the natural keyint for this FPS, suggest the natural one
        if ([math]::Abs($keyintSec - 1.0) -gt 0.25 -and $detectedKeyint -ne $naturalKeyint) {
            # Keep detected if it's a clean multiple of ~1s (2s, 3s, etc.)
            $isCleanMultiple = [math]::Abs(($keyintSec - [math]::Round($keyintSec))) -lt 0.1
            if (-not $isCleanMultiple) {
                $rc['Keyint_Original'] = $detectedKeyint
                $rc['Keyint'] = $naturalKeyint
            }
        }
        # Special case: detected=20 at 23.976fps → should be 24
        if ($detectedKeyint -eq 20 -and $fps -ge 23.9 -and $fps -le 24.1) {
            $rc['Keyint_Original'] = 20
            $rc['Keyint'] = 24
        }
    }

    # Normalize profile names for hardware encoders (NVEnc/QSVEnc use lowercase, no spaces, no quotes)
    # FFprobe reports "Main 10", NVEnc expects "main10", QSVEnc expects "main10"
    $hwProfile = if ($rc['Profile']) {
        $p = "$($rc['Profile'])".Trim()
        switch -Regex ($p) {
            '(?i)Main\s*10\s*444' { 'main444_10' ; break }
            '(?i)Main\s*444'      { 'main444' ; break }
            '(?i)Main\s*10'       { 'main10' ; break }
            '(?i)^Main$'          { 'main' ; break }
            default               { $p.ToLower() -replace '\s+','' -replace '"','' }
        }
    } else { $null }

    # DV profile string for NVEnc/QSVEnc --dolby-vision-profile
    # NVEnc supports: copy, 5.0, 8.1, 8.2, 8.4
    # QSVEnc supports: copy, 5.0, 8.1, 8.2, 8.4, 10.0, 10.1, 10.2, 10.4
    # Profile 7.x is NOT directly supported — use "copy" for transcode (RPU passthrough)
    # For raw encode, DV injection is done via dovi_tool (not via encoder flag)
    $dvProfileStr = if ($isDV) {
        $dvP = [int]$rc['DV_Profile']; $dvC = $rc['DV_Compat']
        $dvFull = "$dvP.$dvC"
        # Check if the profile.compat combo is directly supported
        $supportedDV = @('5.0','8.1','8.2','8.4','10.0','10.1','10.2','10.4')
        if ($dvFull -in $supportedDV) { $dvFull }
        else { 'copy' }  # use copy for unsupported profiles (e.g. 7.6)
    } else { $null }

    # ── Preset Recommendation ──
    Write-Section $rpt "RECOMMENDED ENCODING SETTINGS"
    $rpt.AppendLine("    Based on bitstream analysis of source file.") | Out-Null

    if ($isHEVC) {
        # Estimate complexity from detected tools
        $complexity = 0
        if ($rc['AMP'] -eq $true)          { $complexity += 2 }
        if ($rc['WeightedPred'] -eq $true)  { $complexity += 1 }
        if ($rc['WeightedBipred'] -eq $true) { $complexity += 1 }
        if ($rc['SignDataHiding'] -eq $true) { $complexity += 1 }
        $tuMax = if ($rc['TUDepthInter']) { [int]$rc['TUDepthInter'] } else { 0 }
        $complexity += $tuMax
        $bframes = if ($rc['BFrames']) { [int]$rc['BFrames'] } else { 0 }
        if ($bframes -ge 6) { $complexity += 2 }

        $presetRec = if ($complexity -ge 6) { "slow or slower" }
                     elseif ($complexity -ge 3) { "medium or slow" }
                     else { "medium" }

        # Build QSVEnc conditional feature set (used in guidance and commands)
        $qsvRm = if ($rc['QSVEnc_Remove']) { $rc['QSVEnc_Remove'] } else { @() }
        $qsvAdv = @()
        if ('extbrc' -notin $qsvRm) { $qsvAdv += "--extbrc" }
        if ('mbbrc' -notin $qsvRm) { $qsvAdv += "--mbbrc" }
        if ('adapt-ref' -notin $qsvRm) { $qsvAdv += "--adapt-ref" }
        if ('adapt-ltr' -notin $qsvRm) { $qsvAdv += "--adapt-ltr" }
        if ('adapt-cqm' -notin $qsvRm) { $qsvAdv += "--adapt-cqm" }
        if ('fade-detect' -notin $qsvRm) { $qsvAdv += "--fade-detect" }

        # Build NVEnc bframes (GPU-adjusted)
        $nvBFrames = if ($rc['NVEnc_MaxBFrames']) { [math]::Min([int]$rc['BFrames'], [int]$rc['NVEnc_MaxBFrames']) } elseif ($rc['BFrames']) { [int]$rc['BFrames'] } else { 0 }

        $rpt.AppendLine("`n    [Preset / Quality Guidance]") | Out-Null
        Write-Field $rpt "Recommended x265 Preset" $presetRec -Indent 8 -Width 34
        Write-Field $rpt "Recommended NVEncC Preset" "P6 with --tune uhq (Turing+ GPU required)" -Indent 8 -Width 34
        Write-Field $rpt "Recommended QSVEncC Quality" "--quality best" -Indent 8 -Width 34
        $rpt.AppendLine("") | Out-Null
        Write-Field $rpt "NVEncC Best Quality RC" "--qvbr 0 --vbr-quality <Q> (QVBR = default, best quality)" -Indent 8 -Width 34
        Write-Field $rpt "NVEncC Bitrate Target RC" "--vbr <kbps> --multipass 2pass-full (multipass only works with VBR/CBR)" -Indent 8 -Width 34
        Write-Field $rpt "NVEncC Lookahead" "--lookahead 32 (max, enables adaptive I/B)" -Indent 8 -Width 34
        $qsvQualRC = if ($rc['QSVEnc_RCFallback']) { $rc['QSVEnc_RCFallback'] } else { "--la-icq 23 --la-depth 40 (lookahead ICQ)" }
        $qsvBitRC  = if ($rc['QSVEnc_LAFallback']) { "$($rc['QSVEnc_LAFallback']) (GPU-adjusted)" } else { "--la <kbps> --la-depth 40 (lookahead bitrate)" }
        Write-Field $rpt "QSVEncC Best Quality RC" $qsvQualRC -Indent 8 -Width 34
        Write-Field $rpt "QSVEncC Bitrate Target RC" $qsvBitRC -Indent 8 -Width 34
        Write-Field $rpt "QSVEncC Advanced" "$(if($qsvAdv.Count -gt 0){$qsvAdv -join ' '}else{'(no advanced features supported)'})" -Indent 8 -Width 34
        $rpt.AppendLine("") | Out-Null
        Write-Field $rpt "Compliance (NVEnc)" "--aud --repeat-headers --pic-struct" -Indent 8 -Width 34
        Write-Field $rpt "Compliance (QSVEnc)" "--aud --repeat-headers --pic-struct --buf-period" -Indent 8 -Width 34
        Write-Field $rpt "Compliance (x265)" "aud:hrd:repeat-headers (hrd enables pic-timing + buf-period)" -Indent 8 -Width 34
        Write-Field $rpt "NVEncC Temporal Filter" "--tf-level 4 (requires bframes >= 4)" -Indent 8 -Width 34
        Write-Field $rpt "QSVEncC SAO" "--sao all / --sao none (auto/none/luma/chroma/all)" -Indent 8 -Width 34

        if ($rc['QP_I_Min'] -and $rc['QP_B_Max']) {
            $rpt.AppendLine("") | Out-Null
            Write-Field $rpt "QP Range (source analysis)" "I=$($rc['QP_I_Min'])-$($rc['QP_I_Max'])  P=$($rc['QP_P_Min'])-$($rc['QP_P_Max'])  B=$($rc['QP_B_Min'])-$($rc['QP_B_Max'])" -Indent 8 -Width 34
            $rpt.AppendLine("        NOT included in commands — constrains VBR too tightly on segments.") | Out-Null
            $rpt.AppendLine("        Add manually for full-file encodes if needed:") | Out-Null
            $rpt.AppendLine("          NVEncC/QSVEncC: --qp-min $($rc['QP_I_Min']):$($rc['QP_P_Min']):$($rc['QP_B_Min']) --qp-max $($rc['QP_I_Max']):$($rc['QP_P_Max']):$($rc['QP_B_Max'])") | Out-Null
            $rpt.AppendLine("          x265: qpmin=$($rc['QP_Min']):qpmax=$($rc['QP_Max'])") | Out-Null
        }

        if ($rc['CUQPDelta'] -eq $true) {
            Write-Field $rpt "AQ (Adaptive QP)" "ENABLED in source (CU QP Delta detected)" -Indent 8 -Width 34
            Write-Field $rpt "x265 AQ Recommendation" "aq-mode=2 (auto-variance), aq-strength=1.0" -Indent 8 -Width 34
            Write-Field $rpt "NVEncC AQ Recommendation" "--aq --aq-temporal --aq-strength 0 (auto)" -Indent 8 -Width 34
            Write-Field $rpt "QSVEncC AQ Recommendation" "--mbbrc (macroblock-level rate control)" -Indent 8 -Width 34
            $rpt.AppendLine("        (NVEncC AQ strength: 1=weak to 15=strong, 0=auto)") | Out-Null
        }

        if ($rc['Keyint_Original']) {
            $rpt.AppendLine("") | Out-Null
            Write-Field $rpt "Keyint Adjustment" "Detected mode GOP=$($rc['Keyint_Original']), adjusted to $($rc['Keyint']) (~1s at ${fps}fps)" -Indent 8 -Width 34
        }

        if ($isDV) {
            $rpt.AppendLine("`n    [Dolby Vision]") | Out-Null
            $dvProf = $rc['DV_Profile']; $dvLevel = $rc['DV_Level']; $dvCompat = $rc['DV_Compat']
            Write-Field $rpt "DV Profile" "$dvProf ($(if($dvCompat -eq 1){'HDR10 compatible'}elseif($dvCompat -eq 2){'SDR compatible'}else{'BL+RPU'}))" -Indent 8 -Width 34
            Write-Field $rpt "DV Level" $dvLevel -Indent 8 -Width 34
            $rpt.AppendLine("") | Out-Null
            $rpt.AppendLine("        NOTE: Dolby Vision RPU data requires special handling:") | Out-Null
            $rpt.AppendLine("        - Extract RPU: dovi_tool extract-rpu -i source.hevc -o rpu.bin") | Out-Null
            $rpt.AppendLine("        - Inject RPU:  dovi_tool inject-rpu -i encoded.hevc --rpu-in rpu.bin -o final.hevc") | Out-Null
            $rpt.AppendLine("        - x265 with RPU: --dolby-vision-rpu rpu.bin --dolby-vision-profile $dvProf.$dvCompat") | Out-Null
            $rpt.AppendLine("        - For Profile 8.1, the HDR10 base layer must be preserved correctly.") | Out-Null
        }

        if ($isHDR) {
            $rpt.AppendLine("`n    [HDR Metadata for Passthrough]") | Out-Null

            # Note source of HDR metadata
            $hdrSources = @()
            if ($rc['MaxCLL_Source'] -eq 'SEI') { $hdrSources += "MaxCLL/MaxFALL from SEI (bitstream)" }
            elseif ($rc['MaxCLL_Source'] -eq 'MediaInfo') { $hdrSources += "MaxCLL/MaxFALL from container metadata" }
            if ($rc['DisplayPrimaries_Source'] -eq 'SEI') { $hdrSources += "Display primaries from SEI (bitstream)" }
            if ($rc['MasterLum_Source'] -eq 'SEI') { $hdrSources += "Mastering luminance from SEI (bitstream)" }
            if ($hdrSources.Count -gt 0) {
                $rpt.AppendLine("        Source: $($hdrSources -join ', ')") | Out-Null
            }

            # Build master-display string (HEVC SEI: [0]=G, [1]=B, [2]=R)
            $hasMD = $rc['dp_x0'] -and $rc['MasterMaxLum']
            if ($hasMD) {
                # x265 master-display format: G(x,y)B(x,y)R(x,y)WP(x,y)L(max,min)
                # Values in 1/50000 units for primaries, luminance in cd/m2 * 10000
                $mdStr = "G($($rc['dp_x0']),$($rc['dp_y0']))B($($rc['dp_x1']),$($rc['dp_y1']))R($($rc['dp_x2']),$($rc['dp_y2']))WP($($rc['wp_x']),$($rc['wp_y']))L($([int]($rc['MasterMaxLum']*10000)),$([int]($rc['MasterMinLum']*10000)))"
                Write-Field $rpt "master-display (x265)" $mdStr -Indent 8 -Width 34
            }
            if ($rc['MaxCLL'] -and $rc['MaxFALL']) {
                Write-Field $rpt "max-cll (x265)" "$($rc['MaxCLL']),$($rc['MaxFALL'])" -Indent 8 -Width 34
            }
            # FFmpeg format for mastering display
            if ($hasMD) {
                $fmtV = { param($v) [math]::Round($v / 50000, 4) }
                $mdFF = "G($(& $fmtV $rc['dp_x0']),$(& $fmtV $rc['dp_y0']))B($(& $fmtV $rc['dp_x1']),$(& $fmtV $rc['dp_y1']))R($(& $fmtV $rc['dp_x2']),$(& $fmtV $rc['dp_y2']))WP($(& $fmtV $rc['wp_x']),$(& $fmtV $rc['wp_y']))L($($rc['MasterMaxLum']),$($rc['MasterMinLum']))"
                Write-Field $rpt "mastering-display (FFmpeg)" $mdFF -Indent 8 -Width 34
            }
            if ($rc['MaxCLL'] -and $rc['MaxFALL']) {
                Write-Field $rpt "content-light (FFmpeg)" "max_content=$($rc['MaxCLL']):max_average=$($rc['MaxFALL'])" -Indent 8 -Width 34
            }
        }
    }

    # ══════════════════════════════════════════════════════════════
    # COMMAND A: Encode from raw/source
    # ══════════════════════════════════════════════════════════════
    Write-Section $rpt "SUGGESTED REPRODUCTION COMMANDS (from raw source)"
    $rpt.AppendLine("    Match analyzed encoding settings. Adjust paths and fine-tune as needed.") | Out-Null

    if ($isHEVC) {
        # ── x265 params ──
        $xp = @()
        if($rc['Keyint'])    {$xp+="keyint=$($rc['Keyint'])"}
        if($rc['MinKeyint']) {$xp+="min-keyint=$($rc['MinKeyint'])"}
        if($rc['BFrames'])   {$xp+="bframes=$($rc['BFrames'])"}
        if($rc['Refs'])      {$xp+="ref=$($rc['Refs'])"}
        if($rc['CTU'])       {$xp+="ctu=$($rc['CTU'])"}
        if($null -ne $rc['SAO'])          {$xp+=if($rc['SAO']){"sao"}else{"no-sao"}}
        if($null -ne $rc['AMP'])          {$xp+=if($rc['AMP']){"amp"}else{"no-amp"}}
        if($null -ne $rc['WeightedPred']) {$xp+=if($rc['WeightedPred']){"weightp"}else{"no-weightp"}}
        if($null -ne $rc['WeightedBipred']){$xp+=if($rc['WeightedBipred']){"weightb"}else{"no-weightb"}}
        if($null -ne $rc['StrongIntra'])  {$xp+=if($rc['StrongIntra']){"strong-intra-smoothing"}else{"no-strong-intra-smoothing"}}
        if($null -ne $rc['Deblock'])      {if($rc['Deblock']){$xp+="deblock=$($rc['DeblockBeta']),$($rc['DeblockTC'])"}else{$xp+="no-deblock"}}
        if($null -ne $rc['WPP'])          {$xp+=if($rc['WPP']){"wpp"}else{"no-wpp"}}
        if($null -ne $rc['SignDataHiding']){$xp+=if($rc['SignDataHiding']){"signhide"}else{"no-signhide"}}
        if($rc['TUDepthInter']){$xp+="tu-inter-depth=$($rc['TUDepthInter'])"}
        if($rc['TUDepthIntra']){$xp+="tu-intra-depth=$($rc['TUDepthIntra'])"}
        if($rc['CUQPDelta']) {$xp+="aq-mode=2"}
        # QP bounds: omitted from default command — VBR/CRF selects QP adaptively
        # Add manually if needed: qpmin=$($rc['QP_Min']):qpmax=$($rc['QP_Max'])
        if($isHDR) {$xp+="hdr10-opt"; $xp+="repeat-headers"; $xp+="aud"; $xp+="hrd"}
        if($rc['MaxCLL'] -and $rc['MaxFALL']) {$xp+="max-cll=$($rc['MaxCLL']),$($rc['MaxFALL'])"}
        if ($rc['dp_x0'] -and $rc['MasterMaxLum']) {
            $mdX265 = "G($($rc['dp_x0']),$($rc['dp_y0']))B($($rc['dp_x1']),$($rc['dp_y1']))R($($rc['dp_x2']),$($rc['dp_y2']))WP($($rc['wp_x']),$($rc['wp_y']))L($([int]($rc['MasterMaxLum']*10000)),$([int]($rc['MasterMinLum']*10000)))"
            $xp += "master-display=$mdX265"
        }

        $rcP = if($rc['RateControl'] -match 'CRF.*~([\d.]+)'){"-crf $($Matches[1])"}
               elseif($rc['RateControl'] -match 'CQP.*=\s*(\d+)'){"-qp $($Matches[1])"}
               elseif($rc['BitrateAvgKbps']){"-b:v $([math]::Round($rc['BitrateAvgKbps']))k"}else{"-crf 20"}

        $rpt.AppendLine("`n    # x265 via FFmpeg:") | Out-Null
        $rpt.AppendLine("    ffmpeg -i input.mkv -c:v libx265 $rcP ``") | Out-Null
        $rpt.AppendLine("      -x265-params `"$($xp -join ':')`" ``") | Out-Null
        if($rc['ColorPrimaries']){$rpt.AppendLine("      -colorprim $($rc['ColorPrimaries']) -transfer $($rc['ColorTransfer']) -colormatrix $($rc['ColorSpace']) ``")|Out-Null}
        $maxrateK = if($rc['MaxrateLikely']){$rc['MaxrateLikely']}elseif($rc['SuggestedMaxrate']){[math]::Round($rc['SuggestedMaxrate'])}else{$null}
        if($maxrateK){$rpt.AppendLine("      -maxrate ${maxrateK}k -bufsize $($maxrateK*2)k ``")|Out-Null}
        $rpt.AppendLine("      -pix_fmt $($rc['PixFmt']) output.mkv") | Out-Null

        # ── NVEncC from raw ──
        $rpt.AppendLine("`n    # NVEncC (from raw/uncompressed source):") | Out-Null
        $nv = @("--codec hevc")
        if($hwProfile){$nv+="--profile $hwProfile"}
        if($is10b){$nv+="--output-depth 10"}
        $nv += "--preset P6 --tune uhq"
        if($rc['Keyint']){$nv+="--gop-len $($rc['Keyint'])"}
        # Use GPU-adjusted bframes if available (e.g. RTX 3080 max=5 for HEVC)
        $nvBFrames = if ($rc['NVEnc_MaxBFrames']) { [math]::Min([int]$rc['BFrames'], [int]$rc['NVEnc_MaxBFrames']) } elseif ($rc['BFrames']) { [int]$rc['BFrames'] } else { 0 }
        if($nvBFrames -gt 0){
            $nv+="--bframes $nvBFrames"
            if ($nvBFrames -ge 3) { $nv += "--bref-mode middle" }
        }
        # Ref frames: use HM-derived actual ref count when available
        # Cap to NVEnc HEVC max (6 on Ampere/Ada), minimum 2 with B-pyramid
        $nvRef = if ($rc['Refs']) { [int]$rc['Refs'] } else { 1 }
        if ($nvBFrames -ge 3 -and $nvRef -lt 2) { $nvRef = 2 }
        $nvRef = [math]::Min($nvRef, 6)  # NVEnc HEVC max ref frames
        $nv += "--ref $nvRef"
        # NVEnc: weighted prediction unsupported with B-frames
        if($rc['WeightedPred'] -and $nvBFrames -eq 0){$nv+="--weightp"}
        # Rate control: QVBR (quality-constrained VBR) is NVEnc's best quality mode
        # --multipass only works with --vbr and --cbr, NOT with --qvbr
        if($rc['BitrateAvgKbps']){
            $nv+="--vbr $([math]::Round($rc['BitrateAvgKbps']))"
            $nv+="--multipass 2pass-full"
            if($maxrateK){$nv+="--max-bitrate $maxrateK"; $nv+="--vbv-bufsize $($maxrateK * 2)"}
        } else {
            $nv+="--qvbr 0 --vbr-quality 20"
        }
        $nv += "--lookahead 32"
        # QP bounds: NOT included in command by default — constrains VBR too tightly
        # Source QP range (19-32) reflects full-film variation; short segments need freedom
        # Add manually if needed: --qp-min $($rc['QP_I_Min']):$($rc['QP_P_Min']):$($rc['QP_B_Min']) --qp-max ...
        if($rc['CUQPDelta']){$nv+="--aq --aq-temporal --aq-strength 0"}
        if($rc['CbQPOffset']){$nv+="--chroma-qp-offset $($rc['CbQPOffset'])"}
        # Temporal filter (requires bframes >= 4)
        if($nvBFrames -ge 4){$nv+="--tf-level 4"}
        if($rc['ColorPrimaries']){$nv+="--colorprim $($rc['ColorPrimaries']) --transfer $($rc['ColorTransfer']) --colormatrix $($rc['ColorSpace'])"}
        if($rc['ColorRange'] -eq 'tv'){$nv+="--colorrange limited"}
        # Compliance: AUD + repeat-headers + pic-struct for proper HEVC/HDR/DV stream
        $nv += "--aud --repeat-headers --pic-struct"
        if($isHDR) {
            if($rc['MaxCLL'] -and $rc['MaxFALL']){$nv+="`n      --max-cll `"$($rc['MaxCLL']),$($rc['MaxFALL'])`""}
            if($rc['dp_x0'] -and $rc['MasterMaxLum']){
                $fmtV = { param($v) [math]::Round($v / 50000, 4) }
                $nv += "`n      --master-display `"G($(& $fmtV $rc['dp_x0']),$(& $fmtV $rc['dp_y0']))B($(& $fmtV $rc['dp_x1']),$(& $fmtV $rc['dp_y1']))R($(& $fmtV $rc['dp_x2']),$(& $fmtV $rc['dp_y2']))WP($(& $fmtV $rc['wp_x']),$(& $fmtV $rc['wp_y']))L($($rc['MasterMaxLum']),$($rc['MasterMinLum']))`""
            }
        }
        $rpt.AppendLine("    NVEncC64 -i input.mkv -o output.mkv ``") | Out-Null
        $rpt.AppendLine("      $($nv -join ' ')") | Out-Null

        # ── QSVEncC from raw ──
        $rpt.AppendLine("`n    # QSVEncC (from raw/uncompressed source, Intel QSV):") | Out-Null
        $qsv = @("--codec hevc")
        if($hwProfile){$qsv+="--profile $hwProfile"}
        if($is10b){$qsv+="--output-depth 10"}
        $qsv += "--quality best"
        if($rc['Keyint']){$qsv+="--gop-len $($rc['Keyint'])"}
        if($rc['BFrames']){
            $qsv+="--bframes $($rc['BFrames'])"
            $qsv+="--b-pyramid --weightb"
        }
        # Ref frames: use HM-derived actual ref count when available
        # QSVEnc HEVC supports up to 16 ref frames, minimum 2 with B-pyramid
        $qsvRef = if ($rc['Refs']) { [int]$rc['Refs'] } else { 1 }
        if ($rc['BFrames'] -and [int]$rc['BFrames'] -ge 3 -and $qsvRef -lt 2) { $qsvRef = 2 }
        $qsvRef = [math]::Min($qsvRef, 16)  # QSVEnc HEVC max ref frames
        $qsv += "--ref $qsvRef"
        if($rc['WeightedPred']){$qsv+="--weightp"}
        # Rate control: use GPU-validated fallbacks if available
        if($rc['BitrateAvgKbps']){
            if ($rc['QSVEnc_LAFallback']) {
                # LA not supported, use fallback (e.g. --vbr)
                $qsv += $rc['QSVEnc_LAFallback']
            } else {
                $qsv+="--la $([math]::Round($rc['BitrateAvgKbps'])) --la-depth 40"
            }
            if($maxrateK){$qsv+="--max-bitrate $maxrateK"; $qsv+="--vbv-bufsize $($maxrateK * 2)"}
        } else {
            if ($rc['QSVEnc_RCFallback']) {
                # LA-ICQ not supported, use fallback (e.g. --icq 23)
                $qsv += $rc['QSVEnc_RCFallback']
            } else {
                $qsv+="--la-icq 23 --la-depth 40"
            }
        }
        if($rc['SAO']){$qsv+="--sao all"} else {$qsv+="--sao none"}
        # QP bounds: omitted — VBR selects QP adaptively
        # Advanced features — conditionally exclude if GPU doesn't support them
        $qsvRm = if ($rc['QSVEnc_Remove']) { $rc['QSVEnc_Remove'] } else { @() }
        $qsvAdv = @()
        if ('extbrc' -notin $qsvRm) { $qsvAdv += "--extbrc" }
        if ('mbbrc' -notin $qsvRm) { $qsvAdv += "--mbbrc" }
        if ('adapt-ref' -notin $qsvRm) { $qsvAdv += "--adapt-ref" }
        if ('adapt-ltr' -notin $qsvRm) { $qsvAdv += "--adapt-ltr" }
        if ('adapt-cqm' -notin $qsvRm) { $qsvAdv += "--adapt-cqm" }
        if ('fade-detect' -notin $qsvRm) { $qsvAdv += "--fade-detect" }
        if ($qsvAdv.Count -gt 0) { $qsv += $qsvAdv -join ' ' }
        if($rc['ColorPrimaries']){$qsv+="--colorprim $($rc['ColorPrimaries']) --transfer $($rc['ColorTransfer']) --colormatrix $($rc['ColorSpace'])"}
        if($rc['ColorRange'] -eq 'tv'){$qsv+="--colorrange limited"}
        # Compliance: AUD + repeat-headers + pic-struct + buf-period
        $qsv += "--aud --repeat-headers --pic-struct --buf-period"
        if($isHDR) {
            if($rc['MaxCLL'] -and $rc['MaxFALL']){$qsv+="`n      --max-cll `"$($rc['MaxCLL']),$($rc['MaxFALL'])`""}
            if($rc['dp_x0'] -and $rc['MasterMaxLum']){
                $fmtV = { param($v) [math]::Round($v / 50000, 4) }
                $qsv += "`n      --master-display `"G($(& $fmtV $rc['dp_x0']),$(& $fmtV $rc['dp_y0']))B($(& $fmtV $rc['dp_x1']),$(& $fmtV $rc['dp_y1']))R($(& $fmtV $rc['dp_x2']),$(& $fmtV $rc['dp_y2']))WP($(& $fmtV $rc['wp_x']),$(& $fmtV $rc['wp_y']))L($($rc['MasterMaxLum']),$($rc['MasterMinLum']))`""
            }
        }
        $rpt.AppendLine("    QSVEncC64 -i input.mkv -o output.mkv ``") | Out-Null
        $rpt.AppendLine("      $($qsv -join ' ')") | Out-Null
    }

    if ($codec -match 'h264|avc|x264') {
        # ── x264 params ──
        $xp = @()
        if($rc['Keyint'])  { $xp += "keyint=$($rc['Keyint'])" }
        if($rc['BFrames']) { $xp += "bframes=$($rc['BFrames'])" }
        if($rc['Refs'])    { $xp += "ref=$($rc['Refs'])" }
        # CABAC: confirmed by MediaInfo (Format_Settings_CABAC). Default on for High profile but explicit is safer.
        if($null -ne $rc['CABAC']) { $xp += if($rc['CABAC']) { "cabac=1" } else { "cabac=0" } }
        # B-pyramid: confirmed by JM decoder (uppercase B-ref frames in decode output)
        if($rc['BPyramid'] -eq $true -and $rc['BFrames'] -gt 1) { $xp += "b-pyramid=normal" }
        $rcP = if($rc['RateControl'] -match 'CRF.*~([\d.]+)') { "-crf $($Matches[1])" }
               elseif($rc['BitrateAvgKbps'])                   { "-b:v $([math]::Round($rc['BitrateAvgKbps']))k" }
               else                                            { "-crf 20" }
        $rpt.AppendLine("`n    # x264 via FFmpeg:") | Out-Null
        $rpt.AppendLine("    ffmpeg -i input.mkv -c:v libx264 $rcP ``") | Out-Null
        if($xp) { $rpt.AppendLine("      -x264-params `"$($xp -join ':')`" ``") | Out-Null }
        $rpt.AppendLine("      -pix_fmt $($rc['PixFmt']) output.mkv") | Out-Null

        # ── NVEncC H.264 ──
        # H.264 profile mapping: 'High'→high, 'Main'→main, 'Baseline'→baseline
        $h264Profile = switch -regex ($rc['Profile']) {
            'High'     { 'high' }
            'Main'     { 'main' }
            'Baseline' { 'baseline' }
            default    { 'high' }
        }
        $rpt.AppendLine("`n    # NVEncC (from raw/uncompressed source):") | Out-Null
        $nvH = @("--codec h264")
        $nvH += "--profile $h264Profile"
        $nvH += "--preset P6 --tune uhq"
        if($rc['Keyint']) { $nvH += "--gop-len $($rc['Keyint'])" }
        # NVEncC H.264 max B-frames = 4 (from check-features B=4)
        $nvHBFrames = if ($rc['BFrames']) { [math]::Min([int]$rc['BFrames'], 4) } else { 0 }
        if($nvHBFrames -gt 0) {
            $nvH += "--bframes $nvHBFrames"
            # --bref-mode middle is the H.264 equivalent of B-pyramid
            # NVEncC H.264 supports 'each' and 'only middle' per check-features
            if($rc['BPyramid'] -eq $true -and $nvHBFrames -ge 2) { $nvH += "--bref-mode middle" }
        }
        # H.264 ref frames: cap to 4 (NVEncC H.264 advertises B=4 meaning max DPB entries ≤ 4)
        $nvHRef = if($rc['Refs']) { [math]::Min([int]$rc['Refs'], 4) } else { 1 }
        if($nvHBFrames -ge 2 -and $nvHRef -lt 2) { $nvHRef = 2 }
        $nvH += "--ref $nvHRef"
        if($rc['BitrateAvgKbps']) {
            $nvH += "--vbr $([math]::Round($rc['BitrateAvgKbps']))"
            $nvH += "--multipass 2pass-full"
            $maxrateKH = if($rc['SuggestedMaxrate']){[math]::Round($rc['SuggestedMaxrate'])}else{$null}
            if($maxrateKH) { $nvH += "--max-bitrate $maxrateKH"; $nvH += "--vbv-bufsize $($maxrateKH * 2)" }
        } else {
            $nvH += "--qvbr 0 --vbr-quality 20"
        }
        $nvH += "--lookahead 32"
        if($rc['CUQPDelta']) { $nvH += "--aq --aq-temporal --aq-strength 0" }
        # Temporal filter: NVEncC H.264 supports TF (check-features TF:+)
        if($nvHBFrames -ge 4) { $nvH += "--tf-level 4" }
        if($rc['ColorPrimaries']) { $nvH += "--colorprim $($rc['ColorPrimaries']) --transfer $($rc['ColorTransfer']) --colormatrix $($rc['ColorSpace'])" }
        if($rc['ColorRange'] -eq 'tv') { $nvH += "--colorrange limited" }
        # Compliance: AUD + repeat-headers (no --pic-struct/buf-period needed for H.264)
        $nvH += "--aud --repeat-headers"
        $rpt.AppendLine("    NVEncC64 -i input.mkv -o output.mkv ``") | Out-Null
        $rpt.AppendLine("      $($nvH -join ' ')") | Out-Null

        # ── QSVEncC H.264 ──
        $rpt.AppendLine("`n    # QSVEncC (from raw/uncompressed source, Intel QSV):") | Out-Null
        $qsvH = @("--codec h264")
        $qsvH += "--profile $h264Profile"
        $qsvH += "--quality best"
        if($rc['Keyint']) { $qsvH += "--gop-len $($rc['Keyint'])" }
        if($rc['BFrames']) {
            $qsvH += "--bframes $($rc['BFrames'])"
            # QSVEncC --b-pyramid is valid for both HEVC and H.264 when B-pyramid confirmed
            if($rc['BPyramid'] -eq $true) { $qsvH += "--b-pyramid --weightb" }
        }
        # H.264: QSVEncC supports up to 16 ref, but cap sensibly to detected ref count
        $qsvHRef = if($rc['Refs']) { [math]::Min([int]$rc['Refs'], 16) } else { 1 }
        if($rc['BFrames'] -and [int]$rc['BFrames'] -ge 2 -and $qsvHRef -lt 2) { $qsvHRef = 2 }
        $qsvH += "--ref $qsvHRef"
        if($rc['WeightedPred']) { $qsvH += "--weightp" }
        # H.264 supports LA-ICQ and LA (lookahead bitrate) — use them if bitrate available
        if($rc['BitrateAvgKbps']) {
            $qsvH += "--la $([math]::Round($rc['BitrateAvgKbps'])) --la-depth 40"
            $maxrateKH = if($rc['SuggestedMaxrate']){[math]::Round($rc['SuggestedMaxrate'])}else{$null}
            if($maxrateKH) { $qsvH += "--max-bitrate $maxrateKH"; $qsvH += "--vbv-bufsize $($maxrateKH * 2)" }
        } else {
            $qsvH += "--la-icq 23 --la-depth 40"
        }
        # Advanced: H.264 supports all of these per check-features
        $qsvH += "--extbrc --mbbrc --adapt-ref --adapt-ltr --fade-detect"
        if($rc['ColorPrimaries']) { $qsvH += "--colorprim $($rc['ColorPrimaries']) --transfer $($rc['ColorTransfer']) --colormatrix $($rc['ColorSpace'])" }
        if($rc['ColorRange'] -eq 'tv') { $qsvH += "--colorrange limited" }
        # Compliance: AUD + repeat-headers (no --buf-period for H.264, that's HEVC-only)
        $qsvH += "--aud --repeat-headers"
        $rpt.AppendLine("    QSVEncC64 -i input.mkv -o output.mkv ``") | Out-Null
        $rpt.AppendLine("      $($qsvH -join ' ')") | Out-Null
    }

    # ══════════════════════════════════════════════════════════════
    # COMMAND B: Transcode from THIS file (copy metadata)
    # ══════════════════════════════════════════════════════════════
    if ($isHEVC) {
        Write-Section $rpt "TRANSCODE COMMANDS (from this file, copy HDR metadata)"
        $rpt.AppendLine("    Re-encode from this file. Copies HDR/color metadata from source.") | Out-Null
        $srcFile = if ($rc['FileName']) { $rc['FileName'] } else { "source.mkv" }

        # ── FFmpeg transcode ──
        $xpT = @()
        if($rc['Keyint']){$xpT+="keyint=$($rc['Keyint'])"}
        if($rc['MinKeyint']){$xpT+="min-keyint=$($rc['MinKeyint'])"}
        if($rc['BFrames']){$xpT+="bframes=$($rc['BFrames'])"}
        if($rc['Refs']){$xpT+="ref=$($rc['Refs'])"}
        if($rc['CTU']){$xpT+="ctu=$($rc['CTU'])"}
        if($null -ne $rc['SAO']){$xpT+=if($rc['SAO']){"sao"}else{"no-sao"}}
        if($null -ne $rc['AMP']){$xpT+=if($rc['AMP']){"amp"}else{"no-amp"}}
        if($null -ne $rc['WPP']){$xpT+=if($rc['WPP']){"wpp"}else{"no-wpp"}}
        if($null -ne $rc['SignDataHiding']){$xpT+=if($rc['SignDataHiding']){"signhide"}else{"no-signhide"}}
        if($rc['CUQPDelta']){$xpT+="aq-mode=2"}
        # QP bounds: omitted — CRF selects QP adaptively
        if($isHDR){$xpT+="hdr10-opt"; $xpT+="repeat-headers"; $xpT+="aud"; $xpT+="hrd"}
        if($rc['MaxCLL'] -and $rc['MaxFALL']){$xpT+="max-cll=$($rc['MaxCLL']),$($rc['MaxFALL'])"}
        if ($rc['dp_x0'] -and $rc['MasterMaxLum']) {
            $mdX265 = "G($($rc['dp_x0']),$($rc['dp_y0']))B($($rc['dp_x1']),$($rc['dp_y1']))R($($rc['dp_x2']),$($rc['dp_y2']))WP($($rc['wp_x']),$($rc['wp_y']))L($([int]($rc['MasterMaxLum']*10000)),$([int]($rc['MasterMinLum']*10000)))"
            $xpT += "master-display=$mdX265"
        }

        $rpt.AppendLine("`n    # x265 via FFmpeg (transcode with HDR passthrough):") | Out-Null
        $rpt.AppendLine("    ffmpeg -i `"$srcFile`" -map 0:v:0 -map 0:a -c:a copy ``") | Out-Null
        $rpt.AppendLine("      -c:v libx265 -crf 18 ``") | Out-Null
        $rpt.AppendLine("      -x265-params `"$($xpT -join ':')`" ``") | Out-Null
        if($rc['ColorPrimaries']){$rpt.AppendLine("      -colorprim $($rc['ColorPrimaries']) -transfer $($rc['ColorTransfer']) -colormatrix $($rc['ColorSpace']) ``")|Out-Null}
        $rpt.AppendLine("      -pix_fmt $($rc['PixFmt']) ``") | Out-Null
        $rpt.AppendLine("      output_transcoded.mkv") | Out-Null

        # ── NVEncC transcode (uses copy flags for HDR/DV metadata) ──
        $nvT = @("--avhw --codec hevc")
        if($hwProfile){$nvT+="--profile $hwProfile"}
        if($is10b){$nvT+="--output-depth 10"}
        $nvT += "--preset P6 --tune uhq"
        if($rc['Keyint']){$nvT+="--gop-len $($rc['Keyint'])"}
        # Use GPU-adjusted bframes
        if($nvBFrames -gt 0){
            $nvT+="--bframes $nvBFrames"
            if ($nvBFrames -ge 3) { $nvT += "--bref-mode middle" }
        }
        if($rc['Refs']){
            $nvRefT = [int]$rc['Refs']
            if ($nvBFrames -ge 3 -and $nvRefT -lt 2) { $nvRefT = 2 }
            $nvRefT = [math]::Min($nvRefT, 6)
            $nvT += "--ref $nvRefT"
        }
        if($rc['WeightedPred'] -and $nvBFrames -eq 0){$nvT+="--weightp"}
        # Rate control: QVBR for quality, VBR+multipass for bitrate targeting
        if($rc['BitrateAvgKbps']){
            $nvT+="--vbr $([math]::Round($rc['BitrateAvgKbps']))"
            $nvT+="--multipass 2pass-full"
            if($maxrateK){$nvT+="--max-bitrate $maxrateK"; $nvT+="--vbv-bufsize $($maxrateK * 2)"}
        } else {
            $nvT+="--qvbr 0 --vbr-quality 20"
        }
        $nvT += "--lookahead 32"
        # QP bounds: omitted — VBR selects QP adaptively
        if($rc['CUQPDelta']){$nvT+="--aq --aq-temporal --aq-strength 0"}
        if($rc['CbQPOffset']){$nvT+="--chroma-qp-offset $($rc['CbQPOffset'])"}
        # Temporal filter (requires bframes >= 4)
        if($nvBFrames -ge 4){$nvT+="--tf-level 4"}
        # Auto-copy color info from source
        $nvT += "--colormatrix auto --transfer auto --colorprim auto --chromaloc auto --colorrange auto"
        # Compliance: AUD + repeat-headers + pic-struct (NVEnc has no --buf-period)
        $nvT += "--aud --repeat-headers --pic-struct"
        # Copy HDR metadata from source (no manual values needed)
        if($isHDR) {
            $nvT += "--max-cll copy"
            $nvT += "--master-display copy"
            $nvT += "--dhdr10-info copy"
        }
        # Copy DV RPU directly from source (no dovi_tool needed for NVEncC!)
        if ($isDV) {
            $nvT += "--dolby-vision-rpu copy"
            $nvT += "--dolby-vision-profile $dvProfileStr"
        }

        $rpt.AppendLine("`n    # NVEncC (transcode - auto-copies HDR/DV metadata from source):") | Out-Null
        $rpt.AppendLine("    NVEncC64 -i `"$srcFile`" -o output_transcoded.mkv ``") | Out-Null
        $rpt.AppendLine("      --audio-copy --sub-copy --chapter-copy ``") | Out-Null
        $rpt.AppendLine("      $($nvT -join ' ')") | Out-Null

        # ── QSVEncC transcode (uses copy flags for HDR/DV metadata) ──
        $qsvT = @("--avhw --codec hevc")
        if($hwProfile){$qsvT+="--profile $hwProfile"}
        if($is10b){$qsvT+="--output-depth 10"}
        $qsvT += "--quality best"
        if($rc['Keyint']){$qsvT+="--gop-len $($rc['Keyint'])"}
        if($rc['BFrames']){
            $qsvT+="--bframes $($rc['BFrames'])"
            $qsvT+="--b-pyramid --weightb"
        }
        if($rc['Refs']){
            $qsvRefT = [int]$rc['Refs']
            if ($rc['BFrames'] -and [int]$rc['BFrames'] -ge 3 -and $qsvRefT -lt 2) { $qsvRefT = 2 }
            $qsvRefT = [math]::Min($qsvRefT, 16)
            $qsvT += "--ref $qsvRefT"
        }
        if($rc['WeightedPred']){$qsvT+="--weightp"}
        # Rate control: use GPU-validated fallbacks
        if($rc['BitrateAvgKbps']){
            if ($rc['QSVEnc_LAFallback']) {
                $qsvT += $rc['QSVEnc_LAFallback']
            } else {
                $qsvT+="--la $([math]::Round($rc['BitrateAvgKbps'])) --la-depth 40"
            }
            if($maxrateK){$qsvT+="--max-bitrate $maxrateK"; $qsvT+="--vbv-bufsize $($maxrateK * 2)"}
        } else {
            if ($rc['QSVEnc_RCFallback']) {
                $qsvT += $rc['QSVEnc_RCFallback']
            } else {
                $qsvT += "--la-icq 23 --la-depth 40"
            }
        }
        if($rc['SAO']){$qsvT+="--sao all"} else {$qsvT+="--sao none"}
        # QP bounds: omitted — VBR selects QP adaptively
        # Advanced features — conditionally exclude unsupported (same as raw)
        if ($qsvAdv.Count -gt 0) { $qsvT += $qsvAdv -join ' ' }
        # Auto-copy color info from source
        $qsvT += "--colormatrix auto --transfer auto --colorprim auto --chromaloc auto --colorrange auto"
        # Compliance: AUD + repeat-headers + pic-struct + buf-period
        $qsvT += "--aud --repeat-headers --pic-struct --buf-period"
        # Copy HDR metadata from source
        if($isHDR) {
            $qsvT += "--max-cll copy"
            $qsvT += "--master-display copy"
            $qsvT += "--dhdr10-info copy"
        }
        # Copy DV RPU directly from source
        if ($isDV) {
            $qsvT += "--dolby-vision-rpu copy"
            $qsvT += "--dolby-vision-profile $dvProfileStr"
        }

        $rpt.AppendLine("`n    # QSVEncC (transcode - auto-copies HDR/DV metadata from source, Intel QSV):") | Out-Null
        $rpt.AppendLine("    QSVEncC64 -i `"$srcFile`" -o output_transcoded.mkv ``") | Out-Null
        $rpt.AppendLine("      --audio-copy --sub-copy --chapter-copy ``") | Out-Null
        $rpt.AppendLine("      $($qsvT -join ' ')") | Out-Null

        # ── DV transcode via x265 + dovi_tool (alternative when NVEncC DV not available) ──
        if ($isDV) {
            $rpt.AppendLine("`n    # Dolby Vision via x265 + dovi_tool (Profile $($rc['DV_Profile']).$($rc['DV_Compat'])):") | Out-Null
            $rpt.AppendLine("    # Step 1: Extract raw HEVC stream and DV RPU:") | Out-Null
            $rpt.AppendLine("    ffmpeg -i `"$srcFile`" -c:v copy -bsf:v hevc_mp4toannexb -an -sn -f hevc source.hevc") | Out-Null
            $rpt.AppendLine("    dovi_tool extract-rpu -i source.hevc -o rpu.bin") | Out-Null
            $rpt.AppendLine("    # Step 2: Encode base layer with x265 (include HDR10 metadata):") | Out-Null
            $rpt.AppendLine("    # (use x265 command from above, add: --dolby-vision-rpu rpu.bin --dolby-vision-profile $($rc['DV_Profile']).$($rc['DV_Compat']))") | Out-Null
            $rpt.AppendLine("    # Step 3: If x265 DV injection fails, use dovi_tool:") | Out-Null
            $rpt.AppendLine("    dovi_tool inject-rpu -i encoded.hevc --rpu-in rpu.bin -o final_dv.hevc") | Out-Null
            $rpt.AppendLine("    # Step 4: Mux back with audio/subs from source:") | Out-Null
            $rpt.AppendLine("    mkvmerge -o output_dv.mkv final_dv.hevc -A -S `"$srcFile`" -D") | Out-Null
        }
    }
}