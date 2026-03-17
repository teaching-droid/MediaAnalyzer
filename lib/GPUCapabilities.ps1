# ─────────────────────────────────────────────────────────────────────────────
# GPUCapabilities.ps1 — GPU hardware encode/decode capability detection
# ─────────────────────────────────────────────────────────────────────────────
# Runs NVEncC64 --check-features and QSVEncC64 --check-features to detect
# what the installed GPU(s) actually support for each codec.
# Returns structured capability data used to validate & adjust commands.
# ─────────────────────────────────────────────────────────────────────────────

function Get-GPUCapabilities {
    <#
    .SYNOPSIS
        Detect GPU encode/decode capabilities via NVEncC and QSVEncC.
    .DESCRIPTION
        Runs --check-features on available encoders and parses the output
        into structured data. Used to validate suggested commands.
    .OUTPUTS
        Hashtable with keys: NVEnc, QSVEnc (each containing parsed caps)
    #>
    param(
        [string]$NVEncPath  = "NVEncC64",
        [string]$QSVEncPath = "QSVEncC64"
    )

    $caps = @{
        NVEnc  = $null
        QSVEnc = $null
        Errors = @()
    }

    # ── NVEncC ──
    if ($NVEncPath -and (Test-Path $NVEncPath -ErrorAction SilentlyContinue)) {
        try {
            $r = Run-Command -Exe $NVEncPath -Arguments @('--check-features') -TimeoutSeconds 60 -StatusLabel "NVEncC check-features"
            if ($null -eq $r) {
                $caps.Errors += "NVEncC64: Run-Command returned null"
            } else {
                $nvStdOut = [string]$r.StdOut
                $nvStdErr = [string]$r.StdErr
                if (-not $nvStdOut) { $nvStdOut = "" }
                if (-not $nvStdErr) { $nvStdErr = "" }
                $nvOut = $nvStdOut + "`n" + $nvStdErr
                if ($script:DebugMode) {
                    Write-Host "      [DEBUG] NVEncC check-features: exit=$($r.ExitCode) stdout=$($nvStdOut.Length)ch stderr=$($nvStdErr.Length)ch total=$($nvOut.Length)ch" -ForegroundColor Magenta
                }
                if ($nvOut.Length -gt 100) {
                    $caps.NVEnc = Parse-NVEncFeatures $nvOut
                } else {
                    $caps.Errors += "NVEncC64: check-features too short (stdout=$($nvStdOut.Length)ch, stderr=$($nvStdErr.Length)ch, exit=$($r.ExitCode))"
                }
            }
        } catch {
            $caps.Errors += "NVEncC64 failed: $($_.Exception.Message) at $($_.InvocationInfo.ScriptLineNumber)"
        }
    } else {
        $caps.Errors += "NVEncC64 not found (not in tools folder or PATH)"
    }

    # ── QSVEncC ──
    if ($QSVEncPath -and (Test-Path $QSVEncPath -ErrorAction SilentlyContinue)) {
        try {
            $r = Run-Command -Exe $QSVEncPath -Arguments @('--check-features') -TimeoutSeconds 30 -StatusLabel "QSVEncC check-features"
            $qsvStdOut = if ($r -and $r.StdOut) { $r.StdOut } else { "" }
            $qsvStdErr = if ($r -and $r.StdErr) { $r.StdErr } else { "" }
            $qsvOut = "${qsvStdOut}`n${qsvStdErr}"
            if ($script:DebugMode) {
                Write-Host "      [DEBUG] QSVEncC check-features: exit=$($r.ExitCode) stdout=$($qsvStdOut.Length)ch stderr=$($qsvStdErr.Length)ch total=$($qsvOut.Length)ch" -ForegroundColor Magenta
            }
            if ($qsvOut.Length -gt 100) {
                $caps.QSVEnc = Parse-QSVEncFeatures $qsvOut
            } else {
                $caps.Errors += "QSVEncC64: check-features returned no useful data (stdout=$($qsvStdOut.Length), stderr=$($qsvStdErr.Length))"
            }
        } catch {
            $caps.Errors += "QSVEncC64 failed: $($_.Exception.Message)"
        }
    } else {
        $caps.Errors += "QSVEncC64 not found (not in tools folder or PATH)"
    }

    return $caps
}

function Parse-NVEncFeatures {
    <#
    .SYNOPSIS
        Parse NVEncC64 --check-features output into structured data.
    .DESCRIPTION
        NVEncC outputs blocks per codec (H.264, HEVC, AV1) with key/value
        feature lines, plus NVDec supported formats.
        
        Example format:
            Codec: H.265/HEVC
            Max Bframes          4
            B Ref Mode           3 (each + only middle)
            RC Modes             63 (CQP, CBR, CBRHQ, VBR, VBRHQ)
            SAO                  yes
            Lookahead            yes
            AQ (temporal)        yes
            Weighted Prediction  yes
            10bit depth          yes
            ...
            NVDec features
            H.265/HEVC:  nv12, yv12, yv12(10bit), yv12(12bit), yuv444, ...
    #>
    param([string]$Output)

    if (-not $Output) { return $null }

    $result = @{
        GPU       = ""
        Driver    = ""
        API       = ""
        Codecs    = @{}
        Decode    = @{}
    }

    try {
        # Extract GPU info
        if ($Output -match '#\d+:\s*(.+?)\s*\((\d+)\s*cores.*?\)\[.*?\]\[([^\]]+)\]') {
            $result.GPU    = [string]$Matches[1]
            $result.Driver = [string]$Matches[3]
        } elseif ($Output -match '#\d+:\s*(.+?)\s*\((\d+)\s*cores') {
            $result.GPU = [string]$Matches[1]
        }
        if ($Output -match 'NVENC API v([\d.]+)') {
            $result.API = [string]$Matches[1]
        }

        # Parse codec blocks
        $currentCodec = $null
        $inDecode = $false
        foreach ($line in ($Output -split "`r?`n")) {
            $trimmed = $line.Trim()
            if (-not $trimmed -or $trimmed.Length -lt 3) { continue }

            # Codec header
            if ($trimmed -match '^Codec:\s*(.+)$') {
                $inDecode = $false
                $codecName = [string]$Matches[1]
                $codecKey = if ($codecName -match 'H\.264|AVC') { 'h264' }
                            elseif ($codecName -match 'H\.265|HEVC') { 'hevc' }
                            elseif ($codecName -match 'AV1') { 'av1' }
                            else { $codecName.ToLower() -replace '[^a-z0-9]','' }
                $currentCodec = @{}
                $result.Codecs[$codecKey] = $currentCodec
                continue
            }

            # NVDec section
            if ($trimmed -match '^NVDec features') {
                $currentCodec = $null
                $inDecode = $true
                continue
            }

            # Skip non-data lines
            if ($trimmed -match '^(NVEnc features|Environment Info|OS\s*:|CPU:|RAM:|reader:|others)') { 
                $currentCodec = $null
                continue
            }
            # Skip library lines like "nvml       : yes"
            if ($trimmed -match '^\w+\s+:\s+(yes|no|enabled|disabled)$') { continue }

            # Decode line: "  H.265/HEVC:  nv12, yv12, yv12(10bit), ..."
            if ($inDecode -and $trimmed -match '^([\w./\-]+):\s+(\w.+)$') {
                $decCodec = [string]$Matches[1]
                $decFormats = [string]$Matches[2]
                $decKey = if ($decCodec -match 'H\.264|AVC') { 'h264' }
                          elseif ($decCodec -match 'H\.265|HEVC') { 'hevc' }
                          elseif ($decCodec -match 'AV1') { 'av1' }
                          elseif ($decCodec -match 'VP9') { 'vp9' }
                          elseif ($decCodec -match 'MPEG2') { 'mpeg2' }
                          else { $decCodec.ToLower() -replace '[^a-z0-9]','' }
                $fmtList = $decFormats -split ',\s*'
                $result.Decode[$decKey] = @{
                    Formats  = $fmtList
                    Has10bit = ($fmtList -match '10bit').Count -gt 0
                    Has12bit = ($fmtList -match '12bit').Count -gt 0
                    Has444   = ($fmtList -match '444').Count -gt 0
                }
                continue
            }

            # Feature line inside codec block: "Max Bframes          4"
            if ($currentCodec -and $trimmed -match '^(.+?)\s{2,}(.+)$') {
                $fName = [string]$Matches[1]
                $fVal  = ([string]$Matches[2]).Trim()
                if (-not $fName -or -not $fVal) { continue }
                $fKey = $fName.ToLower().Trim() -replace '\s+','_' -replace '[^a-z0-9_]',''

                # Parse value — avoid switch -Regex to prevent $Matches contamination
                $parsed = $null
                if ($fVal -eq 'yes') { $parsed = $true }
                elseif ($fVal -eq 'no') { $parsed = $false }
                elseif ($fVal -match '^\d+$') { $parsed = [int]$fVal }
                elseif ($fVal -match '^(\d+)\s*\((.+)\)$') {
                    $parsed = @{ Value = [int]$Matches[1]; Detail = [string]$Matches[2] }
                }
                else { $parsed = $fVal }

                $currentCodec[$fKey] = $parsed
            }
        }
    } catch {
        # Return partial result on parse error
        $result.GPU = "$($result.GPU) [PARSE ERROR: $($_.Exception.Message) line $($_.InvocationInfo.ScriptLineNumber)]"
    }

    return $result
}

function Parse-QSVEncFeatures {
    <#
    .SYNOPSIS
        Parse QSVEncC64 --check-features output into structured data.
    .DESCRIPTION
        QSVEncC outputs a table per codec+mode (e.g. "H.265/HEVC PG", "H.265/HEVC FF")
        with RC mode columns (CBR, VBR, AVBR, QVBR, CQP, LA, LAHRD, ICQ, LAICQ, VCM)
        and feature rows using o (supported) / x (not supported).
        
        Also has decode capability table and VPP features.
    #>
    param([string]$Output)

    $result = @{
        GPU       = ""
        Driver    = ""
        API       = ""
        Codecs    = @{}   # keyed like 'hevc_pg', 'hevc_ff', 'h264_pg', etc.
        CodecBest = @{}   # best capability per codec (merged PG+FF)
        Decode    = @{}
        VPP       = @{}
        RawOutput = $Output
    }

    # Extract device info
    if ($Output -match 'GPU:\s*(.+?)(?:\s*$|\r|\n)') {
        $result.GPU = $Matches[1].Trim()
    }
    if ($Output -match 'Media SDK.*?v([\d.]+)') {
        $result.API = $Matches[1]
    }

    $lines = $Output -split "`n"
    $currentCodec = $null
    $currentKey = $null
    $rcModes = @()
    $inDecode = $false
    $inVPP = $false

    for ($li = 0; $li -lt $lines.Count; $li++) {
        $line = $lines[$li]
        $trimmed = $line.Trim()

        # Codec block header: "Codec: H.265/HEVC PG" or "Codec: H.264/AVC FF"
        if ($trimmed -match '^Codec:\s*(.+)$') {
            $codecFull = $Matches[1].Trim()
            $inDecode = $false; $inVPP = $false

            # Parse codec name and mode (PG/FF)
            $mode = ''
            $codecName = $codecFull
            if ($codecFull -match '(.+?)\s+(PG|FF)\s*$') {
                $codecName = $Matches[1].Trim()
                $mode = $Matches[2].ToLower()
            }
            $baseKey = switch -Regex ($codecName) {
                'H\.264|AVC'  { 'h264' }
                'H\.265|HEVC' { 'hevc' }
                'AV1'         { 'av1'  }
                'VP9'         { 'vp9'  }
                'MPEG2'       { 'mpeg2' }
                default       { $codecName.ToLower() -replace '[^a-z0-9]','' }
            }
            $currentKey = if ($mode) { "${baseKey}_${mode}" } else { $baseKey }
            $currentCodec = @{ _baseCodec = $baseKey; _mode = $mode }
            $result.Codecs[$currentKey] = $currentCodec
            $rcModes = @()
            continue
        }

        # RC mode header line: "             CBR   VBR   AVBR  QVBR  CQP   LA    LAHRD ICQ   LAICQ VCM"
        if ($currentCodec -and $trimmed -match '^\s*(CBR|RC mode)\s' -and $trimmed -match 'CBR') {
            # Extract column names by splitting on whitespace
            $rcModes = @($trimmed -split '\s+' | Where-Object { $_ -ne '' })
            continue
        }

        # Feature row: "WeightP       o     o     x     o     o     x     x     o     x     x"
        if ($currentCodec -and $rcModes.Count -gt 0 -and $trimmed -match '^\s*(\S[\w\s/()+-]+?)\s{2,}([ox](?:\s+[ox])*)\s*$') {
            $featureName = $Matches[1].Trim()
            $valStr = $Matches[2].Trim()
            $fKey = $featureName.ToLower() -replace '\s+','_' -replace '[^a-z0-9_+]',''

            # Parse o/x values per RC mode
            $vals = @($valStr -split '\s+')
            $perRC = @{}
            $anySupported = $false
            for ($vi = 0; $vi -lt [math]::Min($vals.Count, $rcModes.Count); $vi++) {
                $supported = $vals[$vi] -eq 'o'
                $perRC[$rcModes[$vi]] = $supported
                if ($supported) { $anySupported = $true }
            }
            $currentCodec[$fKey] = @{
                Supported = $anySupported
                PerRC     = $perRC
            }
            continue
        }

        # Decode section
        if ($trimmed -match '^Supported Decode features') {
            $inDecode = $true; $inVPP = $false; $currentCodec = $null
            continue
        }
        if ($inDecode -and $trimmed -match '^(H\.264|HEVC|MPEG2|VP[89]|AV1|VVC)\s') {
            # Codec header in decode table — skip, parse the value rows below
            continue
        }
        if ($inDecode -and $trimmed -match '^(yuv\d+)\s+(.+)$') {
            # Decode format row like: "yuv420  8bit  10bit   8bit         10bit  10bit"
            # We'd need the codec column headers to properly parse this
            # For now, just note what formats are available
            $fmt = $Matches[1]
            $bitVals = $Matches[2]
            if ($bitVals -match '10bit') { $result.Decode["${fmt}_10bit"] = $true }
            if ($bitVals -match '12bit') { $result.Decode["${fmt}_12bit"] = $true }
            if ($bitVals -match '8bit')  { $result.Decode["${fmt}_8bit"]  = $true }
            continue
        }

        # VPP section
        if ($trimmed -match '^Supported Vpp features') {
            $inVPP = $true; $inDecode = $false; $currentCodec = $null
            continue
        }
        if ($inVPP -and $trimmed -match '^(.+?)\s{2,}([ox])\s*$') {
            $result.VPP[$Matches[1].Trim()] = ($Matches[2] -eq 'o')
            continue
        }
    }

    # Build "best" merged view per base codec (any feature supported in PG or FF = supported)
    $baseCodecs = @($result.Codecs.Values | ForEach-Object { $_._baseCodec } | Sort-Object -Unique)
    foreach ($bc in $baseCodecs) {
        $merged = @{}
        $variants = $result.Codecs.GetEnumerator() | Where-Object { $_.Value._baseCodec -eq $bc }
        foreach ($v in $variants) {
            foreach ($fk in $v.Value.Keys) {
                if ($fk -match '^_') { continue }  # skip internal keys
                $fData = $v.Value[$fk]
                if ($fData -is [hashtable] -and $fData.ContainsKey('Supported')) {
                    if (-not $merged.ContainsKey($fk) -or $fData.Supported) {
                        $merged[$fk] = $fData.Supported
                    }
                    # Also store best RC modes
                    if ($fData.PerRC) {
                        if (-not $merged.ContainsKey("${fk}_rc")) { $merged["${fk}_rc"] = @{} }
                        foreach ($rcK in $fData.PerRC.Keys) {
                            if ($fData.PerRC[$rcK]) { $merged["${fk}_rc"][$rcK] = $true }
                        }
                    }
                }
            }
        }
        $result.CodecBest[$bc] = $merged
    }

    return $result
}

function Validate-EncoderCommands {
    <#
    .SYNOPSIS
        Validate suggested encoder commands against actual GPU capabilities.
    .DESCRIPTION
        Takes the reconstruction data (rc) and GPU capabilities, returns
        warnings about unsupported features and adjusted command suggestions.
    #>
    param(
        [hashtable]$rc,
        [hashtable]$caps
    )

    $warnings = @()
    $adjustments = @()
    $codec = if ($rc['Codec'] -match 'hevc|h265') { 'hevc' }
             elseif ($rc['Codec'] -match 'h264|avc') { 'h264' }
             else { 'unknown' }

    # ── NVEnc validation ──
    if ($caps.NVEnc -and $caps.NVEnc.Codecs.ContainsKey($codec)) {
        $nvc = $caps.NVEnc.Codecs[$codec]

        # B-frames
        if ($rc['BFrames']) {
            $maxB = if ($nvc['max_bframes'] -is [hashtable]) { $nvc['max_bframes'].Value }
                    elseif ($nvc['max_bframes'] -is [int]) { $nvc['max_bframes'] }
                    else { $null }
            if ($null -ne $maxB -and [int]$rc['BFrames'] -gt $maxB) {
                $warnings += "NVEnc: GPU supports max $maxB B-frames for $codec, source uses $($rc['BFrames']). Will be clamped."
                $adjustments += @{ Encoder='NVEnc'; Param='bframes'; Original=$rc['BFrames']; Adjusted=$maxB }
            }
        }

        # B-Ref Mode
        if ($rc['BFrames'] -and [int]$rc['BFrames'] -ge 3) {
            $brefVal = $nvc['b_ref_mode']
            $brefSupport = if ($brefVal -is [hashtable]) { $brefVal.Value -ge 2 }
                           elseif ($brefVal -is [int]) { $brefVal -ge 2 }
                           else { $false }
            if (-not $brefSupport) {
                $warnings += "NVEnc: --bref-mode middle not supported on this GPU. Removing from command."
                $adjustments += @{ Encoder='NVEnc'; Param='bref-mode'; Remove=$true }
            }
        }

        # SAO
        if ($rc['SAO'] -eq $true -and $nvc['sao'] -eq $false) {
            $warnings += "NVEnc: SAO not supported on this GPU for $codec."
        }

        # Weighted prediction
        if ($rc['WeightedPred'] -and $nvc['weighted_prediction'] -eq $false) {
            $warnings += "NVEnc: Weighted prediction not supported. --weightp will be ignored."
            $adjustments += @{ Encoder='NVEnc'; Param='weightp'; Remove=$true }
        }

        # Lookahead
        if ($nvc['lookahead'] -eq $false) {
            $warnings += "NVEnc: Lookahead not supported on this GPU. --lookahead will be ignored."
            $adjustments += @{ Encoder='NVEnc'; Param='lookahead'; Remove=$true }
        }

        # AQ temporal
        if ($nvc['aq_temporal'] -eq $false -and $rc['CUQPDelta']) {
            $warnings += "NVEnc: Temporal AQ not supported. --aq-temporal will be ignored."
            $adjustments += @{ Encoder='NVEnc'; Param='aq-temporal'; Remove=$true }
        }

        # 10-bit
        if (($rc['BitDepth'] -eq '10' -or $rc['BitDepth'] -eq 10 -or $rc['PixFmt'] -match '10') -and $nvc['10bit_depth'] -eq $false) {
            $warnings += "NVEnc: 10-bit encoding NOT supported on this GPU for $codec! Encode will fail."
            $adjustments += @{ Encoder='NVEnc'; Param='output-depth'; Critical=$true }
        }

        # Temporal filter (requires bframes >= 4)
        if ($nvc['temporal_filter'] -eq $true -and $rc['BFrames'] -and [int]$rc['BFrames'] -ge 4) {
            $adjustments += @{ Encoder='NVEnc'; Param='tf-level'; Suggest='--tf-level 4' }
        }

        # Tune UHQ (Turing+ only)
        if ($nvc.ContainsKey('undirectional_b') -or $caps.NVEnc.GPU -match 'GTX 16[56]0|GTX 10[678]0|GTX 9[5-8]0') {
            # Pre-Turing GPU detected, tune uhq not available
            if ($caps.NVEnc.GPU -match 'GTX 10[678]0|GTX 9[5-8]0|GT[SX] 7[5-8]0') {
                $warnings += "NVEnc: --tune uhq requires Turing+ GPU. Falling back to --tune hq."
                $adjustments += @{ Encoder='NVEnc'; Param='tune'; Original='uhq'; Adjusted='hq' }
            }
        }

        # Decode check
        if ($caps.NVEnc.Decode.ContainsKey($codec)) {
            $dec = $caps.NVEnc.Decode[$codec]
            if (($rc['BitDepth'] -eq '10' -or $rc['BitDepth'] -eq 10 -or $rc['PixFmt'] -match '10') -and -not $dec.Has10bit) {
                $warnings += "NVEnc: HW decode of 10-bit $codec not supported. Use --avsw instead of --avhw."
                $adjustments += @{ Encoder='NVEnc'; Param='decoder'; Suggest='--avsw' }
            }
        }
    }

    # ── QSVEnc validation ──
    if ($caps.QSVEnc -and $caps.QSVEnc.CodecBest.ContainsKey($codec)) {
        $qc = $caps.QSVEnc.CodecBest[$codec]

        # LA-ICQ support
        $laicqSupport = $qc['rc_mode_rc'] -and $qc['rc_mode_rc'].ContainsKey('LAICQ') -and $qc['rc_mode_rc']['LAICQ']
        $icqSupport   = $qc['rc_mode_rc'] -and $qc['rc_mode_rc'].ContainsKey('ICQ') -and $qc['rc_mode_rc']['ICQ']
        if (-not $laicqSupport -and $icqSupport) {
            $warnings += "QSVEnc: LA-ICQ not supported for $codec. Falling back to --icq."
            $adjustments += @{ Encoder='QSVEnc'; Param='la-icq'; Suggest='--icq 23' }
        } elseif (-not $laicqSupport -and -not $icqSupport) {
            $warnings += "QSVEnc: Neither LA-ICQ nor ICQ supported for $codec. Use --cqp or --vbr."
            $adjustments += @{ Encoder='QSVEnc'; Param='la-icq'; Suggest='--cqp 20:23:25' }
        }

        # Helper: check if a QSV feature is supported (handles both bool and hashtable values)
        $qsvSupported = { param($key)
            if (-not $qc.ContainsKey($key)) { return $null }  # unknown
            $v = $qc[$key]
            if ($v -is [hashtable]) { return $v.Supported }
            return [bool]$v
        }

        # ExtBRC
        if ((& $qsvSupported 'extbrc') -eq $false) {
            $warnings += "QSVEnc: ExtBRC not supported for $codec. --extbrc will be removed."
            $adjustments += @{ Encoder='QSVEnc'; Param='extbrc'; Remove=$true }
        }

        # MBBRC
        if ((& $qsvSupported 'mbbrc') -eq $false) {
            $warnings += "QSVEnc: MBBRC not supported for $codec. --mbbrc will be removed."
            $adjustments += @{ Encoder='QSVEnc'; Param='mbbrc'; Remove=$true }
        }

        # WeightP/WeightB
        if ((& $qsvSupported 'weightp') -eq $false) {
            $warnings += "QSVEnc: Weighted P-prediction not supported for $codec."
            $adjustments += @{ Encoder='QSVEnc'; Param='weightp'; Remove=$true }
        }
        if ((& $qsvSupported 'weightb') -eq $false) {
            $adjustments += @{ Encoder='QSVEnc'; Param='weightb'; Remove=$true }
        }

        # B-frames / B-pyramid
        if ((& $qsvSupported 'bframegopref') -eq $false) {
            $warnings += "QSVEnc: B-frames not supported for $codec."
            $adjustments += @{ Encoder='QSVEnc'; Param='bframes'; Remove=$true }
        }
        if ((& $qsvSupported 'b_pyramid') -eq $false) {
            $adjustments += @{ Encoder='QSVEnc'; Param='b-pyramid'; Remove=$true }
        }

        # 10-bit
        if ((& $qsvSupported '10bit_depth') -eq $false -and ($rc['BitDepth'] -eq '10' -or $rc['BitDepth'] -eq 10 -or $rc['PixFmt'] -match '10')) {
            $warnings += "QSVEnc: 10-bit encoding NOT supported for $codec on this GPU!"
            $adjustments += @{ Encoder='QSVEnc'; Param='output-depth'; Critical=$true }
        }

        # Adaptive features
        foreach ($feat in @(
            @{Key='adaptive_i'; Param='i-adapt'; Name='Adaptive I'},
            @{Key='adaptive_b'; Param='b-adapt'; Name='Adaptive B'},
            @{Key='fadedetect'; Param='fade-detect'; Name='Fade detect'},
            @{Key='adaptiveref'; Param='adapt-ref'; Name='Adaptive ref'},
            @{Key='adaptiveltr'; Param='adapt-ltr'; Name='Adaptive LTR'},
            @{Key='adaptivecqm'; Param='adapt-cqm'; Name='Adaptive CQM'}
        )) {
            if ($qc.ContainsKey($feat.Key)) {
                $val = $qc[$feat.Key]
                $isSupported = if ($val -is [hashtable]) { $val.Supported } else { [bool]$val }
                if (-not $isSupported) {
                    $adjustments += @{ Encoder='QSVEnc'; Param=$feat.Param; Remove=$true }
                }
            }
        }

        # SAO
        if ((& $qsvSupported 'sao') -eq $false -and $rc['SAO']) {
            $warnings += "QSVEnc: SAO not supported for $codec on this GPU."
        }

        # LA modes
        $laSupport = $qc['rc_mode_rc'] -and $qc['rc_mode_rc'].ContainsKey('LA') -and $qc['rc_mode_rc']['LA']
        if (-not $laSupport -and $rc['BitrateAvgKbps']) {
            $warnings += "QSVEnc: LA (Lookahead bitrate) not supported for $codec. Use --vbr instead."
            $adjustments += @{ Encoder='QSVEnc'; Param='la'; Suggest="--vbr $([math]::Round($rc['BitrateAvgKbps']))" }
        }
    }

    # ── Decode performance note ──
    if ($caps.QSVEnc) {
        if ($caps.QSVEnc.GPU -match 'Arc\s*[AB]\d|DG[12]') {
            $warnings += "QSVEnc: Intel Arc dGPU detected. For some content, --avsw may be faster than --avhw. Benchmark both."
        }
        if ($caps.NVEnc -and $caps.QSVEnc) {
            $warnings += "Both NVIDIA and Intel GPUs detected. NVEncC uses NVIDIA GPU, QSVEncC uses Intel GPU. Use --device to select specific GPU."
        }
    }

    # ── Build AdjustedRC for shared parameters that affect command generation ──
    # These modify $rc values before reconstruction so commands are correct
    $adjustedRC = @{}
    $qsvRemove = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($adj in $adjustments) {
        if ($adj.Encoder -eq 'NVEnc' -and $adj.Param -eq 'bframes' -and $adj.Adjusted) {
            $adjustedRC['NVEnc_MaxBFrames'] = $adj.Adjusted
        }
        if ($adj.Encoder -eq 'QSVEnc' -and $adj.Param -eq 'la-icq') {
            $adjustedRC['QSVEnc_RCFallback'] = $adj.Suggest
        }
        if ($adj.Encoder -eq 'QSVEnc' -and $adj.Param -eq 'la') {
            $adjustedRC['QSVEnc_LAFallback'] = $adj.Suggest
        }
        # Track QSVEnc features to remove from commands
        if ($adj.Encoder -eq 'QSVEnc' -and $adj.Remove) {
            $qsvRemove.Add($adj.Param) | Out-Null
        }
    }
    if ($qsvRemove.Count -gt 0) {
        $adjustedRC['QSVEnc_Remove'] = $qsvRemove
    }

    return @{
        Warnings    = $warnings
        Adjustments = $adjustments
        AdjustedRC  = $adjustedRC
    }
}

function Write-GPUCapabilityReport {
    <#
    .SYNOPSIS
        Write GPU capability summary to the report.
    #>
    param(
        [System.Text.StringBuilder]$rpt,
        [hashtable]$caps,
        [hashtable]$validation
    )

    Write-Section $rpt "GPU HARDWARE CAPABILITIES"

    # ── NVEnc ──
    if ($caps.NVEnc) {
        $nv = $caps.NVEnc
        $rpt.AppendLine("    [NVIDIA GPU (NVEncC)]") | Out-Null
        if ($nv.GPU)    { Write-Field $rpt "GPU"        $nv.GPU    -Indent 8 -Width 24 }
        if ($nv.Driver) { Write-Field $rpt "Driver"     $nv.Driver -Indent 8 -Width 24 }
        if ($nv.API)    { Write-Field $rpt "NVENC API"  "v$($nv.API)" -Indent 8 -Width 24 }

        foreach ($codecKey in @('hevc','h264','av1')) {
            if ($nv.Codecs.ContainsKey($codecKey)) {
                $c = $nv.Codecs[$codecKey]
                $cName = switch ($codecKey) { 'hevc' {'HEVC'} 'h264' {'H.264'} 'av1' {'AV1'} }
                $rpt.AppendLine("") | Out-Null
                $rpt.AppendLine("        Encode $cName :") | Out-Null

                # Key features
                $feats = @()
                $maxB = if ($c['max_bframes'] -is [hashtable]) { $c['max_bframes'].Value } elseif ($c['max_bframes'] -is [int]) { $c['max_bframes'] } else { '?' }
                $feats += "B=$maxB"
                foreach ($f in @('sao','lookahead','aq_temporal','weighted_prediction','10bit_depth','lossless','temporal_filter')) {
                    $fv = $c[$f]
                    $fn = switch ($f) {
                        'sao' {'SAO'} 'lookahead' {'LA'} 'aq_temporal' {'AQ-T'}
                        'weighted_prediction' {'WP'} '10bit_depth' {'10b'}
                        'lossless' {'LL'} 'temporal_filter' {'TF'}
                    }
                    if ($null -ne $fv) {
                        $sym = if ($fv -eq $true) { '+' } elseif ($fv -eq $false) { '-' } else { '?' }
                        $feats += "${fn}:${sym}"
                    }
                }
                $rpt.AppendLine("            $($feats -join '  ')") | Out-Null

                # B-ref mode detail
                $bref = $c['b_ref_mode']
                if ($bref -is [hashtable]) {
                    $rpt.AppendLine("            B-Ref: $($bref.Detail)") | Out-Null
                }
            }
        }

        # Decode
        if ($nv.Decode.Count -gt 0) {
            $rpt.AppendLine("") | Out-Null
            $rpt.AppendLine("        Decode:") | Out-Null
            foreach ($dk in $nv.Decode.Keys | Sort-Object) {
                $d = $nv.Decode[$dk]
                $rpt.AppendLine("            $($dk.ToUpper()): $($d.Formats -join ', ')") | Out-Null
            }
        }
    } else {
        $rpt.AppendLine("    [NVIDIA GPU (NVEncC)]  Not detected") | Out-Null
    }

    # ── QSVEnc ──
    if ($caps.QSVEnc) {
        $qsv = $caps.QSVEnc
        $rpt.AppendLine("") | Out-Null
        $rpt.AppendLine("    [Intel GPU (QSVEncC)]") | Out-Null
        if ($qsv.GPU) { Write-Field $rpt "GPU" $qsv.GPU -Indent 8 -Width 24 }
        if ($qsv.API) { Write-Field $rpt "Media SDK API" "v$($qsv.API)" -Indent 8 -Width 24 }

        # Show per-codec summary from merged best view
        foreach ($codecKey in @('hevc','h264','av1','vp9')) {
            if ($qsv.CodecBest.ContainsKey($codecKey)) {
                $c = $qsv.CodecBest[$codecKey]
                $cName = switch ($codecKey) { 'hevc' {'HEVC'} 'h264' {'H.264'} 'av1' {'AV1'} 'vp9' {'VP9'} }

                # Collect RC modes that support the base RC mode feature
                $rcModes = @()
                if ($c.ContainsKey('rc_mode_rc')) {
                    foreach ($rcK in @('CBR','VBR','CQP','ICQ','LAICQ','LA','QVBR','AVBR')) {
                        if ($c['rc_mode_rc'].ContainsKey($rcK) -and $c['rc_mode_rc'][$rcK]) { $rcModes += $rcK }
                    }
                }

                # Key features
                $feats = @()
                foreach ($f in @('10bit_depth','weightp','weightb','b_pyramid','+manybframes',
                                 'sao','mbbrc','extbrc','adaptiveref','adaptiveltr','fadedetect')) {
                    $fn = switch ($f) {
                        '10bit_depth' {'10b'} 'weightp' {'WP'} 'weightb' {'WB'}
                        'b_pyramid' {'B-Pyr'} '+manybframes' {'ManyB'}
                        'sao' {'SAO'} 'mbbrc' {'MBBRC'} 'extbrc' {'ExtBRC'}
                        'adaptiveref' {'AdRef'} 'adaptiveltr' {'AdLTR'} 'fadedetect' {'Fade'}
                    }
                    if ($c.ContainsKey($f)) {
                        $sym = if ($c[$f]) { '+' } else { '-' }
                        $feats += "${fn}:${sym}"
                    }
                }

                # Show modes (PG/FF)
                $modes = @()
                foreach ($vk in $qsv.Codecs.Keys) {
                    $v = $qsv.Codecs[$vk]
                    if ($v._baseCodec -eq $codecKey) { $modes += $v._mode.ToUpper() }
                }

                $rpt.AppendLine("") | Out-Null
                $rpt.AppendLine("        Encode $cName (modes: $($modes -join ', ')):") | Out-Null
                if ($rcModes.Count -gt 0) {
                    $rpt.AppendLine("            RC: $($rcModes -join ', ')") | Out-Null
                }
                $rpt.AppendLine("            $($feats -join '  ')") | Out-Null
            }
        }

        # Decode
        if ($qsv.Decode.Count -gt 0) {
            $rpt.AppendLine("") | Out-Null
            $rpt.AppendLine("        Decode:") | Out-Null
            foreach ($dk in $qsv.Decode.Keys | Sort-Object) {
                $rpt.AppendLine("            $dk = supported") | Out-Null
            }
        }

        # VPP
        if ($qsv.VPP.Count -gt 0) {
            $supported = @($qsv.VPP.GetEnumerator() | Where-Object { $_.Value } | ForEach-Object { $_.Key })
            if ($supported.Count -gt 0) {
                $rpt.AppendLine("        VPP: $($supported -join ', ')") | Out-Null
            }
        }
    } else {
        $rpt.AppendLine("") | Out-Null
        $rpt.AppendLine("    [Intel GPU (QSVEncC)]  Not detected") | Out-Null
    }

    # ── Validation warnings ──
    if ($validation -and $validation.Warnings.Count -gt 0) {
        $rpt.AppendLine("") | Out-Null
        $rpt.AppendLine("    [Compatibility Warnings]") | Out-Null
        foreach ($w in $validation.Warnings) {
            $rpt.AppendLine("    ⚠ $w") | Out-Null
        }
    }

    if ($validation -and $validation.Adjustments.Count -gt 0) {
        $criticals = @($validation.Adjustments | Where-Object { $_.Critical })
        if ($criticals.Count -gt 0) {
            $rpt.AppendLine("") | Out-Null
            $rpt.AppendLine("    [CRITICAL: Commands adjusted for GPU compatibility]") | Out-Null
            foreach ($a in $criticals) {
                $rpt.AppendLine("    ✖ $($a.Encoder): $($a.Param) - encoding will FAIL on this hardware") | Out-Null
            }
        }
    }

    # ── Errors ──
    if ($caps.Errors.Count -gt 0) {
        $rpt.AppendLine("") | Out-Null
        foreach ($e in $caps.Errors) {
            $rpt.AppendLine("    Note: $e") | Out-Null
        }
    }

    # ── Performance tips ──
    $rpt.AppendLine("") | Out-Null
    $rpt.AppendLine("    [Performance Tips]") | Out-Null
    $rpt.AppendLine("    - Run 'NVEncC64 --check-features' or 'QSVEncC64 --check-features' for full details") | Out-Null
    $rpt.AppendLine("    - With multi-GPU setups, use --device <N> (NVEnc) or -d <N> (QSVEnc) to select GPU") | Out-Null
    $rpt.AppendLine("    - --parallel can distribute work across multiple Intel GPUs (QSVEnc)") | Out-Null
    $rpt.AppendLine("    - For dGPU (Arc/discrete): --avsw may outperform --avhw for CPU-friendly codecs") | Out-Null
}
