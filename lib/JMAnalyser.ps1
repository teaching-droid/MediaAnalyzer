# JMAnalyser.ps1 - H.264 Reference Decoder (JM ldecod) integration
# Equivalent of HMAnalyser.ps1 for H.264/AVC files.
# Provides: per-frame QP stats, definitive B-pyramid detection, accurate SPS ref count.
#
# ldecod stdout frame line format:
#   NNNNN(TYPE)   POC  Pic#   QP    SnrY  SnrU  SnrV  Y:U:V  Time(ms)
#   00000(IDR)      0     0     3                       4:2:0     16
#   00002( P )      4     1     6                       4:2:0      5
#   00003( B )      6     3    11                       4:2:0      6   <- B-pyramid reference (uppercase)
#   00001( b )      2     2    12                       4:2:0      4   <- leaf B (lowercase)
#
# Key: uppercase B = B-pyramid reference frame (definitive, not heuristic)
#      lowercase b = non-reference leaf B frame

function Get-JMAnalysis {
    param([string]$FilePath, [int]$MaxFrames = 1000)
    if (-not $Tools.JMDecoder -or -not $Tools.FFmpeg) { return $null }

    # Get file duration for multi-segment strategy
    $duration = 0
    if ($Tools.FFprobe) {
        $dr = Run-Command $Tools.FFprobe @(
            '-v','quiet','-show_entries','format=duration','-of','csv=p=0',"`"$FilePath`""
        ) -TimeoutSeconds 15
        if ($dr.ExitCode -eq 0 -and $dr.StdOut.Trim() -match '[\d.]+') {
            $duration = [double]$dr.StdOut.Trim()
        }
    }

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "MediaAnalyzer_JM_$([System.IO.Path]::GetRandomFileName())"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    try {
        # Scale segments to file duration (same strategy as HMAnalyser)
        $fps = if ($rc -and $rc['FPS']) { [double]$rc['FPS'] } else { 25 }
        $numSegs = if ($duration -gt 1200) { 5 }
                   elseif ($duration -gt 300) { 3 }
                   else { 1 }
        $framesPerSeg = [math]::Ceiling($MaxFrames / $numSegs)
        $segDuration  = [math]::Ceiling($framesPerSeg / $fps) + 1

        $segments = @()
        if ($numSegs -ge 5 -and $duration -gt ($segDuration * 6)) {
            $segments = @(
                @{ Start = 0;                                                                            Label = "beginning" },
                @{ Start = [math]::Floor($duration * 0.25) - [math]::Floor($segDuration / 2);           Label = "25%" },
                @{ Start = [math]::Floor($duration * 0.50) - [math]::Floor($segDuration / 2);           Label = "middle"    },
                @{ Start = [math]::Floor($duration * 0.75) - [math]::Floor($segDuration / 2);           Label = "75%"       },
                @{ Start = [math]::Max(0, [math]::Floor($duration) - $segDuration - 5);                 Label = "end"       }
            )
            Write-Host "      Multi-point sampling: $numSegs x ${segDuration}s segments" -ForegroundColor DarkGray
        } elseif ($numSegs -ge 3 -and $duration -gt ($segDuration * 4)) {
            $segments = @(
                @{ Start = 0;                                                                 Label = "beginning" },
                @{ Start = [math]::Floor($duration / 2) - [math]::Floor($segDuration / 2);  Label = "middle"    },
                @{ Start = [math]::Max(0, [math]::Floor($duration) - $segDuration - 5);     Label = "end"       }
            )
            Write-Host "      Multi-point sampling: 3 x ${segDuration}s segments" -ForegroundColor DarkGray
        } else {
            $segDuration = [math]::Ceiling($MaxFrames / $fps) + 2
            $segments = @( @{ Start = 0; Label = "full" } )
        }

        $allLines         = [System.Collections.ArrayList]::new()
        $rawStreamFFprobe = $null   # first segment kept for FFprobe SPS probe
        $segIdx           = 0

        foreach ($seg in $segments) {
            $segIdx++
            $rawStream = Join-Path $tempDir "seg_${segIdx}.264"

            Write-Host "      [$segIdx/$($segments.Count)] Extracting $($seg.Label) (@$($seg.Start)s, ${segDuration}s)..." -ForegroundColor DarkGray
            $r = Run-Command $Tools.FFmpeg @(
                '-ss', $seg.Start.ToString(),
                '-i',  "`"$FilePath`"",
                '-t',  $segDuration.ToString(),
                '-c:v','copy',
                '-bsf:v','h264_mp4toannexb',
                '-an','-sn','-y',
                "`"$rawStream`""
            ) -TimeoutSeconds 60 -StatusLabel "ffmpeg extract $($seg.Label)"

            if ($r.ExitCode -ne 0 -or -not (Test-Path $rawStream)) {
                Write-Host "      Extraction failed for $($seg.Label) (exit=$($r.ExitCode))" -ForegroundColor Yellow
                continue
            }
            if ($segIdx -eq 1) { $rawStreamFFprobe = $rawStream }

            Write-Host "      [$segIdx/$($segments.Count)] Decoding $($seg.Label)..." -ForegroundColor DarkGray

            # ldecod requires -d to suppress 'decoder.cfg not found' noise; use NUL config trick.
            # All params via -p. DecFrmNum caps decode count for speed.
            # OutputFile=NUL discards YUV (saves disk I/O for large frames).
            $jmArgs = @(
                '-p', "InputFile=`"$rawStream`"",
                '-p', 'OutputFile=NUL',
                '-p', "DecFrmNum=$framesPerSeg"
            )

            if ($script:DebugMode) {
                Write-Host "      [DEBUG] JM input: $rawStream ($((Get-Item $rawStream).Length) bytes)" -ForegroundColor Magenta
            }

            $jr = Run-Command $Tools.JMDecoder $jmArgs -TimeoutSeconds 300 -StatusLabel "JM Decoder ($($seg.Label))"
            $segOutput = "$($jr.StdOut)`n$($jr.StdErr)"

            if ($script:DebugMode) {
                $lineCount = ($segOutput -split "`n").Count
                Write-Host "      [DEBUG] JM output: $lineCount lines, $($segOutput.Length) chars" -ForegroundColor Magenta
            }

            foreach ($line in ($segOutput -split "`n")) {
                [void]$allLines.Add($line)
            }
        }

        if ($allLines.Count -eq 0) { return $null }

        $result = @{ Source = "JM 19.0 H.264 Reference Decoder" }

        # ── 1. Parse frame table → QP stats + B-pyramid detection ──
        # Frame line regex: NNNNN(TYPE)   POC  Pic#   QP
        # TYPE examples: IDR, P, B (uppercase=B-ref), b (lowercase=leaf-B), I
        # Note: PowerShell hash keys are case-insensitive, so 'B' and 'b' cannot coexist.
        # Use 'Bref' for uppercase B (B-pyramid reference) and 'Bleaf' for lowercase b (leaf B).
        $allQPs  = [System.Collections.ArrayList]::new()
        $byType  = @{ IDR  = [System.Collections.ArrayList]::new()
                      P    = [System.Collections.ArrayList]::new()
                      Bref = [System.Collections.ArrayList]::new()
                      Bleaf= [System.Collections.ArrayList]::new()
                      I    = [System.Collections.ArrayList]::new() }
        $hasBRef  = $false
        $hasBLeaf = $false

        foreach ($line in $allLines) {
            # Match: 5-digit frame number, parenthesised type, POC, Pic#, QP
            if ($line -match '^\d{5}\(([A-Za-z ]+)\)\s+\d+\s+\d+\s+(\d+)') {
                $rawType = $Matches[1].Trim()
                $qpVal   = [int]$Matches[2]

                [void]$allQPs.Add($qpVal)

                # -ceq is case-sensitive equals — essential to distinguish 'B' from 'b'
                if     ($rawType -ceq 'IDR') { [void]$byType['IDR'].Add($qpVal) }
                elseif ($rawType -ceq 'P')   { [void]$byType['P'].Add($qpVal) }
                elseif ($rawType -ceq 'B')   { [void]$byType['Bref'].Add($qpVal);  $hasBRef  = $true }
                elseif ($rawType -ceq 'b')   { [void]$byType['Bleaf'].Add($qpVal); $hasBLeaf = $true }
                elseif ($rawType -ceq 'I')   { [void]$byType['I'].Add($qpVal) }
            }
        }

        if ($allQPs.Count -eq 0) {
            Write-Host "      JM: No frame QP data parsed (check ldecod output format)" -ForegroundColor Yellow
            return $null
        }

        # Overall stats
        $allArr = @($allQPs)
        $qpMin  = ($allArr | Measure-Object -Minimum).Minimum
        $qpMax  = ($allArr | Measure-Object -Maximum).Maximum
        $qpAvg  = [math]::Round(($allArr | Measure-Object -Average).Average, 1)

        # Std dev
        $sqDiff = ($allArr | ForEach-Object { [math]::Pow($_ - $qpAvg, 2) } | Measure-Object -Average).Average
        $qpSD   = [math]::Round([math]::Sqrt($sqDiff), 1)

        # Median
        $sorted  = $allArr | Sort-Object
        $midIdx  = [math]::Floor($sorted.Count / 2)
        $qpMed   = if ($sorted.Count % 2 -eq 0) { ($sorted[$midIdx-1] + $sorted[$midIdx]) / 2 } else { $sorted[$midIdx] }

        $stats = @{
            Min    = $qpMin
            Max    = $qpMax
            Avg    = $qpAvg
            Median = $qpMed
            StdDev = $qpSD
            Count  = $allQPs.Count
        }

        # Per-type stats helper
        $typeStats = @{}
        foreach ($t in $byType.Keys) {
            $arr = @($byType[$t])
            if ($arr.Count -gt 0) {
                $tAvg = [math]::Round(($arr | Measure-Object -Average).Average, 1)
                $typeStats[$t] = @{
                    Min   = ($arr | Measure-Object -Minimum).Minimum
                    Max   = ($arr | Measure-Object -Maximum).Maximum
                    Avg   = $tAvg
                    Count = $arr.Count
                }
            }
        }

        $result.QP = @{
            Stats      = $stats
            ByType     = $typeStats
            FrameCount = $allQPs.Count
            BPyramid   = $hasBRef    # DEFINITIVE: uppercase B = B-pyramid reference frames exist
            HasBLeaf   = $hasBLeaf
            Source     = "JM 19.0 H.264 Reference Decoder"
        }

        # ── 2. FFprobe on raw Annex B stream → accurate SPS fields ──
        # FFprobe on .264 bitstream gives correct max_num_ref_frames,
        # unlike on the MKV container where refs=1 is often returned.
        $sps = @{}
        if ($rawStreamFFprobe -and (Test-Path $rawStreamFFprobe) -and $Tools.FFprobe) {
            $fpR = Run-Command $Tools.FFprobe @(
                '-v','quiet','-print_format','json',
                '-show_streams','-select_streams','v:0',
                "`"$rawStreamFFprobe`""
            ) -TimeoutSeconds 30
            if ($fpR.ExitCode -eq 0 -and $fpR.StdOut) {
                try {
                    $fpJson = $fpR.StdOut | ConvertFrom-Json
                    $fpVs   = $fpJson.streams | Select-Object -First 1
                    if ($fpVs) {
                        $sps['refs']         = if ($fpVs.refs -and [int]$fpVs.refs -gt 0) { [int]$fpVs.refs } else { $null }
                        $sps['profile']      = $fpVs.profile
                        $sps['level']        = $fpVs.level
                        $sps['pix_fmt']      = $fpVs.pix_fmt
                        $sps['has_b_frames'] = $fpVs.has_b_frames
                        $sps['width']        = $fpVs.width
                        $sps['height']       = $fpVs.height
                        if ($script:DebugMode) {
                            Write-Host "      [DEBUG] JM FFprobe on .264: refs=$($sps['refs']) profile=$($sps['profile']) level=$($sps['level'])" -ForegroundColor Magenta
                        }
                    }
                } catch {
                    if ($script:DebugMode) { Write-Host "      [DEBUG] JM FFprobe parse error: $_" -ForegroundColor Magenta }
                }
            }
        }
        $result.SPS = $sps

        # Summary
        $bPyrStr = if ($hasBRef) { "YES (confirmed)" } else { "NO" }
        Write-Host "      JM: $($allQPs.Count) frames decoded | QP avg=$($stats.Avg) min=$($stats.Min) max=$($stats.Max) | B-pyramid=$bPyrStr" -ForegroundColor Green
        if ($sps['refs']) {
            Write-Host "      JM: SPS ref frames=$($sps['refs']) (overrides FFprobe/container value)" -ForegroundColor Green
        }

        return $result
    }
    catch {
        Write-Host "      JM Analyser error: $($_.Exception.Message)" -ForegroundColor Red
        if ($script:DebugMode) { Write-Host "      [DEBUG] $($_.ScriptStackTrace)" -ForegroundColor Magenta }
        return $null
    }
    finally {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}