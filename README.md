# Media Compression Analyzer v2.0

A PowerShell-based tool for deep analysis of video encoding parameters, bitstream-level inspection, GPU hardware capability detection, encoder command reconstruction, and verification encoding with objective quality metrics.

## What It Does

Given a video file, MediaAnalyzer:

1. **Probes every encoding parameter** — from container metadata down to raw HEVC bitstream (SPS/PPS/VUI/SEI) using the HM Reference Decoder
2. **Analyzes frame structure** — GOP patterns, B-frame distribution, keyframe intervals, scene cuts
3. **Extracts HDR metadata** — Dolby Vision RPU, HDR10 mastering display, MaxCLL/MaxFALL from both container tags and SEI NAL units
4. **Measures bitrate distribution** — statistical analysis with histograms, VBV detection, peak/sustained rates
5. **Detects GPU hardware capabilities** — queries NVIDIA and Intel GPUs for supported features, rate control modes, codec limits
6. **Reconstructs encoder commands** — generates ready-to-use x265, NVEncC, and QSVEncC commands that reproduce the source encoding
7. **Runs verification encodes** — encodes a test segment with each GPU encoder, then verifies the output against the source for correctness
8. **Measures quality objectively** — VMAF, PSNR, and SSIM comparison of encoded output vs source

---

## Requirements

### System
- **Windows 10/11** (PowerShell 5.1+)
- **NVIDIA GPU** (Turing or newer recommended) for NVEncC features
- **Intel GPU** (6th gen or newer) for QSVEncC features — integrated (UHD 770, etc.) or discrete (Arc)
- Either or both GPUs are optional; the tool adapts to what's available

### Tools

All tools go in the `tools\` subfolder (or anywhere on PATH). The analyzer auto-discovers them at startup.

| Tool | Required | Version Tested | Purpose | Download |
|------|----------|----------------|---------|----------|
| **FFprobe** | Yes (core) | 2026-02-09 gyan.dev | Stream analysis, frame probing, GOP structure | [gyan.dev/ffmpeg](https://www.gyan.dev/ffmpeg/builds/) (included with FFmpeg) |
| **FFmpeg** | Yes (core) | 2026-02-09 gyan.dev | Segment extraction, scene detection, VMAF/PSNR/SSIM | [gyan.dev/ffmpeg](https://www.gyan.dev/ffmpeg/builds/) — get the **full build** for libvmaf support |
| **MediaInfo** | Recommended | v26.01 | Container metadata, HDR format detection, encoder identification | [mediaarea.net](https://mediaarea.net/en/MediaInfo/Download/Windows) — CLI version |
| **CheckBitrate** | Recommended | v0.06 by rigaya | Per-second bitrate distribution, VBV analysis | [github.com/rigaya/CheckBitrate](https://github.com/rigaya/CheckBitrate) |
| **TAppDecoderAnalyser** | Recommended | HM v18.0 | HEVC reference decoder — reads raw SPS/PPS/VUI/SEI/QP from bitstream | [vcgit.hhi.fraunhofer.de](https://vcgit.hhi.fraunhofer.de/jvet/HM) — build the `TAppDecoderAnalyser` target |
| **ldecod** (JM) | For H.264 | JM 19.0 | H.264/AVC reference decoder — QP stats, B-pyramid confirmation, SPS ref count | [iphome.hhi.de/suehring](https://iphome.hhi.de/suehring/tml.htm) or build from JM source |
| **dovi_tool** | For DV content | v2.3.1 | Dolby Vision RPU stripping (required before HM can decode DV streams) | [github.com/quietvoid/dovi_tool](https://github.com/quietvoid/dovi_tool/releases) |
| **NVEncC64** | For NVIDIA GPU | v9.10 (r3505) | GPU encode verification, feature detection | [github.com/rigaya/NVEnc](https://github.com/rigaya/NVEnc/releases) |
| **QSVEncC64** | For Intel GPU | v8.04 (r3864) | GPU encode verification, feature detection | [github.com/rigaya/QSVEnc](https://github.com/rigaya/QSVEnc/releases) |

**Minimum for basic analysis:** FFprobe + FFmpeg
**Full analysis with verification:** All tools above

### FFmpeg Build Notes

For VMAF quality metrics, your FFmpeg build must include `--enable-libvmaf`. The full builds from gyan.dev include it. You can verify with:
```
ffmpeg -filters 2>&1 | findstr libvmaf
```

---

## Folder Structure

```
MediaAnalyzer\
├── Analyze-Media.ps1          # Main entry point — orchestrates the entire pipeline
├── Compare-Videos.ps1         # Multi-encode comparison — side-by-side quality analysis
├── README.md                  # This file
├── lib\
│   ├── Helpers.ps1            # Utility functions (Run-Command, Find-Tool, Write-Field, etc.)
│   ├── Collectors.ps1         # Data collection — FFprobe, MediaInfo, CheckBitrate, scene cuts
│   ├── HMAnalyser.ps1         # HM Reference Decoder integration — bitstream-level analysis
│   ├── JMAnalyser.ps1         # JM Reference Decoder integration — H.264/AVC bitstream analysis
│   ├── ReportWriter.ps1       # Report formatting — writes all analysis sections to text
│   ├── Reconstruction.ps1     # Encoder command reconstruction — generates x265/NVEnc/QSVEnc commands
│   ├── GPUCapabilities.ps1    # GPU feature detection — parses --check-features output
│   └── Verification.ps1       # Verification encodes — test encode + 6-step validation pipeline
└── tools\                     # Place all external tools here
    ├── ffmpeg.exe
    ├── ffprobe.exe
    ├── mediainfo.exe
    ├── CheckBitrate.exe
    ├── dovi_tool.exe
    ├── TAppDecoderAnalyser.exe
    ├── ldecod.exe
    ├── NVEncC64\
    │   └── NVEncC64.exe       # With its DLLs in same folder
    └── QSVEncC64\
        └── QSVEncC64.exe      # With its DLLs in same folder
```

---

## File Descriptions

### `Analyze-Media.ps1` — Main Script
The entry point. Parses parameters, discovers tools, orchestrates the 9-step analysis pipeline, runs GPU detection, triggers verification, and writes the final report.

### `Compare-Videos.ps1` — Multi-Encode Comparison
Standalone tool for comparing multiple encodes of the same source side-by-side. Runs quality metrics (VMAF/PSNR/SSIM) and bitrate analysis across different encoder configurations to help choose optimal settings.

### `lib\Helpers.ps1` — Utilities
Core helper functions used throughout:
- `Run-Command` — Executes external tools with timeout, captures stdout/stderr, shows elapsed time
- `Find-Tool` — Searches the tools folder and PATH for executables
- `Write-Field` / `Write-Section` — Consistent report formatting
- `Format-Bitrate` / `Format-Size` — Human-readable bitrate and file size display

### `lib\Collectors.ps1` — Data Collection
Gathers raw data from external tools:
- `Get-ProbeJson` — FFprobe stream/format metadata as JSON
- `Get-FrameData` — Frame-by-frame analysis with distributed multi-segment sampling
- `Get-MultiPointFrames` — Alternative 5-point sampling (beginning, 25%, 50%, 75%, end)
- `Get-EncoderInfo` — Encoder identification from container metadata
- `Get-MediaInfoData` — Full MediaInfo JSON + text output
- `Get-CheckBitrateData` — Bitrate-over-time CSV with statistics
- `Get-SceneCuts` — Scene change detection via FFmpeg

### `lib\HMAnalyser.ps1` — HM Reference Decoder
The deepest analysis layer. Extracts a raw HEVC stream, strips Dolby Vision RPU if present, and feeds it through the HM Reference Decoder Analyser to get:
- **SPS** — Sequence parameters (CTU size, bit depth, SAO, AMP, scaling lists, DPB)
- **PPS** — Picture parameters (QP delta, weighted prediction, tiles, WPP, sign hiding)
- **VUI** — Color primaries, transfer characteristics, matrix coefficients, frame rate
- **SEI** — HDR10 mastering display, MaxCLL/MaxFALL, content light level
- **Slice** — Per-slice parameters, merge candidates, temporal MVP
- **QP** — Per-frame quantization parameters with I/P/B breakdown
- **CABAC** — Coding unit statistics (skip mode, merge mode, SAO overhead, MV cost)

### `lib\JMAnalyser.ps1` — JM Reference Decoder (H.264/AVC)
The H.264/AVC counterpart to HMAnalyser. Uses the JM Reference Decoder to extract bitstream-level parameters from AVC content, including SPS/PPS headers, slice parameters, and quantization data.

### `lib\ReportWriter.ps1` — Report Formatting
Transforms collected data into the human-readable text report. Handles all sections from container format through HDR metadata to compression quality estimates. Also populates the `$rc` (reconstruction context) hashtable that feeds into command generation.

### `lib\Reconstruction.ps1` — Command Generation
The "big payoff" — generates encoder commands that reproduce the source encoding:
- **Bitstream-derived parameters** — confidence-tagged parameter list (HIGH/MED/LOW)
- **Encoding recommendations** — presets, rate control modes, AQ settings
- **Raw source commands** — x265, NVEncC, QSVEncC for encoding from uncompressed source
- **Transcode commands** — commands for re-encoding from this file with HDR/DV passthrough
- **Dolby Vision workflow** — step-by-step RPU extraction, encoding, injection, muxing

### `lib\GPUCapabilities.ps1` — GPU Detection
Runs `--check-features` on NVEncC and QSVEncC, parses the output into structured capability data:
- Supported codecs and profiles per GPU
- Maximum B-frames, reference frames, lookahead depth
- Feature flags (SAO, AQ-T, weighted prediction, temporal filter, 10-bit)
- Rate control modes (CBR, VBR, CQP, ICQ, QVBR, etc.)
- Decode capabilities and VPP (video processing) features
- Validates generated commands against actual GPU limits

### `lib\Verification.ps1` — Verification Pipeline
Encodes a test segment with each available GPU encoder, then runs a 6-step validation:

| Step | What It Checks |
|------|---------------|
| **[1/6] FFprobe** | Codec, resolution, bit depth, color primaries, HDR side data, DV RPU |
| **[2/6] Frame/GOP** | Frame type distribution, GOP structure, B-frame patterns |
| **[3/6] MediaInfo** | HDR format, mastering display, MaxCLL, encoder identification |
| **[4/6] HM Decoder** | QP range, reference structure, CTU size, SAO, VUI color, weighted prediction |
| **[5/6] CheckBitrate** | Average/peak bitrate accuracy, VBR pattern, VBV ceiling |
| **[6/6] Quality** | VMAF (perceptual), SSIM (structural), PSNR (signal-level) vs source |

---

## Usage

### Basic Analysis
```powershell
.\Analyze-Media.ps1 -Path "D:\Movie.mkv"
```

### Full Analysis with GPU Verification
```powershell
.\Analyze-Media.ps1 -Path "D:\Movie.mkv" -CheckGPU -Verify -VerifyDuration 120
```

### Batch Processing
```powershell
.\Analyze-Media.ps1 -Path "D:\Movies" -Recurse -CheckGPU
```

### Debug Mode
```powershell
.\Analyze-Media.ps1 -Path "D:\Movie.mkv" -CheckGPU -Verify -VerifyDuration 120 -DebugMode
```

### Compare Multiple Encodes
```powershell
# Compare original vs one encode
.\Compare-Videos.ps1 original.mkv encode_nvenc.mkv

# Compare original vs multiple encodes side-by-side
.\Compare-Videos.ps1 original.mkv enc_nvenc.mkv enc_qsvenc.mkv enc_x265.mkv

# Extended metrics sample (120s instead of default 60s)
.\Compare-Videos.ps1 original.mkv enc1.mkv enc2.mkv -SampleDuration 120

# Parameter comparison only (skip VMAF/PSNR/SSIM)
.\Compare-Videos.ps1 original.mkv enc1.mkv enc2.mkv -SkipMetrics
```

---

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Path` | string | *(required)* | Path to a single media file or a folder containing media files |
| `-Recurse` | switch | off | When `-Path` is a folder, also scan subfolders |
| `-OutputDir` | string | auto | Override output directory. Default: `MediaAnalysis` subfolder next to the source file. Falls back to `%USERPROFILE%\Documents\MediaAnalysis` if source is on a read-only drive (optical, ISO mount, read-only network share) |
| `-MaxAnalysisFrames` | int | 1000 | Number of frames to analyze via distributed sampling. Higher values give more accurate GOP/QP statistics but take longer. The frames are spread evenly across the file |
| `-FullFrameScan` | switch | off | Analyze every frame in the file (can be very slow for long files). Overrides `-MaxAnalysisFrames` |
| `-CheckGPU` | switch | off | Detect GPU capabilities via `--check-features` and validate generated commands against hardware limits. Automatically enabled when `-Verify` is used |
| `-Verify` | switch | off | Run verification encodes with each available GPU encoder and validate the output. Implies `-CheckGPU` |
| `-VerifyDuration` | int | 0 | Duration in seconds for the verification encode segment. `0` = quick mode (~13s). `120` = extended mode (recommended for thorough testing). Implies `-Verify` |
| `-CheckBitrateInterval` | double | 0 | Override the CheckBitrate sampling interval in seconds. `0` = auto (file duration / 1000, typically ~4s for a 1-hour file) |
| `-SkipCheckBitrate` | switch | off | Skip the CheckBitrate analysis step entirely |
| `-SkipQP` | switch | off | Skip QP analysis via HM Reference Decoder (saves significant time but loses bitstream-level detail) |
| `-ExportJson` | switch | off | Export collected raw data as a JSON file alongside the text report |
| `-DebugMode` | switch | off | Show detailed debug output for every analysis step — tool exit codes, data sizes, intermediate values. Useful for troubleshooting |

---

## Output

### Text Report
Written to `<OutputDir>\<filename>_analysis.txt`. Contains all analysis sections with formatted tables.

### Bitrate CSV
Written to `<OutputDir>\<filename>_bitrate.csv`. Per-interval bitrate data from CheckBitrate, suitable for graphing.

### JSON Export (optional)
Written to `<OutputDir>\<filename>_analysis.json` when `-ExportJson` is used. Raw collected data from all tools.

---

## Quality Metrics Thresholds

The verification pipeline uses these thresholds to evaluate encode quality:

| Metric | PASS | WARN/INFO | FAIL |
|--------|------|-----------|------|
| **VMAF** | ≥ 93 | ≥ 80 | < 80 |
| **PSNR** | ≥ 40 dB | ≥ 35 dB | < 35 dB |
| **SSIM** | ≥ 0.95 | ≥ 0.90 | < 0.90 |

VMAF is the most meaningful metric — it models human perceptual quality. A score of 93+ indicates transparent (visually lossless) quality. PSNR and SSIM provide complementary signal-level and structural measures.

---

## HDR Metadata Sourcing

HDR metadata (MaxCLL, MaxFALL, mastering display primaries/luminance) can come from multiple sources. The analyzer uses this priority order:

1. **HM SEI messages** (bitstream) — Most reliable. Read directly from Content Light Level and Mastering Display Colour Volume SEI NAL units in the HEVC stream. Available even when container metadata is missing.
2. **MediaInfo** (container) — Fallback. Reads from MKV/MP4 container tags. Only used if HM didn't find the values.

The report's "HDR Metadata for Passthrough" section notes where each value came from (SEI vs MediaInfo), so you know the provenance.

---

## Multi-GPU Support

The tool detects all available GPU devices at startup:

```
  Detecting GPU capabilities...
    NVIDIA Device #0: NVIDIA GeForce RTX 3080
    NVIDIA Device #1: NVIDIA GeForce RTX 3080
      Multi-GPU: Use --device <N> to select
    Intel  Device #1: Intel UHD Graphics 770
```

For verification encodes, the default GPU is used. The generated commands include notes about device selection:
- **NVEncC**: `--device <N>` to select a specific NVIDIA GPU
- **QSVEncC**: `-d <N>` to select a specific Intel GPU

---

## Typical Runtime

For a ~1 hour 4K HDR file with full verification:

| Step | Time |
|------|------|
| Tool discovery + versions | ~5s |
| GPU capability detection | ~10s |
| FFprobe + Frame/GOP | ~25s |
| HM Reference Decoder (5 segments) | ~3 min |
| MediaInfo + Scene cuts | ~1.5 min |
| CheckBitrate | ~30s |
| NVEncC verification encode (120s) | ~2.5 min |
| NVEncC verification analysis | ~4 min (incl. HM + VMAF) |
| QSVEncC verification encode (120s) | ~35s |
| QSVEncC verification analysis | ~4 min (incl. HM + VMAF) |
| **Total** | **~17-21 min** |

Without verification (`-CheckGPU` only): ~5-6 minutes.
Without GPU detection (basic analysis): ~4-5 minutes.

---

## Confidence Levels

Reconstructed parameters are tagged with confidence levels:

| Level | Meaning | Source |
|-------|---------|--------|
| **HIGH** | Exact value from bitstream or metadata | SPS/PPS/VUI, FFprobe, MediaInfo |
| **MED** | Derived from frame analysis | GOP patterns, QP statistics, B-frame runs |
| **LOW** | Heuristic estimate | Rate control mode, AQ settings, preset |

---

## Supported Codecs

The full analysis pipeline (HM decoder, bitstream parsing) works with **HEVC/H.265**. The JM decoder extends deep analysis to **H.264/AVC** (QP stats, B-pyramid, SPS reference counts). Basic analysis (FFprobe, MediaInfo, frame structure, bitrate) works with any codec FFprobe supports.

---

## Troubleshooting

### HM Decoder fails with Dolby Vision content
The HM decoder cannot parse DV RPU NAL units. The tool automatically strips them using `dovi_tool` before decoding. Make sure `dovi_tool.exe` is in the tools folder.

### VMAF returns no score
Check that your FFmpeg build includes libvmaf (`ffmpeg -filters 2>&1 | findstr libvmaf`). If not, the tool falls back to SSIM + PSNR only.

### CheckBitrate creates CSV in wrong location
CheckBitrate writes CSVs next to the input file. The tool copies the input to the work directory first to control output location.

### Read-only source (ISO mount, optical drive)
The tool detects read-only sources and automatically redirects output to `%USERPROFILE%\Documents\MediaAnalysis`. You can override with `-OutputDir`.

### NVEncC/QSVEncC not detected
Place the encoder executables (with their DLLs) in subfolders under `tools\`. The tool searches: `tools\NVEncC64\NVEncC64.exe`, `tools\NVEncC64.exe`, and system PATH.

### Use `-DebugMode` for diagnostics
Every step prints detailed debug output (magenta text) showing tool exit codes, data sizes, parsed values, and intermediate results.
