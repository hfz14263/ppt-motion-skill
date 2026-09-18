# dsh-ppt-office-motion :: build the known-answer calibration deck + video
#
# Produces a deck whose animation timing is HARD-CODED, so analyze_video.py can be
# checked against ground truth instead of against intuition.
#
#   slide 1 : three circles fade in for 1.5s / 1.0s / 0.5s, separated by 0.5s gaps
#   slide 2 : a text box, to exercise slide splitting
#
# The gaps come from Timing.TriggerDelayTime. Without it, afterEffect chains the
# fades back-to-back into ONE continuous 3.0s motion, which the analyser then
# correctly reports as a single burst -- easily mistaken for a measurement bug.
# See references/video-analysis-limits.md.
#
# Output: <out>/calib.pptx and <out>/calib.mp4
#
# Usage:
#   powershell -NoProfile -File scripts/make_calibration.ps1 [-OutDir <dir>] [-Height 720]

[CmdletBinding()]
param(
  [string]$OutDir = '',
  [int]$Height = 720
)

$ErrorActionPreference = 'Stop'

# scripts/ -> repo root; explicit concatenation to avoid Join-Path surprises
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $OutDir) {
  $OutDir = $repoRoot + '\tests\calib'
} elseif (-not [System.IO.Path]::IsPathRooted($OutDir)) {
  $OutDir = (Get-Location).Path + '\' + $OutDir
}
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = [System.IO.Path]::GetFullPath($OutDir)
$deck = $OutDir + '\calib.pptx'
$video = $OutDir + '\calib.mp4'
Remove-Item $deck, $video -Force -ErrorAction SilentlyContinue
Write-Host ("outdir : {0}" -f $OutDir)

# keep Office's scratch files inside the target dir
$tmp = Join-Path $OutDir '.office-tmp'
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$env:TEMP = $tmp; $env:TMP = $tmp
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add(1)
$pres.PageSetup.SlideSize = 15          # 16:9, 960x540 pt

$durations = @(1.5, 1.0, 0.5)           # GROUND TRUTH
$gap = 0.5

$s1 = $pres.Slides.Add(1, 12)
$s1.FollowMasterBackground = 0
$s1.Background.Fill.Solid()
$s1.Background.Fill.ForeColor.RGB = 0xFFFFFF
$seq = $s1.TimeLine.MainSequence
for ($i = 0; $i -lt 3; $i++) {
  $x = 80 + $i * 280
  $sh = $s1.Shapes.AddShape(9, $x, 180, 220, 220)          # oval
  $sh.Name = "circle$i"
  # high contrast on purpose: low-contrast fades are NOT measurable from video
  # (see references/video-analysis-limits.md section 2)
  $sh.Fill.ForeColor.RGB = 0x000000                         # BGR black
  $sh.Line.Visible = 0
  $e = $seq.AddEffect($sh, 10, 0, 3)                        # msoAnimEffectFade, afterEffect
  $e.Timing.Duration = $durations[$i]
  if ($i -gt 0) { $e.Timing.TriggerDelayTime = $gap }        # real gap, not chained
}

$s2 = $pres.Slides.Add(2, 12)
$s2.FollowMasterBackground = 0
$s2.Background.Fill.Solid()
$s2.Background.Fill.ForeColor.RGB = 0xFFFFFF
$t = $s2.Shapes.AddTextbox(1, 100, 200, 760, 120)
$t.TextFrame.TextRange.Text = 'SECOND SLIDE'
$t.TextFrame.TextRange.Font.Size = 48
$t.TextFrame.TextRange.Font.Color.RGB = 0x000000

$pres.SaveAs($deck)
Write-Host ("deck   : {0}" -f $deck)
Write-Host ("truth  : fades {0} s, gaps {1} s" -f ($durations -join ' / '), $gap)
# create the video (asynchronous; poll until done)
$pres.CreateVideo($video, $true, 2, $Height, 30, 85)
$sw = [System.Diagnostics.Stopwatch]::StartNew()
while ($pres.CreateVideoStatus -ne 3 -and $sw.Elapsed.TotalSeconds -lt 600) {
  Start-Sleep -Milliseconds 1000
}
$status = $pres.CreateVideoStatus
$pres.Close(); $ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue

if (Test-Path $video) {
  Write-Host ("video  : {0} ({1:N2} MB) status={2}" -f $video, ((Get-Item $video).Length / 1MB), $status)
  Write-Host ""
  Write-Host "now run: python scripts/analyze_video.py --video `"$video`" --json"
} else {
  Write-Host "video export FAILED (status=$status)"
  exit 1
}
