# dsh-ppt-office-motion :: probe whether Presentation.CreateVideo works HERE
#
# Upstream observed CreateVideo failing for every parameter combination; another
# machine exports fine. The capability is build/COM-driver dependent, so this
# script measures the CURRENT machine instead of trusting either report.
#
# It prints the Office build fingerprint first, then tries each parameter set and
# reports which ones actually produced a file.
#
# Usage:
#   powershell -NoProfile -File scripts/probe_createvideo.ps1 [-OutDir <dir>]
#
# Exit code: 0 if at least one combination produced a video, 1 if none did.

[CmdletBinding()]
param(
  [string]$OutDir = ''
)

$ErrorActionPreference = 'Continue'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $OutDir) { $OutDir = $repoRoot + '\probe-createvideo' }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = [System.IO.Path]::GetFullPath($OutDir)
$tmp = $OutDir + '\.office-tmp'
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$env:TEMP = $tmp; $env:TMP = $tmp

function Say($m) { Write-Host $m }

# ---------------------------------------------------------------- fingerprint
Say "== Office build fingerprint =="
$c2r = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
if (Test-Path $c2r) {
  $k = Get-ItemProperty $c2r
  foreach ($n in 'VersionToReport', 'ProductReleaseIds', 'Platform', 'UpdateChannel') {
    if ($k.$n) { Say ("  {0,-20} {1}" -f $n, $k.$n) }
  }
}
foreach ($p in @(
    "$env:ProgramFiles\Microsoft Office\root\Office16\POWERPNT.EXE",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\POWERPNT.EXE")) {
  if (Test-Path $p) {
    $vi = (Get-Item $p).VersionInfo
    Say ("  POWERPNT.exe         {0}" -f $vi.FileVersion)
    Say ("  path                 {0}" -f $p)
  }
}
Say ("  OS                   {0}" -f [System.Environment]::OSVersion.VersionString)
Say ""

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

# ------------------------------------------------------------------ test deck
# A busy deck on purpose: a one-shape deck compresses to identical bytes whatever
# the quality, so it cannot show whether the parameters were honoured.
$deck = $OutDir + '\probe-src.pptx'
Remove-Item $deck -Force -ErrorAction SilentlyContinue
$ppt = New-Object -ComObject PowerPoint.Application
Say ("  PowerPoint COM       version {0}" -f $ppt.Version)
$pres = $ppt.Presentations.Add(1)
$pres.PageSetup.SlideSize = 15
for ($p = 1; $p -le 2; $p++) {
  $s = $pres.Slides.Add($p, 12)
  $s.FollowMasterBackground = 0
  $s.Background.Fill.Solid()
  $s.Background.Fill.ForeColor.RGB = 0xFFFFFF
  $seq = $s.TimeLine.MainSequence
  for ($i = 0; $i -lt 12; $i++) {
    $x = 40 + ($i % 4) * 230
    $y = 40 + [math]::Floor($i / 4) * 160
    $sh = $s.Shapes.AddShape(9, $x, $y, 200, 130)
    $sh.Fill.ForeColor.RGB = (($i * 17) % 256) * 65536 + (($i * 43) % 256) * 256 + (($i * 91) % 256)
    $sh.Line.Visible = 0
    $tb = $s.Shapes.AddTextbox(1, $x + 6, $y + 6, 188, 118)
    $tb.TextFrame.TextRange.Text = ("Block {0} line one`rline two, more glyphs to compress" -f $i)
    $tb.TextFrame.TextRange.Font.Size = 11
    ($seq.AddEffect($sh, 10, 0, 3)).Timing.Duration = 0.4
    ($seq.AddEffect($tb, 2, 0, 3)).Timing.Duration = 0.4
  }
}
$pres.SaveAs($deck)
$pres.Close()
Say ("  probe deck           {0:N0} bytes" -f (Get-Item $deck).Length)
Say ""

# ------------------------------------------------------------------- attempts
function Try-Video {
  param([string]$Label, [int]$ReadOnly, [int]$Res, [int]$Fps, [int]$Q)
  $file = $OutDir + '\' + ($Label -replace '[^A-Za-z0-9]', '_') + '.mp4'
  Remove-Item $file -Force -ErrorAction SilentlyContinue
  $p = $null
  try {
    $p = $ppt.Presentations.Open($deck, $ReadOnly, 0, 1)
  } catch {
    Say ("  {0,-30} OPEN-FAIL   {1}" -f $Label, $_.Exception.Message)
    return $false
  }
  try {
    $p.CreateVideo.Invoke(@($file, $true, 2, $Res, $Fps, $Q)) | Out-Null
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $st = 'n/a'
    while ($sw.Elapsed.TotalSeconds -lt 180) {
      Start-Sleep -Milliseconds 1000
      try { $st = $p.CreateVideoStatus } catch { $st = 'err' }
      if ($st -eq 3 -or $st -eq 4) { break }
    }
    if (Test-Path $file) {
      Say ("  {0,-30} OK  status={1} {2,10:N0} bytes  {3}s" -f $Label, $st, (Get-Item $file).Length, [int]$sw.Elapsed.TotalSeconds)
      return $true
    }
    Say ("  {0,-30} NO-FILE status={1}" -f $Label, $st)
    return $false
  } catch {
    Say ("  {0,-30} CALL-FAIL {1}" -f $Label, $_.Exception.Message)
    return $false
  } finally {
    try { $p.Close() } catch { }
  }
}

Say "== CreateVideo attempts (Resolution / Quality / Fps / open mode) =="
$ok = $false
$ok = (Try-Video -Label 'RO 480p 30fps q85'  -ReadOnly 1 -Res 480  -Fps 30 -Q 85) -or $ok
$ok = (Try-Video -Label 'RO 720p 30fps q85'  -ReadOnly 1 -Res 720  -Fps 30 -Q 85) -or $ok
$ok = (Try-Video -Label 'RO 1080p 30fps q85' -ReadOnly 1 -Res 1080 -Fps 30 -Q 85) -or $ok
$ok = (Try-Video -Label 'RO 720p 30fps q1'   -ReadOnly 1 -Res 720  -Fps 30 -Q 1)  -or $ok
$ok = (Try-Video -Label 'RO 720p 30fps q0'   -ReadOnly 1 -Res 720  -Fps 30 -Q 0)  -or $ok
$ok = (Try-Video -Label 'RW 720p 30fps q85'  -ReadOnly 0 -Res 720  -Fps 30 -Q 85) -or $ok

$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue

Say ""
if ($ok) {
  Say "RESULT: CreateVideo WORKS on this machine -- you can export MP4 to review motion."
  Say "        (player.py is still better for per-shape layers and exact scrubbing.)"
  Say ("        artifacts: {0}" -f $OutDir)
  exit 0
} else {
  Say "RESULT: CreateVideo produced NO file for any combination on this machine."
  Say "        Use the layer approach (com-pitfalls.md 16.1-16.3) or motion.py player."
  exit 1
}
