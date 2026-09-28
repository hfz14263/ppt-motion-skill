# dsh-ppt-office-motion :: render every shape-probe deck to a video
#
# probe_anchor.ps1 answers "WHERE does a transition take effect" -- it exports
# a handful of decks and also carries a lot of round-trip bookkeeping. The
# shape probe needs the bulk path instead: one deck per effect (48 of them),
# one mp4 each, no XML readback (the table already proved the names; what is
# in question here is what a viewer SEES).
#
# Kept separate on purpose: mixing the two jobs made the anchor script drag
# ~48 slide XML blobs into its report for no reason.
#
# Usage:
#   powershell -NoProfile -File scripts/probe_shapes.ps1 -Dir C:\t1w\shape
#
# Output: <Dir>\<spec>.mp4 for every <Dir>\<spec>.pptx
# Read frames with: python scripts/build_transition_table.py shapes <Dir> --out shapes.json

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$Dir,
  [double]$SlideSeconds = 2.0,
  [int]$Res = 720,
  [int]$Fps = 30,
  [int]$Quality = 85,
  [int]$TimeoutSeconds = 180
)

$ErrorActionPreference = 'Continue'
$Dir = [System.IO.Path]::GetFullPath($Dir)
if (-not (Test-Path $Dir)) { throw "no such dir: $Dir" }

# CreateVideo writes scratch next to TEMP; a long/garbled TEMP is one of the
# documented failure modes, so pin it to a short path we own.
$tmp = Join-Path $Dir 'tmp'
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$env:TEMP = $tmp; $env:TMP = $tmp

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$ppt = New-Object -ComObject PowerPoint.Application
$report = [ordered]@{ fps = $Fps; slideSeconds = $SlideSeconds; decks = @() }

$decks = Get-ChildItem -Path $Dir -Filter '*.pptx' |
         Where-Object { $_.Name -notlike '*-saved.pptx' } |
         Sort-Object Name

Write-Host "== shape probe: rendering $($decks.Count) decks to video =="
Write-Host ("   fps={0}  dwell={1}s  res={2}  -> boundary every {3} frames" -f $Fps, $SlideSeconds, $Res, ($Fps * $SlideSeconds))

foreach ($d in $decks) {
  $name = [System.IO.Path]::GetFileNameWithoutExtension($d.Name)
  $out = Join-Path $Dir ($name + '.mp4')
  Remove-Item $out -Force -ErrorAction SilentlyContinue
  $entry = [ordered]@{ deck = $name; opened = $false; video = $null; status = $null; ms = 0 }

  $pres = $null
  try {
    # Open($FileName, ReadOnly, Untitled, WithWindow) -- WithWindow MUST be 1.
    $pres = $ppt.Presentations.Open($d.FullName, 0, 0, 1)
    $entry.opened = $true
  } catch {
    $entry.error = $_.Exception.Message
    Write-Host ("  {0,-16} OPEN-FAIL  {1}" -f $name, $_.Exception.Message)
    $report.decks += $entry
    continue
  }

  try {
    # Pin dwell so slide boundaries land on known frames.
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
      $t = $pres.Slides.Item($i).SlideShowTransition
      $t.AdvanceOnTime = -1
      $t.AdvanceOnClick = 0
      $t.AdvanceTime = $SlideSeconds
    }
    # CreateVideo takes six arguments and must go through .Invoke().
    $pres.CreateVideo.Invoke(@($out, $true, $SlideSeconds, $Res, $Fps, $Quality)) | Out-Null
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
      Start-Sleep -Milliseconds 1000
      try { $st = $pres.CreateVideoStatus } catch { $st = 'err' }
      if ($st -eq 3 -or $st -eq 4) { break }     # 3/4 done, 2 failed
    }
    $entry.status = $st
    $entry.ms = [int]$sw.Elapsed.TotalMilliseconds
    if (Test-Path $out) {
      $entry.video = $out
      Write-Host ("  {0,-16} OK  status={1}  {2,9:N0} bytes  {3}s" -f $name, $st, (Get-Item $out).Length, [int]$sw.Elapsed.TotalSeconds)
    } else {
      Write-Host ("  {0,-16} NO-FILE status={1}" -f $name, $st)
    }
  } catch {
    $entry.error = $_.Exception.Message
    Write-Host ("  {0,-16} FAIL  {1}" -f $name, $_.Exception.Message)
  } finally {
    try { $pres.Close() } catch { }
  }
  $report.decks += $entry
}

$okCount = @($report.decks | Where-Object { $_.video }).Count
Write-Host ""
Write-Host ("== done: {0}/{1} rendered ==" -f $okCount, $decks.Count)

$json = $report | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText((Join-Path $Dir 'shape-report.json'), $json, (New-Object System.Text.UTF8Encoding($false)))

try { $ppt.Quit() } catch { }
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
