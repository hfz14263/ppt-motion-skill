# dsh-ppt-office-motion :: measure WHERE a slide transition actually takes effect
#
# Two candidate stories disagree about the anchor point of <p:transition>:
#
#   A. the element on slide N animates the move INTO slide N
#      (slide N is the second endpoint, N-1 the first)
#   B. the element on slide N animates the move OUT OF slide N
#
# Both stories predict "the XML is present on slide N", so reading XML settles
# nothing. Only rendered frames settle it. anchor.pptx carries the transition on
# slide 2 of 3, so the two stories predict a change burst at different slide
# boundaries. first.pptx carries it on slide 1, which has no predecessor at all.
#
# This script:
#   1. forces every slide to auto-advance in exactly $SlideSeconds second(s)
#      (CreateVideo respects default timings only if we set them)
#   2. exports a video per deck            -> read by `anchordeck`... see below
#   3. round-trips twochild.pptx through PowerPoint
#      -> whichever of the two children survives tells us the slot count
#
# Usage:
#   powershell -NoProfile -File scripts/probe_anchor.ps1 -Dir C:\t1w\anchor
#
# Output: <Dir>\<deck>.mp4, <Dir>\anchor-report.json
# Read frames with:
#   python scripts/build_transition_table.py anchors <Dir>\anchor.mp4 <Dir>\anchor.manifest.json

[CmdletBinding()]
param(
  [string]$Dir = '',
  [double]$SlideSeconds = 2.0,
  [int]$Res = 720,
  [int]$Fps = 30,
  [int]$Quality = 85
)

$ErrorActionPreference = 'Continue'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $Dir) { $Dir = $repoRoot + '\anchor-probe' }
$Dir = [System.IO.Path]::GetFullPath($Dir)

# PowerPoint writes its scratch files next to TEMP; a long/garbled TEMP is one of
# the documented CreateVideo failure modes, so pin it to a short path we own.
$tmp = $Dir + '\tmp'
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$env:TEMP = $tmp; $env:TMP = $tmp

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$ppt = New-Object -ComObject PowerPoint.Application
$report = [ordered]@{ fps = $Fps; slideSeconds = $SlideSeconds; decks = @() }

function Get-SlideXml {
  param($pres, [int]$idx)
  $zip = Join-Path $env:TEMP ('ac' + [guid]::NewGuid().ToString('N') + '.zip')
  $dir = $zip -replace '\.zip$', ''
  Copy-Item $pres.FullName $zip -Force
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $a = [System.IO.Compression.ZipFile]::OpenRead($zip)
  $e = $a.Entries | Where-Object { $_.FullName -eq ('ppt/slides/slide' + $idx + '.xml') }
  $xml = ''
  if ($e) {
    $r = New-Object System.IO.StreamReader($e.Open())
    $xml = $r.ReadToEnd(); $r.Close()
  }
  $a.Dispose()
  Remove-Item $zip -Force -ErrorAction SilentlyContinue
  Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
  return $xml
}

# ------------------------------------------------------------- video export
function Export-Deck {
  param([string]$Src, [string]$Name)

  $entry = [ordered]@{ deck = $Name; opened = $false; transitionSeen = @(); video = $null; status = $null }
  $out = Join-Path $Dir ($Name + '.mp4')
  Remove-Item $out -Force -ErrorAction SilentlyContinue

  # Open($FileName, ReadOnly, Untitled, WithWindow) -- WithWindow MUST be 1;
  # the headless path is documented to fail media export.
  $pres = $null
  try {
    $pres = $ppt.Presentations.Open($Src, 0, 0, 1)
    $entry.opened = $true
  } catch {
    $entry.error = $_.Exception.Message
    $report.decks += $entry
    Write-Host ("  {0,-14} OPEN-FAIL   {1}" -f $Name, $_.Exception.Message)
    return
  }

  try {
    # Pin the dwell time so the slide boundaries land on known frames.
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
      $t = $pres.Slides.Item($i).SlideShowTransition
      $t.AdvanceOnTime = -1          # msoTrue
      $t.AdvanceOnClick = 0
      $t.AdvanceTime = $SlideSeconds
    }
    # Read back what PowerPoint thinks each slide carries. This is the COM view,
    # not ground truth about XML, but it tells us whether our block survived.
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
      $t = $pres.Slides.Item($i).SlideShowTransition
      $entry.transitionSeen += [ordered]@{
        slide = $i
        entryEffect = [int]$t.EntryEffect
        duration = [double]$t.Duration
        advanceTime = [double]$t.AdvanceTime
      }
    }
    $entry.slideXmlBefore = @()
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
      $entry.slideXmlBefore += (Get-SlideXml -pres $pres -idx $i)
    }

    # CreateVideo takes six arguments and must be called through .Invoke().
    $pres.CreateVideo.Invoke(@($out, $true, $SlideSeconds, $Res, $Fps, $Quality)) | Out-Null
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt 240) {
      Start-Sleep -Milliseconds 1000
      try { $st = $pres.CreateVideoStatus } catch { $st = 'err' }
      if ($st -eq 3 -or $st -eq 4) { break }     # 3/4 done, 2 failed
    }
    $entry.status = $st
    if (Test-Path $out) {
      $entry.video = $out
      Write-Host ("  {0,-14} OK  status={1}  {2,10:N0} bytes  {3}s" -f $Name, $st, (Get-Item $out).Length, [int]$sw.Elapsed.TotalSeconds)
    } else {
      Write-Host ("  {0,-14} NO-FILE status={1}" -f $Name, $st)
    }
  } catch {
    $entry.error = $_.Exception.Message
    Write-Host ("  {0,-14} FAIL  {1}" -f $Name, $_.Exception.Message)
  } finally {
    try { $pres.Close() } catch { }
  }
  $report.decks += $entry
}

Write-Host "== anchor probe: does <p:transition> belong to the page it sits on? =="
Write-Host ("   fps={0}  dwell={1}s  -> slide boundary every {2} frames" -f $Fps, $SlideSeconds, ($Fps * $SlideSeconds))
# Every deck the generator wrote gets a video; the twochild_* decks are open
# tests instead (they are expected to be refused) and must not be exported.
$decks = Get-ChildItem -Path $Dir -Filter '*.pptx' |
         Where-Object { $_.Name -notlike 'twochild*' -and $_.Name -notlike '*-saved.pptx' } |
         Sort-Object Name
foreach ($d in $decks) {
  Export-Deck -Src $d.FullName -Name ([System.IO.Path]::GetFileNameWithoutExtension($d.Name))
}

# ------------------------------------------------- claim B: one child or two?
Write-Host ""
Write-Host "== two children in one <p:transition>: which one survives? =="
foreach ($two in (Get-ChildItem -Path $Dir -Filter 'twochild*.pptx')) {
  $name = [System.IO.Path]::GetFileNameWithoutExtension($two.Name)
  $twoCopy = Join-Path $Dir ($name + '-saved.pptx')
  Remove-Item $twoCopy -Force -ErrorAction SilentlyContinue
  $twoEntry = [ordered]@{ deck = $name; opened = $false }
  try {
    $p = $ppt.Presentations.Open($two.FullName, 0, 0, 1)
    $twoEntry.opened = $true
    $before = Get-SlideXml -pres $p -idx 2
    $p.SaveAs($twoCopy)
    $twoEntry.entryEffectSlide2 = [int]$p.Slides.Item(2).SlideShowTransition.EntryEffect
    $p.Close()
    $twoEntry.savedTo = $twoCopy
    $twoEntry.slide2Before = $before
  } catch {
    $twoEntry.error = $_.Exception.Message
    Write-Host ("  {0,-16} OPEN-FAIL  {1}" -f $name, $_.Exception.Message)
  }
  if ($twoEntry.opened) {
    Write-Host ("  {0,-16} opened OK; slide 2 EntryEffect now {1}" -f $name, $twoEntry.entryEffectSlide2)
  }
  $report.decks += $twoEntry
}

$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$json = Join-Path $Dir 'anchor-report.json'
ConvertTo-Json -InputObject $report -Depth 6 | Set-Content -Path $json -Encoding UTF8
Write-Host ("report -> {0}" -f $json)
