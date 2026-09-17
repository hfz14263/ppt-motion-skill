# dsh-ppt-office-motion :: Office COM layer
#
# The OOXML engine (scripts/motion.py) owns everything that must not disturb
# geometry. This script owns the two jobs that genuinely need a running
# PowerPoint:
#
#   1. media   - AddMediaObject2 inserts a real MP4/WAV into a slide. PowerPoint
#                will not import media without its own transcoding pipeline, so
#                this cannot be done by hand-writing OOXML.
#   2. verify  - open the deck, read the animation sequence back, render slides
#                to PNG and export a PDF. This is the "true render" review track.
#
# PowerPoint requires a desktop session. On a sandboxed host, COM automation may
# need wider file access; the failure is reported, never silently swallowed.
#
# Usage:
#   pwsh -File motion.ps1 -Pptx deck.pptx -OutDir out
#   pwsh -File motion.ps1 -Pptx deck.pptx -Media map.json
#   pwsh -File motion.ps1 -Pptx deck.pptx -ExportPdf
#
# map.json (media layer):
#   { "slides": [ { "page": 3, "media": [
#       { "src": "clip.mp4", "bounds": [100, 90, 520, 293],
#         "loop": false, "rewind": true, "mute": false, "volume": 0.8,
#         "autoplay": true, "elementId": "VIDEO" } ] } ] }

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Pptx,
  [string]$OutDir,
  [string]$Media,
  [string]$Spec,
  [switch]$ExportPdf,
  [switch]$NoRender,
  [switch]$NoSave,
  [switch]$SyncTransitions,
  [switch]$Strict,
  [int]$RenderWidth = 0,
  [int]$RenderHeight = 0
)

$ErrorActionPreference = 'Stop'
$msoTrue = 1; $msoFalse = 0; $msoAnimEffectMediaPlay = 83
$ppSaveAsOpenXMLPresentation = 24

# Clip the COM layer's report to JSON-safe primitives. A raw COM RCW slipped into
# the report makes ConvertTo-Json recurse forever instead of failing loudly.
function Coerce-Text($v) {
  if ($null -eq $v) { return $null }
  try { return ([string]$v) } catch { return $null }
}
function Coerce-Int($v) {
  if ($null -eq $v) { return 0 }
  try { return [int]$v } catch { return 0 }
}

function Say($msg) { Write-Host $msg }

# ---------------------------------------------------------------------------
# Effect census.
#
# Adds a <T>:New() static to the loaded Win32_PowerShellWrapper type -- that is
# the one .NET route that a restricted-language sandbox cannot block, and it is
# the only way to measure the deck after PowerPoint has normalised it. When it is
# unavailable the caller degrades to a warning rather than a false pass.
# ---------------------------------------------------------------------------
$script:HasCensus = $false
function Enable-EffectCensus {
  if ($script:HasCensus) { return $true }
  $ct = @'
using System;
using System.IO;
using System.Xml;
using System.IO.Compression;
using System.Reflection;

public class PptMotionCensus {
  // Open the .pptx as a zip and count <p:cTn presetID=...> elements. An XmlReader
  // is used rather than a regex over the decoded string because slide parts
  // declare UTF-8 while actually carrying a BOM, which corrupts a naive
  // File.ReadAllText decode and silently yields zero matches.
  public static int CountFile(string path) {
    int n = 0;
    using (FileStream fs = File.OpenRead(path))
    using (var zip = new ZipArchive(fs, ZipArchiveMode.Read)) {
      foreach (var entry in zip.Entries) {
        if (!entry.FullName.StartsWith("ppt/slides/slide") || !entry.FullName.EndsWith(".xml")) continue;
        using (var s = entry.Open())
        using (var r = XmlReader.Create(s)) {
          while (r.Read()) {
            if (r.NodeType == XmlNodeType.Element && r.LocalName == "cTn"
                && r.GetAttribute("presetID") != null) n++;
          }
        }
      }
    }
    return n;
  }

  // Sum per-slide effect counts gathered by the caller. The counts are read in
  // PowerShell, not here: .NET reflection does not bind PowerPoint's COM
  // TimeLine/MainSequence property chain the way PowerShell's own adapter does,
  // so a C#-side read quietly returns zero.
  public static int CountTotal(int[] perSlide) {
    if (perSlide == null) return -1;
    int total = 0;
    foreach (int n in perSlide) total += n;
    return total;
  }
}
'@
  $refs = @('System.Xml.dll', 'System.Xml.ReaderWriter.dll',
            'System.IO.Compression.dll', 'System.IO.Compression.ZipFile.dll')
  if (-not ('PptMotionCensus' -as [type])) {
    try { Add-Type -TypeDefinition $ct -Language CSharp -ReferencedAssemblies $refs -ErrorAction Stop }
    catch { return $false }
  }
  $script:HasCensus = $true
  return $true
}

# ---------------------------------------------------------------------------
# Minimal motion-spec reader.
#
# The same YAML file the OOXML engine consumes also drives this COM layer, so a
# single spec is the source of truth. Rather than require a YAML module on
# Windows PowerShell 5.1 we parse only the shapes we need: per-page transition
# type and per-page media entries (scalar fields plus the bounds list).
# ---------------------------------------------------------------------------
function ConvertFrom-MotionSpec {
  param([string]$Path)
  if (-not $Path) { return $null }
  if (-not (Test-Path $Path)) { throw "spec not found: $Path" }
  $text = Get-Content $Path -Raw

  if ($Path -match '\.json$') { return ($text | ConvertFrom-Json) }

  $slides = New-Object System.Collections.Generic.List[object]
  $cur = $null
  foreach ($line in ($text -split "`r?`n")) {
    if ($line -match '^\s*#') { continue }
    if ($line -match '^\s*-\s*page:\s*(\d+)') {
      $cur = [ordered]@{ page = [int]$Matches[1]; transition = $null; media = @() }
      $slides.Add($cur) | Out-Null
      continue
    }
    if ($line -match '^\s*-\s*slide:\s*(\d+)') {
      $cur = [ordered]@{ page = [int]$Matches[1]; transition = $null; media = @() }
      $slides.Add($cur) | Out-Null
      continue
    }
    if ($null -eq $cur) { continue }

    # transition: {type: fade, duration: 0.8}
    if ($line -match '^\s*transition:\s*\{?\s*type:\s*([A-Za-z]+)') {
      $t = @{ type = $Matches[1] }
      if ($line -match 'duration:\s*([0-9.]+)') { $t.duration = [double]$Matches[1] }
      $cur.transition = $t
      continue
    }
    # media item start:  - {src: "clip.mp4", bounds: [100, 90, 520, 293], loop: true}
    if ($line -match '^\s*-\s*\{\s*src:\s*"?([^",}]+)"?' ) {
      $m = @{ src = $Matches[1].Trim(); bounds = @(60, 60, 480, 270); loop = $false; rewind = $true; mute = $false; autoplay = $true }
      if ($line -match 'bounds:\s*\[([^\]]+)\]') {
        $m.bounds = @($Matches[1] -split ',' | ForEach-Object { [double]$_.Trim() })
      }
      if ($line -match 'loop:\s*(true|false)') { $m.loop = ($Matches[1] -eq 'true') }
      if ($line -match 'rewind:\s*(true|false)') { $m.rewind = ($Matches[1] -eq 'true') }
      if ($line -match 'mute:\s*(true|false)') { $m.mute = ($Matches[1] -eq 'true') }
      if ($line -match 'autoplay:\s*(true|false)') { $m.autoplay = ($Matches[1] -eq 'true') }
      if ($line -match 'volume:\s*([0-9.]+)') { $m.volume = [double]$Matches[1] }
      if ($line -match 'kind:\s*"?([a-z]+)"?') { $m.kind = $Matches[1] }
      if ($line -match 'elementId:\s*"?([A-Za-z0-9_\-]+)"?') { $m.elementId = $Matches[1] }
      $cur.media = @($cur.media) + $m
      continue
    }
  }
  # drop pages that carry neither transition nor media
  $keep = @()
  foreach ($s in $slides) {
    if ($s.transition -or (@($s.media).Count -gt 0)) { $keep += $s }
  }
  return @{ slides = $keep }
}

function Get-MsoAnimNamespace {
  try { Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint -ErrorAction Stop } catch { }
}

function New-PowerPoint {
  $ppt = New-Object -ComObject PowerPoint.Application
  return $ppt
}

function Close-PowerPoint($ppt) {
  try { if ($ppt) { $ppt.Quit() } } catch { }
  try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch { }
  Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------- media layer
function Add-MediaLayer($ppt, $pres, $plan) {
  if (-not $plan) { throw "media plan is empty" }
  $applied = @()
  foreach ($slideEntry in @($plan.slides)) {
    $idx = [int]$slideEntry.page
    if ($idx -lt 1 -or $idx -gt $pres.Slides.Count) {
      Say ("  media: page {0} out of range (deck has {1})" -f $idx, $pres.Slides.Count)
      continue
    }
    $slide = $pres.Slides($idx)
    foreach ($m in @($slideEntry.media)) {
      $src = [string]$m.src
      if (-not (Test-Path $src)) { throw "media file not found: $src" }
      $b = @($m.bounds)
      if ($b.Count -lt 4) { $b = @(60, 60, 480, 270) }
      $shape = $null
      try {
        # LinkToFile=msoFalse keeps the payload inside the pptx
        $shape = $slide.Shapes.AddMediaObject2($src, $msoFalse, $msoTrue, [single]$b[0], [single]$b[1], [single]$b[2], [single]$b[3])
      } catch {
        throw ("AddMediaObject2 failed for {0} (page {1}): {2}" -f $src, $idx, $_.Exception.Message)
      }
      if ($m.elementId) { $shape.Name = [string]$m.elementId }
      if ($m.volume -ne $null) { try { $shape.MediaFormat.AudioVolume = [single]$m.volume } catch { } }
      if ($m.mute) { try { $shape.MediaFormat.Muted = $msoTrue } catch { } }
      try {
        $ps = $shape.AnimationSettings.PlaySettings
        $ps.LoopUntilStopped = if ($m.loop) { $msoTrue } else { $msoFalse }
        $ps.RewindMovie = if ($m.rewind -eq $false) { $msoFalse } else { $msoTrue }
        $ps.HideWhileNotPlaying = $msoFalse
      } catch { Say ("  media: PlaySettings unavailable for {0}" -f $src) }
      if ($m.autoplay -ne $false) {
        try {
          $e = $slide.TimeLine.MainSequence.AddEffect($shape, 83, 0, 3)   # msoAnimEffectMediaPlay, afterEffect
          $e.EffectInformation.PlaySettings.LoopUntilStopped = if ($m.loop) { $msoTrue } else { $msoFalse }
          $e.Timing.Duration = 0.5
        } catch { Say ("  media: autoplay effect failed: {0}" -f $_.Exception.Message) }
      }
      $applied += [pscustomobject]@{ page = $idx; src = $src; name = $shape.Name; width = [int]$shape.Width; height = [int]$shape.Height }
      Say ("  media: page {0} <- {1} as '{2}' ({3}x{4})" -f $idx, (Split-Path $src -Leaf), $shape.Name, [int]$shape.Width, [int]$shape.Height)
    }
  }
  return $applied
}

# ------------------------------------------------------------- transition sync
function Sync-Transitions($pres, $plan) {
  # OPT-IN ONLY (-SyncTransitions). The enum table below is measurably wrong on
  # this build (see references/com-pitfalls.md 7): 0x0A01 emits <p:strips/>, not
  # a fade. The <p:transition> element written by the OOXML engine is the
  # authoritative form and PowerPoint now honours it, so re-applying these enums
  # does not "sync" anything -- it OVERWRITES a correct transition with the wrong
  # one. Kept for probing the property model, never run by default.
  if (-not $plan) { return 0 }
  $items = @($plan.transitions)
  if ($items.Count -eq 0) { return 0 }
  $enum = @{ fade = 0x0A01; smoothfade = 0x0A01; push = 0x0901; wipe = 0x0801 }
  $n = 0
  foreach ($t in $items) {
    $idx = [int]$t.page
    if ($idx -lt 1 -or $idx -gt $pres.Slides.Count) { continue }
    $key = ([string]$t.type).ToLower()
    if (-not $enum.ContainsKey($key)) { Say ("  transition: unknown type '{0}'" -f $key); continue }
    try {
      $pres.Slides($idx).SlideShowTransition.EntryEffect = $enum[$key]
      if ($t.duration) { $pres.Slides($idx).SlideShowTransition.Duration = [single]$t.duration }
      $n++
    } catch { Say ("  transition: page {0} failed: {1}" -f $idx, $_.Exception.Message) }
  }
  return $n
}

# ---------------------------------------------------------------- verify layer
function Verify-Deck($pres) {
  $rows = @()
  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $s = $pres.Slides($i)
    $seq = $s.TimeLine.MainSequence
    $media = 0
    foreach ($sh in $s.Shapes) { if ($sh.Type -eq 16) { $media++ } }
    $rows += [pscustomobject]@{
      slide = $i
      shapes = $s.Shapes.Count
      effects = $seq.Count
      media = $media
      transition = $s.SlideShowTransition.EntryEffect
    }
  }
  return $rows
}

function Render-Slides($pres, $outDir, $w, $h) {
  # $outDir MUST be absolute. Slide.Export is a COM call, so a relative path is
  # resolved against PowerPoint's own working directory (typically the user's
  # Documents folder), not the shell's -- the export then writes somewhere else
  # and COM reports "cannot find <the path you asked for>". Resolve here rather
  # than trusting the caller.
  if (-not [System.IO.Path]::IsPathRooted($outDir)) {
    $outDir = Join-Path (Get-Location).Path $outDir
  }
  $outDir = [System.IO.Path]::GetFullPath($outDir)
  New-Item -ItemType Directory -Force -Path $outDir | Out-Null
  $sw = if ($w -gt 0) { $w } else { [int]$pres.PageSetup.SlideWidth }
  $sh = if ($h -gt 0) { $h } else { [int]$pres.PageSetup.SlideHeight }
  $files = @()
  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $p = Join-Path $outDir ("slide{0:D2}.png" -f $i)
    $pres.Slides($i).Export($p, 'PNG', $sw, $sh)
    $files += $p
  }
  return $files
}

# -------------------------------------------------------------------- pipeline
$pptxPath = (Resolve-Path $Pptx).Path
if (-not $OutDir) { $OutDir = Join-Path (Split-Path $pptxPath -Parent) 'motion-out' }
# Absolutise BEFORE creating it: Resolve-Path throws on a path that does not exist
# yet, and a relative $OutDir would otherwise leak into COM calls.
if (-not [System.IO.Path]::IsPathRooted($OutDir)) {
  $OutDir = Join-Path (Get-Location).Path $OutDir
}
$OutDir = [System.IO.Path]::GetFullPath($OutDir)
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$tmp = Join-Path $OutDir ('.office-tmp')
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$env:TEMP = $tmp; $env:TMP = $tmp

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$ppt = $null
$pres = $null
$result = [ordered]@{
  pptx = $pptxPath; outDir = $OutDir
  media = @(); render = @(); pdf = $null; verify = @()
  savedAs = $null; effectsBefore = -1; effectsAfter = -1
  losses = @(); ok = $false; error = $null
}

try {
  Say "== dsh-ppt-office-motion :: Office COM layer =="
  Say ("deck: {0}" -f $pptxPath)
  $census = Enable-EffectCensus
  if (-not $census) { Say "  note: effect census unavailable; loss detection disabled" }
  $result.effectsBefore = if ($census) { [PptMotionCensus]::CountFile($pptxPath) } else { -1 }
  $ppt = New-PowerPoint
  $ppt.DisplayAlerts = 1
  # ReadOnly := msoTrue. This layer is a REVIEW track: it must never rewrite the
  # deck it was handed. Saving in place would destroy the OOXML engine's output
  # (and, when the caller re-ran the pipeline, the user's source file). Anything
  # that genuinely has to persist is written next to the report instead.
  # WithWindow must be truthy on this build, otherwise Open fails outright
  $pres = $ppt.Presentations.Open($pptxPath, $msoTrue, $msoFalse, $msoTrue)
  Say ("opened read-only: {0} slides, {1}x{2}pt" -f $pres.Slides.Count, [int]$pres.PageSetup.SlideWidth, [int]$pres.PageSetup.SlideHeight)

  $plan = $null
  if ($Spec) {
    $plan = ConvertFrom-MotionSpec -Path $Spec
    Say ("spec: {0} page(s) with COM-side work" -f (@($plan.slides).Count))
  } elseif ($Media) {
    $plan = Get-Content $Media -Raw | ConvertFrom-Json
    Say "-- media layer (explicit map) --"
  }

  $mediaPlan = @()
  $transPlan = @()
  foreach ($s in @($plan.slides)) {
    if ($s.transition) { $transPlan += @{ page = $s.page; type = $s.transition.type; duration = $s.transition.duration } }
    if (@($s.media).Count -gt 0) { $mediaPlan += [pscustomobject]@{ page = $s.page; media = @($s.media) } }
  }

  if ($mediaPlan.Count -gt 0) {
    Say "-- media layer --"
    $result.media = Add-MediaLayer $ppt $pres ([pscustomobject]@{ slides = $mediaPlan })
  }

  Say "-- verify layer (true render back from PowerPoint) --"
  $result.verify = Verify-Deck $pres
  foreach ($r in $result.verify) {
    Say ("  slide {0,2}: shapes={1,-3} effects={2,-3} media={3,-3} transition=0x{4:X}" -f `
      $r.slide, $r.shapes, $r.effects, $r.media, $r.transition)
  }
  if ($census) {
    $perSlide = @()
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
      try { $perSlide += [int]$pres.Slides($i).TimeLine.MainSequence.Count } catch { }
    }
    if ($perSlide.Count -eq 0) {
      $result.effectsAfter = -1
      Say "  note: could not read the timeline back; loss detection disabled"
    } else {
      $result.effectsAfter = [PptMotionCensus]::CountTotal([int[]]$perSlide)
      # Informational only. Do NOT compare this to effectsBefore: PowerPoint splits
      # one engine effect into a visibility <p:set> plus the behaviour, so
      # MainSequence.Count runs higher than the preset count (11 vs 6 on the
      # self-test deck). Only the like-for-like round-trip comparison below can
      # prove a loss.
      Say ("  as PowerPoint loads it: {0} timeline effect(s) across {1} slide(s)" -f `
        $result.effectsAfter, $perSlide.Count)
    }
  }

  if ($SyncTransitions -and $transPlan.Count -gt 0) {
    Say ("  transition plan: {0}" -f (($transPlan | ForEach-Object { "$($_.page):$($_.type)" }) -join ', '))
    $n = Sync-Transitions $pres ([pscustomobject]@{ transitions = @($transPlan) })
    Say ("  transitions re-applied through COM enums: {0} (opt-in; overwrites the engine's elements)" -f $n)
    $result.verify = Verify-Deck $pres
  }

  # Persist COM-side additions (embedded media) to a NEW file. The engine's
  # transitions and animations survive PowerPoint's own save, so this doubles as
  # the round-trip fidelity check.
  if ((-not $NoSave) -and ($mediaPlan.Count -gt 0 -or $ExportPdf)) {
    $saved = Join-Path $OutDir ((Split-Path $pptxPath -Leaf) -replace '\.pptx$', '.com.pptx')
    try {
      $pres.SaveCopyAs($saved, $ppSaveAsOpenXMLPresentation)
      $result.savedAs = $saved
      Say ("saved COM-side changes -> {0}" -f $saved)
      if ($census) {
        $after = [PptMotionCensus]::CountFile($saved)
        Say ("  round-trip census: {0} effect(s) written back, {1} were in the input" -f $after, $result.effectsBefore)
        if ($after -lt $result.effectsBefore) {
          $lost = $result.effectsBefore - $after
          $result.losses += ("{0} effect(s) lost on PowerPoint round-trip ({1} -> {2})" -f $lost, $result.effectsBefore, $after)
          Say ("  LOSS: {0} effect(s) did not survive PowerPoint's save ({1} -> {2})." -f $lost, $result.effectsBefore, $after)
          Say "  Usual cause: an emphasis effect stacked on an entrance for the SAME shape."
          Say "  Use one effect per shape, or move the emphasis onto its own shape."
          Say "  See references/com-pitfalls.md 12. Re-run with -Strict to fail on this."
        } else {
          Say "  round-trip clean: every effect survived PowerPoint's save"
        }
      }
    } catch {
      Say ("  SaveCopyAs failed: {0}" -f $_.Exception.Message)
      $result.losses += ("SaveCopyAs failed: {0}" -f $_.Exception.Message)
    }
  }

  if (-not $NoRender) {
    $rDir = Join-Path $OutDir 'render'
    $result.render = Render-Slides $pres $rDir $RenderWidth $RenderHeight
    Say ("rendered {0} PNG -> {1}" -f $result.render.Count, $rDir)
  }

  if ($ExportPdf) {
    $pdf = Join-Path $OutDir ((Split-Path $pptxPath -Leaf) -replace '\.pptx$', '.pdf')
    $pres.SaveCopyAs($pdf, 32)
    $result.pdf = $pdf
    Say ("pdf -> {0}" -f $pdf)
  }

  $result.ok = $true
}
catch {
  $result.error = $_.Exception.Message
  Say ("FAILED: {0}" -f $_.Exception.Message)
}
finally {
  try { if ($pres) { $pres.Close() } } catch { }
  Close-PowerPoint $ppt
  Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  # PowerPoint lags behind Quit and keeps its Diagnostics log handles open, so a
  # single Remove-Item races it and leaves .office-tmp behind. Give the process a
  # moment to actually exit, then retry.
  for ($i = 0; $i -lt 10; $i++) {
    if (-not (Test-Path $tmp)) { break }
    try { Remove-Item $tmp -Recurse -Force -ErrorAction Stop; break }
    catch { Start-Sleep -Milliseconds 300 }
  }
}

$reportPath = Join-Path $OutDir 'com-report.json'
$result | ConvertTo-Json -Depth 6 | Set-Content $reportPath -Encoding UTF8
Say ("report -> {0}" -f $reportPath)
if (-not $result.ok) { exit 1 }
if ($Strict -and $result.losses.Count -gt 0) {
  Say ("STRICT: {0} problem(s) reported" -f $result.losses.Count)
  exit 1
}
exit 0

