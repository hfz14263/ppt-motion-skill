# dsh-ppt-office-motion :: transition probe -- PowerPoint roundtrip + enum readback
#
# Takes a directory of probe decks (one candidate each, built by
# build_transition_table.py build), opens each in real PowerPoint, SAVES it to
# -OutDir, and records SlideShowTransition.EntryEffect per slide.
#
# Why SaveAs is the whole point: whatever PowerPoint writes back IS the
# authority on what a gallery item means in OOXML on this build. A candidate
# that comes back different is a FINDING, not something to fix up.
#
# Why one deck per candidate: a single schema-invalid <p:transition> child makes
# PowerPoint refuse the WHOLE file (com-pitfalls §26). Batching all candidates
# into one deck destroys that signal -- every result becomes "could not open"
# and nothing can be attributed. So "this deck would not open" is recorded as
# a first-class verdict here, per candidate.
#
# Open pattern copied from motion.ps1: WithWindow must be truthy on this build
# or Open fails outright (com-pitfalls §32).
#
# Usage:
#   powershell -NoProfile -File scripts/probe_transitions.ps1 `
#       -Cases C:\t1w\trans_probe\cases -OutDir C:\t1w\trans_probe\out
#       -Status C:\t1w\trans_probe\enum.json

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Cases,
  [Parameter(Mandatory = $true)][string]$OutDir,
  [Parameter(Mandatory = $true)][string]$Status
)

$ErrorActionPreference = 'Continue'
$msoTrue  = -1
$msoFalse = 0

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }

$all = @()
$okCount = 0
$failCount = 0
$openFail = @()

$ppt = $null
try {
  $ppt = New-Object -ComObject PowerPoint.Application

  foreach ($src in (Get-ChildItem -Path $Cases -Filter *.pptx | Sort-Object Name)) {
    $dst = Join-Path $OutDir $src.Name
    $pres = $null
    try {
      # ReadOnly=false, Untitled=false, WithWindow=true
      $pres = $ppt.Presentations.Open($src.FullName, $msoFalse, $msoFalse, $msoTrue)
      $slides = @()
      foreach ($s in $pres.Slides) {
        $t = $s.SlideShowTransition
        $entry = -1; $dur = -1
        try { $entry = $t.EntryEffect } catch { $entry = -2 }
        try { $dur = $t.Duration }     catch { $dur = -2 }
        $slides += [ordered]@{ slide = $s.SlideIndex; entryEffect = $entry; duration = $dur }
      }
      if (Test-Path $dst) { Remove-Item $dst -Force }
      $pres.SaveAs($dst)
      $all += [ordered]@{ deck = $src.Name; opened = $true; slides = $slides }
      $okCount++
      Write-Host ("  ok   {0}" -f $src.Name)
    }
    catch {
      $all += [ordered]@{ deck = $src.Name; opened = $false; error = $_.Exception.Message }
      $failCount++
      $openFail += $src.Name
      Write-Host ("  FAIL {0}: {1}" -f $src.Name, $_.Exception.Message)
    }
    finally {
      try { if ($pres) { $pres.Close() } } catch { }
    }
  }
}
finally {
  try { if ($ppt) { $ppt.Quit() } } catch { }
}

# Built into $first, THEN serialised: passing a multi-line [ordered]@{} literal
# straight to -InputObject binds each following line as a separate positional
# argument and dies with ParameterBindingException (seen on PS 5.1), which
# leaves $Status a 3-byte stub and loses the whole run.
$first = [ordered]@{
  build     = (Get-Date -Format o)
  probed    = $all.Count
  opened    = $okCount
  failed    = $failCount
  openFailed= $openFail
  decks     = $all
}
$json = ConvertTo-Json -InputObject $first -Depth 6
Set-Content -Path $Status -Value $json -Encoding UTF8
Write-Host ("probe done: {0} decks, {1} opened, {2} refused" -f $all.Count, $okCount, $failCount)
