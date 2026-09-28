# dsh-ppt-office-motion :: enum scan -- let PowerPoint name the transitions
#
# build_transition_table.py enumdeck makes one blank slide per PpEntryEffect
# value; this script sets slide k to value (lo+k) over COM and saves. Reading
# the saved deck back gives the value -> XML pairing that PowerPoint itself
# uses. That is the only way to recover a transition whose element name we
# guessed WRONG: the roundtrip probe can only answer "valid / refused", and
# "refused" does not say what the right name is.
#
# Open/Save pattern copied from motion.ps1: WithWindow must be truthy on this
# build (com-pitfalls §32); keep output paths under ~200 chars (§32 rule 4).
#
# Usage:
#   powershell -NoProfile -File scripts/probe_enum_scan.ps1 `
#       -Deck C:\t1w\trans_probe\enumdeck.pptx `
#       -Manifest C:\t1w\trans_probe\enumdeck.manifest.json `
#       -Out C:\t1w\trans_probe\enumscan.pptx

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Deck,
  [Parameter(Mandatory = $true)][string]$Manifest,
  [Parameter(Mandatory = $true)][string]$Out,
  [string]$Status
)

$ErrorActionPreference = 'Continue'
$msoTrue  = -1
$msoFalse = 0

$man = Get-Content -Path $Manifest -Raw | ConvertFrom-Json
$lo = [int]$man.lo
$hi = [int]$man.hi

$assigned = 0
$bad      = @()

$ppt = $null
try {
  $ppt = New-Object -ComObject PowerPoint.Application
  # ReadOnly=false, Untitled=false, WithWindow=true
  $pres = $ppt.Presentations.Open($Deck, $msoFalse, $msoFalse, $msoTrue)
  try {
    for ($v = $lo; $v -le $hi; $v++) {
      $idx = $v - $lo + 1
      try {
        $s = $pres.Slides.Item($idx)
        $s.SlideShowTransition.EntryEffect = $v
        $assigned++
      }
      catch {
        $bad += [ordered]@{ value = $v; error = $_.Exception.Message }
      }
    }
    if (Test-Path $Out) { Remove-Item $Out -Force }
    $pres.SaveAs($Out)
  }
  finally {
    try { $pres.Close() } catch { }
  }
}
finally {
  try { if ($ppt) { $ppt.Quit() } } catch { }
}

if ($Status) {
  $out = [ordered]@{ build = (Get-Date -Format o); lo = $lo; hi = $hi; assigned = $assigned; rejected = $bad }
  Set-Content -Path $Status -Value (ConvertTo-Json -InputObject $out -Depth 4) -Encoding UTF8
}
Write-Host ("enum scan: {0} values, {1} assigned, {2} rejected by COM -> {3}" -f ($hi - $lo + 1), $assigned, $bad.Count, $Out)
