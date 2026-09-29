# dsh-ppt-office-motion :: roundtrip every deck in a dir and report what survived
#
# WHY THIS EXISTS (separate from probe_roundtrip.ps1)
# ---------------------------------------------------
# probe_roundtrip.ps1 answers "did PowerPoint KEEP the transition I wrote" for a
# single deck. The timing probe needs a different question answered, across a
# whole folder:
#
#     a slide that carries BOTH <p:transition> and <p:timing> -- does PowerPoint
#     keep BOTH, or does one silently eat the other?
#
# Neither failure raises anywhere: a dropped <p:timing> just means the animation
# is gone after a save, and a dropped <p:transition> means the slide cut instead
# of slid. Both look fine in the file we wrote. Only reading PowerPoint's own
# save tells the truth.
#
# The transition readback also reports the ORDER, because the xsd:sequence
# (cSld, clrMapOvr?, transition?, timing?, extLst?) allows exactly one order.
# If PowerPoint rewrites it the other way round, our writer has it backwards.
#
# Usage:
#   powershell -NoProfile -File scripts/probe_timing_roundtrip.ps1 -Dir C:\t1w\timing
#
# Output: <Dir>\<name>-saved.pptx for each input, plus timing-rt-report.json

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$Dir,
  [int]$TimeoutSeconds = 120
)

$ErrorActionPreference = 'Continue'
$Dir = [System.IO.Path]::GetFullPath($Dir)
if (-not (Test-Path $Dir)) { throw "no such dir: $Dir" }

$tmp = Join-Path $Dir 'tmp-rt'
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$env:TEMP = $tmp; $env:TMP = $tmp

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-SlideXml {
  param([string]$Path, [string]$Entry = 'ppt/slides/slide2.xml')
  $a = [System.IO.Compression.ZipFile]::OpenRead($Path)
  $e = $a.Entries | Where-Object { $_.FullName -eq $Entry }
  $xml = ''
  if ($e) {
    $r = New-Object System.IO.StreamReader($e.Open())
    $xml = $r.ReadToEnd(); $r.Close()
  }
  $a.Dispose()
  return $xml
}

$ppt = New-Object -ComObject PowerPoint.Application
$report = [ordered]@{ decks = @() }

$decks = Get-ChildItem -Path $Dir -Filter '*.pptx' |
         Where-Object { $_.Name -notlike '*-saved.pptx' } |
         Sort-Object Name

Write-Host "== timing round-trip: $($decks.Count) decks =="
foreach ($d in $decks) {
  $name = [System.IO.Path]::GetFileNameWithoutExtension($d.Name)
  $out  = Join-Path $Dir ($name + '-saved.pptx')
  Remove-Item $out -Force -ErrorAction SilentlyContinue

  $before = Read-SlideXml $d.FullName
  $entry = [ordered]@{
    deck = $name
    before_transition = [bool]($before -match '<p:transition')
    before_timing     = [bool]($before -match '<p:timing>')
    opened = $false; saved = $false
    after_transition = $false; after_timing = $false
    order = 'n/a'; timing_lost = $false; trans_lost = $false
  }

  $pres = $null
  try {
    $pres = $ppt.Presentations.Open($d.FullName, 0, 0, 1)
    $entry.opened = $true
  } catch {
    $entry.error = $_.Exception.Message
    Write-Host ("  {0,-24} OPEN-FAIL  {1}" -f $name, $_.Exception.Message)
    $report.decks += $entry
    continue
  }

  try {
    $pres.SaveAs($out)
    $entry.saved = $true
  } catch {
    $entry.error = "SaveAs: " + $_.Exception.Message
  } finally {
    try { $pres.Close() } catch { }
  }

  if ($entry.saved -and (Test-Path $out)) {
    $after = Read-SlideXml $out
    $entry.after_transition = [bool]($after -match '<p:transition')
    $entry.after_timing     = [bool]($after -match '<p:timing>')
    $it = $after.IndexOf('<p:timing>')
    $ix = $after.IndexOf('<p:transition')
    if ($it -ge 0 -and $ix -ge 0) {
      $entry.order = if ($ix -lt $it) { 'transition-before-timing' } else { 'timing-before-transition' }
    } elseif ($it -ge 0) { $entry.order = 'timing-only' }
    elseif ($ix -ge 0) { $entry.order = 'transition-only' }
    # The failure this probe exists for:
    $entry.timing_lost = ($entry.before_timing -and -not $entry.after_timing)
    $entry.trans_lost  = ($entry.before_transition -and -not $entry.after_transition)
    $flag = if ($entry.timing_lost -or $entry.trans_lost) { ' <== LOST' } else { '' }
    Write-Host ("  {0,-24} before(t={1},a={2})  after(t={3},a={4})  order={5}{6}" -f `
      $name, [int]$entry.before_transition, [int]$entry.before_timing, `
      [int]$entry.after_transition, [int]$entry.after_timing, $entry.order, $flag)
  } else {
    Write-Host ("  {0,-24} SAVE-FAIL" -f $name)
  }
  $report.decks += $entry
}

$lost = @($report.decks | Where-Object { $_.timing_lost -or $_.trans_lost }).Count
Write-Host ""
Write-Host ("== done: {0}/{1} round-tripped, {2} lost a block ==" -f `
  @($report.decks | Where-Object { $_.saved }).Count, $decks.Count, $lost)

$json = $report | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText((Join-Path $Dir 'timing-rt-report.json'), $json,
  (New-Object System.Text.UTF8Encoding($false)))

try { $ppt.Quit() } catch { }
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
