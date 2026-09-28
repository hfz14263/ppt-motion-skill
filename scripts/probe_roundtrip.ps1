# dsh-ppt-office-motion :: roundtrip an applied deck and read the transition back
#
# WHY: motion.py writes transitions in what the reference table says is PowerPoint's
# own form. The only way to know PowerPoint AGREES is to let it save the file again
# and read back what it kept. A mismatch here does not raise an error anywhere --
# it produces a file that looks fine and renders as something else.
#
# Usage:
#   powershell -NoProfile -File scripts/probe_roundtrip.ps1 -Src <in.pptx> -Out <out.pptx>
#
# Prints the COM view (EntryEffect / Duration) per slide and dumps each slide's
# transition XML to stdout so the spd/dur can be diffed against what was written.

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Src,
  [Parameter(Mandatory = $true)][string]$Out
)

$ErrorActionPreference = 'Continue'
$Out = [System.IO.Path]::GetFullPath($Out)
$dir = Split-Path -Parent $Out
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$tmp = Join-Path $dir 'tmp'
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$env:TEMP = $tmp; $env:TMP = $tmp

function Get-SlideXml {
  param($pres, [int]$idx)
  $zip = Join-Path $env:TEMP ('rt' + [guid]::NewGuid().ToString('N') + '.zip')
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
  return $xml
}

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
$ppt = New-Object -ComObject PowerPoint.Application

$pres = $ppt.Presentations.Open($Src, 0, 0, 1)
Write-Host ("slide  entryEffect  duration")
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
  $t = $pres.Slides.Item($i).SlideShowTransition
  Write-Host ("{0,-6} {1,-12} {2}" -f $i, [int]$t.EntryEffect, [double]$t.Duration)
}
$pres.SaveAs($Out)
$pres.Close()
$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "== slide XML after PowerPoint saved it =="
Add-Type -AssemblyName System.IO.Compression.FileSystem
$a = [System.IO.Compression.ZipFile]::OpenRead($Out)
foreach ($e in ($a.Entries | Where-Object { $_.FullName -like 'ppt/slides/slide*.xml' } | Sort-Object FullName)) {
  $r = New-Object System.IO.StreamReader($e.Open())
  $x = $r.ReadToEnd(); $r.Close()
  if ($x -notmatch '<p:transition') { continue }
  $spd = ([regex]::Matches($x, '<p:transition spd="(\w+)"') | ForEach-Object { $_.Groups[1].Value }) -join ','
  $dur = ([regex]::Matches($x, 'p14:dur="(\d+)"') | ForEach-Object { $_.Groups[1].Value }) -join ','
  $kid = ([regex]::Matches($x, '<p:transition[^>]*>(<[^>]+/>)') | ForEach-Object { $_.Groups[1].Value }) -join ' | '
  Write-Host ("  {0,-24} spd=[{1}] dur=[{2}] child={3}" -f $e.Name, $spd, $dur, $kid)
}
$a.Dispose()
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
Write-Host ""
Write-Host "roundtrip saved -> $Out"
