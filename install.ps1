# dsh-ppt-office-motion :: installer
#
# Installs this skill folder into a DSH skill root so the agent can load it.
# Discovery roots (verified against @deepseek-ai/dsh-skill-filesystem):
#   <projectRoot>/.dsh/skills      project-local
#   <projectRoot>/.agents/skills   project-local
#   <DSH_HOME>/skills              user-wide   (default target)
#   ~/.agents/skills               user-wide
#
# A directory skill is any folder containing SKILL.md.
#
# Usage:
#   powershell -NoProfile -File install.ps1                 # user-wide
#   powershell -NoProfile -File install.ps1 -Project        # into ./.dsh/skills
#   powershell -NoProfile -File install.ps1 -Target <dir>   # explicit root
#   powershell -NoProfile -File install.ps1 -Uninstall

[CmdletBinding()]
param(
  [switch]$Project,
  [string]$Target,
  [switch]$Uninstall,
  [switch]$Force
)

$ErrorActionPreference = 'Stop'
$skillName = 'dsh-ppt-office-motion'
$src = $PSScriptRoot

function Say($m) { Write-Host $m }

if ($Uninstall) {
  $roots = @()
  if ($Target) { $roots += $Target }
  $roots += (Join-Path (Join-Path $env:USERPROFILE '.dsh') 'skills')
  $roots += (Join-Path $env:USERPROFILE (Join-Path '.agents' 'skills'))
  $roots += (Join-Path (Get-Location) (Join-Path '.dsh' 'skills'))
  $n = 0
  foreach ($r in $roots) {
    $p = Join-Path $r $skillName
    if (Test-Path $p) { Remove-Item $p -Recurse -Force; Say "removed $p"; $n++ }
  }
  Say "uninstalled from $n location(s)"
  exit 0
}

if ($Target) {
  $root = $Target
} elseif ($Project) {
  $root = Join-Path (Get-Location) (Join-Path '.dsh' 'skills')
} else {
  $root = Join-Path (Join-Path $env:USERPROFILE '.dsh') 'skills'
}

$dest = Join-Path $root $skillName
Say "== dsh-ppt-office-motion installer =="
Say "source: $src"
Say "target: $dest"

# Guard: installing from the target directory itself makes $src and $dest the same
# folder, so "Remove-Item $dest -Recurse" deletes the source and then copies from
# something that no longer exists -- the skill is left empty. Refuse instead.
$srcFull = [System.IO.Path]::GetFullPath($src).TrimEnd('\')
$destFull = [System.IO.Path]::GetFullPath($dest).TrimEnd('\')
if ($srcFull -eq $destFull) {
  Say "ERROR: source and target are the same directory."
  Say "       Run install.ps1 from the unpacked skill folder, not from the installed copy."
  exit 2
}
if ($destFull.StartsWith($srcFull + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
  Say "ERROR: target is inside the source folder; refusing to recurse into itself."
  exit 2
}

if ((Test-Path $dest) -and -not $Force) {
  Say "already installed. Re-run with -Force to overwrite."
} else {
  if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
  New-Item -ItemType Directory -Force -Path $dest | Out-Null
  foreach ($item in 'SKILL.md', 'scripts', 'references', 'examples', 'install.ps1', 'requirements.txt') {
    $s = Join-Path $src $item
    if (Test-Path $s) { Copy-Item $s -Destination $dest -Recurse -Force }
  }
  # never ship test artefacts or bytecode caches
  Remove-Item (Join-Path $dest 'tests') -Recurse -Force -ErrorAction SilentlyContinue
  Get-ChildItem $dest -Recurse -Directory -Filter '__pycache__' -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
  Say "installed."
}

# ------------------------------------------------------------------ preflight
Say ""
Say "-- preflight --"
$ok = $true

$skillMd = Join-Path $dest 'SKILL.md'
if (Test-Path $skillMd) {
  $head = Get-Content $skillMd -TotalCount 6
  if (($head -join "`n") -match 'name:\s*ppt-office-motion') { Say "  SKILL.md frontmatter: name OK" }
  else { Say "  SKILL.md frontmatter: MISSING name"; $ok = $false }

  # A single unquoted ": " inside a plain scalar breaks the whole frontmatter and
  # the skill silently disappears from the catalog. That is exactly what happened
  # once when the description contained "dir: " -- so validate the YAML itself,
  # not just the presence of the name.
  $yamlOk = $false
  $pyForYaml = Get-Command python -ErrorAction SilentlyContinue
  if ($pyForYaml) {
    $probe = & $pyForYaml.Source -c @"
import io, re, sys
try:
    import yaml
except ImportError:
    print('SKIP'); sys.exit(0)
raw = io.open(r'$skillMd', encoding='utf-8').read()
m = re.match(r'^---\r?\n(.*?)\r?\n---', raw, re.S)
if not m:
    print('NOFRONT'); sys.exit(0)
try:
    d = yaml.safe_load(m.group(1)) or {}
except Exception as e:
    print('BAD:' + str(e).replace('\n', ' ')[:160]); sys.exit(0)
miss = [k for k in ('name', 'description', 'whenToUse') if not d.get(k)]
print('BAD:missing ' + ','.join(miss) if miss else 'OK')
"@ 2>$null
    switch -Regex ($probe) {
      '^OK'       { Say "  SKILL.md frontmatter: YAML valid, name/description/whenToUse present"; $yamlOk = $true }
      '^SKIP'     { Say "  SKILL.md frontmatter: YAML check skipped (PyYAML absent)"; $yamlOk = $true }
      '^NOFRONT'  { Say "  SKILL.md frontmatter: NO --- block found"; $ok = $false }
      default     { Say ("  SKILL.md frontmatter: {0}" -f $probe); $ok = $false }
    }
  } else {
    Say "  SKILL.md frontmatter: YAML check skipped (no python)"
    $yamlOk = $true
  }
} else { Say "  SKILL.md: MISSING"; $ok = $false }

$catalog = Join-Path $dest 'scripts/motion_catalog.json'
if (Test-Path $catalog) {
  $c = Get-Content $catalog -Raw | ConvertFrom-Json
  $n = @($c.PSObject.Properties).Count
  Say ("  effect catalogue: {0} aliases" -f $n)
  $suspicious = @($c.PSObject.Properties | Where-Object {
      $v = $_.Value
      ($v.kind -eq 'path' -and -not $v.hasMotionPath) -or
      ($v.kind -eq 'entrance' -and -not $v.hasVisibilitySet)
    })
  if ($suspicious.Count -gt 0) { Say ("  catalogue SANITY: {0} entries look wrong" -f $suspicious.Count); $ok = $false }
  else { Say "  catalogue sanity: OK (paths have animMotion, entrances have visibility)" }
} else { Say "  motion_catalog.json: MISSING"; $ok = $false }

foreach ($f in 'scripts/motion.py', 'scripts/verify_motion.py', 'scripts/motion.ps1',
               'scripts/player.py', 'scripts/check_coverage.py', 'scripts/selftest.py',
               'scripts/analyze_video.py', 'scripts/make_calibration.ps1',
               'references/com-pitfalls.md', 'references/template-patterns.md',
               'references/authoring-rules.md',
               'references/motion-design-spec.md', 'references/video-analysis-limits.md',
               'references/design-system/README.md') {
  if (Test-Path (Join-Path $dest $f)) { Say "  $f : OK" } else { Say "  $f : MISSING"; $ok = $false }
}

$py = Get-Command python -ErrorAction SilentlyContinue
if ($py) {
  Say ("  python: {0}" -f $py.Source)
  $ver = & $py.Source -c "import sys;print('%d.%d' % sys.version_info[:2])" 2>$null
  Say ("  python version: {0} (need 3.7+)" -f $ver)
  $deps = & $py.Source -c "import importlib.util as u;print(','.join(m for m in ('yaml','lxml') if u.find_spec(m)))" 2>$null
  if ($deps) { Say ("  python deps present: {0}" -f $deps) } else { Say "  python deps: none of yaml/lxml found (YAML specs and verify need them)" }
  # video analysis is optional but needs cv2 + numpy
  $cv = & $py.Source -c "import importlib.util as u;print(','.join(m for m in ('cv2','numpy') if u.find_spec(m)))" 2>$null
  if ($cv -eq 'cv2,numpy') { Say "  video analysis deps: cv2 + numpy OK (analyze_video.py available)" }
  else { Say "  video analysis deps: missing cv2/numpy (analyze_video.py unavailable; optional)" }
} else { Say "  python: NOT FOUND (the OOXML engine needs it)"; $ok = $false }

try {
  $ppt = New-Object -ComObject PowerPoint.Application
  Say ("  PowerPoint COM: OK (version {0})" -f $ppt.Version)
  $ppt.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
} catch { Say "  PowerPoint COM: unavailable (media insert and true-render review need it)" }

try {
  $edge = @(
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
  ) | Where-Object { Test-Path $_ } | Select-Object -First 1
  if ($edge) { Say "  Edge (optional, for external screenshots): found" } else { Say "  Edge: not found" }
} catch { }

Say ""
if ($ok) {
  Say "preflight OK."
  Say "Restart dsh web (or start a new session) so the skill catalog picks it up."
  Say "Smoke test:"
  Say "  python `"$dest\scripts\motion.py`" catalog"
} else {
  Say "preflight reported problems (see above)."
  exit 1
}
