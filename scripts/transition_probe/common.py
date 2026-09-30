"""transition_probe.common — 共享常量与通用工具

共享常量与通用小工具 —— 不依赖本包任何东西
"""

# 本模块由 tools/split_transition_probe.py 从单文件的 build_transition_table.py
# 搬出来（原文照搬，未改逻辑）。对外接口由 __init__.py 重新导出。


import io
import json
import os
import re

# `scripts/` 目录本身。探针模块住在 `scripts/transition_probe/`，而要读的
# `motion_catalog.json` 在**上一层**的 `scripts/`。
#
# ⚠️ **不要用 `__file__` 直接拼数据文件路径。** 2026-09-30 拆包时这里断过一次：
# 原代码写 `os.path.dirname(__file__) + "/motion_catalog.json"`，在那份单文件里
# 它恰好住在 scripts/ 下所以是对的；一旦把函数搬进 `transition_probe/`，
# 同一个表达式就指向了 `transition_probe/motion_catalog.json` —— 一个不存在的地方。
# **代码搬家会改变相对路径的含义**，而静态检查看不见这种错。
SCRIPTS_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "p159": "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
}


NSDECL = " ".join('xmlns:%s="%s"' % kv for kv in sorted(NS.items()))


DUR_MS = 800          # long enough to see, short enough for clean boundaries


PROBE_SPD = "slow"    # what the probe writes; parameterised back out in table


SLIDE_SECONDS = 2     # CreateVideo DefaultSlideDuration


FPS = 30              # must match what probe_anchor.ps1 exports at


# --------------------------------------------------------------------------
# collect: what did PowerPoint do to each candidate?
# --------------------------------------------------------------------------
def _extract_transition(xml):
    """Normalised <p:transition> content of a slide (AlternateContent kept
    whole: Choice + Fallback is BY DESIGN, not a duplicate -- §20/§31)."""
    m = re.search(r"<mc:AlternateContent(?=[\s/>]).*?</mc:AlternateContent>",
                  xml, re.S)
    if m:
        return re.sub(r">\s+<", "><", m.group(0))
    m = re.search(r"<p:transition\b.*?</p:transition>|<p:transition\b[^>]*/>",
                  xml, re.S)
    return re.sub(r">\s+<", "><", m.group(0)) if m else ""


def _load_json(path):
    """PowerShell's Set-Content -Encoding UTF8 writes a BOM; json.load chokes
    on it unless the codec is utf-8-sig. Every file this script reads may have
    come from either side, so always open with utf-8-sig (it is identical to
    utf-8 when there is no BOM)."""
    return json.load(io.open(path, encoding="utf-8-sig"))


def _norm(s):
    """Compare semantics, not namespace-declaration placement.

    PowerPoint moves the xmlns:* declarations off <p:sld> and onto
    <mc:AlternateContent> (and sometimes onto <p:transition>, and it adds a
    stray xmlns="" on <mc:Fallback>). None of that changes what the block
    MEANS, so comparing raw strings reports every single candidate as
    "rewritten" -- which is noise, not a finding."""
    s = re.sub(r'\s+xmlns:\w+="[^"]*"', "", s)
    s = re.sub(r'\s+xmlns=""', "", s)
    return re.sub(r">\s+<", "><", s).strip()


# --------------------------------------------------------------------------
# shapedeck / shapeanalyze: WHAT does each transition look like as motion?
#
# The 48-item table answers "what element does PowerPoint write". It says
# nothing about what the viewer SEES. transition-model.md §七 already found
# that a whole-frame mean cannot see direction: a 1.5s wipe merely moves one
# edge, so the frame mean barely moves, and the same detector reads a wipe as
# a fade. Measuring SHAPE therefore requires a spatial profile, and the probe
# deck must carry spatial structure for that profile to have anything to read.
#
# Design of one probe deck (2 slides):
#   slide 1 = "FROM": a fine grid of distinct cells (each cell its own colour)
#   slide 2 = "TO"  : the same grid, each cell shifted by one step in the
#                     palette -- so every cell differs, and the difference is
#                     uniform in space. A transition that sweeps will reveal
#                     the grid in a spatial order we can recover.
# Both slides carry a large centred disc in a contrasting colour so the
# spatial centroid of change is well defined even for centre-out effects.
# --------------------------------------------------------------------------
GRID_COLS = 16


GRID_ROWS = 9
