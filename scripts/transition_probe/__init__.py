"""transition_probe — 切换实测探针工具包（从单文件拆出）。

**对外接口由本文件维持**：`import build_transition_table as B` 里能用的
每个名字，在这里都能通过 `from transition_probe import *` 拿到。
拆分只是搬家，不改接口。
"""

from .common import NS, NSDECL, DUR_MS, PROBE_SPD, SLIDE_SECONDS, FPS, _extract_transition, _load_json, _norm, GRID_COLS, GRID_ROWS  # noqa: F401
from .data import HYPOTHESES, MOTION_SPECS, RULES, DEFAULT_ATTRS, DECK2_SPECS, DIR_PROBE, ATTR_PROBE, TIMING_ANIM, TIMING_PROBE  # noqa: F401
from .decks import _with_dir, wrap_transition, _slide_xml, _two_slide_deck, _flat_deck, _shape_deck, _shape_deck2, _set_attrs, _probe_deck, _anim_deck  # noqa: F401
from .analysis import _child_tag, _child_attrs, _child_render, _lookup_enum, _template, _dir_pairs, _profile_metrics, mirror_verdict, _read_frames, _window, _energy_trace, _hot_runs, _entry_trace, _timing_chart  # noqa: F401
from .commands import cmd_build, cmd_collect, cmd_enumdeck, cmd_enumread, cmd_table, cmd_anchordeck, cmd_anchors, cmd_video, cmd_sheets, cmd_shapedeck2, cmd_shapedeck, cmd_attrdeck, cmd_timingdeck, cmd_dirdeck, cmd_dirmirror, cmd_attrdiff, cmd_timingdiff, cmd_shapeanalyze, cmd_shapes  # noqa: F401

__all__ = [
    "NS", "NSDECL", "DUR_MS", "PROBE_SPD", "SLIDE_SECONDS", "FPS", "_extract_transition", "_load_json", "_norm", "GRID_COLS", "GRID_ROWS",
    "HYPOTHESES", "MOTION_SPECS", "RULES", "DEFAULT_ATTRS", "DECK2_SPECS", "DIR_PROBE", "ATTR_PROBE", "TIMING_ANIM", "TIMING_PROBE",
    "_with_dir", "wrap_transition", "_slide_xml", "_two_slide_deck", "_flat_deck", "_shape_deck", "_shape_deck2", "_set_attrs", "_probe_deck", "_anim_deck",
    "_child_tag", "_child_attrs", "_child_render", "_lookup_enum", "_template", "_dir_pairs", "_profile_metrics", "mirror_verdict", "_read_frames", "_window", "_energy_trace", "_hot_runs", "_entry_trace", "_timing_chart",
    "cmd_build", "cmd_collect", "cmd_enumdeck", "cmd_enumread", "cmd_table", "cmd_anchordeck", "cmd_anchors", "cmd_video", "cmd_sheets", "cmd_shapedeck2", "cmd_shapedeck", "cmd_attrdeck", "cmd_timingdeck", "cmd_dirdeck", "cmd_dirmirror", "cmd_attrdiff", "cmd_timingdiff", "cmd_shapeanalyze", "cmd_shapes",
]
