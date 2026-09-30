"""transition_probe.data — 探测数据表

**探测数据** —— 假设表与各探测名单。改数据不该碰到分析代码
"""

# 本模块由 tools/split_transition_probe.py 从单文件的 build_transition_table.py
# 搬出来（原文照搬，未改逻辑）。对外接口由 __init__.py 重新导出。


HYPOTHESES = [
    # ---- 细微 (12, excluding 无) -----------------------------------------
    {"ui": "淡入/淡出", "group": "细微", "en": "Fade", "family": "core",
     "spec": "fade", "child": "<p:fade/>"},
    {"ui": "推入", "group": "细微", "en": "Push", "family": "core",
     "spec": "push", "child": "<p:push/>"},
    {"ui": "擦除", "group": "细微", "en": "Wipe", "family": "core",
     "spec": "wipe", "child": "<p:wipe/>"},
    {"ui": "分割", "group": "细微", "en": "Split", "family": "core",
     "spec": "split", "child": "<p:split/>"},
    {"ui": "显示", "group": "细微", "en": "Reveal", "family": "p14",
     "spec": "reveal", "child": "<p14:reveal/>"},
    {"ui": "切入", "group": "细微", "en": "Cut", "family": "core",
     "spec": "cut", "child": "<p:cut/>"},
    {"ui": "随机线条", "group": "细微", "en": "Random Bars", "family": "core",
     "spec": "randombar", "child": "<p:randomBar/>"},
    {"ui": "形状", "group": "细微", "en": "Shape", "family": "core",
     "spec": "shape", "child": "<p:circle/>"},
    {"ui": "揭开", "group": "细微", "en": "Uncover", "family": "core",
     "spec": "uncover", "child": "<p:pull/>"},
    {"ui": "覆盖", "group": "细微", "en": "Cover", "family": "core",
     "spec": "cover", "child": "<p:cover/>"},
    {"ui": "闪光", "group": "细微", "en": "Flash", "family": "p14",
     "spec": "flash", "child": "<p14:flash/>"},
    {"ui": "平滑", "group": "细微", "en": "Morph", "family": "p159",
     "spec": "morph", "child": '<p159:morph option="byObject"/>'},
    # ---- 华丽 (29) --------------------------------------------------------
    {"ui": "跌落", "group": "华丽", "en": "Fall Over", "family": "p15",
     "spec": "fall_over", "child": '<p15:prstTrans prst="fallOver"/>'},
    {"ui": "悬挂", "group": "华丽", "en": "Drape", "family": "p15",
     "spec": "drape", "child": '<p15:prstTrans prst="drape"/>'},
    {"ui": "帘式", "group": "华丽", "en": "Curtains", "family": "p15",
     "spec": "curtains", "child": '<p15:prstTrans prst="curtains"/>'},
    {"ui": "风", "group": "华丽", "en": "Wind", "family": "p15",
     "spec": "wind", "child": '<p15:prstTrans prst="wind"/>'},
    {"ui": "上拉帷幕", "group": "华丽", "en": "Prestige", "family": "p15",
     "spec": "prestige", "child": '<p15:prstTrans prst="prestige"/>'},
    {"ui": "折断", "group": "华丽", "en": "Fracture", "family": "p15",
     "spec": "fracture", "child": '<p15:prstTrans prst="fracture"/>'},
    {"ui": "压碎", "group": "华丽", "en": "Crush", "family": "p15",
     "spec": "crush", "child": '<p15:prstTrans prst="crush"/>'},
    {"ui": "剥离", "group": "华丽", "en": "Peel Off", "family": "p15",
     "spec": "peel_off", "child": '<p15:prstTrans prst="peelOff"/>'},
    {"ui": "页面卷曲", "group": "华丽", "en": "Page Curl", "family": "p15",
     "spec": "page_curl", "child": '<p15:prstTrans prst="pageCurlSingle"/>'},
    {"ui": "飞机", "group": "华丽", "en": "Airplane", "family": "p15",
     "spec": "airplane", "child": '<p15:prstTrans prst="airplane"/>'},
    {"ui": "日式折纸", "group": "华丽", "en": "Origami", "family": "p15",
     "spec": "origami", "child": '<p15:prstTrans prst="origami"/>'},
    {"ui": "溶解", "group": "华丽", "en": "Dissolve", "family": "core",
     "spec": "dissolve", "child": "<p:dissolve/>"},
    {"ui": "棋盘", "group": "华丽", "en": "Checkerboard", "family": "core",
     "spec": "checkerboard", "child": "<p:checker/>"},
    {"ui": "百叶窗", "group": "华丽", "en": "Blinds", "family": "core",
     "spec": "blinds", "child": "<p:blinds/>"},
    {"ui": "时钟", "group": "华丽", "en": "Clock", "family": "core",
     "spec": "clock", "child": '<p:wheel spokes="1"/>'},
    {"ui": "涟漪", "group": "华丽", "en": "Ripple", "family": "p14",
     "spec": "ripple", "child": "<p14:ripple/>"},
    {"ui": "蜂巢", "group": "华丽", "en": "Honeycomb", "family": "p14",
     "spec": "honeycomb", "child": "<p14:honeycomb/>"},
    {"ui": "闪罐", "group": "华丽", "en": "Glitter", "family": "p14",
     "spec": "glitter", "child": "<p14:glitter/>"},
    {"ui": "涡流", "group": "华丽", "en": "Vortex", "family": "p14",
     "spec": "vortex", "child": "<p14:vortex/>"},
    {"ui": "碎片", "group": "华丽", "en": "Shred", "family": "p14",
     "spec": "shred", "child": "<p14:shred/>"},
    {"ui": "切换", "group": "华丽", "en": "Switch", "family": "p14",
     "spec": "switch", "child": '<p14:switch dir="l"/>'},
    {"ui": "翻转", "group": "华丽", "en": "Flip", "family": "p14",
     "spec": "flip", "child": '<p14:flip dir="l"/>'},
    {"ui": "库", "group": "华丽", "en": "Gallery", "family": "p14",
     "spec": "gallery", "child": '<p14:gallery dir="l"/>'},
    {"ui": "立方体", "group": "华丽", "en": "Cube", "family": "p14",
     "spec": "cube", "child": "<p14:prism/>"},
    {"ui": "门", "group": "华丽", "en": "Doors", "family": "p14",
     "spec": "doors", "child": "<p14:doors/>"},
    {"ui": "框", "group": "华丽", "en": "Box", "family": "core",
     "spec": "box", "child": "<p:zoom/>"},
    {"ui": "梳理", "group": "华丽", "en": "Comb", "family": "core",
     "spec": "comb", "child": "<p:comb/>"},
    {"ui": "缩放", "group": "华丽", "en": "Zoom", "family": "p14",
     "spec": "zoom2", "child": "<p14:warp/>"},
    {"ui": "随机", "group": "华丽", "en": "Random", "family": "core",
     "spec": "random", "child": "<p:random/>"},
    # ---- 动态内容 (7) ------------------------------------------------------
    {"ui": "平移", "group": "动态内容", "en": "Pan", "family": "p14",
     "spec": "pan", "child": "<p14:pan/>"},
    {"ui": "摩天轮", "group": "动态内容", "en": "Ferris Wheel", "family": "p14",
     "spec": "ferris_wheel", "child": '<p14:ferris dir="l"/>'},
    {"ui": "传送带", "group": "动态内容", "en": "Conveyor", "family": "p14",
     "spec": "conveyor", "child": '<p14:conveyor dir="l"/>'},
    {"ui": "旋转", "group": "动态内容", "en": "Rotate", "family": "p14",
     "spec": "rotate", "child": '<p14:prism isContent="1"/>'},
    {"ui": "窗口", "group": "动态内容", "en": "Window", "family": "p14",
     "spec": "window", "child": "<p14:window/>"},
    {"ui": "轨道", "group": "动态内容", "en": "Orbit", "family": "p14",
     "spec": "orbit", "child": '<p14:prism isContent="1" isInverted="1"/>'},
    {"ui": "飞过", "group": "动态内容", "en": "Fly Through", "family": "p14",
     "spec": "fly_through", "child": "<p14:flythrough/>"},
]


# --------------------------------------------------------------------------
# The 13 templates motion.py ships today (plus morph, which it builds
# separately). They are probed alongside the gallery items because several of
# them are NOT the gallery default: motion's "push" is dir="u" while the
# gallery's 推入 is the bare (dir="l") form, and its "zoom" is dir="in" while
# the gallery's 框 is dir="out". Refactoring motion.py onto the measured table
# must not change what any of these names produce, so each one gets its own
# deck and its own measured readback -- not a guess derived from the gallery.
# --------------------------------------------------------------------------
MOTION_SPECS = [
    {"ui": "淡入/淡出", "group": "motion", "en": "Fade", "family": "core",
     "spec": "fade", "child": "<p:fade/>"},
    {"ui": "淡入/淡出", "group": "motion", "en": "Fade", "family": "core",
     "spec": "smoothfade", "child": "<p:fade/>"},
    {"ui": "淡出为黑", "group": "motion", "en": "Fade Through Black",
     "family": "core", "spec": "fadeblack", "child": '<p:fade thruBlk="1"/>'},
    {"ui": "推入(上)", "group": "motion", "en": "Push Up", "family": "core",
     "spec": "push", "child": '<p:push dir="u"/>'},
    {"ui": "推入(左)", "group": "motion", "en": "Push Left", "family": "core",
     "spec": "pushleft", "child": '<p:push dir="l"/>'},
    {"ui": "擦除(左)", "group": "motion", "en": "Wipe Left", "family": "core",
     "spec": "wipe", "child": '<p:wipe dir="l"/>'},
    {"ui": "覆盖(左)", "group": "motion", "en": "Cover Left", "family": "core",
     "spec": "cover", "child": '<p:cover dir="l"/>'},
    {"ui": "分割(外)", "group": "motion", "en": "Split Out", "family": "core",
     "spec": "split", "child": '<p:split orient="horz" dir="out"/>'},
    {"ui": "缩放(内)", "group": "motion", "en": "Zoom In", "family": "core",
     "spec": "zoom", "child": '<p:zoom dir="in"/>'},
    {"ui": "溶解", "group": "motion", "en": "Dissolve", "family": "core",
     "spec": "dissolve", "child": "<p:dissolve/>"},
    {"ui": "条带", "group": "motion", "en": "Strips", "family": "core",
     "spec": "strips", "child": "<p:strips/>"},
    {"ui": "揭开(左)", "group": "motion", "en": "Uncover Left", "family": "core",
     "spec": "pull", "child": '<p:pull dir="l"/>'},
    {"ui": "随机线条", "group": "motion", "en": "Random Bars", "family": "core",
     "spec": "randombar", "child": '<p:randomBar/>'},
    {"ui": "平滑", "group": "motion", "en": "Morph", "family": "p159",
     "spec": "morph", "child": '<p159:morph option="byObject"/>'},
]


# Every rule here was measured, not read out of a spec. They travel with the
# table because a table of element names without them invites exactly the
# mistakes that produced them.
RULES = [
    "A transition lives in one of four namespaces and the UI grouping does not "
    "tell you which: 细微 mixes core ECMA-376 elements with p14 (显示/闪光) and "
    "p159 (平滑); all of 华丽's 2013+ items are p15:prstTrans; the 2010 items in "
    "华丽/动态内容 are p14.",
    "Every transition ends up inside mc:AlternateContent -- even a plain core "
    "one -- because p14:dur is a 2010 attribute and PowerPoint wraps the whole "
    "<p:transition> rather than drop the duration.",
    "The mc:Fallback is NOT always <p:fade/>. A core element repeats itself in "
    "the Fallback (it is already backwards-compatible); only a p14/p15/p159 "
    "child degrades to fade.",
    "PowerPoint deletes attributes that carry the default value: <p:push "
    "dir=\"l\"/> is saved as <p:push/>, <p:split orient=\"horz\" dir=\"out\"/> "
    "as <p:split/>. Writing them is harmless but never survives a save.",
    "切换 / 翻转 / 库 / 摩天轮 / 传送带 are REFUSED without a dir attribute -- "
    "a bare <p14:flip/> makes PowerPoint reject the entire file, not just that "
    "slide. PowerPoint itself always writes dir=\"l\".",
    "Namespace declarations may ride on <mc:AlternateContent>, <mc:Choice> or "
    "<p:transition>; PowerPoint moves them around freely and adds a stray "
    "xmlns=\"\" on <mc:Fallback>. Comparing saved XML as raw strings therefore "
    "reports false differences -- compare the element structure, not the text.",
    "Duplicate xmlns declarations on <p:sld> make the part not well-formed and "
    "PowerPoint refuses the whole deck. Add only the prefixes that are missing.",
]


# --------------------------------------------------------------------------
# table: merge the two measurements into the shipped reference file
# --------------------------------------------------------------------------
# Attributes PowerPoint drops because they carry the default value. The enum
# scan cannot show these (it only ever writes the value PowerPoint chose), but
# the roundtrip can: writing <p:push dir="l"/> and saving gives back <p:push/>.
# Needed to look a motion-spec child up in the enum scan, which is keyed by
# PowerPoint's OWN normalized form.
DEFAULT_ATTRS = {("dir", "l"), ("orient", "horz"), ("dir", "out")}


# Effects whose radial profile in deck 1 lands within 0.03 of another's, i.e.
# the first probe cannot tell them apart. Found by clustering the deck-1
# results; the list is data, not a guess.
DECK2_SPECS = ("fade", "wipe", "dissolve", "split", "shape", "randombar",
               "box", "push", "pan", "clock", "random", "cut")


# Which (spec, dir, axis) triples to render for the DIRECTION probe.
#
# Scope: the members of the two directional families in
# reference/transition-shapes.md §一 -- **整页平移** (push/pan/switch/
# page_curl/window) and **方向揭示** (19 members) -- MINUS the ones §四 of
# transition-choice.md forbids in business decks (checkerboard / honeycomb /
# window). Testing an effect nobody is allowed to use buys no decision value;
# this is a scope rule, not an oversight.
#
# Why this exists at all: transition-choice.md §三 carried an honest ⚠️ that
# the shape layer had only ever measured each effect in its DEFAULT form, so
# "dir flips the motion" was an assumption, not a measurement. This probe
# turns that assumption into data (or refutes it).
#
# ⚠️ TWO NOTATIONS, and they must never be mixed up again:
#   * `dir` here is the RAW OOXML ATTRIBUTE VALUE (`l`/`r`/`u`/`d`).
#   * `画面 x→y` is the OBSERVED MOTION, which is the OPPOSITE reading:
#     `dir=l` makes the picture travel r→l (l is the TARGET side the picture
#     moves toward, not the side it enters from). Measured on push:
#     dir=l → tx -1.004 (画面 r→l); dir=r → +1.009 (画面 l→r).
#   * An early version of this table annotated pairs as "push l→r" using the
#     SEMANTIC arrow for the ATTRIBUTE value -- literally backwards. The
#     `why` strings below now spell out the direction each pair is FOR.
#
# `axis` picks which trace mirror_verdict correlates: horizontal effects
# (default 画面 r→l / l→r) use "x"; vertical ones (comb/prestige 画面 b→t,
# airplane/crush/drape/fall_over 画面 t→b) use "y" and are probed with u/d --
# flipping an l/r attribute on a vertical effect tests nothing.
DIR_PROBE = (
    # ---- 整页平移族：水平轴 --------------------------------------------
    ("push", "l", "x", "默认值。画面 r→l"),
    ("push", "r", "x", "画面 l→r（§二 递进要的就是这个）"),
    ("pan", "l", "x", "同 push 族：默认 画面 r→l"),
    ("pan", "r", "x", "画面 l→r"),
    ("switch", "l", "x", "同 push 族：默认 画面 r→l"),
    ("switch", "r", "x", "画面 l→r"),
    ("page_curl", "l", "x", "卷曲带方向；默认 画面 r→l"),
    ("page_curl", "r", "x", "画面 l→r"),
    # ---- 方向揭示族：水平轴 --------------------------------------------
    ("wipe", "l", "x", "默认值。画面 r→l（形态层 band_travel −0.747）"),
    ("wipe", "r", "x", "画面 l→r（§二 并列首选）"),
    ("cover", "l", "x", "默认值。画面 r→l"),
    ("cover", "r", "x", "画面 l→r"),
    ("uncover", "l", "x", "默认值。画面 r→l"),
    ("uncover", "r", "x", "画面 l→r"),
    ("reveal", "l", "x", "§三 语义表标默认 r→l"),
    ("reveal", "r", "x", "画面 l→r（对照）"),
    ("box", "l", "x", "§三 语义表标默认 l→r，形态层却量到 stationary —— 必测"),
    ("box", "r", "x", "同上（对照）"),
    ("shape", "l", "x", "同 box：文档与形态层不一致"),
    ("shape", "r", "x", "同上（对照）"),
    ("wind", "l", "x", "§三 语义表标默认 r→l（形态层实测 画面 r→l ✓）"),
    ("wind", "r", "x", "画面 l→r（对照）"),
    ("glitter", "l", "x", "§三 语义表标默认 r→l（形态层实测 画面 r→l ✓）"),
    ("glitter", "r", "x", "画面 l→r（对照）"),
    ("peel_off", "l", "x", "§三 语义表标默认 r→l（形态层实测 画面 l→r，待校）"),
    ("peel_off", "r", "x", "画面 r→l（对照）"),
    ("curtains", "l", "x", "§三 声明『对称，加 dir 无方向』—— 验证是否真被忽略"),
    ("curtains", "r", "x", "同上：有了 l/r 才能配对"),
    # ---- 方向揭示族：垂直轴（默认就不是水平的）-------------------------
    ("comb", "u", "y", "默认 画面 b→t（形态层实测）—— 垂直轴，不能测 l/r"),
    ("comb", "d", "y", "画面 t→b"),
    ("prestige", "u", "y", "默认 画面 b→t（形态层实测）"),
    ("prestige", "d", "y", "画面 t→b"),
    ("airplane", "u", "y", "默认 画面 t→b（形态层实测）"),
    ("airplane", "d", "y", "画面 b→t"),
    ("crush", "u", "y", "默认 画面 t→b"),
    ("crush", "d", "y", "画面 b→t"),
    ("drape", "u", "y", "默认 画面 t→b"),
    ("drape", "d", "y", "画面 b→t"),
    ("fall_over", "u", "y", "默认 画面 t→b"),
    ("fall_over", "d", "y", "画面 b→t"),
    ("origami", "u", "y", "形态层实测 画面 t→b（§三 语义表把它归在 l→r 一组，待校）"),
    ("origami", "d", "y", "画面 b→t"),
)


# ---------------------------------------------------------------------------
# ATTR_PROBE: the "attribute value set" the mechanism layer never covered.
#
# HANDOVER 3a left this explicitly open: the name layer records ONE child per
# effect (clock -> <p:wheel spokes="1"/>), but several effects take FURTHER
# attributes, and nobody had measured what those values actually DO. The
# PowerPoint-written presetID enum (scripts/transition_reference.json,
# enum_by_child) turns out to enumerate every variant PowerPoint itself emits
# in its UI, so it tells us the **domain** of each attribute exactly -- we do
# not have to guess which values are legal.
#
# Each row is (spec, label, overrides, why):
#   * `spec`  -- the effect's spec in HYPOTHESES (keeps the UI name + family)
#   * `label` -- short tag used in the filename, e.g. "spokes1"
#   * `overrides` -- (attr, value) pairs to FORCE onto the child element.
#                    value=None means "strip the attribute" (the default form).
#   * `why`   -- what decision this variant informs.
ATTR_PROBE = (
    # -- wheel spokes: interface「时钟」+「转盘」。enum lists 1/2/3/8, and
    #    `<p:wheel/>` (no attr) is a SEPARATE presetID, so 4 is the default.
    ("clock", "spokes1", (("spokes", "1"),), "名字层记的形式；一条线扫一圈"),
    ("clock", "spokes2", (("spokes", "2"),), "界面「转盘 2 根」—— 真的是两根吗"),
    ("clock", "spokes3", (("spokes", "3"),), "界面「转盘 3 根」"),
    ("clock", "spokes4", (("spokes", "4"),), "显式写 4 —— 是否等于不写（默认）"),
    ("clock", "spokes8", (("spokes", "8"),), "界面「转盘 8 根」"),
    ("clock", "plain", (("spokes", None),), "不写 spokes —— 默认到底几根"),
    ("clock", "spokes6", (("spokes", "6"),), "enum 里没有的值 —— 合法吗、等同什么"),
    # -- split orient: horz/vert 在 enum 里是两个 presetID，必测
    ("split", "horz", (("orient", "horz"), ("dir", "out")),
     "名字层记的默认形式"),
    ("split", "vert", (("orient", "vert"), ("dir", "out")),
     "竖着开 —— 与横着开是否只是转了 90°"),
    ("split", "in_horz", (("orient", "horz"), ("dir", "in")),
     "反向：从边缘合拢"),
    ("split", "in_vert", (("orient", "vert"), ("dir", "in")),
     "竖着合拢"),
    # -- comb 的 dir：enum 里 comb 的第二个 presetID 是 dir="vert"，
    #    所以**不写 dir 才是默认**，必须把默认单独烘一份才能比。
    ("comb", "plain", (("dir", None),), "默认（不写 dir）—— 基线"),
    ("comb", "horz", (("dir", "horz"),), "显式 horz —— 是否等于默认"),
    ("comb", "vert", (("dir", "vert"),), "竖向梳理（enum 记的第二个值）"),
    # -- glitter pattern
    ("glitter", "default", (("pattern", None),), "不写 pattern 的默认花纹"),
    ("glitter", "hexagon", (("pattern", "hexagon"),), "菱形花纹是否只是换粒子形状"),
    # -- prism isContent / isInverted（名字层 §3 已记 cube/rotate/orbit 三项，
    #    但没点明是哪两个位；第四个组合只有 XML）
    ("cube", "plain", (("isContent", None), ("isInverted", None)),
     "裸 <p14:prism/> —— 界面「立方体」(3914)"),
    ("cube", "content", (("isContent", "1"),),
     "isContent=1 —— 界面「旋转」(3918)"),
    ("cube", "inverted", (("isInverted", "1"),),
     "isInverted=1 单独置位 —— 界面上没有这一项 (3922)"),
    ("cube", "both", (("isContent", "1"), ("isInverted", "1")),
     "两个都写 —— 界面「轨道」(3926)"),
    # -- p15 prstTrans invX：与 dir 完全不同的一个反向属性
    ("wind", "plain", (("invX", None),), "默认方向"),
    ("wind", "invX", (("invX", "1"),), "invX=1 是不是横向镜像（与 dir 有别）"),
    ("peel_off", "plain", (("invX", None),), "默认方向"),
    ("peel_off", "invX", (("invX", "1"),), "invX=1 的反向"),
    ("fall_over", "plain", (("invX", None),), "默认方向"),
    ("fall_over", "invX", (("invX", "1"),), "invX=1 的反向"),
)


# --------------------------------------------------------------------------
# TIMING_PROBE: <p:transition> and <p:timing> on the SAME slide.
#
# The mechanism layer asserts, at three separate places, that the two are
# orthogonal: "两者正交，互不占用，可以同时存在于一页" (transition-model.md §一/§四).
# That assertion has never been measured. It is exactly the shape of claim this
# project refuses to accept unmeasured -- and it is the load-bearing one, because
# the whole "two mechanisms" architecture rests on it.
#
# Three things have to be told apart, and they need DIFFERENT instruments:
#
#   (a) STRUCTURAL  -- does PowerPoint keep BOTH after a round-trip, or does one
#       silently eat the other?  Same failure class as §12 (misplaced child is
#       dropped on save) -- invisible in the file you wrote.
#   (b) TIMELINE    -- when the animation fires, does it land in the same frame
#       window as the transition, or after it?  A per-frame energy trace answers
#       this; the shape metrics (direction/mode) do not -- they describe HOW the
#       change looks, not WHEN it happens.
#   (c) INTERFERENCE-- with both present, does either one change shape vs. its
#       solo rendering?  Only a pixel diff against the solo control can say.
#
# So every row below comes in a PAIR: the combination, and its solo control.
# Row layout: (spec, tspec, label, why)
#   tspec is the transition written on the same slide, or None for the control.
# --------------------------------------------------------------------------
# The animation is deliberately the SAME everywhere: a fade on one shape,
# 0.5 s, trigger=with. `with` is the interesting trigger -- an `after` animation
# waits for a click and CreateVideo would never reach it, which would look like
# "the animation was eaten" when nothing of the sort happened.
TIMING_ANIM = {"effect": "fade", "duration": 0.5, "trigger": "with"}


TIMING_PROBE = (
    # ---- the pair that answers (a) and (b) -------------------------------
    ("push", "push", "both_push",
     "同页写 push 切换 + fade 动画 —— 两者是否都在、动画何时播"),
    ("push", None, "solo_anim",
     "只有动画、没有切换 —— 对照组：动画本身的帧窗口"),
    ("push", "push_only", "solo_trans",
     "只有 push 切换、没有动画 —— 同样是 push，排除切换种类这个变量"),
    # ---- a second transition family, to check the answer is not push-specific
    ("wipe", "wipe", "both_wipe",
     "换成方向揭示族 —— 结论是否只对 push 成立"),
    ("wipe", None, "solo_anim_wipe", "对照组：动画本身（wipe 组）"),
    ("wipe", "wipe_only", "solo_trans_wipe", "对照组：wipe 本身（同样是 wipe）"),
    # ---- no transition vs. a SLOW one: does transition duration delay the
    #      animation?  If the animation waited on the transition we would see
    #      its window slide right as the transition lengthens.  The transition
    #      is held FIXED (fade) across the three rows so only duration varies.
    ("push", "fade", "both_fade_short",
     "短切换（0.3s fade）+ 动画 —— 动画是否被切换推后"),
    ("push", "fade_slow", "both_fade_long",
     "长切换（1.5s fade）+ 动画 —— 动画窗口是否跟着右移"),
    ("push", "fade_only", "solo_fade_800",
     "fade(默认 0.8s) 单独 —— fade 系列自己的基线，供上面两行比"),
)
