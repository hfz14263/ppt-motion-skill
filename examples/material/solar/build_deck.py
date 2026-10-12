#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成「太阳系漫游指南」示例 deck 的**版面**（不含动效）。

分工（这个项目的一条铁律：版面与动效分开做）：
    这一层只负责"长什么样" —— 静态版面 + 稳定的 shape name/id
    动效由 examples/solar-motion.yaml 描述，交给 scripts/motion.py 注入
这样任何时候改版面都不影响动效定义，反之亦然。

结构（9 页）：
    P1  总览：太阳从左侧升起 + 8 行星等距排列 + 引线标注（复刻演示视频首页）
    P2..P9  逐行星详情：同一个位置放这颗行星的放大版 + 引线 + 数据

**name/id 稳定性是这套方案的地基**：P1 与各详情页里同名形状必须
name 与 id 都一致，Morph 才能跨页配对补间。所以这里的命名规则是硬约定：

    PLANET_<name>   行星图（总览页小、详情页大）
    LABEL_<name>    英文名标签
    CN_<name>       中文名标签
    LEAD_<name>     引线（直线连接符）
    <NAME>_H1       该页大标题
    <NAME>_BODY     该页正文段落

改名字 = 断掉 Morph 配对。verify_motion 会抓，但那时已经晚了。

用法：
    python examples/material/solar/build_deck.py
"""
from __future__ import annotations

import pathlib
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

HERE = pathlib.Path(__file__).resolve().parent
OUT = HERE / "solar-system.pptx"

# 16:9
SLIDE_W, SLIDE_H = Emu(12192000), Emu(6858000)

# 配色（取自 reference/design-system/moon-white-imagery.md 的深空基调）
INK = RGBColor(0xEC, 0xF0, 0xF6)      # 主文字
DIM = RGBColor(0x8F, 0x9E, 0xB4)      # 次级文字
GOLD = RGBColor(0xE8, 0xB8, 0x4A)     # 强调

# 行星按距离排列（自左向右，与太阳系真实顺序一致）
PLANETS = [
    # key,      中文,   英文,      总览直径(pt), 详情直径(pt), 高亮文案
    ("mercury", "水星", "MERCURY", 46, 150, "距太阳 5790 万公里 · 公转一周 88 天"),
    ("venus",   "金星", "VENUS",   62, 190, "表面约 465 ℃ · 太阳系最热的行星"),
    ("earth",   "地球", "EARTH",   70, 230, "距太阳 1.5 亿 km · 已知唯一的生命星球"),
    ("mars",    "火星", "MARS",    58, 180, "红色来自氧化铁 · 拥有太阳系最高的火山"),
    ("jupiter", "木星", "JUPITER", 116, 320, "质量是其余 7 颗行星总和的 2.5 倍"),
    ("saturn",  "土星", "SATURN",  104, 290, "光环宽 28 万 km · 主要由冰块构成"),
    ("uranus",  "天王星", "URANUS",  76, 220, "躺着自转 · 南北极各有 42 年的极昼"),
    ("neptune", "海王星", "NEPTUNE", 72, 215, "风速可达 2100 km/h · 太阳系最冷行星"),
]

# --------------------------------------------------------------------------
# 版面常量：总览页的「行星带」与详情页的「画布外队列」共用同一套相对关系
# --------------------------------------------------------------------------
BAND_Y = 250.0                      # 行星带中心 y（pt）
BAND_X0 = 250.0                     # 第一颗行星中心 x（太阳占 0..223pt，让开）
BAND_SPAN = 600.0                   # 第一颗到最后一颗的跨度
STEP = BAND_SPAN / (len(PLANETS) - 1)
LABEL_Y = BAND_Y + 116 * 0.5 + 46   # 标签基线（按最大行星的底缘统一，八组对齐）

DETAIL_CX, DETAIL_CY = 480.0, 270.0  # 详情页：聚焦行星的中心（画布正中）

# 「画布外」的两个停靠位 —— 本样例的设计核心（见 group_geom 的说明）
LEFT_EXIT = -140.0                  # 已看过的行星：最近一颗的停靠位（中心 x）
RIGHT_WAIT = 1100.0                 # 还没看到的行星：下一颗的候场位（中心 x）


def _set_bg(slide, img: pathlib.Path):
    """背景铺满整页。用 page background（不是形状）——
    这样它**不参与动画**，也不会出现在 Morph 的配对候选里。"""
    # python-pptx 没有直接设 slide 背景图的 API，走 XML。
    from pptx.oxml.ns import qn
    sld = slide._element
    cSld = sld.find(qn("p:cSld"))
    # pic 作为 bg 的 blipFill
    part = slide.part
    image_part, rId = part.get_or_add_image_part(str(img))

    from lxml import etree
    ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
          "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    bg = etree.SubElement(cSld, "{%s}bg" % ns["p"])
    bgPr = etree.SubElement(bg, "{%s}bgPr" % ns["p"])
    blipFill = etree.SubElement(bgPr, "{%s}blipFill" % ns["a"])
    blip = etree.SubElement(blipFill, "{%s}blip" % ns["a"])
    blip.set("{%s}embed" % ns["r"], rId)
    stretch = etree.SubElement(blipFill, "{%s}stretch" % ns["a"])
    etree.SubElement(stretch, "{%s}fillRect" % ns["a"])
    etree.SubElement(bgPr, "{%s}effectLst" % ns["a"])
    # bg 必须是 cSld 的**第一个**子元素
    cSld.remove(bg)
    cSld.insert(0, bg)
    return rId


def _txt(slide, name, text, x, y, w, h, size, color, bold=False,
         align=PP_ALIGN.LEFT, spacing=None):
    box = slide.shapes.add_textbox(x, y, w, h)
    box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = "Segoe UI"
    if spacing is not None:
        # 字间距（OOXML 的 spc，单位 1/100 pt）
        r.font._rPr.set("spc", str(int(spacing * 100)))
    return box


def _planet_img(slide, name, key, cx, cy, dia):
    """行星图：按直径在 (cx,cy) 居中放置。"""
    path = HERE / f"{key}.png"
    d = Emu(int(dia * 12700))
    pic = slide.shapes.add_picture(str(path), Emu(int(cx - dia * 6350)),
                                   Emu(int(cy - dia * 6350)), d, d)
    pic.name = name
    return pic


def _lead(slide, name, x1, y1, x2, y2, color=DIM):
    """引线（直线连接符）。演示视频里每个星球都有一根细线连到它的标注。"""
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    ln.name = name
    ln.line.color.rgb = color
    ln.line.width = Pt(0.75)
    return ln


def _header(slide, title, sub=None):
    _txt(slide, "HDR_TITLE", title, Emu(640000), Emu(360000),
         Emu(6000000), Emu(420000), 22, INK, bold=True, spacing=2)
    if sub:
        _txt(slide, "HDR_SUB", sub, Emu(640000), Emu(760000),
             Emu(6000000), Emu(300000), 11, DIM)


# --------------------------------------------------------------------------
# 组几何：一颗行星 + 它的引线 + 中英标签，作为**一组**在状态之间切换
# --------------------------------------------------------------------------
def _band_geom(i, d_o):
    """总览页（带上）的组几何：八组等距排成一条带，标签基线统一。"""
    cx = BAND_X0 + (i - 1) * STEP
    return {
        "planet": (cx, BAND_Y, d_o),
        "lead": (cx, BAND_Y + d_o * 0.5 + 6, cx, LABEL_Y - 4),
        "label": (cx - 50, LABEL_Y, 100, 200),
        "cn": (cx - 50, LABEL_Y + 17, 100, 200),
    }


def _focus_geom(d_d):
    """详情页（聚焦）的组几何：行星居中放大，金色引线 + 标签在下方。"""
    cx, cy = DETAIL_CX, DETAIL_CY
    return {
        "planet": (cx, cy, d_d),
        "lead": (cx, cy + d_d * 0.5 + 4, cx, cy + d_d * 0.5 + 46),
        "label": (cx - 90, cy + d_d * 0.5 + 48, 180, 240),
        "cn": (cx - 90, cy + d_d * 0.5 + 70, 180, 220),
    }


def _shift_geom(g, dx):
    """整组水平平移 dx（组内所有元素的相对关系不变）。"""
    px, py, pd = g["planet"]
    x1, y1, x2, y2 = g["lead"]
    lx, ly, lw, lh = g["label"]
    cxx, cyy, cw, ch = g["cn"]
    return {
        "planet": (px + dx, py, pd),
        "lead": (x1 + dx, y1, x2 + dx, y2),
        "label": (lx + dx, ly, lw, lh),
        "cn": (cxx + dx, cyy, cw, ch),
    }


def group_geom(i, page_j):
    """第 i 颗行星在【第 page_j 页】的组几何（page_j=0 为总览页）。

    这就是「那么多不同的元素，它们之间如何切换」的答案 ——
    每个组在任意时刻只有三种状态：

      1. 在带里（总览页）   : 八组等距排成一条带，y=BAND_Y 一行
      2. 聚焦（当前行星）   : 画布中央放大，引线加长、变金色
      3. 画布外（其余行星） : **整组平移到画布外**
         - 已看过的（i < j）→ 左外侧排队，最近看过的停在 LEFT_EXIT
         - 没看到的（i > j）→ 右外侧候场，下一颗停在 RIGHT_WAIT

    第 3 种是这份版面的设计关键。两件容易做错的事：

    * **不要把它们缩小后堆在画布内的角落里。** Morph 会忠实执行这个指令，
      于是观众看到七颗行星带着引线标签一起飞向角落、缩成一团小点
      （实测就是这样 —— 那不是「退场」，是「乱」）。退场的正确指令是
      「挪出画布」：滑出画面的动作在任何尺寸下都成立，且方向即语义。
    * **平移量按「整组」算，不要让每颗各自找目标。** 同一侧的行星共享同一个
      dx（由「最近的一颗恰好停在画布外」定出），带上的间距因此保持 ——
      整队行星像一列火车平移出站，而不是各飞各的。位移一致 = 速度一致，
      这正是 Morph 页面看起来「稳」的原因。
    """
    d_o = PLANETS[i - 1][3]
    d_d = PLANETS[i - 1][4]
    if page_j == 0:
        return _band_geom(i, d_o)
    if i == page_j:
        return _focus_geom(d_d)
    g = _band_geom(i, d_o)
    if i < page_j:                                # 已看过 → 左外侧
        ref = BAND_X0 + (page_j - 2) * STEP       # 最近看过的那颗的带位置
        return _shift_geom(g, LEFT_EXIT - ref)
    ref = BAND_X0 + page_j * STEP                 # 下一颗的带位置
    return _shift_geom(g, RIGHT_WAIT - ref)


def build_overview(prs, bg):
    """第 1 页 = 总览：复刻演示视频首屏（太阳 + 8 行星带 + 引线标注）。

    ## 与详情页的关系：形状集合必须**完全相同**

    Morph 按 shape **id** 配对，id 在一页内唯一。所以两页之间要补间，
    两页得是**同一批形状**：同名、同 id、只是状态不同。

    本页与详情页共 39 个形状：SUN + 6 个文字 + 8 组（行星/引线/中英标签），
    且两页都按【同一顺序】建出来 —— id 才天然一致（_align_overview_ids 校验）。
    """
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(s, bg)

    # 太阳：贴左边缘（左边缘 -13pt，球心 105pt）。详情页里它滑到画布外，
    # 于是总览→详情那一跳里，太阳是「向左滑出」的。
    sun = s.shapes.add_picture(str(HERE / "sun.png"),
                               Emu(-166500), Emu(1750000),
                               Emu(3000000), Emu(3000000))
    sun.name = "SUN"

    # 与详情页同名的文字形状（顺序与详情页严格一致）
    _txt(s, "HDR_TITLE", "SOLAR SYSTEM · 太阳系漫游", Emu(640000), Emu(360000),
         Emu(6000000), Emu(420000), 22, INK, bold=True, spacing=2)
    _txt(s, "HDR_CN", "八颗行星，一条旅程，以及 46 亿年的旅程", Emu(640000),
         Emu(760000), Emu(6000000), Emu(300000), 11, DIM)
    _txt(s, "TAGLINE", "从太阳出发，依次经过八颗行星", Emu(640000),
         Emu(1080000), Emu(4000000), Emu(300000), 11.5, INK)
    # BODY 的位置与详情页对齐（115pt ↔ 205pt，位移只有 90pt）。
    # 早先它放在页脚（484pt）—— 进入详情页时要从底部飞越整个画面，
    # 那是「没有设计」的位移。对齐后它只是轻轻下沉。
    _txt(s, "BODY", "每一站停一次，看清一颗行星，再继续向前。", Emu(640000),
         Emu(1460500), Emu(5000000), Emu(300000), 10.5, DIM)
    _txt(s, "INDEX", "序", Emu(11200000), Emu(430000),
         Emu(700000), Emu(300000), 11, DIM, align=PP_ALIGN.RIGHT)
    _txt(s, "FOOT", "7 月 15 日 · 银河可见 · 夜间观测", Emu(640000),
         Emu(6180000), Emu(4000000), Emu(260000), 10, DIM)

    # 8 组：行星 + 引线 + 中英标签（带上状态）
    for i, (key, cn, en, _d_o, _dd, _tg) in enumerate(PLANETS, start=1):
        g = group_geom(i, 0)
        cx, cy, d = g["planet"]
        _planet_img(s, f"PLANET_{i}", key, cx * 12700, cy * 12700, d)
        x1, y1, x2, y2 = g["lead"]
        _lead(s, f"LEAD_{i}", Emu(int(x1 * 12700)), Emu(int(y1 * 12700)),
              Emu(int(x2 * 12700)), Emu(int(y2 * 12700)))
        lx, ly, lw, lh = g["label"]
        _txt(s, f"LABEL_{i}", en, Emu(int(lx * 12700)), Emu(int(ly * 12700)),
             Emu(int(lw * 12700)), Emu(int(lh * 12700)), 9, INK,
             bold=True, align=PP_ALIGN.CENTER, spacing=1)
        cxx, cyy, cw, ch = g["cn"]
        _txt(s, f"CN_{i}", cn, Emu(int(cxx * 12700)), Emu(int(cyy * 12700)),
             Emu(int(cw * 12700)), Emu(int(ch * 12700)), 8, DIM,
             align=PP_ALIGN.CENTER)
    return s


def _set_geom(slide, name, x=None, y=None, cx=None, cy=None):
    """改一个已存在形状的几何（只改给了值的那些）。"""
    for shp in slide.shapes:
        if shp.name != name:
            continue
        if x is not None:
            shp.left = Emu(int(x))
        if y is not None:
            shp.top = Emu(int(y))
        if cx is not None:
            shp.width = Emu(int(cx))
        if cy is not None:
            shp.height = Emu(int(cy))
        return shp
    raise KeyError("形状不存在: %s" % name)


def _swap_image(slide, name, png_path):
    """把某个图片形状的贴图换成另一张（**保持 name 与 id 不变**）。

    这是让兄弟页也能 Morph 的关键：8 个详情页里的行星形状**同名同 id**，
    只有贴图不同 —— Morph 于是能把"水星"补间成"金星"。
    """
    from pptx.oxml.ns import qn
    for shp in slide.shapes:
        if shp.name != name:
            continue
        part, rId = slide.part.get_or_add_image_part(str(png_path))
        blip = shp._element.find(".//" + qn("a:blip"))
        blip.set(qn("r:embed"), rId)
        return shp
    raise KeyError("形状不存在: %s" % name)


def _set_text(slide, name, text):
    """改一个形状的文字。**幂等** —— 可以对同一形状重复调用。

    幂等是硬要求：这个函数在构建流程里会被调用多次（stage1 模板一次、
    stage3 每页一次；而 stage2 的复制页继承模板的文字）。多段文字如果
    只追加不清理，第二次调用就会写出两份（实测：详情页正文重复两遍）。

    另一个坑是空串：复制/新建的文本框里**可能一个 run 都没有**
    （python-pptx 的 add_textbox 默认给一个空段落）。清空文字时不能假设
    runs[0] 存在，否则 IndexError —— 实测踩过。
    """
    for shp in slide.shapes:
        if shp.name != name:
            continue
        tf = shp.text_frame
        lines = text.split("\n")

        # 第 1 步：删掉首段以外的**全部历史段落**（幂等的前提）。
        for extra in list(tf.paragraphs[1:]):
            extra._element.getparent().remove(extra._element)

        para = tf.paragraphs[0]
        if para.runs:
            para.runs[0].text = lines[0]
            for extra_run in para.runs[1:]:
                extra_run.text = ""
        elif text:
            para.text = lines[0]
        # 换行必须落成**真实段落**：把 "\n" 直接塞进 run 的话，
        # PowerPoint 会把它渲染成同一行里叠在一起的字（实测踩过）。
        for line in lines[1:]:
            new_para = tf.add_paragraph()
            new_para.text = line
            # 继承首段的字号与颜色，否则续行变默认字号
            if para.runs:
                src = para.runs[0]
                r = new_para.runs[0] if new_para.runs else new_para.add_run()
                r.text = line
                r.font.size = src.font.size
                r.font.bold = src.font.bold
                r.font.color.rgb = src.font.color.rgb
                r.font.name = src.font.name
        return shp
    raise KeyError("形状不存在: %s" % name)


# --------------------------------------------------------------------------
# 详情页的形状模板 —— 8 页共用一套 shape（同名同 id），只有状态不同
# --------------------------------------------------------------------------
def build_detail_template(prs, bg):
    """建【详情页模板】：39 个形状，与总览页**一一对应、同一顺序**。

    ## 为什么是「同构页」

    Morph 按 shape **id** 配对，而**一个 id 在一页里只能属于一个形状**。
    所以「总览页 8 行星 / 详情页 1 行星」在原理上不可能 8 颗全配上 ——
    8 个不同 id 映射不到 1 个 id（实测：最多 1 颗配上，其余退化成淡出）。

    **同构页**才能全配上：两页都是同样 39 个形状、同样的 name 与 id，
    只是同一批形状换了状态。于是：

        总览 → 水星页：水星升到中央放大，其余 7 组整队滑出画布，
                       标题块轻轻下沉、太阳向左滑出
        水星页 → 金星页：水星向左滑出、金星从右侧滑入，其余整队前移一格

    状态规则见 group_geom —— 三种状态：带里 / 聚焦 / 画布外。
    """
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(s, bg)

    # 太阳：详情页里它停在画布外（相对总览页左移 400pt，滑出画面）
    sun = s.shapes.add_picture(str(HERE / "sun.png"),
                               Emu(int((-295 - 118) * 12700)), Emu(1750000),
                               Emu(3000000), Emu(3000000))
    sun.name = "SUN"

    # 6 个文字形状（顺序与总览页严格一致）
    _txt(s, "HDR_TITLE", "", Emu(640000), Emu(430000),
         Emu(3000000), Emu(500000), 26, INK, bold=True, spacing=3)
    _txt(s, "HDR_CN", "", Emu(640000), Emu(920000),
         Emu(2000000), Emu(360000), 15, GOLD, bold=True)
    _txt(s, "TAGLINE", "", Emu(640000), Emu(1420000),
         Emu(2300000), Emu(900000), 11, INK)
    _txt(s, "BODY", "", Emu(640000), Emu(2600000),
         Emu(2400000), Emu(2400000), 10.5, DIM)
    _txt(s, "INDEX", "", Emu(11200000), Emu(430000),
         Emu(700000), Emu(300000), 11, DIM, align=PP_ALIGN.RIGHT)
    _txt(s, "FOOT", "太阳系漫游指南", Emu(640000),
         Emu(6180000), Emu(4000000), Emu(260000), 10, DIM)

    # 8 组：行星 + 引线 + 中英标签（模板页 = 第 1 页的状态，水星聚焦）
    for i, (k, c, e, _d_o, _d_d, _tg) in enumerate(PLANETS, start=1):
        g = group_geom(i, 1)
        cx, cy, d = g["planet"]
        _planet_img(s, f"PLANET_{i}", k, cx * 12700, cy * 12700, d)
        x1, y1, x2, y2 = g["lead"]
        _lead(s, f"LEAD_{i}", Emu(int(x1 * 12700)), Emu(int(y1 * 12700)),
              Emu(int(x2 * 12700)), Emu(int(y2 * 12700)), color=GOLD)
        lx, ly, lw, lh = g["label"]
        _txt(s, f"LABEL_{i}", e, Emu(int(lx * 12700)), Emu(int(ly * 12700)),
             Emu(int(lw * 12700)), Emu(int(lh * 12700)), 11, INK,
             bold=True, align=PP_ALIGN.CENTER, spacing=2)
        cxx, cyy, cw, ch = g["cn"]
        _txt(s, f"CN_{i}", c, Emu(int(cxx * 12700)), Emu(int(cyy * 12700)),
             Emu(int(cw * 12700)), Emu(int(ch * 12700)), 9.5, GOLD,
             align=PP_ALIGN.CENTER)
    return s


def fill_detail(slide, idx):
    """把（复制来的）详情页改成「第 idx 颗行星是主体」的状态。

    只改几何与文字，**name 与 id 一律不动** —— 这是 Morph 能配对的前提。
    每一条几何都来自 group_geom(i, idx)：把 8 个组各自摆到它们在这页的
    状态（聚焦 / 左外侧 / 右外侧）。
    """
    key, cn, en, _d_o, d_d, tag = PLANETS[idx - 1]
    _set_text(slide, "HDR_TITLE", en)
    _set_text(slide, "HDR_CN", cn)
    _set_text(slide, "TAGLINE", tag)
    _set_text(slide, "INDEX", "%d / %d" % (idx, len(PLANETS)))
    _set_text(slide, "BODY",
              f"第 {idx} 站 · 距离与构成\n\n"
              f"它在太阳系中排第 {idx} 位，"
              f"绕太阳一周的时间随轨道半径变化。"
              f"要看清它的细节，需要的不是飞得更近，"
              f"而是更好的观测方法 —— 这正是下一页要做的事。")

    for i, (k, c, e, _do, dd, _tg) in enumerate(PLANETS, start=1):
        g = group_geom(i, idx)
        cx, cy, d = g["planet"]
        # 贴图：第 i 个形状永远显示第 i 颗行星（形状与行星绑定，不随页面变）
        _swap_image(slide, f"PLANET_{i}", HERE / f"{k}.png")
        _set_geom(slide, f"PLANET_{i}",
                  x=cx * 12700 - d * 6350, y=cy * 12700 - d * 6350,
                  cx=d * 12700, cy=d * 12700)
        x1, y1, x2, y2 = g["lead"]
        _set_geom(slide, f"LEAD_{i}",
                  x=x1 * 12700, y=y1 * 12700,
                  cx=(x2 - x1) * 12700, cy=(y2 - y1) * 12700)
        lx, ly, lw, lh = g["label"]
        _set_geom(slide, f"LABEL_{i}",
                  x=lx * 12700, y=ly * 12700, cx=lw * 12700, cy=lh * 12700)
        cxx, cyy, cw, ch = g["cn"]
        _set_geom(slide, f"CN_{i}",
                  x=cxx * 12700, y=cyy * 12700, cx=cw * 12700, cy=ch * 12700)
        # 标签文字**每页都填**：画布外的组观众看不到，不需要清空；
        # 清空反而会在它滑回画面时闪一下（文字内容不参与补间）。
        _set_text(slide, f"LABEL_{i}", e)
        _set_text(slide, f"CN_{i}", c)
    return slide


def _stamp_tool_author(prs) -> None:
    """把作者字段盖成工具名（dsh-ppt-studio）。

    **为什么每趟 python-pptx 保存前都要盖**：中间那趟 COM 复制
    （`Slide.Duplicate()` 手工步）会把 Windows 用户名写进 `lastModifiedBy`
    —— 2026-10-12 privacy_audit 实测抓到的就是这个（真人名）。
    stage3 跑在 COM 之后，它这趟的盖章是最终生效的那次；stage1 也盖，
    是为了任何中途产物都不带人名。约定同 `scripts/build_dual_photo.py`，
    放行名单见 `tests/privacy_audit.py` 的 AUTHOR_ALLOW。
    """
    prs.core_properties.author = "dsh-ppt-studio"
    prs.core_properties.last_modified_by = "dsh-ppt-studio"


def _verify_authorship(path) -> None:
    """产物判据：core.xml 的作者字段必须正好是工具名。

    privacy_audit 的 container scan 事后会读这两个字段；这里在**构建时**
    就把「盖住了」变成可核对的断言 —— 而不是等审计发现时，文件已经发出去。
    """
    import re as _re
    import zipfile
    core = zipfile.ZipFile(str(path)).read("docProps/core.xml").decode("utf-8")
    bad = []
    for tag, label in (("dc:creator", "creator"),
                       ("cp:lastModifiedBy", "lastModifiedBy")):
        m = _re.search(r"<%s>(.*?)</%s>" % (tag, tag), core)
        val = m.group(1).strip() if m else ""
        if val != "dsh-ppt-studio":
            bad.append("%s = %r" % (label, val))
    if bad:
        raise SystemExit("作者字段没盖住：%s（privacy_audit 会报人名泄漏）"
                         % "; ".join(bad))
    print("authorship: creator / lastModifiedBy = dsh-ppt-studio")


def main() -> None:
    """三阶段：python-pptx 建模板 → COM 复制页 → python-pptx 改数值。

    为什么中间一定要过 PowerPoint 的 Slide.Duplicate()：
    **复制是唯一能保证 shape id 一致的做法**（2026-10-10 实测）。
    从零分别建页时 id 按 spTree 顺序发放，形状数一变就错位，
    而 Morph 按 id 配对 —— 于是 name 相同也不配对，放映时看不到补间。
    python-pptx 没有"复制页"接口；自造 XML 克隆会在 sldIdLst 上留下假关系
    （实测报 'Part' object has no attribute 'slide'）。

    `--stage 1|2|3` 分开跑：阶段 2 需要真 PowerPoint，单独一步便于排错。
    """
    stage = 3
    for i, a in enumerate(sys.argv):
        if a.startswith("--stage"):
            stage = int(a.split("=")[1])

    bg = HERE / "background.png"
    for need in [bg] + [HERE / f"{k}.png" for k, *_ in PLANETS] + [HERE / "sun.png"]:
        if not need.exists():
            raise SystemExit("缺素材：%s（先跑 make_assets.py）" % need)

    # ---- 阶段 1：只建【一个】详情页模板 -----------------------------------
    # 模板页的行星形状统一叫 PLANET_FOCUS，8 颗行星复用它。
    # 统一名字是让兄弟页之间也能 Morph 的关键（同一形状换贴图）。
    if stage == 1:
        prs = Presentation()
        _stamp_tool_author(prs)
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        # 第 1 页 = 总览（带完整 8 行星 + 引线），第 2 页 = 详情模板。
        # 顺序在这里就定好，后面只做「复制第 2 页」，不去挪动任何引用。
        build_overview(prs, bg)
        build_detail_template(prs, bg)
        fill_detail(prs.slides[1], 1)
        prs.save(str(OUT))
        print("stage1: 总览 + 1 页详情模板（2 页）->", OUT.name)
        return

    prs = Presentation(str(OUT))
    _stamp_tool_author(prs)     # COM 那一趟会写进用户名，这里盖回工具名
    if stage == 2:
        need = len(PLANETS) + 1        # 总览 1 + 详情 8
        print("stage2: 当前 %d 页。**请复制【第 2 页】**（详情模板）到 %d 页 —— "
              "复制第 1 页（总览）会把 8 个行星也一起复制，id 就乱了。"
              % (len(prs.slides), need))
        return

    # ---- 阶段 3：改每页数值与贴图 ------------------------------------------
    need = len(PLANETS) + 1
    if len(prs.slides) != need:
        raise SystemExit("页数不对：%d（期望 %d）。先跑 stage1，"
                         "再复制【第 2 页】到 %d 页。"
                         % (len(prs.slides), need, need))
    # 第 1 页是总览，第 2..9 页是详情
    for i in range(len(PLANETS)):
        fill_detail(prs.slides[i + 1], i + 1)

    # 总览页**先建**（成为第 1 页），再把 8 个详情页追加在后面 ——
    # 放映顺序 = 总览 → 水星 → … → 海王星。
    #
    # 走过一条弯路：先建 8 个详情页、再把总览的 sldId 用 lst.insert(0, lst[-1])
    # 顶到最前。结果 pptx 里出现**重复的 ppt/slides/slide9.xml**（同一个 part
    # 被写了两遍，zipfile 直接报 Duplicate name）。原因是 python-pptx 的
    # sldIdLst 只是**引用表**，移动引用不会搬动 part，append 时又新建了一个同号 part。
    # 结论：**顺序要在建页时就定好**，不要事后挪引用。
    _align_overview_ids(prs)

    prs.save(str(OUT))
    print("saved: %s" % OUT)
    print("slides: %d" % len(prs.slides._sldIdLst))
    _verify_authorship(OUT)


def _align_overview_ids(prs):
    """总览页与详情页的 shape **id** 对齐（同 name → 同 id）。

    ## 这不是可选项，是 Morph 生效的前提

    2026-10-10 实测：Morph 按 shape **id** 配对，name 相同但 id 不同就配不上 ——
    症状是所有结构检查、覆盖率审计、坐标差异检查全绿，放映却看不到补间。

    ## 现状：两页的 id 本来就一致

    两页都由同一套模板函数按**相同顺序**建出来（同样的 HDR_*、PLANET_1..8、
    LEAD_1..8…），python-pptx 按 spTree 顺序发 id，于是天然对齐。
    本函数只做一件事：**逐对校验**，发现不一致就报出来并修。

    ## 踩过的坑

    1. **同页内 id 必须唯一**。第一版强行把总览页 6 个文字形状改成详情页的 id，
       其中 FOOT 从 40 改成 7 —— 而 INDEX 已经占了 7，同页两个形状共用 id，
       Morph 会把它们互相配对（错得离谱）。所以改 id 前必须查占用。
    2. **同页内不能有两个形状同名**。否则 name 映射到哪个 id 说不清。
    """
    from pptx.oxml.ns import qn

    def table(slide):
        out, byid = {}, {}
        for shp in slide.shapes:
            cid = shp._element.find(".//" + qn("p:cNvPr")).get("id")
            out[shp.name] = cid
            byid.setdefault(cid, []).append(shp.name)   # 坑 1 的探针
        return out, byid

    d_ids, d_dup = table(prs.slides[1])     # 详情页模板
    o_ids, o_dup = table(prs.slides[0])     # 总览页

    # 第 1 步：总览页**独有**的形状（如 SUN）让出 id。
    #   它们在详情页没有对应形状，不该占着 id 挡住同名形状。
    #   必须在重复检测【之前】做 —— 否则检测先报出还没修的冲突（实测如此）。
    top = max(int(c) for c in o_dup)
    for shp in prs.slides[0].shapes:
        if shp.name not in d_ids:
            top += 1
            shp._element.find(".//" + qn("p:cNvPr")).set("id", str(top))
    o_ids, o_dup = table(prs.slides[0])     # 让位后重新采集

    # 第 2 步：页内 id 唯一性自检。共用 id 会让 Morph 把形状互相配对。
    for cid, names in list(d_dup.items()) + list(o_dup.items()):
        if len(names) > 1:
            raise SystemExit("同页 id 重复: id=%s 被 %s 共用 —— Morph 会把它们互相配对"
                             % (cid, names))

    fixed, already, missing = [], [], []
    for shp in prs.slides[0].shapes:
        want = d_ids.get(shp.name)
        if want is None:
            continue
        cur = shp._element.find(".//" + qn("p:cNvPr")).get("id")
        if cur == want:
            already.append(shp.name)
            continue
        # 改之前查占用（坑 1）
        taken = {c for c, ns in o_dup.items() if shp.name in ns}
        if want in taken:
            continue                       # 目标 id 已被别的形状占用，跳过并保留现状
        # 逐个腾挪：先把自己挪到一个空位
        used = set(o_dup)
        spare = max(int(c) for c in used) + 1
        shp._element.find(".//" + qn("p:cNvPr")).set("id", str(spare))
        shp._element.find(".//" + qn("p:cNvPr")).set("id", want)
        fixed.append("%s %s→%s" % (shp.name, cur, want))

    for nm in d_ids:
        if nm not in o_ids:
            missing.append(nm)

    print("id 对齐：已一致 %d 个，修正 %d 个" % (len(already), len(fixed)))
    if fixed:
        print("  修正: " + "; ".join(fixed))
    if missing:
        print("  ⚠ 总览页缺这些形状（进入时淡入）: %s" % ", ".join(missing))
    print("  → 同名同 id 的形状可在 Morph 中互相补间；总览页独有的形状走淡出")


def qn_cNvPr():
    from pptx.oxml.ns import qn
    return qn("p:cNvPr")


if __name__ == "__main__":
    main()
