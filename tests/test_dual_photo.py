#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Regression for the "同图双版本" layout validator (verify_dual_photo.py).

版式定义：同一张照片的两个版本 —— 淡化整图铺底 + 原色塞进形状当前景，只有两层。

为什么需要它（每一条断言都对应一个真实踩过的坑）：
  A. 前景被拉伸。视频里专门给过失败对照"拉伸人物就长胖"。肉眼很难量化，
     但从 XML 能算出失真倍数，且实测能量出注入值的 ±0.5%（见 calibration 段）。
  B. 底图和前景不是同一张照片，那就不是这个版式。两条独立判据：长宽比 + 低频结构相关。
     实测：同图 corr=0.950，换图 corr=0.535，阈值取 0.80（有 margin）。
  C. 遗留翻转副本。视频里"复制+翻转+重叠"是过程演示形状，成片前要删；
     我复现时把它当成品层留了下来，用户一眼看出"开口向右的燕尾没删"。
     检测器必须能报出这类残留。
  D. Windows 下 os.path.join 会产出反斜杠，而 zip 内部路径恒为正斜杠 ——
     曾因此读不到 media，校验器"读不到源图尺寸"静默降级成 WARN 而没报错。

对照设计（检测器要能同时看见"对"和"错"）：
    正对照 ok        : 无变形、同图         -> 0 error
    负对照 fat135    : 注入 1.35 倍拉胖      -> 报错，且量出 1.35
    负对照 fat115    : 注入 1.15 倍拉胖      -> 报错，且量出 1.15
    负对照 diffsrc   : 底图换另一张照片      -> 报"不是同图"
    负对照 plain     : 页面上什么都没有      -> class=none，不误报
    真样例 repro.pptx: 视频复现成品          -> 0 error
"""
import os
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "scripts"))

import build_dual_photo as B                                   # noqa: E402
import verify_dual_photo as V                                  # noqa: E402

T1 = os.path.join(ROOT, "examples", "material", "template1")
PERSON = os.path.join(T1, "template1.2.png")     # 人像（抠图，320x370 带透明通道）
SCENE = os.path.join(T1, "template1.PNG")        # 另一张：场景图，用来造"非同图"

PASS = FAIL = 0


def check(label, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print("  %-58s ok" % label)
    else:
        FAIL += 1
        print("  %-58s FAIL %s" % (label, detail))


def one(name, **kw):
    out = os.path.join(TMP, name + ".pptx")
    B.build(PERSON, out, **kw)
    res = V.check(out)
    return res


print("同图双版本 :: 正负对照（素材 = template1 人像 template1.2.png）")
TMP = tempfile.mkdtemp(prefix="dualphoto_test_")

# ---------------------------------------------------------------- 正对照
res = one("ok")
s = res["slides"][0]
fg = s["foreground"][0]
check("ok: 识别为 dual-photo", s["class"] == "dual-photo", s["class"])
check("ok: 前景不变形 (ratio≈1.00)", abs(fg["ratio"] - 1.0) <= 0.01,
      "ratio=%.4f" % fg["ratio"])
check("ok: 底图与前景判定为同图", fg["same_source"]["verdict"] == "same",
      fg["same_source"])
check("ok: 同图相关系数 ≥ 0.80", fg["same_source"]["corr"] >= 0.80,
      "corr=%.3f" % fg["same_source"]["corr"])
check("ok: 无翻转残留", not (fg["flipH"] or fg["flipV"]))
check("ok: 0 error / 0 warn", not s["errors"] and not s["warnings"],
      "%s %s" % (s["errors"], s["warnings"]))

# ---------------------------------------------------------------- 负对照：拉伸
for f in (1.15, 1.35):
    res = one("fat%d" % int(f * 100), stretch=f)
    s = res["slides"][0]
    fg = s["foreground"][0]
    check("stretch %.2f: 检出变形" % f,
          any("失真倍数" in e for e in s["errors"]), s["errors"])
    check("stretch %.2f: 量出的倍数 ≈ 注入值(±2%%)" % f,
          abs(fg["ratio"] - f) <= 0.02 * f,
          "量出 %.4f，注入 %.2f" % (fg["ratio"], f))
    check("stretch %.2f: 不影响同图判定" % f,
          fg["same_source"]["verdict"] == "same", fg["same_source"]["verdict"])

# 亚阈值：1.05 倍（< 默认容差 0.08）应放行，证明容差不是"一律报错"
res = one("fat105", stretch=1.05)
s = res["slides"][0]
check("stretch 1.05: 低于容差 0.08 -> 放行", not s["errors"], s["errors"])

# ---------------------------------------------------------------- 负对照：非同图
res = one("diffsrc", bg_src=SCENE)
s = res["slides"][0]
fg = s["foreground"][0]
check("diffsrc: 判定为不同图", fg["same_source"]["verdict"] == "different",
      fg["same_source"])
check("diffsrc: 低频相关明显下降 (<0.80)", fg["same_source"]["corr"] < 0.80,
      "corr=%.3f" % fg["same_source"]["corr"])
check("diffsrc: 报出'不是同图双版本'",
      any("不是同一张照片" in e for e in s["errors"]), s["errors"])
check("diffsrc: 两条判据都指向不同",
      fg["same_source"]["ar_equal"] is False and fg["same_source"]["corr_ok"] is False,
      fg["same_source"])

# ---------------------------------------------------------------- 负对照：空页
from pptx import Presentation                                  # noqa: E402
from pptx.util import Pt                                       # noqa: E402
plain = os.path.join(TMP, "plain.pptx")
prs = Presentation()
prs.slide_width = Pt(960)
prs.slide_height = Pt(540)
prs.slides.add_slide(prs.slide_layouts[6])
prs.save(plain)
res = V.check(plain)
s = res["slides"][0]
check("plain: 空页 class=none 且不误报", s["class"] == "none" and not res["errors"],
      s["class"])

# ---------------------------------------------------------------- 残留翻转层
flip = os.path.join(TMP, "flip.pptx")
info = B.build(PERSON, flip)
from pptx import Presentation as _P                            # noqa: E402
_p = _P(flip)
_sp = _p.slides[0].shapes[-1]
_sp._element.spPr.xfrm.set("flipH", "1")
_p.save(flip)
s = V.check(flip)["slides"][0]
check("flip: 报出翻转残留 WARN", any("翻转" in w for w in s["warnings"]), s["warnings"])

# ---------------------------------------------------------------- 真样例（视频复现成品）
REPRO = r"D:\workbuddy\idea\material\video\repro\repro.pptx"
if os.path.isfile(REPRO):
    s = V.check(REPRO)["slides"][0]
    check("repro.pptx: 识别为 dual-photo", s["class"] == "dual-photo", s["class"])
    if s["foreground"]:
        fg = s["foreground"][0]
        check("repro.pptx: 前景不变形", abs(fg["ratio"] - 1.0) <= 0.08,
              "ratio=%s" % fg["ratio"])
        check("repro.pptx: 底图与前景同图", fg["same_source"]["verdict"] == "same",
              fg["same_source"])
    check("repro.pptx: 0 error", not s["errors"], s["errors"])
else:
    print("  (跳过 repro.pptx 真样例：文件不存在)")

shutil.rmtree(TMP, ignore_errors=True)
print("\n%d passed, %d failed" % (PASS, FAIL))
sys.exit(1 if FAIL else 0)
