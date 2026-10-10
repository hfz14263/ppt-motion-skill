#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从用户的莲花 deck 制作「结构复核副本」——入库用（原件不入库）。

为什么需要副本而不是原件（2026-10-10 定）：
  1. 原件 docProps/core.xml 里有**用户真名** —— 本仓库是公开的，真名不能出门；
  2. 两页背景图是**第三方版权素材**（带水印的预览图）—— 放进公开仓库不行；
  3. 而配方的证据要能「重新核对」—— 所以保留**全部结构**，只换掉上面两类。

副本与原件的差异（**除以下三处外逐字节不变**）：
  · ppt/media/image1.jpeg  -> 生成的中性底图（同尺寸 JPEG）
  · ppt/media/image2.png   -> 生成的中性底图（同尺寸 PNG）
  · docProps/core.xml      -> creator / lastModifiedBy **清空**（沿用入库夹具先例；
                              正则洗，不在脚本里出现原名）
  · （若有）docProps/thumbnail.jpeg -> 换中性缩略图（缩略图会泄露背景图！）

结构（2 页、每页 8 片 pie、rot 0→-105、useBgFill、innerShdw、p:bg）逐值保留 ——
搬完**当场断言**（原件与副本的结构快照必须完全相等）。之后由
`tests/test_docs.py` 第 24 组继续守着，改坏了会被拦。

用法：
    python tools/make_lotus_fixture.py <原件路径> <输出路径>
"""

import io
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

from PIL import Image, ImageDraw

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
SLIDES = ("ppt/slides/slide1.xml", "ppt/slides/slide2.xml")


def structure(path):
    """结构快照：两份 slide 里每片 pie 的值 + 过渡计数（p:transition / p159:morph）。

    过渡也要进来 —— 它藏在 `mc:AlternateContent` 里（兼容写法），
    不数它的话，"副本没动结构"的断言会漏掉动画那一半。
    """
    z = zipfile.ZipFile(path)
    snap = {}
    for sn in SLIDES:
        raw = z.read(sn).decode("utf-8")
        root = ET.fromstring(raw)
        rows = []
        for sp in root.iter(P + "sp"):
            pg = sp.find(".//" + A + "prstGeom")
            if pg is None or pg.get("prst") != "pie":
                continue
            xf = sp.find(".//" + A + "xfrm")
            o, e = xf.find(A + "off"), xf.find(A + "ext")
            rows.append((sp.get("useBgFill"), xf.get("rot"),
                         int(o.get("x")), int(o.get("y")),
                         int(e.get("cx")), int(e.get("cy")),
                         sp.find(".//" + A + "innerShdw") is not None))
        snap[sn] = {"pies": rows,
                    "transitions": raw.count("<p:transition"),
                    "morph": raw.count("p159:morph")}
    return snap


def neutral(size, tint):
    """中性底图：浅灰 + 网格 + 对角纹 + 标注 —— 让「pie 继承背景」仍可辨。"""
    im = Image.new("RGB", size, tint)
    d = ImageDraw.Draw(im)
    w, h = size
    for i in range(-h, w, 48):
        d.line([(i, h), (i + h, 0)], fill=(tint[0] - 12, tint[1] - 12, tint[2] - 12))
    for x in range(0, w, 96):
        d.line([(x, 0), (x, h)], fill=(tint[0] - 6, tint[1] - 6, tint[2] - 6))
    for y in range(0, h, 96):
        d.line([(0, y), (w, y)], fill=(tint[0] - 6, tint[1] - 6, tint[2] - 6))
    d.rectangle([0, 0, w - 1, h - 1], outline=(170, 178, 188))
    d.text((14, 12), "structure evidence copy - ppt-motion-skill", fill=(96, 104, 116))
    return im


def scrub_core(text):
    """把 creator / lastModifiedBy **清空** —— 沿用入库夹具的既有先例
    （`powerpoint-native.pptx`、material 两份模板的作者值都是空的）。
    不引原名，也不留任何名义值；来源说明在工具本文档与第 24 组里。"""
    text = re.sub(r"(<dc:creator>).*?(</dc:creator>)", r"\1\2", text, flags=re.S)
    text = re.sub(r"(<cp:lastModifiedBy>).*?(</cp:lastModifiedBy>)", r"\1\2",
                  text, flags=re.S)
    return text


def main():
    if len(sys.argv) != 3:
        raise SystemExit(__doc__.strip().splitlines()[-1])
    src, dst = sys.argv[1], sys.argv[2]

    zim = zipfile.ZipFile(src)
    im1 = Image.open(io.BytesIO(zim.read("ppt/media/image1.jpeg")))
    im2 = Image.open(io.BytesIO(zim.read("ppt/media/image2.png")))
    bj = io.BytesIO(); neutral(im1.size, (238, 240, 244)).save(bj, "JPEG", quality=82)
    bp = io.BytesIO(); neutral(im2.size, (240, 244, 238)).save(bp, "PNG")

    os.makedirs(os.path.dirname(dst), exist_ok=True)
    replaced = [0]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zim.namelist():
            data = zim.read(n)
            if n == "ppt/media/image1.jpeg":
                data = bj.getvalue(); replaced[0] += 1
            elif n == "ppt/media/image2.png":
                data = bp.getvalue(); replaced[0] += 1
            elif n == "docProps/core.xml":
                data = scrub_core(data.decode("utf-8")).encode("utf-8")
            elif n == "docProps/thumbnail.jpeg":
                tb = io.BytesIO(); neutral((256, 144), (238, 240, 244)).save(tb, "JPEG", quality=80)
                data = tb.getvalue(); replaced[0] += 1
            zout.writestr(n, data)

    # —— 搬完当场断言：结构与原件逐值一致（§六.4.2：断言放进执行路径）——
    a, b = structure(src), structure(dst)
    assert a == b, "结构断言失败：副本与原件的结构快照不一致！"
    n_pie = sum(len(v["pies"]) for v in b.values())

    print("原件 : %s（%d 字节）" % (src, os.path.getsize(src)))
    print("副本 : %s（%d 字节，替换了 %d 个实体）" % (dst, os.path.getsize(dst), replaced[0]))
    print("结构 : 2 页 · 每页 %d 片 pie · 过渡 morph=%d · 断言与原件逐值相等 ✓"
          % (n_pie // 2, b[SLIDES[1]]["morph"]))
    print("      rot 序列:", sorted(int(r) // 60000 if r else 0
                                   for r in [x[1] for x in b[SLIDES[1]]["pies"]]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
