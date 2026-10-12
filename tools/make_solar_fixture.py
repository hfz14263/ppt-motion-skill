#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把注入动效后的 deck 洗成**结构复核副本**（fixture），供 test_docs 逐值核对。

为什么需要副本而不是直接测 out/solar.pptx：
- `out/` 是构建产物，不入库（每次重建都会变）
- 入库的是**断言的对象**。判据必须钉在一个稳定文件上，否则测试在
  "产物没构建" 与 "构建了但结构不对" 两种情况下给出同样的失败。

洗什么：
1. `docProps/core.xml` 与 `app.xml` 的作者/公司字段清空
   （python-pptx 建的 deck 本来就是空的，但**必须断言它为空** ——
   不写这一条，将来有人用 PowerPoint 打开另存就会把真名带进来）
2. 媒体原样保留（都是程序化生成的，无版权问题）

不改任何结构。洗完当场断言"结构与源逐值相等"。

用法：
    python tools/make_solar_fixture.py
"""
from __future__ import annotations

import pathlib
import re
import sys
import zipfile

ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC = ROOT / "out" / "solar.pptx"
DST = ROOT / "tests" / "fixtures" / "solar-system-structure.pptx"

# 必须清空的元数据字段（沿用 make_lotus_fixture.py 的先例）
SCRUB = {
    "docProps/core.xml": [
        r"<dc:creator>.*?</dc:creator>",
        r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>",
    ],
    "docProps/app.xml": [
        r"<Company>.*?</Company>",
        r"<Manager>.*?</Manager>",
    ],
}

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def structure(z: zipfile.ZipFile) -> list:
    """抽出可比的结构特征：每页的 shape（名 + 类型 + 几何 + 效果数 + morph）。"""
    import xml.etree.ElementTree as ET
    out = []
    names = sorted(n for n in z.namelist()
                   if re.fullmatch(r"ppt/slides/slide\d+\.xml", n))
    for n in names:
        xml = z.read(n)
        root = ET.fromstring(xml)
        shapes = []
        for sp in root.iter():
            if not sp.tag.endswith("}sp") and not sp.tag.endswith("}pic") \
                    and not sp.tag.endswith("}cxnSp"):
                continue
            nv = sp.find(f".//{P}cNvPr")
            if nv is None:
                continue
            xfrm = sp.find(f".//{A}xfrm")
            off = xfrm.find(A + "off") if xfrm is not None else None
            ext = xfrm.find(A + "ext") if xfrm is not None else None
            shapes.append({
                "name": nv.get("name"),
                "kind": sp.tag.split("}")[1],
                "x": int(off.get("x")) if off is not None else None,
                "y": int(off.get("y")) if off is not None else None,
                "cx": int(ext.get("cx")) if ext is not None else None,
                "cy": int(ext.get("cy")) if ext is not None else None,
            })
        shapes.sort(key=lambda s: (s["name"] or ""))
        timing = root.find(f"{P}timing")
        effects = 0
        if timing is not None:
            effects = len(timing.findall(f".//{P}par"))
        raw = xml.decode("utf-8", "replace")
        out.append({
            "slide": n,
            "shapes": shapes,
            "effects_present": timing is not None,
            "par_count": effects,
            "morph": "p159:morph" in raw,
        })
    return out


def main() -> int:
    if not SRC.exists():
        sys.exit("先构建：python examples/material/solar/build_deck.py && "
                 "python scripts/motion.py apply --pptx examples/material/solar/"
                 "solar-system.pptx --spec examples/solar-motion.yaml "
                 "--out out/solar.pptx")
    DST.parent.mkdir(parents=True, exist_ok=True)

    zin = zipfile.ZipFile(SRC)
    patched = {}
    for name, pats in SCRUB.items():
        if name not in zin.namelist():
            continue
        raw = zin.read(name).decode("utf-8")
        for pat in pats:
            raw = re.sub(pat, "", raw)
        patched[name] = raw.encode("utf-8")

    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, patched.get(item.filename,
                                             zin.read(item.filename)))

    # 结构自证：洗完必须与源逐值相等（字段已删，但形状/几何/效果数不能变）
    a = structure(zipfile.ZipFile(SRC))
    b = structure(zipfile.ZipFile(DST))
    assert a == b, "洗元数据时结构被改动了 —— 这正是莲花那条教训"
    print("structure identical after scrub: %d slide(s)" % len(a))
    print("saved: %s (%d bytes)" % (DST.relative_to(ROOT), DST.stat().st_size))

    # 顺带断言元数据确实空了
    z = zipfile.ZipFile(DST)
    core = z.read("docProps/core.xml").decode("utf-8")
    for pat in (r"<dc:creator>\s*\S", r"<cp:lastModifiedBy>\s*\S"):
        assert not re.search(pat, core), "元数据没洗干净: %s" % pat
    print("docProps clean")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
