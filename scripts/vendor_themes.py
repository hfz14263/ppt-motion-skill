# -*- coding: utf-8 -*-
"""Vendor a curated set of design systems from open-kimi-ppt-skill (MIT).

Upstream: https://github.com/acnlie/open-kimi-ppt-skill
License : MIT, Copyright (c) 2026 Binaryify Zhuang

Only the design.md text files are vendored (no images, no code). The upstream
LICENSE text is copied alongside them as references/design-system/LICENSE-upstream.txt
so the redistribution condition ("include the copyright notice") is met.
"""
import json
import os
import sys
import urllib.request

RAW = "https://raw.githubusercontent.com/acnlie/open-kimi-ppt-skill/main/skills/open-kimi-ppt/reference"
DEST = r"D:\idea\dsh-ppt-office-motion\references\design-system"

# id, upstream relative path, category, one-line purpose (Chinese, used for matching)
THEMES = [
    ("pine-green-strategy",        "design_system/consulting/pine-green-strategy/design.md",        "consulting", "深绿科技题头带 + 薄荷强调；战略咨询、尽调、高管汇报，高信息密度"),
    ("indigo-due-diligence",       "design_system/consulting/indigo-due-diligence/design.md",       "consulting", "靛蓝尽调风；麦肯锡式信息密度与框架图"),
    ("black-gold-ledger",          "design_system/finance/black-gold-ledger/design.md",             "finance",    "黑金台账风；机构级投研报告、年报"),
    ("rice-paper-annual",          "design_system/finance/rice-paper-annual/design.md",             "finance",    "宣纸年报风；米白纸感大标题编辑风"),
    ("blue-flame-brand",           "design_system/work/blue-flame-brand/design.md",                 "work",       "蓝焰品牌风；结论先行的业务复盘、经营汇报"),
    ("moon-white-imagery",         "design_system/work/moon-white-imagery/design.md",               "work",       "月白影像风；大图叙事 + 经营结论"),
    ("silk-yellow-magazine",       "design_system/promotion/silk-yellow-magazine/design.md",        "promotion",  "丝黄杂志风；高密度杂志编辑排版、品牌刊"),
    ("silver-gray-luxury-magazine","design_system/promotion/silver-gray-luxury-magazine/design.md", "promotion",  "银灰奢华杂志风；高端品牌刊感"),
    ("paper-white-courseware",     "design_system/academic/paper-white-courseware/design.md",       "academic",   "纸白课件风；论文答辩级排版"),
    ("wine-red-data",              "design_system/academic/wine-red-data/design.md",                "academic",   "酒红数据风；学术数据可视化"),
]


def get(url):
    req = urllib.request.Request(url, headers={"User-Agent": "dsh-ppt-office-motion/vendor"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def main():
    os.makedirs(DEST, exist_ok=True)
    index = []
    ok = 0
    for tid, rel, cat, purpose in THEMES:
        url = "%s/%s" % (RAW, rel)
        out = os.path.join(DEST, tid + ".md")
        try:
            data = get(url)
        except Exception as exc:
            print("FAIL %-30s %s" % (tid, exc))
            continue
        header = (
            "<!--\n"
            "Vendored from open-kimi-ppt-skill (MIT, Copyright (c) 2026 Binaryify Zhuang)\n"
            "Source: %s\n"
            "This file is upstream content, kept verbatim except for this header.\n"
            "Upstream design systems govern STATIC layout only; motion is governed by\n"
            "references/motion-design-spec.md in this skill.\n"
            "-->\n\n" % url
        ).encode("utf-8")
        with open(out, "wb") as fh:
            fh.write(header + data)
        size = os.path.getsize(out)
        print("OK   %-30s %6d bytes  (%s)" % (tid, size, cat))
        ok += 1
        index.append({"id": tid, "category": cat, "purpose": purpose,
                      "file": tid + ".md", "upstream": rel, "bytes": size})

    # upstream LICENSE, to satisfy the redistribution condition
    try:
        lic = get("https://raw.githubusercontent.com/acnlie/open-kimi-ppt-skill/main/LICENSE")
        with open(os.path.join(DEST, "LICENSE-upstream.txt"), "wb") as fh:
            fh.write(lic)
        print("OK   LICENSE-upstream.txt")
    except Exception as exc:
        print("FAIL LICENSE :: %s" % exc)

    with open(os.path.join(DEST, "themes.json"), "w", encoding="utf-8") as fh:
        json.dump({"source": "https://github.com/acnlie/open-kimi-ppt-skill",
                   "license": "MIT (c) 2026 Binaryify Zhuang",
                   "note": "static layout only; motion governed by references/motion-design-spec.md",
                   "themes": index}, fh, ensure_ascii=False, indent=1)
    print("\nvendored %d/%d themes -> %s" % (ok, len(THEMES), DEST))


if __name__ == "__main__":
    sys.exit(main())
