# -*- coding: utf-8 -*-
"""Generate the design-system index for this skill.

Vendored files are listed with their local path; the rest of the upstream
catalogue is listed with its raw URL so the agent can fetch one on demand
without cloning the repo.

本文件是 README 的**唯一来源**：正文改动要改这里，不要在 README 里手改 ——
`--check` 会对比生成器输出与磁盘文件，手改的部分会被判为分叉
（2026-10-09 实际丢过一次：手工加的 design_compose 章节被重跑吞掉 35 行）。
"""
import argparse
import json
import os
import sys

# this file lives IN the design-system directory
DEST = os.path.dirname(os.path.abspath(__file__))
RAW = ("https://raw.githubusercontent.com/acnlie/open-kimi-ppt-skill/main/"
       "skills/open-kimi-ppt/reference/design_system")

# full upstream catalogue: (id, category, upstream relative path, one-line, vendored?)
CATALOG = [
    # ---- consulting ----
    ("apricot-white-brief", "consulting", "consulting/apricot-white-brief/design.md",
     "杏白底咨询简报风，高密度论证、细线图表", False),
    ("indigo-due-diligence", "consulting", "consulting/indigo-due-diligence/design.md",
     "靛蓝尽调风，麦肯锡式信息密度与框架图", True),
    ("marine-blue-research", "consulting", "consulting/marine-blue-research/design.md",
     "海蓝研究风，专业咨询研究报告版式", False),
    ("moss-green-transformation", "consulting", "consulting/moss-green-transformation/design.md",
     "苔绿转型风，战略落地与变革叙事", False),
    ("pine-green-strategy", "consulting", "consulting/pine-green-strategy/design.md",
     "松绿战略风，深色题头条 + 薄荷绿强调", True),
    ("red-black-growth", "consulting", "consulting/red-black-growth/design.md",
     "红黑增长风，强对比结论导向", False),
    # ---- finance ----
    ("black-gold-ledger", "finance", "finance/black-gold-ledger/design.md",
     "黑金台账风，机构级投研报告", True),
    ("ebony-ledger", "finance", "finance/ebony-ledger/design.md",
     "乌木台账风，深色金融研究气质", False),
    ("honey-orange-memo", "finance", "finance/honey-orange-memo/design.md",
     "蜜橙备忘录风，温暖商务投研备忘", False),
    ("lake-blue-memo", "finance", "finance/lake-blue-memo/design.md",
     "湖蓝备忘录风，清爽财务研究备忘", False),
    ("prospect-annual", "finance", "finance/prospect-annual/design.md",
     "展望年报风，机构年度研究出版感", False),
    ("rice-paper-annual", "finance", "finance/rice-paper-annual/design.md",
     "宣纸年报风，米白纸感大标题编辑风", True),
    # ---- work ----
    ("blue-flame-brand", "work", "work/blue-flame-brand/design.md",
     "蓝焰品牌风，结论先行的业务复盘", True),
    ("electric-violet-business", "work", "work/electric-violet-business/design.md",
     "电紫商务风，经营复盘与进展汇报", False),
    ("moon-white-imagery", "work", "work/moon-white-imagery/design.md",
     "月白影像风，大图叙事 + 经营结论", True),
    ("sky-blue-wayfinding", "work", "work/sky-blue-wayfinding/design.md",
     "天蓝导视风，清晰导航式工作汇报", False),
    ("warm-clay-works", "work", "work/warm-clay-works/design.md",
     "暖陶作品风，高完成度作品集式汇报", False),
    ("warm-jade-annual-report", "work", "work/warm-jade-annual-report/design.md",
     "暖玉年报风，经营年报与复盘", False),
    # ---- promotion ----
    ("aqua-charity-report", "promotion", "promotion/aqua-charity-report/design.md",
     "水色公益报告风，品牌公益传播", False),
    ("cream-collage", "promotion", "promotion/cream-collage/design.md",
     "奶油拼贴风，情绪化品牌拼贴叙事", False),
    ("pine-soot-pictorial", "promotion", "promotion/pine-soot-pictorial/design.md",
     "松烟画报风，画报感品牌年刊", False),
    ("silk-yellow-magazine", "promotion", "promotion/silk-yellow-magazine/design.md",
     "丝黄杂志风，高密度杂志编辑排版", True),
    ("silver-gray-luxury-magazine", "promotion", "promotion/silver-gray-luxury-magazine/design.md",
     "银灰奢华杂志风，高端品牌刊感", True),
    ("travel-green-handbook", "promotion", "promotion/travel-green-handbook/design.md",
     "旅行绿手册风，旅行/目的地手册气质", False),
    # ---- academic ----
    ("blue-line-courseware", "academic", "academic/blue-line-courseware/design.md",
     "蓝线课件风，严谨学术课件排版", False),
    ("deep-blue-atlas", "academic", "academic/deep-blue-atlas/design.md",
     "深蓝图集风，研究讲座与图集叙事", False),
    ("paper-white-courseware", "academic", "academic/paper-white-courseware/design.md",
     "纸白课件风，论文答辩级排版", True),
    ("pastel-derivation", "academic", "academic/pastel-derivation/design.md",
     "粉彩推导风，公式/推导渐进揭示", False),
    ("teal-green-academic-defense", "academic", "academic/teal-green-academic-defense/design.md",
     "青绿答辩风，硕博答辩正式感", False),
    ("wine-red-data", "academic", "academic/wine-red-data/design.md",
     "酒红数据风，学术数据可视化", True),
    # ---- extra (numbered English set) ----
    ("dusk-violet-consulting", "extra/strategy", "01_strategy/01/en/dusk-violet-consulting.md",
     "暮紫咨询思想领导力报告风", False),
    ("red-black-business", "extra/strategy", "01_strategy/04/en/red-black-business.md",
     "红黑管理咨询全案模板风", False),
    ("map-strategy", "extra/strategy", "01_strategy/06/en/map-strategy.md",
     "地图战略风，国际化/区域布局", False),
    ("xuan-paper-annual", "extra/business", "02_business/03/en/xuan-paper-annual.md",
     "宣纸年刊风，超重标题编辑研究刊", False),
    ("lead-gray-quarterly", "extra/business", "02_business/04/en/lead-gray-quarterly.md",
     "铅灰季报风，财经报纸式数据监测", False),
    ("ink-green-market-trends", "extra/business", "02_business/05/en/ink-green-market-trends.md",
     "墨绿行情风，深色加密/市场季报", False),
    ("orange-tech", "extra/business", "02_business/06/en/orange-tech.md",
     "橙色科技年报与商业洞察风", False),
    ("mist-blue-travelogue", "extra/work", "03_work/02/en/mist-blue-travelogue.md",
     "雾蓝游记风，旅行酒店行业研究", False),
    ("color-stripes-documentary", "extra/work", "03_work/04/en/color-stripes-documentary.md",
     "彩条纪实风，影响力/ESG 长叙事", False),
    ("red-white-business", "extra/work", "03_work/06/en/red-white-business.md",
     "红白商务风，雇主品牌与组织能力", False),
    ("gold-orange-type-journal", "extra/promotion", "04_promotion/04/en/gold-orange-type-journal.md",
     "金橙字体趋势刊风", False),
    ("fresh-brand", "extra/promotion", "04_promotion/06/en/fresh-brand.md",
     "清新品牌权益/影响力报告风", False),
    ("pink-purple-diagnosis", "extra/academic", "05_academic/01/en/pink-purple-diagnosis.md",
     "粉紫云层诊断叙事风", False),
    ("dark-themed-data", "extra/academic", "05_academic/06/en/dark-themed-data.md",
     "深色数据主题，加密/市场研究仪表盘", False),
]

CATEGORY_CN = {
    "consulting": "咨询策略", "finance": "商务财务", "work": "工作汇报",
    "promotion": "品牌推广", "academic": "学术教育",
    "extra/strategy": "补充·策略", "extra/business": "补充·商务",
    "extra/work": "补充·工作", "extra/promotion": "补充·推广",
    "extra/academic": "补充·学术",
}

# design_compose 章节原文（2026-10-09 从 README 搬入生成器）。
# 原文只存在于 README —— 重跑生成器会静默丢掉它。搬进来之后，生成器才是唯一来源。
SECTION_COMPOSE = [
    '## 让设计系统可执行（`scripts/design_compose.py`）',
    '',
    '这些文件是散文 —— 人能读，程序读不了。所以**每一份"按 X 风格做的 deck"其实都在跟着人对 X 的记忆走**，',
    '这正是不同风格最后长得都差不多的原因。',
    '',
    '`design_compose.py` 做最小可用的一步：**把调色板从散文里解析出来，再用它生成页面。**',
    '',
    '```bash',
    'python scripts/design_compose.py --list',
    'python scripts/design_compose.py --theme pine-green-strategy --show   # 看解析结果',
    'python scripts/design_compose.py --all --outdir composed              # 10 套各出一页',
    '```',
    '',
    '实测：**10 套全部定位到调色板分区，共提取 160 个 hex 码**，带角色标注',
    '（`background` / `structural` / `accent` / `negative` / `neutral`）。',
    '同一份内容套 10 套，两两平均色距中位 ~33、最远 400 —— **风格差异是量出来的，不是感觉。**',
    '',
    '### 解析调色板时踩的两个坑（都是"过度校正"）',
    '',
    '1. **取了 `neutral[0]` 当正文色** → 拿到 `#F1F1F1`，那是系统指定给「解释栏**底色**」的，',
    '   压在纯白页面上**正文直接看不见**。角色名不等于用途，位置不等于语义。',
    '2. **改成"取对比度最高的"** → 黑白永远胜出，`#04512C` 深林绿被换成 `#111111`，',
    '   **可读性满分、品牌识别归零**。现在的规则：可读是**下限**不是目标，',
    '   结构色与强调色优先保留**有色相且过对比度**的候选。',
    '',
    '还有一个是**项目自己的审计工具抓出来的**：早期把 `other` 桶（图表系列色）也算进强调色候选，',
    '一页上出现 5 个饱和色，被 `review_assist` 判为 **rainbow risk** ——',
    '而"一页主导色 ≤2"正是这些系统明写的禁止项。**生成器的输出必须过自己的检查。**',
    '',
    '### 明确未实现（不假装）',
    '',
    '- **摄影题头带**：多套要求"变暗的人物/科技照片横带"，此环境无图像生成',
    '- **图表语言**：每套对图表形态有详细规定，尚未生成',
    '- **密度仍偏低**：实测非背景像素约 8.6–9.0%，而系统要求"正文不留实心空白块"。',
    '  当前版面已不空，但离真正的高密度咨询页还有距离',
]


def build_text():
    """产出 README 全文。**生成器是唯一来源** —— 不要在产物里手改：
    手改的内容要么搬进这里，要么会被 `--check` 判为分叉。"""
    lines = []
    lines.append("# 设计系统索引（静态版面）")
    lines.append("")
    lines.append("本目录为**静态版面**设计资产，来源 [open-kimi-ppt-skill]"
                 "(https://github.com/acnlie/open-kimi-ppt-skill)（MIT，"
                 "Copyright (c) 2026 Binaryify Zhuang，全文见 `LICENSE-upstream.txt`）。")
    lines.append("")
    lines.append("分工：**这些文件管静态版面**（色板、字体、布局骨架、图表语言、页型版式、禁止项）；")
    lines.append("**动效归 `reference/motion-design-spec.md`**。上游文件本身不含动效指导，"
                 "不要指望从里面读出节奏建议。")
    lines.append("")
    vendored = [c for c in CATALOG if c[4]]
    remote = [c for c in CATALOG if not c[4]]
    lines.append("已内置 **%d 套**（离线可用），其余 **%d 套**按需从上游抓取。"
                 % (len(vendored), len(remote)))
    lines.append("")

    # ---------- vendored ----------
    lines.append("## 已内置（离线可用）")
    lines.append("")
    lines.append("| id | 分类 | 适用 | 文件 |")
    lines.append("| --- | --- | --- | --- |")
    for tid, cat, rel, purpose, _ in sorted(vendored, key=lambda c: (c[1], c[0])):
        lines.append("| `%s` | %s | %s | `%s.md` |" % (tid, CATEGORY_CN.get(cat, cat), purpose, tid))
    lines.append("")
    lines.append("用法：先按「分类 + 适用」挑一套，再读对应 `.md` 全文作为**唯一风格源**，"
                 "不要混搭多套（上游明确要求不得混用其他风格）。")
    lines.append("")
    lines.extend(SECTION_COMPOSE)   # design_compose 那一节（2026-10-09 从 README 搬进来）
    lines.append("")

    # ---------- remote ----------
    lines.append("## 上游其余（按需抓取）")
    lines.append("")
    lines.append("用 `web_fetch` 取下表 URL 的原文即可，无需 clone 仓库：")
    lines.append("")
    lines.append("| id | 分类 | 适用 | 上游路径 |")
    lines.append("| --- | --- | --- | --- |")
    for tid, cat, rel, purpose, _ in sorted(remote, key=lambda c: (c[1], c[0])):
        lines.append("| `%s` | %s | %s | `%s` |" % (tid, CATEGORY_CN.get(cat, cat), purpose, rel))
    lines.append("")
    lines.append("URL 前缀：")
    lines.append("")
    lines.append("```")
    lines.append(RAW + "/<上游路径>")
    lines.append("```")
    lines.append("")
    lines.append("例如要取 `apricot-white-brief`：")
    lines.append("")
    lines.append("```")
    lines.append(RAW + "/consulting/apricot-white-brief/design.md")
    lines.append("```")
    lines.append("")
    lines.append("想批量内置，改 `scripts/vendor_themes.py` 的 `THEMES` 表后重跑即可。")
    lines.append("")

    # ---------- companion docs ----------
    lines.append("## 配套文档（同一上游）")
    lines.append("")
    lines.append("| 主题 | 上游路径 | 用途 |")
    lines.append("| --- | --- | --- |")
    for name, rel, use in (
        ("7 类场景版面基线", "slides_categories.md", "先定场景再选风格；7 类：分析决策/商业提案/管理汇报/学术研究/教育培训/技术工程/品牌创意"),
        ("场景明细", "slides_categories/analysis-decision.md 等 7 个文件", "每类的读者任务与版面做法"),
        ("字体规范", "fonts.md", "字体族与字号层级"),
        ("形状库", "shapes.md", "可用形状清单"),
        ("PPTX 动画语法", "pptd.md §6", "上游自己的动画 DSL，可与本 skill 对照"),
    ):
        lines.append("| %s | `%s` | %s |" % (name, rel, use))
    lines.append("")
    lines.append("这些也走同一个 URL 前缀。")
    lines.append("")

    return "\n".join(lines)


def missing_vendored():
    """「已内置」表里声明了的文件，是否真的在磁盘上。"""
    return [tid for tid, _cat, _rel, _p, ven in CATALOG
            if ven and not os.path.exists(os.path.join(DEST, tid + ".md"))]


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="生成 / 校验 design-system 的 README 索引（--check 只读）")
    ap.add_argument("--check", action="store_true",
                    help="只对比 README 与生成器输出，不写文件")
    ns = ap.parse_args(argv)

    text = build_text()
    out = os.path.join(DEST, "README.md")
    miss = missing_vendored()
    n_ven = len([c for c in CATALOG if c[4]])
    n_rem = len(CATALOG) - n_ven

    if ns.check:
        bad = []
        if not os.path.exists(out):
            bad.append("README.md 不存在")
        else:
            disk = open(out, encoding="utf-8", newline="").read().replace("\r\n", "\n")
            if disk != text:
                bad.append("README.md 与生成器输出不一致"
                           "（有人手改过，或生成器变了而没重跑）")
        for tid in miss:
            bad.append("「已内置」声明了 %s，但 %s.md 不在磁盘上" % (tid, tid))
        if bad:
            for b in bad:
                print("!! " + b)
            print("修法：重跑 `python reference/design-system/build_index.py`；")
            print("      若 README 里有手工内容 —— 搬进生成器（它才是唯一来源），别留着。")
            return 1
        print("OK  README 与生成器一致（vendored=%d remote=%d）" % (n_ven, n_rem))
        return 0

    # 仓库约定 LF（.gitattributes text=auto eol=lf）——
    # 不显式指定的话 Windows 上会写成 CRLF，重跑一次整个文件全行 diff。
    with open(out, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(text)
    print("wrote", out, os.path.getsize(out), "bytes")
    print("vendored=%d remote=%d total=%d" % (n_ven, n_rem, len(CATALOG)))
    for tid in miss:
        print("WARN 「已内置」声明了 %s，但 %s.md 不在磁盘上" % (tid, tid))
    return 0


if __name__ == "__main__":
    sys.exit(main())
