# -*- coding: utf-8 -*-
"""Generate the design-system index for this skill.

Vendored files are listed with their local path; the rest of the upstream
catalogue is listed with its raw URL so the agent can fetch one on demand
without cloning the repo.
"""
import json
import os

DEST = r"D:\idea\dsh-ppt-office-motion\references\design-system"
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


def main():
    lines = []
    lines.append("# 设计系统索引（静态版面）")
    lines.append("")
    lines.append("本目录为**静态版面**设计资产，来源 [open-kimi-ppt-skill]"
                 "(https://github.com/acnlie/open-kimi-ppt-skill)（MIT，"
                 "Copyright (c) 2026 Binaryify Zhuang，全文见 `LICENSE-upstream.txt`）。")
    lines.append("")
    lines.append("分工：**这些文件管静态版面**（色板、字体、布局骨架、图表语言、页型版式、禁止项）；")
    lines.append("**动效归 `references/motion-design-spec.md`**。上游文件本身不含动效指导，"
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

    out = os.path.join(DEST, "README.md")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))
    print("wrote", out, os.path.getsize(out), "bytes")
    print("vendored=%d remote=%d total=%d" % (len(vendored), len(remote), len(CATALOG)))


if __name__ == "__main__":
    main()
