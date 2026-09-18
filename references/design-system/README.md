# 设计系统索引（静态版面）

本目录为**静态版面**设计资产，来源 [open-kimi-ppt-skill](https://github.com/acnlie/open-kimi-ppt-skill)（MIT，Copyright (c) 2026 Binaryify Zhuang，全文见 `LICENSE-upstream.txt`）。

分工：**这些文件管静态版面**（色板、字体、布局骨架、图表语言、页型版式、禁止项）；
**动效归 `references/motion-design-spec.md`**。上游文件本身不含动效指导，不要指望从里面读出节奏建议。

已内置 **10 套**（离线可用），其余 **34 套**按需从上游抓取。

## 已内置（离线可用）

| id | 分类 | 适用 | 文件 |
| --- | --- | --- | --- |
| `paper-white-courseware` | 学术教育 | 纸白课件风，论文答辩级排版 | `paper-white-courseware.md` |
| `wine-red-data` | 学术教育 | 酒红数据风，学术数据可视化 | `wine-red-data.md` |
| `indigo-due-diligence` | 咨询策略 | 靛蓝尽调风，麦肯锡式信息密度与框架图 | `indigo-due-diligence.md` |
| `pine-green-strategy` | 咨询策略 | 松绿战略风，深色题头条 + 薄荷绿强调 | `pine-green-strategy.md` |
| `black-gold-ledger` | 商务财务 | 黑金台账风，机构级投研报告 | `black-gold-ledger.md` |
| `rice-paper-annual` | 商务财务 | 宣纸年报风，米白纸感大标题编辑风 | `rice-paper-annual.md` |
| `silk-yellow-magazine` | 品牌推广 | 丝黄杂志风，高密度杂志编辑排版 | `silk-yellow-magazine.md` |
| `silver-gray-luxury-magazine` | 品牌推广 | 银灰奢华杂志风，高端品牌刊感 | `silver-gray-luxury-magazine.md` |
| `blue-flame-brand` | 工作汇报 | 蓝焰品牌风，结论先行的业务复盘 | `blue-flame-brand.md` |
| `moon-white-imagery` | 工作汇报 | 月白影像风，大图叙事 + 经营结论 | `moon-white-imagery.md` |

用法：先按「分类 + 适用」挑一套，再读对应 `.md` 全文作为**唯一风格源**，不要混搭多套（上游明确要求不得混用其他风格）。

## 上游其余（按需抓取）

用 `web_fetch` 取下表 URL 的原文即可，无需 clone 仓库：

| id | 分类 | 适用 | 上游路径 |
| --- | --- | --- | --- |
| `blue-line-courseware` | 学术教育 | 蓝线课件风，严谨学术课件排版 | `academic/blue-line-courseware/design.md` |
| `deep-blue-atlas` | 学术教育 | 深蓝图集风，研究讲座与图集叙事 | `academic/deep-blue-atlas/design.md` |
| `pastel-derivation` | 学术教育 | 粉彩推导风，公式/推导渐进揭示 | `academic/pastel-derivation/design.md` |
| `teal-green-academic-defense` | 学术教育 | 青绿答辩风，硕博答辩正式感 | `academic/teal-green-academic-defense/design.md` |
| `apricot-white-brief` | 咨询策略 | 杏白底咨询简报风，高密度论证、细线图表 | `consulting/apricot-white-brief/design.md` |
| `marine-blue-research` | 咨询策略 | 海蓝研究风，专业咨询研究报告版式 | `consulting/marine-blue-research/design.md` |
| `moss-green-transformation` | 咨询策略 | 苔绿转型风，战略落地与变革叙事 | `consulting/moss-green-transformation/design.md` |
| `red-black-growth` | 咨询策略 | 红黑增长风，强对比结论导向 | `consulting/red-black-growth/design.md` |
| `dark-themed-data` | 补充·学术 | 深色数据主题，加密/市场研究仪表盘 | `05_academic/06/en/dark-themed-data.md` |
| `pink-purple-diagnosis` | 补充·学术 | 粉紫云层诊断叙事风 | `05_academic/01/en/pink-purple-diagnosis.md` |
| `ink-green-market-trends` | 补充·商务 | 墨绿行情风，深色加密/市场季报 | `02_business/05/en/ink-green-market-trends.md` |
| `lead-gray-quarterly` | 补充·商务 | 铅灰季报风，财经报纸式数据监测 | `02_business/04/en/lead-gray-quarterly.md` |
| `orange-tech` | 补充·商务 | 橙色科技年报与商业洞察风 | `02_business/06/en/orange-tech.md` |
| `xuan-paper-annual` | 补充·商务 | 宣纸年刊风，超重标题编辑研究刊 | `02_business/03/en/xuan-paper-annual.md` |
| `fresh-brand` | 补充·推广 | 清新品牌权益/影响力报告风 | `04_promotion/06/en/fresh-brand.md` |
| `gold-orange-type-journal` | 补充·推广 | 金橙字体趋势刊风 | `04_promotion/04/en/gold-orange-type-journal.md` |
| `dusk-violet-consulting` | 补充·策略 | 暮紫咨询思想领导力报告风 | `01_strategy/01/en/dusk-violet-consulting.md` |
| `map-strategy` | 补充·策略 | 地图战略风，国际化/区域布局 | `01_strategy/06/en/map-strategy.md` |
| `red-black-business` | 补充·策略 | 红黑管理咨询全案模板风 | `01_strategy/04/en/red-black-business.md` |
| `color-stripes-documentary` | 补充·工作 | 彩条纪实风，影响力/ESG 长叙事 | `03_work/04/en/color-stripes-documentary.md` |
| `mist-blue-travelogue` | 补充·工作 | 雾蓝游记风，旅行酒店行业研究 | `03_work/02/en/mist-blue-travelogue.md` |
| `red-white-business` | 补充·工作 | 红白商务风，雇主品牌与组织能力 | `03_work/06/en/red-white-business.md` |
| `ebony-ledger` | 商务财务 | 乌木台账风，深色金融研究气质 | `finance/ebony-ledger/design.md` |
| `honey-orange-memo` | 商务财务 | 蜜橙备忘录风，温暖商务投研备忘 | `finance/honey-orange-memo/design.md` |
| `lake-blue-memo` | 商务财务 | 湖蓝备忘录风，清爽财务研究备忘 | `finance/lake-blue-memo/design.md` |
| `prospect-annual` | 商务财务 | 展望年报风，机构年度研究出版感 | `finance/prospect-annual/design.md` |
| `aqua-charity-report` | 品牌推广 | 水色公益报告风，品牌公益传播 | `promotion/aqua-charity-report/design.md` |
| `cream-collage` | 品牌推广 | 奶油拼贴风，情绪化品牌拼贴叙事 | `promotion/cream-collage/design.md` |
| `pine-soot-pictorial` | 品牌推广 | 松烟画报风，画报感品牌年刊 | `promotion/pine-soot-pictorial/design.md` |
| `travel-green-handbook` | 品牌推广 | 旅行绿手册风，旅行/目的地手册气质 | `promotion/travel-green-handbook/design.md` |
| `electric-violet-business` | 工作汇报 | 电紫商务风，经营复盘与进展汇报 | `work/electric-violet-business/design.md` |
| `sky-blue-wayfinding` | 工作汇报 | 天蓝导视风，清晰导航式工作汇报 | `work/sky-blue-wayfinding/design.md` |
| `warm-clay-works` | 工作汇报 | 暖陶作品风，高完成度作品集式汇报 | `work/warm-clay-works/design.md` |
| `warm-jade-annual-report` | 工作汇报 | 暖玉年报风，经营年报与复盘 | `work/warm-jade-annual-report/design.md` |

URL 前缀：

```
https://raw.githubusercontent.com/acnlie/open-kimi-ppt-skill/main/skills/open-kimi-ppt/reference/design_system/<上游路径>
```

例如要取 `apricot-white-brief`：

```
https://raw.githubusercontent.com/acnlie/open-kimi-ppt-skill/main/skills/open-kimi-ppt/reference/design_system/consulting/apricot-white-brief/design.md
```

想批量内置，改 `scripts/vendor_themes.py` 的 `THEMES` 表后重跑即可。

## 配套文档（同一上游）

| 主题 | 上游路径 | 用途 |
| --- | --- | --- |
| 7 类场景版面基线 | `slides_categories.md` | 先定场景再选风格；7 类：分析决策/商业提案/管理汇报/学术研究/教育培训/技术工程/品牌创意 |
| 场景明细 | `slides_categories/analysis-decision.md 等 7 个文件` | 每类的读者任务与版面做法 |
| 字体规范 | `fonts.md` | 字体族与字号层级 |
| 形状库 | `shapes.md` | 可用形状清单 |
| PPTX 动画语法 | `pptd.md §6` | 上游自己的动画 DSL，可与本 skill 对照 |

这些也走同一个 URL 前缀。
