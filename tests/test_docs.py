#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""文档与规范的回归测试（16–24 组）。

为什么需要这份测试：这一组守的不是代码，是**文档与实现的一致性**。
本项目里最危险的失败模式是**静默丢失**—— 文件能开、能跑、报告全绿，
但根因编号对不上了、目录漏了、规范里的数字和实现里的常量漂了。

这一组因此**只做机械可判的事**：编号完整/唯一、映射表一致、目录存在、
阈值三处同值、生成器产物不漂、证据结构逐值可核对。它**不**判断
"哪一节该在哪份文件" —— 那是设计判断，会变。

> **2026-10-08 从 test_transition_table.py 拆出。** 拆分依据是**被测对象**：
> 16–22 组守的是文档与规范，与切换实测表无关。组号**保持原编号**。
> 23 组为 2026-10-09 新增（生成器产物判据，§十.7）。
> 24 组为 2026-10-10 新增（配方证据的结构复核件逐值核对）。

模块内容表（9 组；组号是稳定接口，与拆分前一致）
------------------------------------------------
  16  踩坑手册拆分：编号完整、唯一、映射表一致
  17  症状文档拆分：现象栏可路由
  18  配方文档拆分：名实相符 + §n 可解析
  19  长文档的目录（导航不该靠滚）
  20  规范与实现：同一组数字，三处一致
  21  更正守则（打补丁，不重写）
  22  根因收敛后的两族结构
  23  设计系统索引：生成器产物可复现
  24  配方证据（莲花 fan）：结构复核件逐值可核对
"""
import io
import os
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from helpers import HERE, ROOT, read_text, make_checker, report

TITLE = "文档与规范回归"


def main():
    fails, check = make_checker(TITLE)

    # ---- 16. 踩坑手册拆分后，§编号必须仍然可解析 -----------------------------
    # 为什么需要这一组：com-pitfalls.md 从 1 份 / 51 节拆成 10 份之后，
    # 全仓库 90+ 处引用还写着 `com-pitfalls.md §44`。**编号是接口**，
    # 拆分不能让它失效 —— 一旦某个编号在两份文件里同时出现，或者某节
    # 漏进回收站，那些引用就会静默指向错误的地方（比 404 更难发现）。
    #
    # 这一组不检查"哪一节该在哪份文件"（那是设计判断，会变），只钉三件
    # 机械可判的事：编号完整、编号唯一、映射表与真实文件一致。
    print()
    print("== 16. 踩坑手册拆分：编号完整性 ==")
    ref_dir = os.path.join(ROOT, "reference")
    pit_files = sorted(f for f in os.listdir(ref_dir)
                       if f.startswith("pitfall-") and f.endswith(".md"))
    check("踩坑已拆成多份（>=8）", len(pit_files) >= 8,
          "只找到 %d 份: %s" % (len(pit_files), pit_files))

    sec_re = re.compile(r"^## (\d+)\. (.+)$", re.M)
    owner = {}
    dups = []
    for f in pit_files:
        body = io.open(os.path.join(ref_dir, f), encoding="utf-8").read()
        for m in sec_re.finditer(body):
            n = int(m.group(1))
            if n in owner:
                dups.append("§%d 同时在 %s 和 %s" % (n, owner[n], f))
            owner[n] = f
    check("同一条坑不会出现在两份文件里", not dups, "; ".join(dups[:4]))

    # 拆分只能搬家，不能丢节。51 是拆分发生时的节数；日后只会增。
    check("拆分没有丢节（§1..§51 全部还在）",
          all(n in owner for n in range(1, 52)),
          "缺: %s" % [n for n in range(1, 52) if n not in owner])
    check("编号连续无空洞（§1..§max）",
          sorted(owner) == list(range(1, max(owner) + 1)) if owner else False,
          "现有: %s" % sorted(owner))

    # 原文件名必须留下，且必须变成目录 —— 90+ 处引用指着这个名字
    hub = os.path.join(ref_dir, "com-pitfalls.md")
    check("com-pitfalls.md 仍在（旧引用不能死）", os.path.exists(hub))
    if os.path.exists(hub):
        hub_body = io.open(hub, encoding="utf-8").read()
        hub_links = set(re.findall(r"\]\((pitfall-[a-z0-9-]+\.md)\)", hub_body))
        check("com-pitfalls.md 链到全部拆分文件",
              hub_links >= set(pit_files),
              "漏链: %s" % sorted(set(pit_files) - hub_links))
        check("com-pitfalls.md 已不再是正文（没有 §n 小节标题）",
              not sec_re.search(hub_body),
              "仍含 %d 个 `## n.` 标题" % len(sec_re.findall(hub_body)))

    # 映射表必须存在、覆盖全部编号，且与真实文件一致（不是手抄的）
    pmap = os.path.join(ref_dir, "pitfall-map.md")
    check("pitfall-map.md 存在（§n 的解释器）", os.path.exists(pmap))
    if os.path.exists(pmap):
        ptext = io.open(pmap, encoding="utf-8").read()
        missing = [n for n in sorted(owner) if ("| §%d |" % n) not in ptext]
        check("映射表覆盖所有编号", not missing,
              "漏: %s" % missing[:6])
        # 映射表说的文件必须和真实归属一致 —— 这条最值钱：它防止
        # "文件挪了、表忘了改"，那会让读者按表去另一份文件里找，找不到。
        wrong = []
        for n, f in sorted(owner.items()):
            m = re.search(r"\| §%d \|.*?\]\(([a-z0-9_.-]+)\)" % n, ptext)
            if m and m.group(1) != f:
                wrong.append("§%d 表说 %s，实际在 %s" % (n, m.group(1), f))
        check("映射表与真实归属一致", not wrong, "; ".join(wrong[:4]))

    # 生成器必须可用，否则映射表会慢慢腐烂成手抄
    gen = os.path.join(ROOT, "scripts", "build_pitfall_map.py")
    check("有生成映射表的脚本（不是手抄的）", os.path.exists(gen))

    # ---- 17. 症状文档拆分后，现象栏必须仍能找到 -----------------------------
    # 与第 16 组同一件事，但症状这一族用的不是编号而是**现象栏名**
    # （INDEX 和各处写着 `symptoms.md → 「文件损坏」`）。所以守的是：
    #   ① 每个现象栏有归属；② 目录页链到全部拆分文件；③ 原文件名还在。
    print()
    print("== 17. 症状文档拆分：现象栏可路由 ==")
    sym_files = sorted(f for f in os.listdir(ref_dir)
                       if f.startswith("symptom-") and f.endswith(".md"))
    check("症状已拆成多份（>=5）", len(sym_files) >= 5,
          "只找到 %d 份: %s" % (len(sym_files), sym_files))

    # ① 原来那 10 个现象栏，内容必须仍在某一族里（栏名本身可以并进新文件）
    KEY_PHENOMENA = [
        ("文件损坏", ["0x80070570"], "symptom-file-corruption.md"),
        ("静默丢弃", ["静默丢弃"], "symptom-silent-drop.md"),
        ("结果不对", ["渲染出来的不是你要的", "只能靠**真渲染**发现"],
         "symptom-wrong-result.md"),
        ("往返丢失", ["往返"], "symptom-roundtrip.md"),
        ("工具与环境", ["PowerShell", "编码"], "symptom-tooling.md"),
    ]
    for name, needles, fname in KEY_PHENOMENA:
        p = os.path.join(ref_dir, fname)
        if not os.path.exists(p):
            check("现象「%s」有归属文件 %s" % (name, fname), False, "文件不存在")
            continue
        body = io.open(p, encoding="utf-8").read()
        check("现象「%s」的内容落在 %s" % (name, fname),
              any(n in body for n in needles),
              "找不到任何标志串: %s" % needles)

    # ② 目录页必须链到全部拆分文件，且自己不再是正文
    shub = os.path.join(ref_dir, "symptoms.md")
    check("symptoms.md 仍在（旧引用不能死）", os.path.exists(shub))
    if os.path.exists(shub):
        sh = io.open(shub, encoding="utf-8").read()
        s_links = set(re.findall(r"\]\((symptom-[a-z0-9-]+\.md)\)", sh))
        check("symptoms.md 链到全部拆分文件",
              s_links >= set(sym_files),
              "漏链: %s" % sorted(set(sym_files) - s_links))
        # 目录页不该再含正文。判据用**落款的实际形态**：正文里每节结尾是
        # `<p align="right"><sub>来源 com-pitfalls §n</sub></p>`。
        # ⚠️ 不能只查子串 "来源 com-pitfalls" —— 目录页的说明文字里会写到
        # 这个落款（就是本段这句），那是描述，不是正文。
        footer = re.compile(r"<sub>\s*来源 com-pitfalls")
        check("symptoms.md 已不再是正文（不含各节来源落款）",
              not footer.search(sh),
              "仍含 %d 条落款" % len(footer.findall(sh)))

    # ③ 拆分文件必须保留 `来源 com-pitfalls §n` 落款 —— 那是回到根因的路
    no_src = []
    for f in sym_files:
        b = io.open(os.path.join(ref_dir, f), encoding="utf-8").read()
        # 允许个别文件（如"证据"栏）没有编号落款，但那也要有可回溯的文字
        if "来源 com-pitfalls" not in b and "见 [" not in b and "见 §" not in b:
            no_src.append(f)
    check("拆分文件保留了回到根因的线索（来源 §n 或链接）",
          not no_src, "缺线索: %s" % no_src)

    # ---- 18. 配方文档拆分：文件名必须对上内容 -------------------------------
    # 为什么需要这一组：`morph-and-3d-recipes.md` 曾 683 行，但真正的毛病
    # **不是篇幅，是名实不符** —— 名字只提 Morph 与 3D 相机，实际装着
    # 素材派生（§4/§6/§7）和窗口化填充（§8）。**找"形状切割"的人不会想到
    # 打开这个文件名。** 这是导航失败，光加目录治不了。
    #
    # 守三件事：① §5 之外全部搬走；② §1..§8 一条不丢、不重复；
    # ③ 目录页有一张能查 §n 的表（它就是这一族的"编号解释器"）。
    print()
    print("== 18. 配方文档拆分：名实相符 + §n 可解析 ==")
    hub_p = os.path.join(ref_dir, "morph-and-3d-recipes.md")
    check("配方目录页仍在（21 处旧引用不能死）", os.path.exists(hub_p))
    if os.path.exists(hub_p):
        hub = io.open(hub_p, encoding="utf-8").read()
        hub_lines = hub.count("\n")
        # 663 行搬走之后，目录页应该是"一屏级"的
        check("目录页已不再是正文（<=150 行）", hub_lines <= 150,
              "还有 %d 行" % hub_lines)

        recipe_files = sorted(f for f in os.listdir(ref_dir)
                              if f.startswith("recipe-")
                              and f.endswith(".md")
                              and f != "recipe-library.md")
        check("配方已拆成 3 份", len(recipe_files) == 3, str(recipe_files))
        check("目录页链到全部 3 份",
              set(recipe_files) <= set(re.findall(
                  r"\]\((recipe-[a-z0-9-]+\.md)\)", hub)),
              "漏链: %s" % sorted(set(recipe_files) - set(re.findall(
                  r"\]\((recipe-[a-z0-9-]+\.md)\)", hub))))

        # ① §1..§8 一条不丢、不重复（§5 留在目录页）
        h2 = re.compile(r"^## (\d+)\. ", re.M)
        seen = {}
        dup = []
        for n in [int(x) for x in h2.findall(hub)]:
            if n in seen:
                dup.append(n)
            seen[n] = "morph-and-3d-recipes.md"
        for f in recipe_files:
            body = io.open(os.path.join(ref_dir, f), encoding="utf-8").read()
            for n in [int(x) for x in h2.findall(body)]:
                if n in seen:
                    dup.append("§%d 同时在 %s 和 %s" % (n, seen[n], f))
                seen[n] = f
        check("§1..§8 一条不丢", all(n in seen for n in range(1, 9)),
              "缺: %s" % [n for n in range(1, 9) if n not in seen])
        check("§n 不会出现在两份文件里", not dup, "; ".join(dup[:4]))
        check("§5（横跨三份的整合视图）留在目录页", seen.get(5) ==
              "morph-and-3d-recipes.md", str(seen.get(5)))

        # ② 目录页必须能当 §n 的解释器：8 个编号每个都出现在表里
        missing = [n for n in range(1, 9) if ("§%d" % n) not in hub]
        check("目录页的表覆盖全部 §1..§8（能当编号解释器）",
              not missing, "表里没有: %s" % missing)

        # ③ 目录页要说明"为什么按想做的事排、不按编号排"
        check("目录页解释了导航主线是「想做的事」",
              "想做的事" in hub)

    # ---- 19. 长文档必须有目录（否则等于没有地图）----------------------------
    # 为什么需要这一组：拆分解决了"文件太大"，但**内聚的长文档不该拆** ——
    # 形态层/选择层/机制层是 L1–L4 架构的骨架，拆开就把架构打散了。
    # 它们要的是**加目录**：读者打开第一屏就该知道里面有什么。
    #
    # 门槛取 300 行 + 至少 4 个标题：低于这个规模加目录是噪音。
    # 判据接受两种形态：本工具生成的（有 `<!-- toc -->` 标记），
    # 或人写的（前 40 行有「表格行 + 站内锚点」—— `transition-model.md` 是这种）。
    print()
    print("== 19. 长文档的目录（导航不该靠滚）==")
    need = []
    have = 0
    for f in sorted(os.listdir(ref_dir)):
        if not f.endswith(".md"):
            continue
        body = io.open(os.path.join(ref_dir, f), encoding="utf-8").read()
        if body.count("\n") < 300:
            continue
        stripped = re.sub(r"```.*?```", "", body, flags=re.S)
        n_heads = len(re.findall(r"^#{2,3} (?!#)", stripped, re.M))
        if n_heads < 4:
            continue
        marked = "<!-- toc -->" in body
        handwritten = bool(re.search(
            r"^\|.*\]\(#", "\n".join(body.splitlines()[:40]), re.M))
        if marked or handwritten:
            have += 1
        else:
            need.append("%s（%d 行 / %d 标题）" % (f, body.count("\n"), n_heads))
    check(">=300 行的文档都有目录（%d 份）" % have, not need,
          "缺目录: %s" % need)

    # 生成器必须存在 —— 否则下次加新长文档时会手写，而手写的会漏
    toc_tool = os.path.join(ROOT, "tools", "add_toc.py")
    check("有加目录的脚本（不要手写目录）", os.path.exists(toc_tool))
    if os.path.exists(toc_tool):
        src = io.open(toc_tool, encoding="utf-8").read()
        # 必须幂等：已有标记就替换，而不是叠加
        check("加目录脚本是幂等的（有 toc 标记就替换）",
              "<!-- toc -->" in src and "--check" in src)

    # ---- 20. 规范与实现不能漂 ----------------------------------------------
    # 为什么需要这一组：`CONTRIBUTING.md` 是一份**声明式**规范，而
    # `scripts/check_structure.py` 是它的**可执行实现**。两者写了同一组数字
    # （10000 / 20000 / 800 / 2000 / 6000）—— 这正是"细节只写一次"要禁的事，
    # 但规范必须给人读、实现必须给机器读，只能各写一份。
    # **那就用测试把它们钉在一起。** 数字漂了，这里立刻报。
    print()
    print("== 20. 规范与实现：同一组数字，三处一致 ==")
    std_p = os.path.join(ROOT, "CONTRIBUTING.md")
    chk_p = os.path.join(ROOT, "scripts", "check_structure.py")
    check("变更有规范（CONTRIBUTING.md）", os.path.exists(std_p))
    check("规范有可执行实现（check_structure.py）", os.path.exists(chk_p))

    if os.path.exists(std_p) and os.path.exists(chk_p):
        std = io.open(std_p, encoding="utf-8").read()
        chk = io.open(chk_p, encoding="utf-8").read()
        # 规范里写的是 `20,000`（给人读加千分位），实现里写 `20_000`（Python 数字分隔符）。
        # 比对前把两种都归一化掉 —— 要比的是**数值**，不是排版。
        std_n = std.replace(",", "")

        # ① 规范必须声明决策树的三种治法 —— 缺一种，读者就会只想到"拆"
        for word, why in (("拆", "装了多件事"), ("加目录", "内聚长文档"),
                          ("归档", "会无限增长的内容")):
            check("规范声明了「%s」这条治法（%s）" % (word, why), word in std)
        check("规范声明了「拆分产物不再二次拆」", "二次拆" in std)

        # ② 实现必须**只读** —— 这是它能无人值守的唯一前提
        writes = re.findall(r"open\([^)]*,\s*[\"']w", chk)
        check("结构体检是只读的（能无人值守）", not writes,
              "发现写操作: %s" % writes[:3])
        check("结构体检是只读的（注释也这么写）",
              "只读" in chk or "不改文件" in chk)

        # ③ 阈值：规范里的数字必须与实现里的常量一致
        def _nums(text):
            return set(re.findall(r"\b(\d{1,3}(?:_\d{3})+|\d{4,6})\b",
                                  text.replace(",", "")))
        chk_consts = dict(re.findall(
            r"^(DOC_TOC_CHARS|DOC_SPLIT_CHARS|INDEX_BUDGET|"
            r"CODE_MAP_LINES|CODE_SPLIT_LINES)\s*=\s*([\d_]+)",
            chk, re.M))
        check("实现里五个阈值常量都在", len(chk_consts) == 5,
              str(sorted(chk_consts)))
        for name, val in sorted(chk_consts.items()):
            v = val.replace("_", "")
            check("规范写到了 %s=%s" % (name, v), v in std_n,
                  "CONTRIBUTING.md 里找不到这个数字")

        # ④ INDEX 预算是三处共用：规范 / 体检 / 文档校验器 —— 必须同值
        vd = io.open(os.path.join(ROOT, "scripts", "verify_docs.py"),
                     encoding="utf-8").read()
        vd_budget = re.search(r"INDEX_BUDGET\s*=\s*([\d_]+)", vd)
        chk_budget = chk_consts.get("INDEX_BUDGET", "").replace("_", "")
        check("INDEX 预算在体检与文档校验器里同值",
              vd_budget and vd_budget.group(1).replace("_", "") == chk_budget,
              "verify_docs=%s check_structure=%s"
              % (vd_budget and vd_budget.group(1), chk_budget))

        # ⑤ 豁免必须带理由，且必须带待办（否则下一个看的人只会照抄）
        #
        # ⚠️ 判据要认「**真的有条目**」，不能只看花括号里非空 ——
        # 这个块的注释会留档"已被撤销的豁免"（如 2026-09-30 撤销的
        # build_transition_table 那条），那是有价值的记录，不是条目。
        # **豁免全清空是健康状态**，不该被判失败。
        m = re.search(r"CODE_SPLIT_EXEMPT\s*=\s*\{(.*?)\n\}", chk, re.S)
        body = m.group(1) if m else ""
        entries = re.findall(r'^\s{4}"([^"]+)":\s*\{', body, re.M)
        if entries:
            for name in entries:
                blk = re.search(r'"%s":\s*\{(.*?)\n\s{4}\}' % re.escape(name),
                                body, re.S)
                seg = blk.group(1) if blk else ""
                check("例外「%s」写了理由" % name, '"why"' in seg)
                check("例外「%s」带待办（豁免不等于遗忘）" % name, '"todo"' in seg)
        else:
            check("代码例外清单为空（豁免都已清掉 —— 健康状态）", True,
                  "条目: %s" % entries)

    # ── 第 21 组：更正守则（2026-10-08 用户定的） ─────────────────────
    # 「发现之前写错了就打补丁」这条规则本身也会腐烂：如果它只躺在
    # CONTRIBUTING.md 里而没人检查，半年后就没人记得为什么 §8.3 和 §8.4
    # 会互相冲突。这里守住两件事：
    #   ① 守则本身在规范里，且 HANDOVER 也指得到（两处不许各说各话）
    #   ② 规范与交接的措辞不许出现「修正版 / 新开一份」这类反模式
    print("== 21. 更正守则（打补丁，不重写）==")
    std_p = os.path.join(ROOT, "CONTRIBUTING.md")
    hov_p = os.path.join(ROOT, "HANDOVER.md")
    std = io.open(std_p, encoding="utf-8").read() if os.path.exists(std_p) else ""
    hov = io.open(hov_p, encoding="utf-8").read() if os.path.exists(hov_p) else ""

    check("规范里有「打补丁」这条守则",
          "打补丁" in std and "就地" in std)
    check("守则说清了边界（改错→补丁 / 位置错→分诊）",
          "位置不对" in std or "位置不对" in std.replace("**", ""))
    check("守则标出了唯一例外（history/ 归档不补丁）",
          "history/" in std and "毁证" in std)
    check("HANDOVER 的验证文化里也写了这条（两处不许各说各话）",
          "打补丁" in hov and "就地" in hov)
    # 反模式：不能出现"新开一份修正版"这类指导 —— 那正是它要禁的
    for kw in ("修正版", "重写整节"):
        check("规范没有把「%s」列为做法" % kw,
              not re.search(r"(?<!不要%s)(修正版)" % kw, std)
              or "两份说法比一份错更糟" in std,
              "「%s」只应作为被禁做法出现" % kw)

    # 实测：本仓库真的按这条守则做过一次就地更正（§8.3 vs §8.4 的冲突），
    # 确认更正留下了「当时为什么判断错了」的痕迹，而不是悄悄改掉。
    if hov:
        check("§8.4 的更正留了原因（不是悄悄改掉的）",
              "更正" in hov and "把两者混为一谈" in hov,
              "就地更正必须留下当时判断错在哪")

    # ── 第 22 组：根因收敛后的两族结构（2026-10-08） ─────────────────────
    # symptom/pitfall 按根因合并后，最容易腐烂的是「同一段解释又写了两遍」。
    # 这里守住合并的三个结论：
    #   ① 每个 pitfall 顶层节都有「现象」导语 —— 否则"按现象查根因"在根因侧不可用
    #   ② 每个 symptom 编号节都能走到根因（链接或编号），不会变成孤岛
    #   ③ 编号与文件名一个都没动 —— 改了会让 90+ 处历史引用静默指向别处
    print("== 22. 根因收敛后的两族结构 ==")
    ref_p = os.path.join(ROOT, "reference")
    pit_files = sorted(f for f in os.listdir(ref_p)
                       if f.startswith("pitfall-") and f != "pitfall-map.md")
    sym_files = sorted(f for f in os.listdir(ref_p)
                       if f.startswith("symptom-"))
    check("两族文件都在（10 份 pitfall + 6 份 symptom）",
          len(pit_files) == 10 and len(sym_files) == 6,
          "pitfall=%d symptom=%d" % (len(pit_files), len(sym_files)))

    # ① 现象导语覆盖度
    top_re = re.compile(r"^## (\d+)\.\s+", re.M)
    pit_nums = set()
    pit_missing = []
    for name in pit_files:
        text = io.open(os.path.join(ref_p, name), encoding="utf-8").read()
        parts = top_re.split(text)
        # split 结果形如 ['', '1', '标题\n正文', '2', '标题\n正文', ...]
        for k in range(1, len(parts) - 1, 2):
            num = int(parts[k])
            body = parts[k + 1]
            pit_nums.add(num)
            if not re.search(r"^\*\*现象\*\*[:：]", body, re.M):
                pit_missing.append("§%d(%s)" % (num, name))
    check("51 个 pitfall 顶层节都有编号", len(pit_nums) == 51,
          "实际 %d 个" % len(pit_nums))
    check("每个 pitfall 顶层节都有「现象」导语",
          not pit_missing,
          "缺: %s" % ", ".join(pit_missing[:6]))

    # ② symptom 每节可达根因
    sec_re = re.compile(r"^#{2,4}\s*§\s*(\d+(?:\.\d+)*)", re.M)
    orphan = []
    total_sec = 0
    for name in sym_files:
        text = io.open(os.path.join(ref_p, name), encoding="utf-8").read()
        for raw in sec_re.findall(text):
            total_sec += 1
            top = int(raw.split(".")[0])
            if top not in pit_nums:
                orphan.append("§%s(%s)" % (raw, name))
    check("symptom 的每个编号在 pitfall 侧都有对应", not orphan,
          "孤立: %s" % ", ".join(orphan[:6]))
    check("symptom 编号节数量合理（48）", total_sec == 48,
          "实际 %d 节" % total_sec)

    # ③ 编号集合没变：1..51 连续且唯一
    missing_nums = sorted(set(range(1, 52)) - pit_nums)
    check("§1~§51 一个不缺（编号是接口，不能动）", not missing_nums,
          "缺: %s" % missing_nums)

    # 根因唯一性：索引节（正文已收敛）里不该再出现代码块。
    # 判据要分档 —— 档 3 保留了现象侧独有的正文，**本来就该有代码块**，
    # 拿"有没有代码块"去判全部节会误报（曾误报 §3/§35/§41）。
    # 区分办法：索引节的正文只有「根因与修法」一行，没有其他散文。
    for name in sym_files:
        text = io.open(os.path.join(ref_p, name), encoding="utf-8").read()
        for m in re.finditer(r"^#{2,4}\s*§(\d+)[^\n]*\n(.*?)(?=^#{2,4}\s|\Z)",
                             text, re.M | re.S):
            body = m.group(2)
            prose = [
                b for b in body.split("\n")
                if b.strip() and "根因与修法" not in b
                and not b.strip().startswith("**现象**")
                and not b.strip().startswith("<p align=")
            ]
            if not prose and "```" in body:
                check("§%s 是索引节，不该留代码块" % m.group(1), False)
                break
        else:
            continue
        break
    else:
        check("索引节（已收敛）里没有残留代码块", True)

    # 导航与真相源同步
    cap_p = os.path.join(ROOT, "facts", "capabilities.json")
    if os.path.exists(cap_p):
        cap = io.open(cap_p, encoding="utf-8").read()
        check("capabilities.json 声明了根因唯一",
              "根因正文只存在于此族" in cap and "structure_note" in cap)
    idx_p = os.path.join(ROOT, "INDEX.md")
    if os.path.exists(idx_p):
        idx = io.open(idx_p, encoding="utf-8").read()
        check("INDEX 说清了两族是一根轴的两个方向",
              "同一根轴" in idx or "一根轴" in idx)

    # ---- 23. 设计系统索引：生成器产物不漂 ------------------------------------
    # 为什么需要这一组：design-system/README.md 既是**随包文档**，又是
    # build_index.py 的**产物**。2026-10-09 调研实测到：README 里有一节
    # （design_compose）是后加的手工内容，生成器不知道它 —— 重跑一次
    # 就静默丢 35 行。§十.7「生成型工具的产物必须有判据」，这一组就是它。
    print()
    print("== 23. 设计系统索引：生成器产物可复现 ==")
    ds_dir = os.path.join(ROOT, "reference", "design-system")
    gen = os.path.join(ds_dir, "build_index.py")
    check("生成器存在（README 是产物，不是手写的）", os.path.exists(gen))
    if os.path.exists(gen):
        r = subprocess.run([sys.executable, gen, "--check"],
                           capture_output=True, text=True, encoding="utf-8")
        check("--check 通过（README 与生成器输出一致）", r.returncode == 0,
              ((r.stdout or "") + (r.stderr or "")).strip()[-160:])
    # 独立口径：不经过生成器，直接验「README 写的」和「磁盘上有的」是否同。
    # —— 与 --check 是两个角度：一个验"生成器与磁盘一致"，
    # 一个验"README 声称内置的 10 份文件真的在"。共享盲区会一起错（§十.4.2）。
    readme_p = os.path.join(ds_dir, "README.md")
    if os.path.exists(readme_p):
        txt = read_text(readme_p)
        rows = re.findall(r"^\| `([\w.-]+)` \| [^|]+ \| [^|]+ \| `([\w.-]+\.md)` \|$",
                          txt, re.M)
        missing = [f for _i, f in rows
                   if not os.path.exists(os.path.join(ds_dir, f))]
        check("「已内置」表里的 %d 份文件都在磁盘上" % len(rows),
              bool(rows) and not missing,
              "缺: " + ", ".join(missing[:4]))

    # ---- 24. 配方证据（莲花 fan）：结构复核件逐值可核对 ------------------------
    # 为什么需要这一组：recipe「move-the-mask-not-the-image」的证据说
    # "八片 pie、位置尺寸一致、唯一变量 rot 0→-105、第 2 页 morph"。
    # 原件（用户莲花 deck）含私人元数据与第三方版权素材、**不能入库**，
    # 入库的是**结构复核副本**（tools/make_lotus_fixture.py 制作）——
    # 那么"副本结构与原件一致"这件事必须有判据，否则副本只是传闻的替身。
    # 这一组把配方声明逐值钉在副本上（含藏在 mc:AlternateContent 里的 morph）。
    print()
    print("== 24. 配方证据（莲花 fan）：结构复核件逐值可核对 ==")
    fix_p = os.path.join(ROOT, "tests", "fixtures", "lotus-fan-structure.pptx")
    check("结构复核件存在", os.path.exists(fix_p))
    if os.path.exists(fix_p):
        A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
        P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
        z = zipfile.ZipFile(fix_p)

        def _pies(slide):
            root = ET.fromstring(z.read(slide))
            out = []
            for sp in root.iter(P + "sp"):
                pg = sp.find(".//" + A + "prstGeom")
                if pg is None or pg.get("prst") != "pie":
                    continue
                xf = sp.find(".//" + A + "xfrm")
                o, e = xf.find(A + "off"), xf.find(A + "ext")
                rot = xf.get("rot")
                out.append({"x": int(o.get("x")), "y": int(o.get("y")),
                            "cx": int(e.get("cx")), "cy": int(e.get("cy")),
                            "rot": int(rot) // 60000 if rot else 0,
                            "useBg": sp.get("useBgFill"),
                            "shdw": sp.find(".//" + A + "innerShdw") is not None})
            return out

        p1, p2 = _pies("ppt/slides/slide1.xml"), _pies("ppt/slides/slide2.xml")
        check("两页各 8 片 pie", len(p1) == 8 and len(p2) == 8,
              "slide1=%d slide2=%d" % (len(p1), len(p2)))
        check("8 片尺寸完全一致",
              len({(p["cx"], p["cy"]) for p in p1 + p2}) == 1)
        xs = [p["x"] for p in p1 + p2]
        check("位置一致（x 跨度 ≤ 100000 EMU，含 2 片手工微调）",
              max(xs) - min(xs) <= 100000, "x 跨度 %d" % (max(xs) - min(xs)))
        check("y 全部一致", len({p["y"] for p in p1 + p2}) == 1)
        check("第 1 页全 rot=0、第 2 页 rot = {0,-15,…,-105}",
              all(p["rot"] == 0 for p in p1) and
              sorted(p["rot"] for p in p2) ==
              [-105, -90, -75, -60, -45, -30, -15, 0])
        check("全部 useBgFill=1（窗口继承背景）",
              all(p["useBg"] == "1" for p in p1 + p2))
        check("全部带 innerShdw（瓣边可见）",
              all(p["shdw"] for p in p1 + p2))
        s1 = z.read("ppt/slides/slide1.xml").decode("utf-8")
        s2 = z.read("ppt/slides/slide2.xml").decode("utf-8")
        check("两页都有 p:bg（要继承的照片背景）",
              ET.fromstring(s1).find(P + "cSld/" + P + "bg") is not None and
              ET.fromstring(s2).find(P + "cSld/" + P + "bg") is not None)
        check("slide2 带 p159:morph（平滑；mc:AlternateContent + fade 兜底）",
              "p159:morph" in s2 and "mc:AlternateContent" in s2 and
              "<p:fade/>" in s2)
        z.close()

    return report(TITLE, fails)


if __name__ == "__main__":
    sys.exit(main())
