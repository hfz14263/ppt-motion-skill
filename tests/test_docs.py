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
> 27 组为 2026-10-10 追加（**Morph 配对必须查 shape id** —— 只查 name 会全绿而实际失效）。

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
  27  配方证据（太阳系）：Morph 配对逐值可核对（含 **id**）
"""
import io
import json
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
    # 2026-10-10：HANDOVER 的「验证文化」整节搬进了 history/verification-culture.md
    # （它已经超了 20k 硬上限）。这两条守则跟着搬走了，判据还只读 HANDOVER.md，
    # 于是变成"两处不许各说各话"自己先说了不算 —— 判据要跟着内容搬家。
    cult_p = os.path.join(ROOT, "history", "verification-culture.md")
    cult = io.open(cult_p, encoding="utf-8").read() if os.path.exists(cult_p) else ""
    hov_all = hov + "\n" + cult

    check("HANDOVER 的验证文化里也写了这条（两处不许各说各话）",
          "打补丁" in hov_all and "就地" in hov_all)
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
              "更正" in hov_all and "把两者混为一谈" in hov_all,
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

    # ---- 25. 决策树：生成器产物可复现 + 配方字段齐全 -------------------------
    # 为什么需要这一组：reference/decision-tree.md 是 build_decision_tree.py
    # 的**产物**，数据源是 recipes.json 的 applicable_when / breaks_how。
    # 它最容易腐烂的地方不是内容过时，而是「有人直接手改了 md」——
    # 下次重跑静默覆盖，读者以为自己改的东西生效了。
    # §十.7「生成型工具的产物必须有判据」，这一组就是它。
    print()
    print("== 25. 决策树：生成器产物可复现 + 配方字段齐全 ==")
    dt_p = os.path.join(ROOT, "reference", "decision-tree.md")
    dt_gen = os.path.join(ROOT, "reference", "design-system", "build_decision_tree.py")
    check("决策树存在", os.path.exists(dt_p))
    check("生成器存在（决策树是产物，不是手写的）", os.path.exists(dt_gen))
    if os.path.exists(dt_gen):
        r = subprocess.run([sys.executable, dt_gen, "--check"],
                           capture_output=True, text=True, encoding="utf-8")
        check("--check 通过（决策树与 recipes.json 一致）", r.returncode == 0,
              ((r.stdout or "") + (r.stderr or "")).strip()[-160:])
    # 独立口径：不看生成器，直接验「配方表自己声明齐了没有」。
    # —— 与 --check 两个角度：一个验产物一致，一个验数据源完整。
    rj_p = os.path.join(ROOT, "reference", "recipes.json")
    if os.path.exists(rj_p):
        entries = json.loads(read_text(rj_p)).get("entries", [])
        need = ("applicable_when", "breaks_how", "constraints", "evidence")
        incomplete = [e.get("id", "?") for e in entries
                      for f in need if not e.get(f)]
        check("%d 条配方的 applicable_when / breaks_how 齐全" % len(entries),
              bool(entries) and not incomplete,
              "缺字段: " + ", ".join(sorted(set(incomplete))[:4]))
        # 每条配方至少要有一个失败信号 —— 否则它无法被验收，不该进决策树
        nobreak = [e["id"] for e in entries if len(e.get("breaks_how") or []) < 1]
        check("没有「零失败信号」的配方", not nobreak,
              ", ".join(nobreak[:4]))
    if os.path.exists(dt_p):
        dt = read_text(dt_p)
        for label in ("gen:preflight", "gen:tree", "gen:table", "gen:breaks"):
            s, e = f"<!-- {label}:start -->", f"<!-- {label}:end -->"
            seg = dt[dt.index(s) + len(s):dt.index(e)].strip() \
                if s in dt and e in dt else ""
            check(f"[{label}] 区间已生成（非空非占位）",
                  bool(seg) and seg != "> 待生成",
                  "区间为空或仍是占位符")
        check("决策树说明了自己的腐烂方式并指向 recipes.json",
              "recipes.json" in dt and "生成" in dt)

    # ---- 26. SVG 载荷：机制判据（能力边界写进代码，不只写进文档） -------------
    # 为什么需要这一组：reference/svg-in-ppt.md 声称「SVG 不能自己动、
    # Morph 会让它发虚、注入走 COM 不走 XML」。这些是**会过期的事实**——
    # 哪天 PowerPoint 支持了 SVG 动画，文档就该改。
    # 把「文档说的」和「代码做的」钉在一起，文档才不会变成传说。
    print()
    print("== 26. SVG 载荷：机制与文档声明一致 ==")
    svg_doc = os.path.join(ROOT, "reference", "svg-in-ppt.md")
    check("SVG 边界文档存在", os.path.exists(svg_doc))
    ps1 = read_text(os.path.join(ROOT, "scripts", "motion.ps1"))
    check("COM 层有 Add-SvgLayer", "function Add-SvgLayer" in ps1)
    check("COM 层真的用 AddPicture 注入（不是手写 XML）",
          "AddPicture" in ps1)
    check("svg 与 media 分开注入（分类错误的教训）",
          "$bucket" in ps1 and "svgPlan" in ps1)
    check("路径已规范化（AddPicture 不接受正斜杠）",
          ps1.count("GetFullPath") >= 2)
    mxml = read_text(os.path.join(ROOT, "scripts", "motion_xml.py"))
    check("[Content_Types].xml 已声明 svg", '"svg": "image/svg+xml"' in mxml)
    mpy = read_text(os.path.join(ROOT, "scripts", "motion.py"))
    check("XML 层解析 svg 段", "svg_plan" in mpy)
    if os.path.exists(svg_doc):
        d = read_text(svg_doc)
        check("文档写明「SVG 不能自己动」（微软官方口径）",
              "do not support SVG Animation" in d)
        check("文档写明 Morph 会发虚（三个坑之一）",
              "Morph" in d and "糊" in d)
        # 三件事各落各处：morph 在 <p:transition>、svg 在 asvg:svgBlip、
        # zoom 在独立 part。三者混谈是这个 topic 最常见的误解，所以钉住。
        # （判据用文档里真出现的字符串 —— 上一版查 96DAC541 时整组 FAIL，
        #   而文档写的是全称 uri 不是短 id。那次是判据错了，不是文档缺。）
        check("文档把三件事分开（morph / svg / zoom 各落各处）",
              "asvg:svgBlip" in d and "zoomM" in d
              and "<p:transition>" in d)

    # ---- 27. 配方证据（太阳系）：结构复核件逐值可核对 --------------------------
    # 为什么需要这一组：recipe「motion-lives-between-pages」原先只有 test_morph
    # 的机制断言，**没有一份"多页连贯叙事"的端到端样例**。这条配方说
    # "两页同名同 id、只改数值，Morph 自动补间" —— 听起来显然，实际最容易
    # 坏在【跨页配对失败】上：id 被重新排、名字对不上、某一页漏改。
    # 这一组把配方声明逐值钉在副本上（含藏在 mc:AlternateContent 里的 morph、
    # 以及总览页↔详情页的同名形状对）。
    print()
    print("== 27. 配方证据（太阳系）：结构复核件逐值可核对 ==")
    solar = os.path.join(ROOT, "tests", "fixtures", "solar-system-structure.pptx")
    spec_p = os.path.join(ROOT, "examples", "solar-motion.yaml")
    check("结构复核件存在", os.path.exists(solar))
    check("可跑 spec 存在（三件套的第一件）", os.path.exists(spec_p))
    if os.path.exists(solar):
        z = zipfile.ZipFile(solar)
        slides = sorted((n for n in z.namelist()
                         if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                        key=lambda n: int(re.search(r"\d+", n.split("/")[-1]).group()))
        check("9 页（1 总览 + 8 行星）", len(slides) == 9, "实得 %d" % len(slides))

        def _shapes(xml_bytes):
            """{name: (x, y, cx, cy, shape_id)} —— id 必须一起读出来。"""
            root = ET.fromstring(xml_bytes)
            out = {}
            for sp in root.iter():
                if not (sp.tag.endswith("}sp") or sp.tag.endswith("}pic")
                        or sp.tag.endswith("}cxnSp")):
                    continue
                nv = sp.find(".//" + P + "cNvPr")
                xf = sp.find(".//" + A + "xfrm")
                if nv is None or xf is None:
                    continue
                o, e = xf.find(A + "off"), xf.find(A + "ext")
                if o is None or e is None:
                    continue
                out[nv.get("name")] = (int(o.get("x")), int(o.get("y")),
                                       int(e.get("cx")), int(e.get("cy")),
                                       nv.get("id"))
            return out

        p1 = _shapes(z.read("ppt/slides/slide1.xml"))
        # 配方 1 的核心断言：总览页有 8 行星 + 8 引线 + 16 个标签 + 太阳
        keys = [str(j) for j in range(1, 9)]      # 形状名是序号式：PLANET_1..8
        check("总览页有 8 颗行星（PLANET_1..8，两页同名同 id）",
              all("PLANET_%s" % k in p1 for k in keys),
              "缺: " + ", ".join(k for k in keys if "PLANET_%s" % k not in p1))
        check("总览页有 8 根引线（LEAD_1..8，视频里的细线）",
              all("LEAD_%s" % k in p1 for k in keys))
        check("总览页有中英双标签（LABEL_1..8 + CN_1..8）",
              all(("LABEL_%s" % k in p1 and "CN_%s" % k in p1) for k in keys))
        check("总览页有太阳（SUN）", "SUN" in p1)

        # ── Morph 配对的三条硬判据 ──────────────────────────────────
        # ① 跨页同名形状全部存在
        # ② **同名形状的 shape id 必须逐值相同** ← 这一条是 2026-10-10 补的，
        #    因为漏了它整套判据全绿而放映看不到任何补间。
        #    实测：Morph 按 **id** 配对，name 相同但 id 错位就配不上。
        # ③ 同页内 id 必须唯一（两个形状共用 id 会让 Morph 把它们互相配对）
        pair_ok, pair_missing = [], []
        id_mismatch = []
        for i in range(2, 10):
            pg = _shapes(z.read("ppt/slides/slide%d.xml" % i))
            for nm in pg:
                if nm.startswith(("PLANET_", "LEAD_", "LABEL_", "CN_")) \
                        or nm in ("HDR_TITLE", "HDR_CN", "TAGLINE",
                                  "BODY", "INDEX", "FOOT", "SUN"):
                    if nm in p1:
                        pair_ok.append((i, nm))
                        if p1[nm][4] != pg[nm][4]:
                            id_mismatch.append("s%d:%s id=%s≠s1:%s"
                                               % (i, nm, pg[nm][4], p1[nm][4]))
                    else:
                        pair_missing.append("s%d:%s" % (i, nm))
        check("跨页同名形状全部成对（8 详情页 × 39 形状 = 312 对）",
              len(pair_ok) == 312 and not pair_missing,
              "共 %d 对，缺 %s" % (len(pair_ok), ", ".join(pair_missing[:4])))
        check("**同名形状的 shape id 逐值相同**（Morph 按 id 配对，不是 name）",
              not id_mismatch,
              "共 %d 处错位: %s" % (len(id_mismatch), "; ".join(id_mismatch[:3])))

        # ③ 页内 id 唯一 —— 共用会让 Morph 把两个形状互相配对（实测踩过）
        dup_pages = []
        for i in range(1, 10):
            ids = [v[4] for v in _shapes(z.read("ppt/slides/slide%d.xml" % i)).values()]
            seen = set()
            for cid in ids:
                if cid in seen:
                    dup_pages.append("P%d:id=%s" % (i, cid))
                seen.add(cid)
        check("同页内 shape id 唯一（共用会让 Morph 互相对配）",
              not dup_pages, "; ".join(dup_pages[:4]))

        # 唯一的变量：位置/尺寸必须**不同**（相同 = 没东西可补间 = 硬切）
        # 这是配方 1 最容易被"看起来对"掩盖的失败模式。
        diff_pos = []
        for i, k in enumerate(keys, start=2):
            pg = _shapes(z.read("ppt/slides/slide%d.xml" % i))
            if "PLANET_%s" % k in pg and "PLANET_%s" % k in p1:
                # 只比几何（前 4 项），id 在上面单独判
                diff_pos.append(pg["PLANET_%s" % k][:4] != p1["PLANET_%s" % k][:4])
        check("行星在总览页与详情页位置/尺寸**都不同**（有差异才有补间）",
              len(diff_pos) == 8 and all(diff_pos),
              "%d/8 有差异" % sum(diff_pos))

        # ── 「挪出画布」判据（2026-10-12 新增）────────────────────────────
        # 详情页里，非当前行星的那 7 组必须**完全在画布外**。
        # 这是本样例的退场设计：滑出画面，而不是在画布内缩小、堆到角落。
        # 前一种做法（缩到 2pt 停在左下角）实测的观感是"一团东西飞向角落
        # 缩成小点" —— 那不是退场，是乱。判据钉的就是这个差别。
        CW, CH = 12192000, 6858000      # 画布尺寸（EMU，16:9）

        def _outside(geom):
            x, y, cx, cy = geom[:4]
            return (x + cx <= 0) or (x >= CW) or (y + cy <= 0) or (y >= CH)

        outside_ok, outside_bad, focus_bad = 0, [], []
        for j in range(2, 10):          # 详情页 P2..P9
            pg = _shapes(z.read("ppt/slides/slide%d.xml" % j))
            focus = j - 1               # 页 j 聚焦的是第 j-1 颗行星
            for i in range(1, 9):
                members = ("PLANET_%d" % i, "LEAD_%d" % i,
                           "LABEL_%d" % i, "CN_%d" % i)
                if i == focus:
                    if "PLANET_%d" % i in pg and _outside(pg["PLANET_%d" % i]):
                        focus_bad.append("s%d:PLANET_%d" % (j, i))
                    continue
                for nm in members:
                    if nm not in pg:
                        continue
                    if _outside(pg[nm]):
                        outside_ok += 1
                    else:
                        outside_bad.append("s%d:%s" % (j, nm))
        check("详情页里非当前行星组**全部在画布外**（挪出画布，不是缩小堆角落）",
              not outside_bad and outside_ok == 7 * 4 * 8,
              "在外 %d/224；留在画布内: %s" % (outside_ok, ", ".join(outside_bad[:5])))
        check("每页聚焦的行星在画布内（不然「聚焦」无从谈起）",
              not focus_bad, ", ".join(focus_bad))

        # ── 跨页差异判据：可见形状在相邻页之间必须有差异 ──────────────────
        # morph 页面**不带页内 timing**（实测：带 entrance 的形状会被踢出
        # 补间，先淡出再重新入场）。所以过渡完全靠两页的差异 ——
        # 某形状若在两页完全一样（几何 + 文字），它在 morph 中一动不动。
        # 只有页脚（每页本来就一样的落款）例外。
        FOOTER_TXT = ("太阳系漫游指南",)

        def _full(xml_bytes):
            """{name: (x, y, cx, cy, id, text)} —— 差异判据要连文字一起看。"""
            root = ET.fromstring(xml_bytes)
            out = {}
            for sp in root.iter():
                if not (sp.tag.endswith("}sp") or sp.tag.endswith("}pic")
                        or sp.tag.endswith("}cxnSp")):
                    continue
                nv = sp.find(".//" + P + "cNvPr")
                xf = sp.find(".//" + A + "xfrm")
                if nv is None or xf is None:
                    continue
                o, e = xf.find(A + "off"), xf.find(A + "ext")
                if o is None or e is None:
                    continue
                txt = "".join(t.text or "" for t in sp.iter(A + "t"))
                out[nv.get("name")] = (int(o.get("x")), int(o.get("y")),
                                       int(e.get("cx")), int(e.get("cy")),
                                       nv.get("id"), txt.strip())
            return out

        full = {i: _full(z.read("ppt/slides/slide%d.xml" % i))
                for i in range(1, 10)}
        static_bad, unpaired_bad = [], []
        for j in range(2, 10):
            cur, prv = full[j], full[j - 1]
            for nm, g in cur.items():
                if any(k in (g[5] or "") for k in FOOTER_TXT):
                    continue                     # 页脚：每页一样的落款，故意静止
                pg2 = prv.get(nm)
                if pg2 is None:
                    unpaired_bad.append("s%d:%s" % (j, nm))
                    continue
                if _outside(g) and _outside(pg2):
                    continue                     # 两页都在画布外：看不见，静止无妨
                same = (g[:4] == pg2[:4]) and ((g[5] or "") == (pg2[5] or ""))
                if same:
                    static_bad.append("s%d:%s" % (j, nm))
        check("相邻页里可见形状**都有差异**（morph 的动画在差异里；无差异=静止）",
              not static_bad, ", ".join(static_bad[:6]))
        check("相邻页里没有无配对形状（否则 morph 只能让它淡入淡出）",
              not unpaired_bad, ", ".join(unpaired_bad[:6]))

        # 每页都要有 morph；timing 只该在第一页
        morphs = 0
        timed = 0
        for s in slides:
            raw = z.read(s).decode("utf-8", "replace")
            if "p159:morph" in raw and "mc:AlternateContent" in raw:
                morphs += 1
            if "<p:timing>" in raw:
                timed += 1
        check("9 页全部带 p159:morph（平滑 3954，藏在 mc:AlternateContent 里）",
              morphs == 9, "%d/9" % morphs)
        # 2026-10-12 实测：morph 页面上的页内 entrance 动画会把形状踢出补间
        # （先淡出、morph 完再重新入场）。所以 timing 只保留在 P1（它是第一页，
        # 没有"上一页"，入场表演只能靠 timing），P2..P9 必须为空。
        check("只有总览页带 <p:timing>（详情页带页内动画会踢出 morph 补间）",
              timed == 1 and "<p:timing>" in z.read(slides[0]).decode("utf-8"),
              "共 %d 页带 timing" % timed)

        s1 = z.read("ppt/slides/slide1.xml").decode("utf-8")
        check("总览页效果数 ≥ 30（39 个：6 文字/太阳/8 行星/8 引线/16 标签/页脚）",
              s1.count("<p:par>") >= 30, "实得 %d" % s1.count("<p:par>"))

        # 正文段落数正确（_set_text 幂等性回归）——模板与 stage3 各写一次，
        # 不幂等就会越写越多（实测：详情页正文曾重复两遍）。
        body_bad = []
        for i in range(2, 10):
            raw = z.read("ppt/slides/slide%d.xml" % i).decode("utf-8", "replace")
            m = re.search(r'name="BODY".*?</p:sp>', raw, re.S)
            n = m.group(0).count("<a:p>") if m else -1
            if n != 3:
                body_bad.append("s%d:%d段" % (i, n))
        check("每页正文恰好 3 段（_set_text 幂等：不重复、不丢失）",
              not body_bad, ", ".join(body_bad))

        # 私有元数据必须为空（副本的意义之一就是能公开入库）
        core = z.read("docProps/core.xml").decode("utf-8")
        check("docProps 无作者残留（可公开入库）",
              not re.search(r"<dc:creator>\s*\S", core)
              and not re.search(r"<cp:lastModifiedBy>\s*\S", core))

        # 版面契约：背景必须是 p:bg（这样它不参与动画、不进 Morph 候选）
        root1 = ET.fromstring(z.read("ppt/slides/slide1.xml"))
        check("总览页用 p:bg 铺星空（背景不参与动画与 Morph 配对）",
              root1.find(P + "cSld/" + P + "bg") is not None)
        z.close()

    if os.path.exists(spec_p):
        sp = read_text(spec_p)
        # spec 里的 target 一律是 name 而非数字 id —— authoring-rules §A1
        targets = re.findall(r"\{target:\s*([^,}]+)", sp)
        numeric = [t for t in targets if t.strip().isdigit()]
        check("spec 的 target 全是 shape name，无手写 id（§A1）",
              bool(targets) and not numeric,
              "有 %d 个数字 target" % len(numeric))
        check("spec 给 9 页都写了 transition",
              sp.count("transition:") == 9, "实得 %d" % sp.count("transition:"))
        check("spec 覆盖 8 行星（PLANET_1..8 每颗都被点名）",
              all(("PLANET_%s" % k) in sp for k in keys))

    return report(TITLE, fails)


if __name__ == "__main__":
    sys.exit(main())
