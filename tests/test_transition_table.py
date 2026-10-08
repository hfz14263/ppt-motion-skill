#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""切换实测表的回归测试（1–15 组）。

为什么需要这份测试：`scripts/transition_reference.json` 是量出来的，而
`scripts/motion.py` 直接读它。两者一旦漂移，症状不是报错，而是**写出一个
PowerPoint 会悄悄改掉的块** —— 看起来一切正常，文件就是不对。所以这里把
"表还成立"和"motion.py 还认这张表"钉住。

不测的东西：本测试**不**打开 PowerPoint。真机验证在 `tests/smoke.py`。

> **2026-10-08 从test_transition_table.py 拆出。** 拆分依据是**被测对象**，
> 不是行数：1–15 组守的是切换实测表与注入引擎的一致性，
> 与文档结构无关。同批拆出的还有 `test_docs.py`（16–22 组）。
> 组号是**稳定接口**，跨文件后仍然保持原编号 —— 别重排。

模块内容表（15 组；组号是稳定接口，别重排）
------------------------------------------------
   1  表本身（48 项齐全、字段完整）
   2  motion.py 认这张表（表和引擎不能漂）
   3  每个切换算"一个"，不是两个
   4  重新生成的表不会重复累积
   5  Choice 说毫秒、Fallback 说 spd，两者讲同一个故事
  5b  附加属性必须和占位符一起合流
   6  机制层文档 + facts 与引擎一致
   7  形态探测 deck 生成器
   8  形态分析器：方向必须真的被量出来
  8b  旋转指标（第四类）—— 前三类都看不见它
   9  形态层文档 + facts 一致
  10  证据 deck 的断言必须与形态文档同源
  11  选择层：判断可以反驳，引用不能漂移
  12  方向探测（补 3b §三 的缺口）
  13  属性取值全集（attrdeck / attrdiff）
  14  切换 × 页内动画：结构平行、时间串行
  15  48 项之间的取舍：引用可回溯 + 无自造效果名
"""
import io
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from helpers import HERE, ROOT, REF, LEGACY_SPECS, load_json, make_checker, report

TITLE = "切换实测表回归"


def main():
    fails, check = make_checker(TITLE)

    with io.open(REF, encoding="utf-8-sig") as fh:
        ref = json.load(fh)

    # ---- 1. 表本身 ------------------------------------------------------
    check("表能读出来", isinstance(ref.get("gallery"), list))
    check("48 个界面切换项", len(ref["gallery"]) == 48,
          "got %d" % len(ref.get("gallery", [])))
    check("14 个 motion spec", len(ref["by_spec"]) == 14,
          "got %d" % len(ref.get("by_spec", [])))

    groups = {}
    for g in ref["gallery"]:
        groups.setdefault(g["group"], 0)
        groups[g["group"]] += 1
    check("界面分组是 12 细微 / 29 华丽 / 7 动态",
          groups.get("细微") == 12 and groups.get("华丽") == 29
          and groups.get("动态内容") == 7, str(groups))

    # 每个 spec 的 xml 必须真的含它自己的子元素 -- 表是手抄的话这里会炸
    for spec, e in sorted(ref["by_spec"].items()):
        check("by_spec[%s] 的 xml 含它的子元素" % spec, e["child"] in e["xml"],
              "%s not in %s" % (e["child"], e["xml"][:60]))
    for g in ref["gallery"]:
        check("gallery[%s] 的 xml 含它的子元素" % g["spec"], g["child"] in g["xml"])

    # 实测枚举：必须全部量到，且互不相同（旧的 0x0A01 那套会让 fade/strips 撞车）
    vals = [g["entryEffect"] for g in ref["gallery"]]
    check("48 项的枚举全部量到", all(v is not None for v in vals),
          str([g["spec"] for g in ref["gallery"] if g["entryEffect"] is None]))
    check("48 个枚举互不相同", len(set(vals)) == 48,
          str(sorted(v for v in set(vals) if vals.count(v) > 1)))

    # ---- 2. motion.py 认这张表 ------------------------------------------
    import motion as M

    missing = [s for s in LEGACY_SPECS if s not in M.TRANSITIONS]
    check("旧 spec 名一个都没丢", not missing, str(missing))

    # 写出来的块必须 well-formed，且带齐它自己用到的命名空间
    try:
        from lxml import etree as ET
    except ImportError:
        ET = None
    # The block is a FRAGMENT, not a document: the `p:` prefix comes from
    # <p:sld>, which every real slide declares. So parse it inside a wrapper
    # that binds p and mc -- parsing it bare fails on an unbound prefix and
    # says nothing about the block.
    WRAP = ('<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
            '>%s</p:sld>')
    if ET is not None:
        for spec in sorted(M.TRANSITIONS):
            if spec == "none":
                continue
            blk = M.build_transition(spec, 0.8)
            try:
                ET.fromstring((WRAP % blk).encode("utf-8"))
                ok = True
                detail = ""
            except Exception as exc:
                ok, detail = False, str(exc)
            check("%s 的块是合法 XML" % spec, ok, detail)
            if not ok:
                continue
            # 自包含：块内部能解析出所有用到的前缀
            used = set(re.findall(r"(?:<|xmlns:)(\w+):", blk))
            declared = set(re.findall(r'xmlns:(\w+)=', blk))
            check("%s 的块自带全部命名空间" % spec, used <= declared | {"p"},
                  "used=%s declared=%s" % (sorted(used), sorted(declared)))

    # 块里不该剩下未替换的占位符
    for spec in sorted(M.TRANSITIONS):
        if spec == "none":
            continue
        blk = M.build_transition(spec, 1.2)
        check("%s 的块没有残留占位符" % spec, "{" not in blk, blk[:80])
        check("%s 的块带上了时长" % spec, 'p14:dur="1200"' in blk, blk[:80])

    # ---- 3. 每个切换都算"一个"，不是两个 --------------------------------
    # Choice + Fallback 会写出两个 <p:transition>，raw 计数会翻倍。
    slide = ('<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
             'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
             '<p:cSld/>%s</p:sld>')
    for spec in ("fade", "push", "wipe", "strips"):
        x = slide % M.build_transition(spec, 0.8)
        raw = len(re.findall(r"<p:transition\b", x))
        eff = len(M.active_transition_blocks(x))
        check("%s: 2 个 raw 元素 = 1 个生效切换" % spec, raw == 2 and eff == 1,
              "raw=%d eff=%d" % (raw, eff))

    # ---- 4. 重新生成的表不会重复累积 ------------------------------------
    # 幂等：同一份 deck 连续处理两次，不能出现两个切换
    x = slide % ""
    once = M.replace_or_insert(x, "transition", M.build_transition("fade", 0.8))
    twice = M.drop_transition_alternate_content(once)
    twice = M.remove_elements(twice, "transition")
    twice = M.replace_or_insert(twice, "transition", M.build_transition("fade", 0.8))
    check("重复注入后仍只有一个生效切换",
          len(M.active_transition_blocks(twice)) == 1,
          str(len(M.active_transition_blocks(twice))))

    # ---- 5. Choice 说毫秒，Fallback 说 spd，两者必须讲同一个故事 --------
    # 实测（build 17928）：p14:dur 压过 spd；而 mc:Fallback 里**没有** p14:dur
    # （2010 属性进不了那个世界），所以降级世界里只有 spd 说话。写死 spd="med"
    # 不但让那里全塌回 ~0.5s，还会被 PowerPoint 当场改写。
    def spd_of(block):
        m = re.search(r'spd="(\w+)"', block)
        return m.group(1) if m else None

    # 取 PowerPoint 自己从 Duration 反算出来的映射（slow>=1000，其余 med/fast）
    # 就近取档：快≈500 / 中≈767 / 慢≈1000（实测）。三个边界因此落在
    #   (500+767)/2 = 633.5  和  (767+1000)/2 = 883.5
    # 也就是：<=633 取 fast，634–883 取 med，>=884 取 slow。
    expect = [("push", 0.3, "fast"), ("push", 0.5, "fast"), ("push", 0.6, "fast"),
              ("push", 0.7, "med"), ("push", 0.8, "med"),
              ("push", 0.9, "slow"), ("push", 2.5, "slow"),
              ("wipe", 0.63, "fast"), ("wipe", 0.64, "med")]
    for spec, secs, want in expect:
        blk = M.build_transition(spec, secs)
        got = spd_of(blk)
        check("%s %.1fs -> spd=%s" % (spec, secs, want), got == want,
              "got %r" % got)

    # Fallback 分支不许带 p14:dur：它需要 xmlns:p14，而 Fallback 的世界正是
    # 「没有 2010 扩展」的那个世界。
    blk = M.build_transition("push", 1.2)
    fb = blk[blk.index("<mc:Fallback"):]
    check("Fallback 不带 p14:dur", "p14:dur" not in fb)
    check("Fallback 的 spd 与 Choice 一致", spd_of(fb) == spd_of(blk))
    check("显式 speed= 会被尊重",
          spd_of(M.build_transition("wipe", 3.0, speed="fast")) == "fast")
    try:
        M.build_transition("wipe", speed="zzz")
        check("非法 speed 被拒绝", False, "no exception")
    except ValueError:
        check("非法 speed 被拒绝", True)

    # ---- 5b. 附加属性必须和占位符一起合流 -------------------------------
    # 这条踩过：先 replace 再 format，会把 {spd} 插进一个位置上看不见的地方，
    # 于是最后一个占位符没被替换 —— 只有在**又给 advTm** 时才发作，
    # 写出的文件里躺着 "spd=""med"" ... {spd}"，而 PowerPoint 会打开它。
    for spec in ("fade", "push", "wipe", "morph"):
        x = M.build_transition(spec, 1.5, advance_after=3.0, on_click=False)
        check("%s: 带 advTm/advClick 后无占位符残留" % spec,
              "{" not in x and "}" not in x)
        # 两个分支都要拿到 advTm（否则自动换页只在其中一套环境生效）
        heads = re.findall(r"<p:transition[^>]*>", x)
        check("%s: 两个 <p:transition> 都带 advTm" % spec,
              len(heads) == 2 and all('advTm="3000"' in h for h in heads),
              str(heads))

    # ---- 6. 机制层文档 + facts 与引擎一致 --------------------------------
    # 这四条是「渲染出来的帧」给的结论（reference/transition-model.md）。
    # 它们不靠读代码复核，只能靠真渲染；能在这里钉住的只有"文档还在、
    # facts 还声明了它、引擎还按结论办事"这三点。
    model = os.path.join(ROOT, "reference", "transition-model.md")
    check("机制层文档存在", os.path.exists(model))
    if os.path.exists(model):
        text = io.open(model, encoding="utf-8").read()
        for token, why in (("终点页", "模型①：写在哪页管的是进入哪页"),
                           ("一个槽位", "模型②：两次子元素会让整份文件拒开"),
                           ("p14:dur", "模型③：毫秒压过 spd"),
                           ("1000", "模型③：spd 三档的实测毫秒")):
            check("机制层文档仍写着「%s」" % why, token in text)

    fx = load_json(os.path.join(ROOT, "facts", "transitions.json"))
    ids = {r["id"] for r in fx["rules"]}
    for want in ("transition-anchors-on-the-target-page", "one-transition-one-slot",
                 "duration-truth-is-milliseconds",
                 "transition-cannot-address-a-shape", "frames-not-total-length"):
        check("facts 仍记着 %s" % want, want in ids)
    declared = " ".join(s.get("location", "") for s in fx["sources"])
    check("facts 已声明机制层为唯一真相源",
          "reference/transition-model.md" in declared)

    # 引擎必须按模型② 办事：切换与动画各占一个位置，互不挤掉对方。
    both = M.build_transition("push", 0.8)
    check("切换块写在自己的位置（不含 timing）", "p:timing" not in both)
    # 模型① 的可测后果：块是"写在哪页"的，所以引擎不该把它写到别页去。
    # 这里只能核对块本身不含页号 —— 页号由 spec 决定，见 normalize_spec。
    check("切换块不自带页号（页号来自 spec）",
          not re.search(r"slide\d", both))

    # ---- 7. 形态探测 deck 生成器 -----------------------------------------
    # 3a 的探测 deck 必须真的能被分析器读出东西来。这里不渲染视频，只保证
    # 生成器：① 每份 deck 都写上了切换（掉一个就会静默少量一个效果）；
    # ② 两页在空间上真的不同（纯色页会让"方向"无从测起，正是 §七 的坑）；
    # ③ manifest 里的字段分析器真的会用。
    import tempfile
    import build_transition_table as B
    tmp = tempfile.mkdtemp(prefix="shapedeck_")
    try:
        rc = B.cmd_shapedeck(tmp)
        check("shapedeck 生成成功", rc == 0)
        idx = load_json(os.path.join(tmp, "_index.json"))
        check("每份 Hypothesis 都有 deck", idx["count"] == len(B.HYPOTHESES),
              "got %d vs %d" % (idx["count"], len(B.HYPOTHESES)))
        specs = {d["spec"] for d in idx["decks"]}
        check("deck 覆盖全部 spec",
              {h["spec"] for h in B.HYPOTHESES} == specs)

        # 抽三种家族各验一份：切换真的写在第 2 页、manifest 字段齐全。
        for spec in ("push", "wind", "morph"):
            spec_obj = next(h for h in B.HYPOTHESES if h["spec"] == spec)
            man = load_json(os.path.join(tmp, spec + ".manifest.json"))
            check("%s manifest 有 spec/ui/family" % spec,
                  man.get("spec") == spec and man.get("family")
                  == spec_obj["family"] and man.get("ui"))
            check("%s manifest 记录网格尺寸" % spec,
                  man.get("grid") == [B.GRID_COLS, B.GRID_ROWS])
            check("%s 切换写在第 2 页" % spec, man.get("transition_on") == 2)
            # deck 里第 2 页必须真的带着这个效果的子元素
            import zipfile
            with zipfile.ZipFile(os.path.join(tmp, spec + ".pptx")) as z:
                x2 = z.read("ppt/slides/slide2.xml").decode("utf-8")
            check("%s deck 第 2 页含 %s" % (spec, spec_obj["child"]),
                  spec_obj["child"] in x2)
    finally:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)

    # ---- 8. 形态分析器：方向必须真的被量出来 ------------------------------
    # §七 的坑是"整帧均值看不出方向"。这里用两张合成的"擦除"帧序列做标定：
    # 一个从左推进、一个从右推进，分析器必须给出相反的方向。
    try:
        import numpy as np
        import cv2
    except ImportError:
        np = cv2 = None
    if np is not None:
        H, W = 9 * 4, 16 * 4

        def sweep_frames(left_to_right, n=12):
            # A directional reveal moves its FRONT across the frame: the newly
            # revealed strip sits at the advancing edge, so the diff centroid
            # travels in the reveal direction. Left-to-right grows the lit
            # region from x=0; right-to-left grows it from x=W. Getting this
            # fixture backwards was the whole reason a bug hid here once --
            # both branches produced identical frames.
            out = []
            for k in range(n):
                f = np.zeros((H, W, 3), np.uint8)
                edge = int(W * (k + 1) / (n + 1))
                if left_to_right:
                    f[:, :edge] = 255          # lit region grows left -> right
                else:
                    f[:, W - edge:] = 255      # lit region grows right -> left
                out.append(f)
            return out

        r1 = B._profile_metrics(sweep_frames(True))
        r2 = B._profile_metrics(sweep_frames(False))
        check("左→右的推进被量成 l->r", r1 and r1["direction"] == "l->r",
              str(r1 and r1["direction"]))
        check("右→左的推进被量成 r->l", r2 and r2["direction"] == "r->l",
              str(r2 and r2["direction"]))
        check("两个方向互为镜像（对称性判据分开）",
              abs(r1["symmetry_v"] - r2["symmetry_v"]) < 0.01)
        # 一个完全静止的序列不能被当成有方向
        still = [np.full((H, W, 3), 128, np.uint8) for _ in range(6)]
        r3 = B._profile_metrics(still)
        check("静止序列判成 none", r3 and r3["direction"] == "none",
              str(r3 and r3["direction"]))
        # 硬切：单帧变化不能被描述成一种"运动形态"
        single = [np.zeros((H, W, 3), np.uint8), np.full((H, W, 3), 255, np.uint8)]
        r4 = B._profile_metrics(single)
        check("单帧硬切判成 cut", r4 and r4["mode"] == "cut",
              str(r4 and r4.get("mode")))

        # ---- 8b. 旋转指标（第四类）--------------------------------------
        # 为什么单独测：前三类指标【看不见旋转】—— 一圈扫下来重心回原点、
        # 不是整页平移、径向剖面又和中心扩散几乎一样。clock 就是这么被错
        # 归到「中心扩散」、方向还错记成 t->b 的。这是用户目视发现的。
        def wheel_frames(cw=True, n=14):
            # 一条从中心射出的扇区，绕中心扫过。cw=False 反向。
            out = []
            for k in range(n):
                f = np.zeros((H, W, 3), np.uint8)
                a0 = (90 - k * 25) if cw else (90 + k * 25)
                yy, xx = np.indices((H, W)).astype(np.float32)
                ang = (np.degrees(np.arctan2(-(yy - H / 2.0), xx - W / 2.0))
                       + 360.0) % 360.0
                delta = np.abs(((ang - (a0 % 360) + 180) % 360) - 180)
                m = delta < 20                      # 当前扫到的那个扇区
                f[m] = 255
                out.append(f)
            return out

        r5 = B._profile_metrics(wheel_frames(cw=True))
        check("顺时针扫一圈被认成旋转（is_rotation）",
              r5 and r5.get("is_rotation") is True,
              "rot_span=%s rot_mono=%s" % (r5 and r5.get("rot_span"),
                                           r5 and r5.get("rot_mono")))
        check("旋转指标给出扇区覆盖数（≥12 才算扫过半圈）",
              r5 and r5.get("rot_span", 0) >= 12,
              str(r5 and r5.get("rot_span")))
        # 直线平移绝不能被认成旋转 —— 否则这个指标会污染整页平移族
        check("平移序列不被误判为旋转",
              r1 and r1.get("is_rotation") is False,
              str(r1 and r1.get("is_rotation")))
        check("静止序列不被误判为旋转",
              r3 and r3.get("is_rotation") is False,
              str(r3 and r3.get("is_rotation")))

    # ---- 9. 形态层文档 + facts 一致 --------------------------------------
    shapes_doc = os.path.join(ROOT, "reference", "transition-shapes.md")
    check("形态层文档存在", os.path.exists(shapes_doc))
    if os.path.exists(shapes_doc):
        text = io.open(shapes_doc, encoding="utf-8").read()
        for token, why in (("整页平移", "六族之一：push/pan 靠位移指标认"),
                           ("方向揭示", "六族之一：wipe/cover 靠漂移指标认"),
                           ("中心扩散", "六族之一：split/ripple"),
                           ("band_travel", "deck2 用来分开 fade 与 wipe 的判据"),
                           ("morph", "形态层必须说明 morph 不适用"),
                           ("旋转", "第七族：clock 靠角度轨迹认（用户目视发现）"),
                           ("角度轨迹", "旋转指标的名字，不能被删掉")):
            check("形态层文档仍写着「%s」" % why, token in text)
        # 48 个效果必须都在表里（避免只写了几族就交差）
        missing = [h["spec"] for h in B.HYPOTHESES
                   if ("`%s`" % h["spec"]) not in text]
        check("形态表覆盖全部 48 个效果", not missing,
              "缺 %s" % missing[:8])

    fx2 = load_json(os.path.join(ROOT, "facts", "transitions.json"))
    ids2 = {r["id"] for r in fx2["rules"]}
    for want in ("shape-needs-two-complementary-probes",
                 "measurement-needs-the-right-instrument-per-effect-family",
                 "random-transition-has-no-stable-shape",
                 "dont-write-i-cant-measure-as-it-has-no-direction"):
        check("facts 仍记着 %s" % want, want in ids2)
    declared2 = " ".join(s.get("location", "") for s in fx2["sources"])
    check("facts 已声明形态层为真相源",
          "reference/transition-shapes.md" in declared2)

    # deck2 的名单必须是 HYPOTHESES 里真实存在的 spec
    bad = [s for s in B.DECK2_SPECS
           if s not in {h["spec"] for h in B.HYPOTHESES}]
    check("DECK2_SPECS 都是真实 spec", not bad, "假 spec: %s" % bad)

    # ---- 10. 证据 deck 的断言必须与形态文档同源 ----------------------------
    # 为什么需要这一组：证据 deck 是给人做"判定 vs 实物"对照用的。如果它引用的
    # 数字/方向与 reference/transition-shapes.md 对不上，那用户核对的就是一份
    # 与真正发货内容不同的东西 —— 比不做还糟（用户会以为已核过）。
    ev = os.path.join(ROOT, "scripts", "build_shape_evidence.py")
    check("证据 deck 生成器存在", os.path.exists(ev))
    if os.path.exists(ev):
        import build_shape_evidence as E
        check("证据 deck 至少覆盖 6 个族",
              len({c["family"] for c in E.CLAIMS}) >= 6,
              "族=%s" % sorted({c["family"] for c in E.CLAIMS}))
        bad_spec = [c["spec"] for c in E.CLAIMS
                    if c["spec"] not in {h["spec"] for h in B.HYPOTHESES}]
        check("证据 deck 的 spec 都是真实效果", not bad_spec, str(bad_spec))
        if os.path.exists(shapes_doc):
            text = io.open(shapes_doc, encoding="utf-8").read()
            drift = []
            for c in E.CLAIMS:
                # 断言里出现的每个小数必须逐字出现在文档里
                for num in re.findall(r"-?\d+\.\d+", c["claim"]):
                    if num not in text:
                        drift.append("%s 的数字 %s 文档里没有" % (c["spec"], num))
                # 方向符号必须同源（用同一套记号，否则读者无法互校）
                for arrow in ("l→r", "r→l", "t→b", "b→t", "l↔r"):
                    if arrow in c["claim"] and arrow not in text:
                        drift.append("%s 的 %s 文档里没有" % (c["spec"], arrow))
            check("证据 deck 与形态文档无漂移（数字/方向同源）",
                  not drift, "; ".join(drift[:5]))

        # ⚠️ 这一条是被用户当场骂出来的：
        #   第一版 deck 贴的是【静态抽帧】，用户说"压根是静态的，看不到动效"。
        #   抽帧只能证明"某一刻有什么"，证明不了"怎么动" —— 而形态层判的就是运动。
        #   现在必须真的嵌入 GIF，且 GIF 必须真的在动。
        check("证据 deck 用循环 GIF（不再贴静态帧）",
              hasattr(E, "make_gif") and hasattr(E, "gif_motion"),
              "缺少 GIF 生成/检测函数")
        check("证据 deck 不再有抽帧函数（grab_frames 已移除）",
              not hasattr(E, "grab_frames"),
              "grab_frames 还在，说明回退成了静态版")
        # 帧数多 ≠ 看得见 —— 必须有可见性阈值，否则又会作出"有 21 帧但不动"的 deck
        check("证据 deck 有可见性阈值（防止再交静态动画）",
              isinstance(getattr(E, "MIN_VISIBLE", None), float)
              and E.MIN_VISIBLE > 0,
              "MIN_VISIBLE=%s" % getattr(E, "MIN_VISIBLE", None))
        # 两个 deck 的分工必须写死在代码里，否则 deck1 看不出的效果会又被静音
        check("证据 deck 知道哪些效果要换 deck2",
              set(E.DECK2_ONLY) & {"wipe", "fade", "split", "random"},
              "DECK2_ONLY=%s" % (E.DECK2_ONLY,))
        bad2 = [s for s in E.DECK2_ONLY
                if s not in {h["spec"] for h in B.HYPOTHESES}]
        check("DECK2_ONLY 都是真实 spec", not bad2, str(bad2))

    # ---- 11. 选择层文档（3b）：判断可以反驳，引用不能漂移 ------------------
    # 为什么需要这一组：3b 是唯一一层**没有测量支撑**的文档 —— 它给的是设计建议。
    # 所以它有两种失效方式，都必须钉住：
    #   ① 建议本身变味（少了某个关系/页型 → 场景字典就残缺了）
    #   ② 更隐蔽：建议里引用的【实测】数字/方向与形态层对不上。
    #      这一层的每条 [实测] 都自称"不能反驳"，那就必须真的能回溯到形态层；
    #      否则读者拿到的是"看起来有证据、其实是我编的" —— 比没有证据更糟。
    choice_doc = os.path.join(ROOT, "reference", "transition-choice.md")
    check("选择层文档存在", os.path.exists(choice_doc))
    if os.path.exists(choice_doc):
        ctext = io.open(choice_doc, encoding="utf-8").read()

        # ① 场景字典的两个维度必须齐全（缺一个就等于没有字典）
        for rel, why in (("递进", "关系之一：顺阅读流"),
                         ("并列", "关系之一：横向对照"),
                         ("转折", "关系之一：换话题，要断"),
                         ("回归", "关系之一：反向，闭环"),
                         ("复位", "关系之一：最常用，无方向")):
            check("选择层覆盖「%s」关系（%s）" % (rel, why), rel in ctext)
        for pt in ("封面", "章节", "金句", "要点正文", "数据", "对比",
                   "流程", "收尾"):
            check("选择层覆盖「%s」页型" % pt, pt in ctext)

        # ② 两种标记都在 —— 这一层的诚实性全押在"哪些能反驳"上
        check("选择层声明了 [实测] 标记（引用，不可反驳）", "[实测]" in ctext)
        check("选择层声明了 ［判断］ 标记（建议，可反驳）", "［判断］" in ctext)
        # 本层不能改上层的硬约束必须写着
        check("选择层声明「本层不能改上层」", "本层不能改上层" in ctext)

        # ③ 反向自证：选择层里出现的每个切换名必须是真实 spec。
        #    引一个不存在的效果 = 读者照着写会写不出来，而且不报错。
        if os.path.exists(REF):
            real = {h["spec"] for h in B.HYPOTHESES}
            # 只认反引号里的单个小写词，避免把 `dir`/`l→r`/代码片段当效果名
            cited = set(re.findall(r"`([a-z][a-z0-9_]{1,20})`", ctext))
            # 那些明确不是效果名的反引号词：属性名、占位符、命令名、指标名，
            # 以及 `none`（"无切换"，合法但不属于 48 个效果集）。注意 `fade`
            # 是**真实** spec，不在此列 —— 混进来会让这条检查失去意义。
            NOT_SPEC = {"dir", "none", "auto", "true", "advtm",
                        "fallback", "choice", "xml", "pptx",
                        # 本轮为 dir 探测新增
                        "dirdeck", "dirmirror",
                        # 指标名（不是切换）
                        "dx", "dy", "tx", "ty",
                        # §三 全量实测引用的指标名与样本目录名
                        "band_travel", "max_frame_diff", "dirprobe2",
                        "page_curl", "peel_off", "shape", "wind"}
            bogus = sorted(s for s in cited
                           if s not in real and s not in NOT_SPEC)
            check("选择层引用的切换名都是真实 spec", not bogus,
                  "疑似假名: %s" % bogus[:8])
            # 反过来：至少得引用到一定数量的真实效果，否则这层是空壳
            used = sorted(s for s in cited if s in real)
            check("选择层至少引用 10 个真实效果", len(used) >= 10,
                  "只引用 %d 个: %s" % (len(used), used[:8]))

        # ④ 最关键：所有 [实测] 引用的数字与方向，必须逐字出现在形态层。
        #    这是"不能反驳"这句承诺的唯一技术保障。
        #
        #    比较前把 U+2212（真正的减号）归一化成 ASCII '-'：要求的是【数值
        #    同源】，不是【字符编码相同】。全仓库排版用的是 U+2212，而 Python
        #    的 \d 正则只认 ASCII '-'，不归一化会制造一堆假失败。
        if os.path.exists(shapes_doc):
            def _norm_neg(t):
                # ONLY U+2212 (the real minus sign). Do NOT fold U+2013
                # (en-dash): it is used as a RANGE separator here ("0.6–0.8s"),
                # and folding it turns the range into a negative number that
                # then has no counterpart in the other document.
                return t.replace("\u2212", "-")

            stext = _norm_neg(io.open(shapes_doc, encoding="utf-8").read())
            ctext_n = _norm_neg(ctext)
            #    只扫**量出来的数字**。必须排除两类"看起来像小数、其实是编号"的串：
            #      ① 本文件自己的小节号（2.1 / 2.6 …）—— 小节号是导航，不是实测值；
            #      ② 形如 `cut`(0.6) 这类**效果名紧邻的括号**写法也不会出现，
            #         因为跨度都写成整数（见下）。
            #    判据：小节号紧跟在"### "或"2." 这样的小节上下文里。用简单的
            #    位置规则：`\d+\.\d+` 前面若是行首/空白且后跟空格+中文标题，则跳过。
            def _measures(t):
                out = []
                for m in re.finditer(r"-?\d+\.\d+", t):
                    # 往前看这一行，若是 markdown 小节标题则不算测量值
                    line_start = t.rfind("\n", 0, m.start()) + 1
                    line = t[line_start:t.find("\n", m.start())]
                    if re.match(r"\s*#{2,4}\s+\d+\.\d+", line):
                        continue
                    out.append(m.group(0))
                return out

            drift = []
            for num in _measures(ctext_n):
                if num not in stext:
                    drift.append("数字 %s 形态层里没有" % num)
            for arrow in ("l→r", "r→l", "t→b", "b→t", "l↔r"):
                if arrow in ctext and arrow not in stext:
                    drift.append("方向 %s 形态层里没有" % arrow)
            check("选择层的实测引用与形态层同源（数字/方向无漂移）",
                  not drift, "; ".join(drift[:5]))

            # ⑤ 形态层的**族名**必须被选择层用同一套说法引（不许自造族名）
            for fam, why in (("整页平移", "选择层讲 push 时该用形态层的族名"),
                             ("方向揭示", "选择层讲 wipe 时该用形态层的族名"),
                             ("中心扩散", "选择层讲 split 时该用形态层的族名"),
                             ("旋转", "选择层讲 clock 时该用形态层的族名")):
                check("选择层沿用形态层族名「%s」（%s）" % (fam, why),
                      fam in ctext)

    # ⑥ facts 必须记着"选择层是判断不是测量"这条元规则。
    #    这不是装饰：它决定了日后有人质疑本文建议时，正确的动作是"改建议"，
    #    而不是"改实测"。
    if os.path.exists(os.path.join(ROOT, "facts", "transitions.json")):
        fx3 = load_json(os.path.join(ROOT, "facts", "transitions.json"))
        ids3 = {r["id"] for r in fx3["rules"]}
        check("facts 记着「选择层是判断不是测量」",
              "choice-layer-is-judgement-not-measurement" in ids3,
              "缺 choice-layer-is-judgement-not-measurement")
        d3 = " ".join(s.get("location", "") for s in fx3["sources"])
        check("facts 已声明选择层为判断来源",
              "reference/transition-choice.md" in d3)

    # ⑦ §5 的方向断层必须修好（旧版没写 dir，导致 push/wipe 默认下分不开）
    spec_doc = os.path.join(ROOT, "reference", "motion-design-spec.md")
    if os.path.exists(spec_doc):
        s5 = io.open(spec_doc, encoding="utf-8").read()
        check("§5 已加方向列（不能只写效果名）", "方向" in s5)
        check("§5 已指向选择层完整版", "transition-choice.md" in s5)
        check("§5 已标出默认 r→l 的陷阱", "r→l" in s5)

    # ⑧ 选择层必须能从 INDEX 到达（否则读者永远找不到这一层）
    index_doc = os.path.join(ROOT, "INDEX.md")
    if os.path.exists(index_doc):
        itext = io.open(index_doc, encoding="utf-8").read()
        check("INDEX 能到达选择层文档", "transition-choice.md" in itext)

    # ---- 12. 方向探测（补 3b §三 的缺口）--------------------------------
    # 为什么需要这一组：选择层 §三 曾挂着一条诚实的 ⚠️ —— 形态层只量了每个
    # 效果的【默认形态】，"把 dir 从 r 换成 l 就会镜像"是个假设，不是测量。
    # 而 §二/§五 的整张推荐表都建在这个假设上（"递进必须显式写 dir=l→r"）。
    # 现在把它变成可测的，并且用合成帧把判据本身先标定住 —— 否则"测了"
    # 和"测对了"还是两回事。
    if np is not None:
        # 标定 1：两个互为镜像的扫过序列，必须被判成 mirror。
        # sweep_frames(True) 从左侧生长，sweep_frames(False) 从右侧 —— 这就是
        # 定义上的镜像，判据必须先在这个已知答案上成立。
        ml = B._profile_metrics(sweep_frames(True))
        mr = B._profile_metrics(sweep_frames(False))
        v_mirror = B.mirror_verdict(ml, mr, axis="x")
        check("镜像标定：左→右 vs 右→左 判成 mirror",
              v_mirror.get("verdict") == "mirror",
              "verdict=%s r=%s dirs=(%s,%s)" % (v_mirror.get("verdict"),
                                                v_mirror.get("r"),
                                                v_mirror.get("dir_a"),
                                                v_mirror.get("dir_b")))
        check("镜像标定：相关系数接近 -1",
              (v_mirror.get("r") is not None) and v_mirror["r"] <= -0.7,
              "r=%s" % v_mirror.get("r"))

        # 标定 2：同一个序列与它自己必须【不】被判成 mirror。
        # 这是负向对照 —— 少了它，"r 永远是负的"这种坏实现也能骗过标定 1。
        v_same = B.mirror_verdict(ml, ml, axis="x")
        check("负向对照：同一序列与自身不判 mirror",
              v_same.get("verdict") != "mirror",
              "verdict=%s r=%s" % (v_same.get("verdict"), v_same.get("r")))

        # 标定 3：轨迹太短 / 无运动时必须说"判不了"，不能硬给一个结论。
        # 形态层吃过这个亏：数据不足时静默给出错误分类（§47 盲区是静默的）。
        still = [np.full((H, W, 3), 128, np.uint8) for _ in range(6)]
        r_still = B._profile_metrics(still)
        v_nm = B.mirror_verdict(r_still, r_still, axis="x")
        check("静止/无轨迹时判成 no-motion（不硬下结论）",
              v_nm.get("verdict") == "no-motion",
              "verdict=%s" % v_nm.get("verdict"))

        # 标定 4：判据必须真的用到轨迹，而不是只看汇总量标签。
        # 若实现退化成"两个 direction 标签不同就算镜像"，这一条会碎。
        check("镜像判据基于逐帧轨迹（cx_trace 被返回且非空）",
              (ml or {}).get("cx_trace") and len(ml["cx_trace"]) >= 3,
              "len=%s" % len((ml or {}).get("cx_trace") or []))

    # DIR_PROBE 的名单必须都是真实 spec，且每条都说得出"为什么测它"
    bad_dp = [s for s, _d, _a, _w in B.DIR_PROBE
              if s not in {h["spec"] for h in B.HYPOTHESES}]
    check("DIR_PROBE 引用的 spec 都真实存在", not bad_dp, str(bad_dp))
    check("DIR_PROBE 每条都写了理由（不能只列组合）",
          all(w.strip() for _s, _d, _a, w in B.DIR_PROBE))
    # 范围纪律：这一组是补 §三 的缺口，不是把 48 个效果全测一遍。
    # 判据不是拍一个"份数上限"，而是【有原则地圈定范围】：
    #   只测两个方向性族（整页平移 + 方向揭示）的成员，减去禁配清单。
    #   禁配的效果商务场景不许用，测它的镜像性没有决策价值。
    # 这样以后加成员时，判据会问"它属于这两族吗、它被禁了吗"，而不是"数量超了吗"。
    #
    # 族名从 transition-shapes.md §一 的族表里【读】出来，不在测试里另抄 ——
    # 抄一份就会漂。行格式：| **族名** | 在做什么 | 成员 | 指标 |
    dir_fams = {"整页平移", "方向揭示"}
    fam_members = set()
    if os.path.exists(shapes_doc):
        stext_f = io.open(shapes_doc, encoding="utf-8").read()
        for line in stext_f.splitlines():
            m = re.match(r"\|\s*\*\*(.+?)\*\*\s*\|[^|]*\|([^|]+)\|", line)
            if m and m.group(1).strip() in dir_fams:
                for tok in re.findall(r"`(\w+)`", m.group(2)):
                    fam_members.add(tok)
    check("从形态层读到了两个方向性族的成员表",
          len(fam_members) >= 10, "读到 %d 个" % len(fam_members))
    ztext = ""
    if os.path.exists(choice_doc):
        ztext = io.open(choice_doc, encoding="utf-8").read()
    # 禁配清单在 §四：从文档里读，不在测试里另抄一份（抄了就会漂）。
    # 行格式是一串反引号词，用空格分隔：`random` `checkerboard` `blinds` ...
    ban = set()
    m_ban = re.search(r"### 商务场合默认禁用[^\n]*\n\n(.+)", ztext)
    if m_ban:
        ban = set(re.findall(r"`(\w+)`", m_ban.group(1)))
    check("从选择层读到了 §四 禁配清单", len(ban) >= 5, "读到 %s" % sorted(ban))
    probe_specs = {s for s, _d, _a, _w in B.DIR_PROBE}
    outside = sorted(s for s in probe_specs if s not in fam_members)
    check("DIR_PROBE 只含两个方向性族的成员（不越界到中心扩散/旋转等）",
          not outside, "越界: %s" % outside)
    banned_in = sorted(probe_specs & ban)
    check("DIR_PROBE 不含 §四 禁配效果（禁配的测了也没决策价值）",
          not banned_in, "混入禁配: %s" % banned_in)
    # 另外：只测了方向的那些效果，必须真的【两个方向都测】才能判镜像
    from collections import Counter as _C
    cnt = _C(s for s, _d, _a, _w in B.DIR_PROBE)
    lonely = sorted(s for s, c in cnt.items() if c < 2)
    check("DIR_PROBE 每个效果都测了两个方向（单方向判不了镜像）",
          not lonely, "只测一个方向: %s" % lonely)
    # 每个 dir 取值都应只出现一次（重复 = 白渲染一份）
    pairs = [(s, d) for s, d, _a, _w in B.DIR_PROBE]
    check("DIR_PROBE 无重复 (spec,dir) 组合",
          len(pairs) == len(set(pairs)),
          "重复: %s" % [p for p in set(pairs) if pairs.count(p) > 1])
    # 垂直效果必须用 u/d 测，水平效果必须用 l/r 测 —— 混了等于没测
    axis_bad = []
    for s, d, a, _w in B.DIR_PROBE:
        if a == "y" and d not in ("u", "d"):
            axis_bad.append("%s dir=%s axis=y" % (s, d))
        if a == "x" and d not in ("l", "r"):
            axis_bad.append("%s dir=%s axis=x" % (s, d))
    check("DIR_PROBE 轴与 dir 取值自洽（垂直用 u/d、水平用 l/r）",
          not axis_bad, str(axis_bad[:4]))
    # 选择层真正用到的四个效果必须在名单里（否则缺口没补到点上）
    need = {"push", "wipe", "cover", "uncover"}
    check("DIR_PROBE 覆盖选择层实际推荐的四个效果",
          need <= probe_specs,
          "缺 %s" % (need - probe_specs))

    # 缺口闭合的**文档证据**：§三 不能再挂着"没有实测证据"。
    # 这条是负向检查 —— 防止有人回退文档、把已闭合的缺口又写回"未知"。
    if os.path.exists(choice_doc):
        ctext2 = io.open(choice_doc, encoding="utf-8").read()
        check("§三 不再自称『没有实测证据』（缺口已闭合）",
              "本文没有实测证据" not in ctext2,
              "文档里还留着『没有实测证据』")
        ctext2n = ctext2.replace("\u2212", "-")
        for token, why in (("镜", "§三 要写出镜像结论"),
                           ("-0.99", "wipe/reveal 的相关系数要可核对"),
                           ("1.004", "push 的累积位移要可核对"),
                           ("+0.406", "wipe 的 band_travel 要可核对"),
                           ("+0.671", "glitter 的 band_travel 要可核对")):
            check("§三 写着 %s（%s）" % (token, why), token in ctext2n)
        # 两路判据的方法论必须留在文档里（不然下次又会只用一路）
        check("§三 说明了为何 push 需要另一路判据",
              "变化重心根本不动" in ctext2 or "重心根本不动" in ctext2)
        # 负向检查：§三 不能再把 split 列进"能定向"（实测它是中心扩散、dir 无效）
        check("§三 不再把 split 说成能定向",
              "`push` `wipe` `cover` `uncover` `split` `reveal`" not in ctext2,
              "§三 还留着把 split 列进能定向的旧表")
        # 全量实测后，必须写出"多少会被忽略"这件事
        check("§三 写出 dir 被忽略的一类（不是所有效果都能定向）",
              "被忽略" in ctext2)
    # 形态层 §九 也要有全量结论（不能只在选择层写）
    if os.path.exists(shapes_doc):
        stext9 = io.open(shapes_doc, encoding="utf-8").read()
        stext9n = stext9.replace("\u2212", "-")
        check("§九 写出被 dir 忽略的成员", "被忽略" in stext9 or "被**忽略**" in stext9)
        check("§九 写出 box/comb 加了 dir 会打不开",
              "打开" in stext9 and ("9.4" in stext9 or "打不开" in stext9))
        check("§九 保留 wipe 的 band_travel 数字", "+0.406" in stext9n)

    # facts 必须记着这几条（否则"缺口已闭合"这件事会随时间丢失）
    if os.path.exists(os.path.join(ROOT, "facts", "transitions.json")):
        fx4 = load_json(os.path.join(ROOT, "facts", "transitions.json"))
        ids4 = {r["id"] for r in fx4["rules"]}
        for want in ("mirroring-needs-two-instruments-not-one",
                     "dir-mirrors-for-the-four-recommended-effects",
                     "dir-ignored-by-most-directional-effects",
                     "adding-dir-can-make-the-file-unopenable",
                     "direction-must-be-measured-not-inferred-from-effect-name"):
            check("facts 记着 %s" % want, want in ids4)

    # 全量实测的归档必须留着 —— 这是 §九 结论的可回溯证据。
    fxfull = os.path.join(ROOT, "tests", "fixtures", "dir", "dirmirror_full.json")
    if os.path.exists(fxfull):
        ffull = load_json(fxfull)
        frows = {r["spec"]: r for r in ffull.get("rows", [])}
        check("归档 dirmirror_full 覆盖 19 个效果", len(frows) >= 19,
              "只有 %d 条" % len(frows))
        # 会镜像的 8 个：像素不同且 band_travel 反号
        mirror8 = {"push", "pan", "switch", "wipe", "cover", "uncover",
                   "reveal", "glitter"}
        got = {s for s, r in frows.items() if r.get("band_flip")}
        check("归档里 band_travel 反号的正是 8 个", got == mirror8,
              "多/少: %s" % ((got - mirror8) | (mirror8 - got)))
        # 被忽略的：像素逐帧相同 —— 这是另一路独立判据，必须与上一条同为 8 个
        same = {s for s, r in frows.items() if r.get("pix_identical")}
        check("归档里像素逐帧相同的正是另外 11 个",
              same == set(frows) - mirror8,
              "不一致: %s" % ((same ^ (set(frows) - mirror8))))
        # 分离度：没有灰色地带 —— 被忽略的全是 0.000
        nofloor = [s for s in same if (frows[s].get("max_frame_diff") or 0) > 0.05]
        check("被忽略的效果逐帧像素差都是 0.000（无灰色地带）", not nofloor,
              str(nofloor))
        # 目视对照图也要在
        check("目视对照图 dir_compare_full.png 已归档",
              os.path.exists(os.path.join(ROOT, "tests", "fixtures", "dir",
                                          "dir_compare_full.png")))

    # 生成的 deck 文件名必须能直接配对（分析器靠文件名分 l/r）
    import glob as _glob
    probe_dir = os.path.join(ROOT, ".workbuddy", "_dirprobe_smoke")
    try:
        if not os.path.isdir(probe_dir):
            B.cmd_dirdeck(probe_dir)
        made = sorted(os.path.basename(p) for p in
                      _glob.glob(os.path.join(probe_dir, "*.pptx")))
        check("dirdeck 生成的文件名可配对（<spec>_dir_<两个 dir 值>）",
              all(re.fullmatch(r"\w+_dir_[lrud]\.pptx", n) for n in made)
              and len(made) == len(B.DIR_PROBE),
              "%d 个: %s" % (len(made), made[:4]))
        # 生成器必须真的把 dir 写进 XML —— 文件名对但属性没写是最坏的假象
        import zipfile as _zip
        bad_xml = []
        for s, d, _a, _w in B.DIR_PROBE:
            fp = os.path.join(probe_dir, "%s_dir_%s.pptx" % (s, d))
            with _zip.ZipFile(fp) as z:
                x = z.read("ppt/slides/slide2.xml").decode("utf-8")
            m = re.search(r"<(p|p14|p15|p159):\w+[^>]*dir=\"%s\"[^>]*/>" % d, x)
            if not m:
                bad_xml.append("%s_dir_%s" % (s, d))
        check("dirdeck 真的把 dir 写进了 slide XML", not bad_xml, str(bad_xml))
    finally:
        import shutil as _sh
        if os.path.isdir(probe_dir):
            _sh.rmtree(probe_dir, ignore_errors=True)

    # ------------------------------------------------------------------
    # 组 13：属性取值全集（attrdeck / attrdiff）—— 3a 遗留的第二件事
    # ------------------------------------------------------------------
    # 名单纪律：每条 (spec,label,overrides,why) 四元组；spec 真实；
    # 同 spec 内 label 不重复（重复 = 白渲染一份）。
    bad_a = [s for s, _l, _o, _w in B.ATTR_PROBE
             if s not in {h["spec"] for h in B.HYPOTHESES}]
    check("ATTR_PROBE 引用的 spec 都真实存在", not bad_a, str(bad_a))
    check("ATTR_PROBE 每条都写了理由",
          all(w.strip() for _s, _l, _o, w in B.ATTR_PROBE))
    labels = [(s, l) for s, l, _o, _w in B.ATTR_PROBE]
    check("ATTR_PROBE 无重复 (spec,label)",
          len(labels) == len(set(labels)),
          str([p for p in set(labels) if labels.count(p) > 1]))
    # 覆盖三件 3a 遗留的具体问题：spokes 取值、prism 组合、invX 反向
    cover = {s for s, _l, _o, _w in B.ATTR_PROBE}
    for need, why in (("clock", "spokes 取值全集"),
                      ("split", "orient × dir 四组合"),
                      ("cube", "prism 属性组合"),
                      ("wind", "invX 反向（与 dir 有别）")):
        check("ATTR_PROBE 覆盖 %s（%s）" % (need, why), need in cover)

    # 每个 override 必须是 (attr, value) 二元组，值要么是字符串要么是 None
    bad_ov = []
    for s, l, ov, _w in B.ATTR_PROBE:
        for pair in ov:
            if not (isinstance(pair, tuple) and len(pair) == 2):
                bad_ov.append("%s/%s" % (s, l)); break
            _a, v = pair
            if not (v is None or isinstance(v, str)):
                bad_ov.append("%s/%s:%r" % (s, l, v)); break
    check("ATTR_PROBE 的 override 都是 (attr, value) 且 value 为 str|None",
          not bad_ov, str(bad_ov[:4]))

    # `_set_attrs` 的删除语义：value=None 必须真的把属性删掉（不是留空值），
    # 否则"默认形态"那条会偷偷带着 HYPOTHESES 里原本的属性，基线就错了。
    check("_set_attrs 能把属性删掉（value=None）",
          B._set_attrs('<p:wheel spokes="1"/>', (("spokes", None),))
          == "<p:wheel/>",
          B._set_attrs('<p:wheel spokes="1"/>', (("spokes", None),)))
    check("_set_attrs 能换成新值",
          B._set_attrs('<p:wheel spokes="1"/>', (("spokes", "8"),))
          == '<p:wheel spokes="8"/>',
          B._set_attrs('<p:wheel spokes="1"/>', (("spokes", "8"),)))
    check("_set_attrs 不叠加同名属性",
          B._set_attrs('<p:zoom dir="in"/>', (("dir", "out"),))
          == '<p:zoom dir="out"/>',
          B._set_attrs('<p:zoom dir="in"/>', (("dir", "out"),)))
    # prst= 这类"身份属性"不能被误删
    check("_set_attrs 不乱动未点名的属性（如 prst=）",
          B._set_attrs('<p15:prstTrans prst="wind"/>', (("invX", "1"),))
          == '<p15:prstTrans prst="wind" invX="1"/>',
          B._set_attrs('<p15:prstTrans prst="wind"/>', (("invX", "1"),)))

    # 归档必须留着（§十 结论的可回溯证据）：spokes 默认=4、非法值=默认
    fxattr = os.path.join(ROOT, "tests", "fixtures", "dir", "attrdiff_full.json")
    if os.path.exists(fxattr):
        fa = load_json(fxattr)
        groups = {r["spec"]: r for r in fa.get("attrs", [])}
        ck = groups.get("clock", {}).get("pairs", [])
        byl = {p["label"]: p for p in ck}
        check("归档 clock 的 spokes4 与默认像素相同",
              byl.get("spokes4", {}).get("max_frame_diff") == 0.0,
              str(byl.get("spokes4")))
        check("归档 clock 的 spokes6（非法值）也像素相同 = 退回默认",
              byl.get("spokes6", {}).get("max_frame_diff") == 0.0,
              str(byl.get("spokes6")))
        check("归档 clock 的 spokes1/2/3/8 都与默认不同",
              all(byl.get(k, {}).get("differs") for k in
                  ("spokes1", "spokes2", "spokes3", "spokes8")),
              str({k: byl.get(k, {}).get("max_frame_diff")
                   for k in ("spokes1", "spokes2", "spokes3", "spokes8")}))
        # comb：显式 horz = 默认（0.000），vert 明显不同
        cmb = {p["label"]: p for p in groups.get("comb", {}).get("pairs", [])}
        check("归档 comb 的 dir=horz 等于默认、dir=vert 不同",
              cmb.get("horz", {}).get("max_frame_diff") == 0.0
              and cmb.get("vert", {}).get("differs"),
              str(cmb))
        # prism 四组合两两不同（三个对照项都 differs）
        cub = groups.get("cube", {}).get("pairs", [])
        check("归档 prism 的其余三个组合都与裸 prism 不同",
              len(cub) == 3 and all(p["differs"] for p in cub),
              str([(p["label"], p["max_frame_diff"]) for p in cub]))
        # invX 真的反向（wind 最明显）
        wnd = {p["label"]: p for p in groups.get("wind", {}).get("pairs", [])}
        check("归档 wind 的 invX=1 与默认不同（真反向）",
              wnd.get("invX", {}).get("differs"), str(wnd))
        check("目视对照图 attr_compare_full.png 已归档",
              os.path.exists(os.path.join(ROOT, "tests", "fixtures", "dir",
                                          "attr_compare_full.png")))

    # 文档证据：名字层 §5.5 + 形态层 §十 + facts + pitfalls §50
    if os.path.exists(os.path.join(ROOT, "reference", "transitions.md")):
        tdoc = io.open(os.path.join(ROOT, "reference", "transitions.md"),
                       encoding="utf-8").read()
        check("名字层写出 §5.5 属性取值全集", "属性取值全集" in tdoc)
        check("名字层写明 clock 默认 4 根", "默认 = 4" in tdoc or "默认=4" in tdoc)
        check("名字层写明非法值退回默认", "静默退回默认" in tdoc or "退回默认" in tdoc)
        check("名字层写明 invX 不是 dir",
              "`invX` 不是 `dir`" in tdoc or "invX" in tdoc)
    # ⚠️ §50 已随踩坑手册拆分搬到 pitfall-silent-drop.md。断言要跟着走 ——
    # "静默失败"这一类归那一份。旧写法读 com-pitfalls.md 会因为文件里不再有
    # 正文而假失败（见第 16 组：manifest 保留，但正文已迁走）。
    _psd = os.path.join(ROOT, "reference", "pitfall-silent-drop.md")
    if os.path.exists(_psd):
        pdoc = io.open(_psd, encoding="utf-8").read()
        check("pitfalls 新增 §50（非法值静默退回默认）",
              "## 50." in pdoc and "静默退回默认" in pdoc)
    if os.path.exists(os.path.join(ROOT, "facts", "transitions.json")):
        fx5 = load_json(os.path.join(ROOT, "facts", "transitions.json"))
        ids5 = {r["id"] for r in fx5["rules"]}
        for want in ("invalid-attribute-value-silently-falls-back-to-default",
                     "clock-default-is-four-spokes",
                     "invX-is-the-reverse-for-p15-prsttrans-not-dir"):
            check("facts 记着 %s" % want, want in ids5)

    # attrdeck 生成器冒烟：文件名可配对 + 属性真写进 XML
    aprobe = os.path.join(ROOT, ".workbuddy", "_attrprobe_smoke")
    try:
        # ⚠️ 判据是**有没有文件**，不是有没有目录 —— 一次中断的测试会留下
        # 一个空目录，而  会因此跳过生成，然后在下面读文件时炸掉，
        # 报出来的错看起来像生成器坏了，其实是残留状态。
        if not _glob.glob(os.path.join(aprobe, "*.pptx")):
            B.cmd_attrdeck(aprobe)
        amade = sorted(os.path.basename(p) for p in
                       _glob.glob(os.path.join(aprobe, "*.pptx")))
        check("attrdeck 生成的文件数正确",
              len(amade) == len(B.ATTR_PROBE),
              "%d 个: %s" % (len(amade), amade[:3]))
        import zipfile as _zip2
        abad = []
        for s, l, ov, _w in B.ATTR_PROBE:
            fp = os.path.join(aprobe, "%s_attr_%s.pptx" % (s, l))
            with _zip2.ZipFile(fp) as z:
                x = z.read("ppt/slides/slide2.xml").decode("utf-8")
            for attr, val in ov:
                if val is None:
                    # 删除语义：该属性不该出现在 child 里
                    if re.search(r'<%s[^>]*\b%s=' % ("\\w+", attr), x):
                        # 可能是别的元素带的，只在 transition 块内判
                        m = re.search(r"<p:transition[^>]*>(.*?)</p:transition>",
                                      x, re.S)
                        inner = m.group(1) if m else ""
                        if re.search(r"\b%s=\"" % attr, inner):
                            abad.append("%s/%s 未删 %s" % (s, l, attr))
                else:
                    if not re.search(r'\b%s="%s"' % (attr, re.escape(val)), x):
                        abad.append("%s/%s 缺 %s=%s" % (s, l, attr, val))
        check("attrdeck 的属性真的写进了 slide XML（含删除语义）",
              not abad, str(abad[:4]))
    finally:
        import shutil as _sh2
        if os.path.isdir(aprobe):
            _sh2.rmtree(aprobe, ignore_errors=True)

    # ------------------------------------------------------------------
    # 14 组：切换 × 页内动画同页（TIMING_PROBE）
    #
    # 这一组守的是一条**被推翻过的断言**：文档曾写「两者正交、不限名额」。
    # 实测推翻了「时间上正交」那一半 —— 页内动画排队等切换演完。
    # 所以这里既守结构（两块都在、顺序对），也守那条实测出来的延迟规律。
    # ------------------------------------------------------------------
    print("\n== 14. 切换 × 页内动画：结构平行、时间串行 ==")

    if os.path.exists(os.path.join(ROOT, "scripts", "build_transition_table.py")):
        # 范围纪律：每一行要么是"组合"，要么是某个组合的"对照"，
        # 且对照必须与它的组合行用同一个切换元素（只差动画这一个变量）。
        labels = [r[2] for r in B.TIMING_PROBE]
        check("TIMING_PROBE 标签唯一", len(labels) == len(set(labels)),
              str([l for l in labels if labels.count(l) > 1]))
        for spec, tspec, label, why in B.TIMING_PROBE:
            check("TIMING_PROBE 行 %s 是四元组且有 why" % label,
                  bool(why) and isinstance(why, str))
        # 每一组组合都要有对照：both_<x> 必须配 solo_anim 与一个 solo_trans
        combos = [l for l in labels if l.startswith("both_")]
        check("有组合行也有对照行", len(combos) >= 2 and
              any(l.startswith("solo_anim") for l in labels) and
              any(l.startswith("solo_trans") for l in labels))
        # push/wipe 两组组合都必须有"同切换的对照"——这是踩过的坑：
        # 对照写成别的切换，会一次动两个变量，低能量的那个被读成"切换消失"
        for grp, want_trans in (("both_push", "solo_trans"),
                                ("both_wipe", "solo_trans_wipe")):
            check("%s 有同切换对照 %s" % (grp, want_trans), want_trans in labels)

    # 归档：实测 JSON 与目视对照图
    tfx = os.path.join(ROOT, "tests", "fixtures", "dir", "timingdiff_full.json")
    if os.path.exists(tfx):
        tf = load_json(tfx)
        trows = {r["label"]: r for r in tf.get("timing", [])}
        check("timingdiff 归档里有组合与对照",
              "both_push" in trows and "solo_trans" in trows)
        # ⭐ 实测规律：同页有切换时，进入形状出现得**更晚**，
        #    且延迟随切换时长单调增。这三条一起守住"串行"这个结论。
        solo = trows.get("solo_anim", {}).get("entry_delay_ms")
        short = trows.get("both_fade_short", {}).get("entry_delay_ms")
        push = trows.get("both_push", {}).get("entry_delay_ms")
        long_ = trows.get("both_fade_long", {}).get("entry_delay_ms")
        check("归档：无切换时进入形状几乎不延迟（<=120ms）",
              solo is not None and solo <= 120, str(solo))
        check("归档：挂了切换后进入形状被推后（>300ms）",
              push is not None and push > 300, str(push))
        check("归档：延迟随切换时长单调增（short < push < long）",
              short is not None and push is not None and long_ is not None
              and short < push < long_,
              "%s < %s < %s" % (short, push, long_))
        # 两族切换（push / wipe）在**同样 800ms** 下延迟应当一致 ——
        # 说明改的是"时长"不是"哪个效果"
        pw = trows.get("both_wipe", {}).get("entry_delay_ms")
        check("归档：push 与 wipe 同为 800ms 时延迟相同",
              pw == push, "%s vs %s" % (pw, push))
        # 动画自身的长度不受切换影响 —— 五种组合的斜坡都是同一个值
        ramps = {r["label"]: r.get("entry_ramp_ms") for r in tf.get("timing", [])
                 if r["label"].startswith("both_") and r.get("entry_ramp_ms")}
        check("归档：动画自身长度不受切换影响（组合间斜坡一致）",
              len(set(ramps.values())) <= 1,
              str(ramps))
        # 结构：往返报告里没有任何一份丢块
        rt = os.path.join(ROOT, "tests", "fixtures", "dir", "timing_rt_full.json")
        if os.path.exists(rt):
            rj = load_json(rt)
            lost = [d["deck"] for d in rj.get("decks", [])
                    if d.get("timing_lost") or d.get("trans_lost")]
            check("归档：往返后没有任何一份丢块", not lost, str(lost))
            orders = {d.get("order") for d in rj.get("decks", [])
                      if d.get("after_transition") and d.get("after_timing")}
            check("归档：两块都在时顺序是 transition 在前",
                  orders and orders == {"transition-before-timing"},
                  str(orders))
    else:
        check("timingdiff 归档存在", False, tfx)

    check("目视对照图 timing_compare_full.png 已归档",
          os.path.exists(os.path.join(ROOT, "tests", "fixtures", "dir",
                                      "timing_compare_full.png")))

    # 文档证据：机制层 §五 + pitfalls §51 + symptoms + facts
    tmodel = os.path.join(ROOT, "reference", "transition-model.md")
    if os.path.exists(tmodel):
        tm = io.open(tmodel, encoding="utf-8").read()
        check("机制层新增 §五（切换与动画串行）",
              "## 五、" in tm and "不是正交" in tm)
        check("机制层写明延迟 ≈ 切换时长 + 133 ms", "133 ms" in tm)
        # 「与切换正交」只允许以**被引用的旧断言**形式存在（§五 的举证），
        # 不允许作为正文断言。判据：该短语每一处都出现在引用块（> 「…」）里。
        bad_ortho = []
        for ln in tm.splitlines():
            if "与切换正交" in ln and not ln.lstrip().startswith(">"):
                bad_ortho.append(ln.strip()[:60])
        check("机制层不再**无条件**断言『与切换正交』（只允许在引用旧断言处）",
              not bad_ortho, str(bad_ortho))
        check("机制层决策表已标注时间串行",
              "时间上排在切换之后" in tm or "串行" in tm)
    # ⚠️ §51 拆到了 pitfall-transition.md（切换类），"全盲"那条也在同一节里。
    _ptr = os.path.join(ROOT, "reference", "pitfall-transition.md")
    if os.path.exists(_ptr):
        pd = io.open(_ptr, encoding="utf-8").read()
        check("pitfalls 新增 §51（动画排队等切换）",
              "## 51." in pd and "不是正交" in pd)
        check("pitfalls §51 记下『能量指标对淡入全盲』",
              "全盲" in pd)
    if os.path.exists(os.path.join(ROOT, "reference", "symptoms.md")):
        sd = io.open(os.path.join(ROOT, "reference", "symptoms.md"),
                     encoding="utf-8").read()
        check("symptoms 新增『动画像慢了半拍』",
              "慢了半拍" in sd)
    if os.path.exists(os.path.join(ROOT, "facts", "transitions.json")):
        fx6 = load_json(os.path.join(ROOT, "facts", "transitions.json"))
        ids6 = {r["id"] for r in fx6["rules"]}
        for want in ("transition-and-timing-are-not-time-orthogonal",
                     "energy-thresholds-are-blind-to-fade-in",
                     "trigger-with-means-with-the-animation-sequence-not-with-the-transition"):
            check("facts 记着 %s" % want, want in ids6)

    # timingdeck 生成器冒烟：文件名可配对 + 两块真写进 XML + 顺序正确
    tprobe = os.path.join(ROOT, ".workbuddy", "_timingprobe_smoke")
    try:
        if not os.path.isdir(tprobe):
            B.cmd_timingdeck(tprobe)
        tmade = sorted(os.path.basename(p) for p in
                       _glob.glob(os.path.join(tprobe, "*.pptx")))
        check("timingdeck 生成的文件数正确",
              len(tmade) == len(B.TIMING_PROBE),
              "%d 个: %s" % (len(tmade), tmade[:3]))
        import zipfile as _zip3
        from lxml import etree as _et
        tbad = []
        for spec, tspec, label, _w in B.TIMING_PROBE:
            fp = os.path.join(tprobe, "timing_%s.pptx" % label)
            with _zip3.ZipFile(fp) as z:
                x = z.read("ppt/slides/slide2.xml").decode("utf-8")
            # 良构
            try:
                root = _et.fromstring(x.encode("utf-8"))
            except Exception as exc:
                tbad.append("%s 不良构: %s" % (label, exc))
                continue
            kids = [_et.QName(k).localname for k in root]
            has_t = "<p:timing>" in x
            has_x = bool(re.search(r"<p:transition\b", x))
            want_t = label != "solo_trans" and label != "solo_trans_wipe" \
                and label != "solo_fade_800"
            want_x = tspec is not None
            if has_t != want_t:
                tbad.append("%s timing=%s 期望 %s" % (label, has_t, want_t))
            if has_x != want_x:
                tbad.append("%s transition=%s 期望 %s" % (label, has_x, want_x))
            # 顺序：cSld → clrMapOvr → transition(AlternateContent) → timing
            if has_t and has_x:
                if kids.index("AlternateContent") > kids.index("timing"):
                    tbad.append("%s 顺序错: %s" % (label, kids))
        check("timingdeck 两块真写进 slide XML 且顺序正确",
              not tbad, str(tbad[:4]))
    finally:
        import shutil as _sh3
        if os.path.isdir(tprobe):
            _sh3.rmtree(tprobe, ignore_errors=True)

    # ------------------------------------------------------------------
    # 15 组：48 项之间的取舍（选择层 §二·补）
    #
    # 这一组守的是"选择层的数字必须可回溯到形态层"这条承诺。
    # 选择层是判断层，但它引用的每一个 [实测] 数字都必须逐字活在形态层里 ——
    # 否则判断就建在编出来的数据上，而这恰恰是判断层最容易出的错。
    # ------------------------------------------------------------------
    print("\n== 15. 48 项之间的取舍：引用可回溯 + 无自造效果名 ==")

    choice_p = os.path.join(ROOT, "reference", "transition-choice.md")
    shapes_p = os.path.join(ROOT, "reference", "transition-shapes.md")
    if os.path.exists(choice_p) and os.path.exists(shapes_p):
        ch = io.open(choice_p, encoding="utf-8").read()
        shp = io.open(shapes_p, encoding="utf-8").read()

        check("选择层新增 §二·补（48 项之间的取舍）",
              "## 二·补" in ch and "48 个名字 ≠ 48 种选择" in ch)

        # 形态层 48 项表的 span 数据（唯一真相源）
        spans = {m[0]: int(m[5]) for m in re.findall(
            r"^\| \`(\w+)\` \| ([^|]+)\| *([^|]*)\| *([^|]*)\| *\*{0,2}([^|*]+?)\*{0,2} *\| *(\d+) \|",
            shp, re.M)}
        check("形态层能解析出跨度数据（>=45 项）", len(spans) >= 45,
              "%d 项" % len(spans))

        # 每个 `spec`(NNN) 形式的跨度断言，必须与形态层逐字一致
        claimed = re.findall(r"`(\w+)`\((\d+)\)", ch)
        mism = []
        for spec, num in claimed:
            if spans.get(spec) != int(num):
                mism.append("%s cited %s actual %s" % (spec, num, spans.get(spec)))
        check("选择层引用的跨度数字全部与形态层一致（%d 处）" % len(claimed),
              not mism, str(mism))

        # 选择层里的每个反引号效果名都必须是真实 spec（不许自造）
        known = set(spans) | set(re.findall(r"^\| \`(\w+)\`", shp, re.M))
        NOT_EFFECT = {"none", "dir", "spec", "spd", "dur", "advTm", "morph",
                      "option", "xml", "l", "r", "u", "d", "t", "b", "x", "y",
                      "w", "h", "box", "comb"} | {"cut"}
        # 收窄到 §二·补 这一段，避免把别处的普通词当效果名
        sec = ch[ch.index("## 二·补"):ch.index("## 三、")]
        sec_specs = set(re.findall(r"`([a-z_0-9]+)`", sec))
        invented = sorted(s for s in sec_specs
                          if s not in known and s not in NOT_EFFECT)
        check("§二·补 里没有自造的效果名", not invented, str(invented))

        # ⚠️ 关键：clock 的跨度是"指标看不见旋转"的伪值，必须被标为不可用
        check("§二·补 标出 clock 跨度是伪值（不可拿它比大小）",
              "指标看不见旋转" in sec and "clock" in sec)

        # 硬约束必须在取舍一节里也出现（不能只在 §三）
        check("§二·补 复述 dir 会让 box/comb 打不开",
              "打不开" in sec and "comb" in sec and "box" in sec)

        # 三层口径一致：族名不许自造
        fam_names = ["硬切", "整页平移", "方向揭示", "中心扩散", "旋转",
                     "均匀淡变", "随机"]
        for f in fam_names:
            check("§二·补 沿用形态层族名「%s」" % f, f in shp)
        used_bad = [f for f in ("扫描族", "淡入族", "平移族")
                    if f in ch and f not in shp]
        check("§二·补 没自造族名", not used_bad, str(used_bad))
    else:
        check("选择层与形态层都可读", False)

    return report(TITLE, fails)


if __name__ == "__main__":
    sys.exit(main())
