#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""切换实测表的回归测试。

为什么需要这份测试：scripts/transition_reference.json 是量出来的，而
scripts/motion.py 直接读它。两者一旦漂移，症状不是报错，而是**写出一个
PowerPoint 会悄悄改掉的块** —— 看起来一切正常，文件就是不对。所以这里把
"表还成立"和"motion.py 还认这张表"钉住。

不测的东西：本测试**不**打开 PowerPoint。真机验证在 tests/smoke.py。

模块内容表（19 组；组号是稳定接口，别重排）
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
  16  踩坑手册拆分后，§编号必须仍然可解析
  17  症状文档拆分后，现象栏必须仍能找到
  18  配方文档拆分：文件名必须对上内容
  19  长文档必须有目录（导航不该靠滚）
"""
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "scripts"))

REF = os.path.join(ROOT, "scripts", "transition_reference.json")

# motion.py 曾经手写的 13 个 spec + none。改名就是 breaking change --
# 用户 spec 里写的 type: push 必须在任意版本里产出同样的东西。
LEGACY_SPECS = ("fade", "smoothfade", "fadeblack", "push", "pushleft", "wipe",
                "cover", "split", "zoom", "dissolve", "strips", "pull",
                "randombar", "none")

fails = []


def _load_json(path):
    # 这些文件有的是 PowerShell 写的（Set-Content -Encoding UTF8 会带 BOM）。
    return json.load(io.open(path, encoding="utf-8-sig"))


def check(name, ok, detail=""):
    print("  %-58s %s" % (name, "ok" if ok else "FAIL " + detail))
    if not ok:
        fails.append(name)


def main():
    print("== 切换实测表回归 ==")
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

    fx = _load_json(os.path.join(ROOT, "facts", "transitions.json"))
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
        idx = _load_json(os.path.join(tmp, "_index.json"))
        check("每份 Hypothesis 都有 deck", idx["count"] == len(B.HYPOTHESES),
              "got %d vs %d" % (idx["count"], len(B.HYPOTHESES)))
        specs = {d["spec"] for d in idx["decks"]}
        check("deck 覆盖全部 spec",
              {h["spec"] for h in B.HYPOTHESES} == specs)

        # 抽三种家族各验一份：切换真的写在第 2 页、manifest 字段齐全。
        for spec in ("push", "wind", "morph"):
            spec_obj = next(h for h in B.HYPOTHESES if h["spec"] == spec)
            man = _load_json(os.path.join(tmp, spec + ".manifest.json"))
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

    fx2 = _load_json(os.path.join(ROOT, "facts", "transitions.json"))
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
        fx3 = _load_json(os.path.join(ROOT, "facts", "transitions.json"))
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
        fx4 = _load_json(os.path.join(ROOT, "facts", "transitions.json"))
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
        ffull = _load_json(fxfull)
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
        fa = _load_json(fxattr)
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
        fx5 = _load_json(os.path.join(ROOT, "facts", "transitions.json"))
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
        tf = _load_json(tfx)
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
            rj = _load_json(rt)
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
        fx6 = _load_json(os.path.join(ROOT, "facts", "transitions.json"))
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

    print()
    if fails:
        print("切换表测试 FAILED (%d):" % len(fails))
        for f in fails:
            print("  - %s" % f)
        return 1
    print("切换表测试 PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
