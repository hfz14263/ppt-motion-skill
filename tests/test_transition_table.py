#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""切换实测表的回归测试。

为什么需要这份测试：scripts/transition_reference.json 是量出来的，而
scripts/motion.py 直接读它。两者一旦漂移，症状不是报错，而是**写出一个
PowerPoint 会悄悄改掉的块** —— 看起来一切正常，文件就是不对。所以这里把
"表还成立"和"motion.py 还认这张表"钉住。

不测的东西：本测试**不**打开 PowerPoint。真机验证在 tests/smoke.py。
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
                        "dx", "dy", "tx", "ty"}
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
            drift = []
            for num in re.findall(r"-?\d+\.\d+", ctext_n):
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
    bad_dp = [s for s, _d, _w in B.DIR_PROBE
              if s not in {h["spec"] for h in B.HYPOTHESES}]
    check("DIR_PROBE 引用的 spec 都真实存在", not bad_dp, str(bad_dp))
    check("DIR_PROBE 每条都写了理由（不能只列组合）",
          all(w.strip() for _s, _d, w in B.DIR_PROBE))
    # 范围纪律：这一组是为了补 §三 的缺口，不是把 48 个效果全测一遍。
    # 失控会让"补缺口"变成"重做形态层"，那就不该叫补缺口了。
    check("DIR_PROBE 范围收敛（≤12 份，只覆盖选择层用到的组合）",
          0 < len(B.DIR_PROBE) <= 12, "len=%d" % len(B.DIR_PROBE))
    # 每个 dir 取值都应只出现一次（重复 = 白渲染一份）
    pairs = [(s, d) for s, d, _w in B.DIR_PROBE]
    check("DIR_PROBE 无重复 (spec,dir) 组合",
          len(pairs) == len(set(pairs)),
          "重复: %s" % [p for p in set(pairs) if pairs.count(p) > 1])
    # 选择层真正用到的四个效果必须在名单里（否则缺口没补到点上）
    need = {"push", "wipe", "cover", "uncover"}
    check("DIR_PROBE 覆盖选择层实际推荐的四个效果",
          need <= {s for s, _d, _w in B.DIR_PROBE},
          "缺 %s" % (need - {s for s, _d, _w in B.DIR_PROBE}))

    # 缺口闭合的**文档证据**：§三 不能再挂着"没有实测证据"。
    # 这条是负向检查 —— 防止有人回退文档、把已闭合的缺口又写回"未知"。
    if os.path.exists(choice_doc):
        ctext2 = io.open(choice_doc, encoding="utf-8").read()
        check("§三 不再自称『没有实测证据』（缺口已闭合）",
              "本文没有实测证据" not in ctext2,
              "文档里还留着『没有实测证据』")
        ctext2n = ctext2.replace("\u2212", "-")
        for token, why in (("镜", "§三 要写出镜像结论"),
                           ("-0.99", "wipe 的相关系数要可核对"),
                           ("1.004", "push 的累积位移要可核对"),
                           ("1.000", "curtains 对称性要可核对")):
            check("§三 写着 %s（%s）" % (token, why), token in ctext2n)
        # 两路判据的方法论必须留在文档里（不然下次又会只用一路）
        check("§三 说明了为何 push 需要另一路判据",
              "变化重心根本不动" in ctext2 or "重心根本不动" in ctext2)

    # facts 必须记着这两条（否则"缺口已闭合"这件事会随时间丢失）
    if os.path.exists(os.path.join(ROOT, "facts", "transitions.json")):
        fx4 = _load_json(os.path.join(ROOT, "facts", "transitions.json"))
        ids4 = {r["id"] for r in fx4["rules"]}
        for want in ("mirroring-needs-two-instruments-not-one",
                     "dir-mirrors-for-the-four-recommended-effects"):
            check("facts 记着 %s" % want, want in ids4)

    # 生成的 deck 文件名必须能直接配对（分析器靠文件名分 l/r）
    import glob as _glob
    probe_dir = os.path.join(ROOT, ".workbuddy", "_dirprobe_smoke")
    try:
        if not os.path.isdir(probe_dir):
            B.cmd_dirdeck(probe_dir)
        made = sorted(os.path.basename(p) for p in
                      _glob.glob(os.path.join(probe_dir, "*.pptx")))
        check("dirdeck 生成的文件名可配对（<spec>_dir_<l|r>）",
              all(re.fullmatch(r"\w+_dir_[lrtb]\.pptx", n) for n in made)
              and len(made) == len(B.DIR_PROBE),
              "%d 个: %s" % (len(made), made[:4]))
        # 生成器必须真的把 dir 写进 XML —— 文件名对但属性没写是最坏的假象
        import zipfile as _zip
        bad_xml = []
        for s, d, _w in B.DIR_PROBE:
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
