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

    # ---- 9. 形态层文档 + facts 一致 --------------------------------------
    shapes_doc = os.path.join(ROOT, "reference", "transition-shapes.md")
    check("形态层文档存在", os.path.exists(shapes_doc))
    if os.path.exists(shapes_doc):
        text = io.open(shapes_doc, encoding="utf-8").read()
        for token, why in (("整页平移", "六族之一：push/pan 靠位移指标认"),
                           ("方向揭示", "六族之一：wipe/cover 靠漂移指标认"),
                           ("中心扩散", "六族之一：split/ripple"),
                           ("band_travel", "deck2 用来分开 fade 与 wipe 的判据"),
                           ("morph", "形态层必须说明 morph 不适用")):
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
                for arrow in ("l→r", "r→l", "t→b", "b→t"):
                    if arrow in c["claim"] and arrow not in text:
                        drift.append("%s 的 %s 文档里没有" % (c["spec"], arrow))
            check("证据 deck 与形态文档无漂移（数字/方向同源）",
                  not drift, "; ".join(drift[:5]))

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
