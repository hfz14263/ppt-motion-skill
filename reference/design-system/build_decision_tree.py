#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从 recipes.json 生成 reference/decision-tree.md 的三块内容。

数据流：recipes.json（唯一真相源）→ 标记区间 → decision-tree.md

    python reference/design-system/build_decision_tree.py            # 写盘
    python reference/design-system/build_decision_tree.py --check     # 只校验，不写

设计要点（每条都对应一次真实教训）：
- 章节正文【不进生成器】：手写段落留在 md 里，只生成标记区间。生成器重跑
  不吞正文（build_index.py 曾经吞掉 35 行，教训在这）。
- --check 报"文件不存在"和"内容不一致"两种不同情况，别混成一个 FAIL。
- 必填字段校验：缺 applicable_when / breaks_how 直接报错。因为没有失败信号的
  配方无法被验收，它不该进决策树。
- LF 写盘：Windows 上默认会写 CRLF，diff 会全文件变红。
"""
from __future__ import annotations

import argparse
import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC = ROOT / "reference" / "recipes.json"
DST = ROOT / "reference" / "decision-tree.md"

BLOCKS = [
    ("preflight", "<!-- gen:preflight:start -->", "<!-- gen:preflight:end -->"),
    ("tree", "<!-- gen:tree:start -->", "<!-- gen:tree:end -->"),
    ("table", "<!-- gen:table:start -->", "<!-- gen:table:end -->"),
    ("breaks", "<!-- gen:breaks:start -->", "<!-- gen:breaks:end -->"),
]

REQUIRED = ("id", "principle", "one_line", "applicable_when", "mechanism",
            "constraints", "breaks_how", "evidence")

# 决策树的分叉顺序 = 接手时能一眼看出的先后顺序。
# 顺序本身是设计：先问"主体是不是一张图"（最常见、最快判断），
# 再问"要不要立体"（要建模/两页，成本高），最后兜底。
ROUTES = [
    ("主体是【一张图】或【一个居中的物体】，想让它展开", ["move-the-mask-not-the-image"]),
    ("想把整图【切成几块】分别控制，或拼回来", ["shapes-reconstruct-one-image"]),
    ("要【立体感】但不想建模", ["flat-then-3d-is-two-pages"]),
    ("任何多页连贯叙事（默认起点）", ["motion-lives-between-pages"]),
    ("目标版本【不支持 Morph】", ["geometry-tween-by-hand"]),
    ("拿不准该动图还是该动形状", ["image-x-shape-is-the-design-layer"]),
]

PREFLIGHT = """**0. 先扫一遍 shape，拿到真实 id —— 不要手写。**

```bash
python scripts/motion.py inspect --pptx deck.pptx
```

**为什么放在最前面**：手写形状 id 列表是本项目最贵的错误来源。
deck 改一次版式，id 就全体重排，spec 于是指向了**别的形状** ——
不是报错，是**静默错位**（曾经整张卡没动画，而且正好差一位）。

**1. 判断素材形态**（打开 deck 一眼能看出的东西，这是唯一已知量）

| 看到什么 | 往哪走 |
| --- | --- |
| 一张照片 / 一个居中物体 | 「主体是张图」那一支 |
| 一张图想分块控制 | 「切成几块」那一支 |
| 一个平面对象（照片、卡片、图表） | 「立体感」那一支 |
| 多页连续内容 | 「多页连贯叙事」那一支 |

**2. 只动一个变量**

这是最容易违反的一条：`motion-lives-between-pages` 要求两页形状
name/id 完全一致，**只改要动的那几个数值**。多改一个，两个页就配不上对，
Morph 退化成淡入 —— 而你看到的是"效果没生效"，不是"配对失败"。

**3. 拿 §3 的失败信号验收，不要自我感觉**

做完对着 §3 逐条看。长成那样就是错了。
"""


def load() -> list[dict]:
    data = json.loads(SRC.read_text(encoding="utf-8"))
    entries = data["entries"]
    # 必填校验：缺字段的配方不许进树
    problems = []
    for e in entries:
        for f in REQUIRED:
            if not e.get(f):
                problems.append(f"{e.get('id', '<无 id>')}: 缺 {f}")
    if problems:
        sys.exit("recipes.json 字段不全：\n  " + "\n  ".join(problems))
    # 路由引用的 id 必须存在，否则树里会指向不存在的配方
    known = {e["id"] for e in entries}
    for label, ids in ROUTES:
        for i in ids:
            if i not in known:
                sys.exit(f"决策树引用了不存在的配方 id: {i}")
    # 反向：有配方没进任何一支 = 孤儿，读者永远看不到它
    routed = {i for _, ids in ROUTES for i in ids}
    orphans = sorted(known - routed)
    if orphans:
        sys.exit("这些配方没被任何一条决策路线引用（孤儿）：\n  " + "\n  ".join(orphans))
    return entries


def by_id(entries: list[dict]) -> dict[str, dict]:
    return {e["id"]: e for e in entries}


def render_preflight(_: list[dict]) -> str:
    return PREFLIGHT


def render_tree(entries: list[dict]) -> str:
    idx = by_id(entries)
    lines = ["```"]
    lines.append("拿到一份 pptx")
    for label, ids in ROUTES:
        lines.append("│")
        lines.append(f"├─ {label}")
        for n, i in enumerate(ids):
            e = idx[i]
            brk = e["breaks_how"][0] if e["breaks_how"] else ""
            # 同一支内多条配方时，只有最后一条收口，其余挂横杠
            tee = "└─" if n == len(ids) - 1 else "├─"
            cont = "│ " if n == len(ids) - 1 else "│ "
            lines.append(f"│  {tee} [{e['principle']}]")
            lines.append(f"│  {cont}机制: {e['mechanism'].split('；')[0]}")
            lines.append(f"│  {cont}⚠ 若出现: {brk}")
    lines.append("│")
    lines.append("└─ 读不出想要的读法 → 回 §3 失败信号速查，对症状定位")
    lines.append("")
    lines.append("**走向由素材形态决定，不由「想做什么效果」决定。**")
    lines.append("效果无穷，按效果分叉必然退化成效果清单；素材形态只有几个分支，")
    lines.append("而素材是接手时唯一【确定已知】的东西 —— 效果是猜的。")
    lines.append("```")
    return "\n".join(lines)


def render_table(entries: list[dict]) -> str:
    head = ("| # | 原理 | 什么时候用 | 一句话 |\n| --- | --- | --- | --- |\n")
    rows = []
    for n, e in enumerate(entries, 1):
        aw = "<br>".join(a.replace("【", "**").replace("】", "**") for a in e["applicable_when"])
        rows.append(f"| {n} | **{e['principle']}** | {aw} | {e['one_line']} |")
    return head + "\n".join(rows)


def render_breaks(entries: list[dict]) -> str:
    out = ["**这些不是建议，是验收标准。** 做出来长成下面这样，就是选错配方了。", ""]
    out.append("| 看到什么 | 真实原因 | 怎么改 |")
    out.append("| --- | --- | --- |")
    for e in entries:
        for b in e["breaks_how"]:
            cause = e["constraints"][0] if e["constraints"] else e["principle"]
            fix = e["mechanism"].split("；")[0]
            b = b.replace("→", "→").replace("|", "\\|")
            out.append(f"| {b} | {cause} | {fix} |")
    out.append("")
    out.append("> 表格由 `recipes.json` 生成。加配方请改 JSON —— 手改这里会在下次重跑时丢失。")
    return "\n".join(out)


RENDERERS = {
    "preflight": render_preflight,
    "tree": render_tree,
    "table": render_table,
    "breaks": render_breaks,
}


def replace_block(text: str, start: str, end: str, body: str, label: str) -> str:
    """替换标记区间。标记缺失时明确报错，而不是静默不写 —— 那会让生成器看起来成功。"""
    if start not in text or end not in text:
        sys.exit(f"[{label}] decision-tree.md 里找不到标记 {start} / {end}；"
                 f"标记被手删了？生成器只管区间，不管正文。")
    i, j = text.index(start), text.index(end)
    return text[: i + len(start)] + "\n" + body + "\n" + text[j:]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="只校验，不写盘")
    a = ap.parse_args()

    entries = load()
    if not DST.exists():
        if a.check:
            print("FAIL: decision-tree.md 不存在")
            return 1
        sys.exit(f"decision-tree.md 不存在: {DST}")
    text = DST.read_text(encoding="utf-8")
    want = text
    for label, s, e in BLOCKS:
        want = replace_block(want, s, e, RENDERERS[label](entries), label)

    if a.check:
        if not DST.exists():
            print("FAIL: decision-tree.md 不存在"); return 1
        cur = DST.read_text(encoding="utf-8")
        if cur != want:
            # 逐块报出是哪一块不一致，否则只能靠 diff 猜
            for label, s, e in BLOCKS:
                def seg(t):
                    return t[t.index(s) + len(s): t.index(e)].strip() if s in t and e in t else "<缺>"
                if seg(cur) != seg(want):
                    print(f"FAIL: [{label}] 区间与 recipes.json 不一致 —— 跑生成器刷新")
            return 1
        print("OK: decision-tree.md 与 recipes.json 一致")
        return 0

    if text == want:
        print("NO-CHANGE: 已是最新")
        return 0
    DST.write_text(want, encoding="utf-8", newline="\n")
    print(f"written: {DST.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
