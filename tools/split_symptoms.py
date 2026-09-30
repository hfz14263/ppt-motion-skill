#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一次性拆分器：把 symptoms.md 按「现象栏」切成多份。

与 split_pitfalls.py 同一套路，差别在**编号规则**：

  symptoms 与踩坑手册是**同一套 §编号**（它按现象重排踩坑，每节结尾写着
  "来源 com-pitfalls §n"）。所以这里**不能重新编号** —— 一重排，两边的 §n
  就对不上了。

  做法：拆完之后保留原有 §n（它们本来就是从 com-pitfalls 抄来的），
  每份文件内部按**出现顺序**排列（现象栏的顺序有意义），
  `symptoms.md` 退化成按现象分的目录页。

  这样做的好处：`symptoms.md → 「文件损坏」` 这类现成的指路语还能用，
  只要目录页保留同名的小节标题。
"""
import io
import os
import re
import sys
from collections import OrderedDict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF = os.path.join(ROOT, "reference")
SRC = os.path.join(REF, "symptoms.md")

# 现象栏 -> 文件。栏名保持原样（INDEX 里的指路语指向它）。
GROUPS = OrderedDict([
    ("symptom-file-corruption.md", {
        "cat": "文件打不开 / 报损坏",
        "duty": "PowerPoint 报 `0x80070570`「文件或目录已损坏」，且**不告诉你是哪个元素**。"
                "这一栏全部是**结构非法**，不是内容错。",
    }),
    ("symptom-silent-drop.md", {
        "cat": "XML 里有，PowerPoint 不认（静默丢弃）",
        "duty": "文件能开、不报错，但 PowerPoint **安静地丢掉**你的元素或属性。"
                "这一栏的共同点就是「不报错」，所以判据只能是「真开一遍看效果在不在」。",
    }),
    ("symptom-wrong-result.md", {
        "cat": "文件正常但结果不对",
        "duty": "文件合法、PowerPoint 采纳了，**但渲染出来的不是你要的**。"
                "这一栏只能靠**真渲染**发现，任何结构检查都查不出来。",
    }),
    ("symptom-roundtrip.md", {
        "cat": "往返后效果被吃掉",
        "duty": "PowerPoint 打开再保存一遍（round-trip），你的效果就变了或没了。"
                "注意：往返**既会救活**你写不对的东西，**也会改掉**你写对的东西。",
    }),
    ("symptom-tooling.md", {
        "cat": "工具与环境",
        "duty": "不是文档写错了，是**工具或环境**在坑你 —— 脚本、编码、导出。",
    }),
    ("symptom-evidence.md", {
        "cat": "证据与判断",
        # 这一组把原「值越界」小栏 + 四个以当事人语气写的坑合并 ——
        # 它们的共同点不是"哪坏了"，而是**"我以为我看见了，其实没看见"**：
        # 值越界是"以为规范就是上界"（实测上界与规范不同），
        # 其余四个是"抽帧证方向 / 单路指标判族 / 仪器选错 / 切换挤动画"。
        "cats": ["值越界 / 写错值",
                 "「我给的证据是静态的，看不到动效」",
                 "「我明明看着是左右开合，你写的是上下」",
                 "「挂了切换以后，入场动画像慢了半拍」",
                 "「动画明明设了，导出视频里却像没跑」"],
        "duty": "「我看到的证据说明不了这件事」这一栏 —— 静态证据证不了运动、"
                "单路指标会误判、仪器会选错、规范的值域不等于实测的值域。"
                "**关于「怎么才算是证据」的坑。**",
    }),
])

HEAD = """# 症状 · {cat}

> {duty}
>
> 本页按**现象**写，不按发现顺序。每节结尾标着它对应的踩坑编号
> （`来源 com-pitfalls §n`）—— 编号表见 [`pitfall-map.md`](pitfall-map.md)。
> 目录见 [`symptoms.md`](symptoms.md)；入口见 [`../INDEX.md`](../INDEX.md)。

---

"""


def main():
    text = io.open(SRC, encoding="utf-8").read()
    # 以 `## `（二级）切栏
    marks = list(re.finditer(r"^## (?!#)(.+)$", text, re.M))
    if not marks:
        raise SystemExit("找不到任何 `## ` 栏")
    intro = text[:marks[0].start()]
    cats = OrderedDict()
    for i, m in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(text)
        cats[m.group(1).strip()] = text[m.start():end].rstrip() + "\n"

    assigned = []
    for fname, spec in GROUPS.items():
        want = spec.get("cats") or [spec["cat"]]
        hit = []
        for w in want:
            if w not in cats:
                raise SystemExit("找不到栏 %r；实际有: %s" % (w, list(cats)))
            hit.append(w)
        assigned += hit
        body = "".join(cats[k] for k in hit)
        out = HEAD.format(cat=spec.get("cat") or want[0], duty=spec["duty"]) + body
        io.open(os.path.join(REF, fname), "w", encoding="utf-8",
                newline="\n").write(out.rstrip() + "\n")
        print("  %-34s %s" % (fname, " + ".join(hit)))

    leftover = [k for k in cats if k not in assigned]
    print()
    print("未分配的栏（会留在目录页，需人工决定）: %s" % leftover if leftover
          else "全部分配")
    return 0


if __name__ == "__main__":
    sys.exit(main())
