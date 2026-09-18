"""Check for duplicate singletons in a deck, and fail loudly.

Adds a structural gate for the one bug class that has corrupted every file this project
has produced: an element inserted where its sequence type already held one. PowerPoint
reports only "file or directory is corrupt" with no indication of the element, so this
finds it from the XML before anyone opens the file.

Usage:
    python scripts/verify_singletons.py --pptx out.pptx
    python scripts/verify_singletons.py --pptx out.pptx --json
"""
import argparse
import json
import re
import sys
import zipfile

HERE = __import__("os").path.dirname(__import__("os").path.abspath(__file__))
sys.path.insert(0, HERE)
import motion                                             # noqa: E402

# Parts and the parent element within which each tag must be unique.
# NOTE ON TAGS: motion.element_spans prefixes the namespace itself -- it searches for
# "<p:" + tag -- so every caller passes the BARE name ("spPr"), never "p:spPr". Passing
# the prefix produced "<p:p:spPr", which matches nothing, and a validator that silently
# matches nothing reports OK on a corrupt file. All 12 existing call sites use the bare
# form; these follow it.
CHECKS = [
    ("slides", "spPr", ("a:effectLst", "a:effectDag", "a:ln", "a:blipFill",
                        "a:solidFill", "a:noFill", "a:gradFill", "a:pattFill",
                        "a:grpFill", "a:scene3d", "a:sp3d", "a:xfrm")),
    # p:transition is checked as a DIRECT child of p:sld only. Counting it with a global
    # regex flagged a correct morph as a duplicate, because morph carries two
    # transitions by design -- one inside mc:Choice and one inside mc:Fallback.
    ("slides", "sld", ("p:cSld", "p:clrMapOvr", "p:timing", "p:transition")),
    ("layouts", "spPr", ("a:effectLst", "a:ln", "a:blipFill", "a:solidFill",
                         "a:noFill", "a:xfrm")),
    ("layouts", "sldLayout", ("p:cSld", "p:clrMapOvr", "p:timing", "p:transition")),
]


def check(pptx):
    findings = []
    with zipfile.ZipFile(pptx) as z:
        names = z.namelist()
        for kind, parent, tags in CHECKS:
            pref = {"slides": "ppt/slides/slide", "layouts": "ppt/slideLayouts/slideLayout",
                    "masters": "ppt/slideMasters/slideMaster"}[kind]
            for n in sorted(x for x in names
                            if x.startswith(pref) and x.endswith(".xml")):
                xml = z.read(n).decode("utf-8", "replace")
                for tag, count in motion.find_duplicate_singletons(xml, parent, tags):
                    findings.append({"part": n, "parent": parent, "tag": tag,
                                     "count": count})
    return findings


def main(argv=None):
    ap = argparse.ArgumentParser(description="重复单例元素检查")
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--json", action="store_true")
    ns = ap.parse_args(argv)

    f = check(ns.pptx)
    if ns.json:
        print(json.dumps({"findings": f, "ok": not f}, ensure_ascii=False, indent=1))
        return 1 if f else 0

    if not f:
        print("OK  没有重复的单例元素（形状属性 / 幻灯片子元素 / 切换）")
        print("注意：这只证明「没有重复」，不证明 PowerPoint 肯打开 —— "
              "值域越界、命名空间未声明同样会让文件损坏。")
        return 0

    print("!! 发现重复的单例元素 —— 这类文件 PowerPoint 会判为「已损坏」，"
          "且不会告诉你是哪个元素")
    for x in f:
        print("  %-42s <%s> x%d  （在 %s 内）"
              % (x["part"], x["tag"], x["count"], x["parent"]))
    print("\n修法：用 motion.set_singleton() 替换而不是插入。")
    return 1


if __name__ == "__main__":
    sys.exit(main())
