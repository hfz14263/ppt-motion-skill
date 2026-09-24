# -*- coding: utf-8 -*-
"""校验"同图双版本"版式：淡化整图铺底 + 原色塞进形状当前景。

这个版式来自抖音教学视频的复现。它有两个很容易翻车的地方，正是本脚本要挡住的：
  1) 底图和前景必须是同一张照片的两个版本 —— 否则就不是这个版式；
  2) 前景不能被拉伸 —— 视频里专门给了失败对照："拉伸人物就长胖"。
     用 a:srcRect 裁剪式放大才是不变形的做法，改 fillRect/直接拉伸都会变形。

判定原则（对齐 references/review-checklist.md §5）：
  * 每条"通过"尽量给两条独立判据；判据互相打架时报 WARN 而不是装作通过；
  * 检测器必须先在已知样本上证明它能看见要找的问题（正负对照见 tests/test_dual_photo.py）。

用法:
    python verify_dual_photo.py FILE [--tol 0.08] [--json] [--expect-class dual-photo]
退出码: 0 通过 / 1 检出错误 / 2 用法或文件问题
"""
import argparse
import hashlib
import io
import json
import math
import os
import posixpath
import sys
import zipfile

import lxml.etree as ET

try:
    from PIL import Image
except Exception:                                   # 可选依赖
    Image = None

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

EMU_PER_PT = 12700.0
AR_TOL = 0.02            # 判据A：两张源图长宽比一致容差
CORR_MIN = 0.80          # 判据B：低频结构相关系数下限（由 tests 实测标定）


# ---------------------------------------------------------------- 图片尺寸
def image_size(data):
    """不依赖 PIL 地读出 PNG/JPEG/BMP/GIF 的像素尺寸。"""
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return (int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big"))
    if data[:2] == b"\xff\xd8":
        i = 2
        while i < len(data) - 9:
            if data[i] != 0xFF:
                i += 1
                continue
            m = data[i + 1]
            if m in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                return (int.from_bytes(data[i + 7:i + 9], "big"),
                        int.from_bytes(data[i + 5:i + 7], "big"))
            if m in (0xD8, 0xD9) or 0xD0 <= m <= 0xD7:
                i += 2
                continue
            i += 2 + int.from_bytes(data[i + 2:i + 4], "big")
        return None
    if data[:2] == b"BM":
        return (int.from_bytes(data[18:22], "little"), int.from_bytes(data[22:26], "little"))
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return (int.from_bytes(data[6:8], "little"), int.from_bytes(data[8:10], "little"))
    return None


def _flat_gray(data, size=32):
    """转 32x32 灰度（RGBA 先压到白底，避免透明区乱入）。"""
    im = Image.open(io.BytesIO(data))
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bg.paste(im, mask=im.split()[-1])
        im = bg
    return im.convert("L").resize((size, size), Image.BILINEAR)


def lowfreq_corr(d1, d2, size=32):
    """低频结构相关（Pearson）。淡化/虚化后低频结构保留，换图则明显下降。"""
    if Image is None:
        return None
    # tobytes() 在 L 模式下每像素 1 字节；getdata() 已在 Pillow 路线图里被移除
    a = list(_flat_gray(d1, size).tobytes())
    b = list(_flat_gray(d2, size).tobytes())
    ma, mb = sum(a) / len(a), sum(b) / len(b)
    num = sum((x - ma) * (y - mb) for x, y in zip(a, b))
    da = math.sqrt(sum((x - ma) ** 2 for x in a))
    db = math.sqrt(sum((y - mb) ** 2 for y in b))
    return num / (da * db) if da and db else 0.0


# ---------------------------------------------------------------- XML 取值
def _frac(el, key):
    """OOXML 里这类值是 percent*1000，100000 = 100%。"""
    v = el.get(key)
    return (int(v) / 100000.0) if v else 0.0


def _rect(el):
    """返回 (kw, kh)：源图/目标矩形保留比例。"""
    return (1.0 - _frac(el, "l") - _frac(el, "r"),
            1.0 - _frac(el, "t") - _frac(el, "b"))


def _geom(spPr):
    xfrm = spPr.find("{%s}xfrm" % NS_A)
    if xfrm is None:
        return None
    off = xfrm.find("{%s}off" % NS_A)
    ext = xfrm.find("{%s}ext" % NS_A)
    if off is None or ext is None:
        return None
    return {
        "x": int(off.get("x")) / EMU_PER_PT, "y": int(off.get("y")) / EMU_PER_PT,
        "w": int(ext.get("cx")) / EMU_PER_PT, "h": int(ext.get("cy")) / EMU_PER_PT,
        "flipH": xfrm.get("flipH") == "1", "flipV": xfrm.get("flipV") == "1",
    }


def _blip(spPr, host=None):
    """返回 (rId, srcRect元素, fillRect元素)。
    形状(p:sp)的 blipFill 在 spPr 里；图片(p:pic)的 p:blipFill 是 p:pic 的直接子节点。"""
    bf = spPr.find("{%s}blipFill" % NS_A) if spPr is not None else None
    if bf is None and host is not None:
        bf = host.find("{%s}blipFill" % NS_P)
    if bf is None:
        return None
    blip = bf.find("{%s}blip" % NS_A)
    if blip is None:
        return None
    return (blip.get("{%s}embed" % NS_R),
            bf.find("{%s}srcRect" % NS_A),
            bf.find("{%s}stretch/{%s}fillRect" % (NS_A, NS_A)))


def _media(z, rels, rId):
    target = rels.get(rId)
    if not target:
        return None, None
    # zip 内部路径恒为正斜杠，不能用 os.path（Windows 会给反斜杠）
    name = target if target.startswith("ppt/") else posixpath.normpath(
        posixpath.join("ppt/slides", target))
    try:
        return name, z.read(name)
    except KeyError:
        return None, None


def _rels(z, slide_name):
    base = slide_name.rsplit("/", 1)[-1]
    rp = slide_name.rsplit("/", 1)[0] + "/_rels/" + base + ".rels"
    out = {}
    try:
        root = ET.fromstring(z.read(rp))
    except KeyError:
        return out
    for rel in root:
        out[rel.get("Id")] = rel.get("Target")
    return out


def distortion(shape_w, shape_h, img_w, img_h, src_rect, fill_rect):
    """失真倍数 = (目标矩形 AR) / (源图有效区 AR)。1.0 = 不变形；>1 = 横向拉胖。"""
    kw, kh = _rect(src_rect) if src_rect is not None else (1.0, 1.0)
    dw, dh = _rect(fill_rect) if fill_rect is not None else (1.0, 1.0)
    dst_ar = (shape_w * dw) / (shape_h * dh)
    src_ar = (img_w * kw) / (img_h * kh)
    if not src_ar:
        return None
    return dst_ar / src_ar


# ---------------------------------------------------------------- 主检查
def check(pptx, tol=0.08, corr_min=CORR_MIN):
    res = {"file": pptx, "errors": [], "warnings": [], "slides": []}
    if not os.path.isfile(pptx):
        res["errors"].append("文件不存在: %s" % pptx)
        return res
    try:
        z = zipfile.ZipFile(pptx)
    except Exception as exc:
        res["errors"].append("打不开 pptx: %s" % exc)
        return res

    try:
        pres = ET.fromstring(z.read("ppt/presentation.xml"))
        sz = pres.find("{%s}sldSz" % NS_P)
        sw = int(sz.get("cx")) / EMU_PER_PT
        sh = int(sz.get("cy")) / EMU_PER_PT
    except Exception:
        sw, sh = 960.0, 540.0
    res["slide_size_pt"] = [sw, sh]

    slides = [n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    slides.sort(key=lambda n: int("".join(c for c in os.path.basename(n) if c.isdigit()) or 0))

    for name in slides:
        root = ET.fromstring(z.read(name))
        rels = _rels(z, name)
        info = {"slide": os.path.basename(name), "errors": [], "warnings": [],
                "background": None, "foreground": []}

        # ---- 背景候选：<p:pic> 且铺满页面
        for pic in root.iter("{%s}pic" % NS_P):
            spPr = pic.find("{%s}spPr" % NS_P)
            g = _geom(spPr) if spPr is not None else None
            b = _blip(spPr, host=pic)
            if not g or not b:
                continue
            mname, mdata = _media(z, rels, b[0])
            size = image_size(mdata) if mdata else None
            if Image is not None and size is None and mdata:
                size = Image.open(io.BytesIO(mdata)).size
            area = (g["w"] * g["h"]) / (sw * sh)
            if area >= 0.90:
                bg = {"kind": "picture", "geometry": g, "media": mname, "size": size}
                bg["ratio"] = distortion(g["w"], g["h"], size[0], size[1], b[1], b[2]) if size else None
                info["background"] = bg

        # ---- 前景候选：<p:sp> 里带 blipFill（形状图片填充）
        for sp in root.iter("{%s}sp" % NS_P):
            spPr = sp.find("{%s}spPr" % NS_P)
            if spPr is None:
                continue
            bf = spPr.find("{%s}blipFill" % NS_A)
            if bf is None:
                continue
            g = _geom(spPr)
            b = _blip(spPr)
            prst = spPr.find("{%s}prstGeom" % NS_A)
            mname, mdata = _media(z, rels, b[0])
            size = image_size(mdata) if mdata else None
            if Image is not None and size is None and mdata:
                size = Image.open(io.BytesIO(mdata)).size
            fg = {
                "prst": prst.get("prst") if prst is not None else None,
                "geometry": g, "media": mname, "size": size,
                "flipH": g["flipH"] if g else False, "flipV": g["flipV"] if g else False,
            }
            fg["ratio"] = distortion(g["w"], g["h"], size[0], size[1], b[1], b[2]) if (g and size) else None
            info["foreground"].append(fg)

        # ---- 版式判定
        if info["background"] and info["foreground"]:
            info["class"] = "dual-photo"
        elif info["foreground"]:
            info["class"] = "foreground-only"
        elif info["background"]:
            info["class"] = "background-only"
        else:
            info["class"] = "none"

        if info["class"] != "dual-photo":
            res["slides"].append(info)
            continue

        # ---- 判据：底图与前景是否同一张照片的两个版本
        bg = info["background"]
        bg_data = z.read(bg["media"]) if bg.get("media") else None
        for fg in info["foreground"]:
            fg_data = z.read(fg["media"]) if fg.get("media") else None
            same = {"md5_equal": False, "ar_equal": None, "corr": None,
                    "verdict": "unknown", "reasons": []}
            if bg_data and fg_data:
                same["md5_equal"] = (hashlib.md5(bg_data).hexdigest()
                                     == hashlib.md5(fg_data).hexdigest())
                if bg.get("size") and fg.get("size"):
                    ar1 = bg["size"][0] / bg["size"][1]
                    ar2 = fg["size"][0] / fg["size"][1]
                    same["ar_equal"] = abs(ar1 - ar2) <= AR_TOL * max(ar1, ar2)
                    same["ar"] = [ar1, ar2]
                corr = lowfreq_corr(bg_data, fg_data) if bg_data else None
                same["corr"] = corr
                same["corr_ok"] = None if corr is None else corr >= corr_min
                if same["md5_equal"]:
                    same["verdict"] = "same"
                    same["reasons"].append("两份媒体字节完全相同")
                elif same["ar_equal"] is not None and same["corr_ok"] is not None:
                    if same["ar_equal"] and same["corr_ok"]:
                        same["verdict"] = "same"
                        same["reasons"].append("判据A 长宽比一致(%.3f vs %.3f)" % (same["ar"][0], same["ar"][1]))
                        same["reasons"].append("判据B 低频结构相关 %.3f ≥ %.2f" % (corr, corr_min))
                    elif not same["ar_equal"] and not same["corr_ok"]:
                        same["verdict"] = "different"
                        same["reasons"].append("判据A 长宽比不一致(%.3f vs %.3f)" % (same["ar"][0], same["ar"][1]))
                        same["reasons"].append("判据B 低频结构相关 %.3f < %.2f" % (corr, corr_min))
                        info["errors"].append(
                            "%s: 底图与前景不是同一张照片 —— 不是'同图双版本'版式" % info["slide"])
                    else:
                        same["verdict"] = "conflict"
                        info["warnings"].append(
                            "%s: 两条判据打架(长宽比=%s, 相关=%.3f) —— 需人工看图确认" %
                            (info["slide"], same["ar_equal"], corr))
                else:
                    info["warnings"].append(
                        "%s: 同图判据不足（缺 PIL 或读不到尺寸），无法判定" % info["slide"])
            fg["same_source"] = same

            # ---- 判据：前景是否被拉伸
            r = fg.get("ratio")
            if r is None:
                info["warnings"].append("%s: 读不到源图尺寸，跳过变形检查" % info["slide"])
            elif abs(r - 1.0) > tol:
                how = "横向拉胖" if r > 1 else "纵向拉长"
                info["errors"].append(
                    "%s: 前景被%s %.1f%%（失真倍数 %.3f）—— 对应失败症状'拉伸人物就长胖'；"
                    "应改用 a:srcRect 裁剪式放大" % (info["slide"], how, abs(r - 1) * 100, r))
            fg["distortion_ok"] = (r is not None and abs(r - 1.0) <= tol)

            # ---- 该版式不应出现翻转形状（过程演示的翻转副本成片前要删）
            if fg.get("flipH") or fg.get("flipV"):
                info["warnings"].append(
                    "%s: 前景形状带翻转(flipH=%s/flipV=%s) —— 疑似遗留的过程演示形状" %
                    (info["slide"], fg["flipH"], fg["flipV"]))

        # ---- 背景自身也不该被压扁
        br = bg.get("ratio")
        if br is not None and abs(br - 1.0) > tol:
            info["errors"].append("%s: 铺底图被拉伸 %.1f%%（失真倍数 %.3f）" %
                                  (info["slide"], abs(br - 1) * 100, br))

        prsts = [f["prst"] for f in info["foreground"]]
        if len(prsts) != len(set(prsts)):
            info["warnings"].append("%s: 出现重复形状 %s —— 确认是否遗留了演示层" %
                                    (info["slide"], prsts))

        res["errors"].extend(info["errors"])
        res["warnings"].extend(info["warnings"])
        res["slides"].append(info)
    return res


def main():
    ap = argparse.ArgumentParser(description="校验'同图双版本'版式")
    ap.add_argument("pptx")
    ap.add_argument("--tol", type=float, default=0.08, help="失真倍数容差，默认 0.08")
    ap.add_argument("--corr-min", type=float, default=CORR_MIN, help="低频相关下限")
    ap.add_argument("--expect-class", default=None, help="期望的版式类别，不一致则报错")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()

    res = check(a.pptx, tol=a.tol, corr_min=a.corr_min)
    if a.expect_class:
        got = res["slides"][0]["class"] if res["slides"] else "none"
        if got != a.expect_class:
            res["errors"].append("版式类别为 %s，期望 %s" % (got, a.expect_class))

    if a.json:
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return 1 if res["errors"] else 0

    print("文件: %s" % res["file"])
    for s in res["slides"]:
        print("  %s: class=%s" % (s["slide"], s["class"]))
        for fg in s["foreground"]:
            print("    前景 %s ratio=%.4f 同图=%s" %
                  (fg["prst"], fg["ratio"] if fg["ratio"] else float("nan"),
                   fg.get("same_source", {}).get("verdict")))
            for r in fg.get("same_source", {}).get("reasons", []):
                print("      - %s" % r)
    for w in res["warnings"]:
        print("  [WARN] %s" % w)
    for e in res["errors"]:
        print("  [ERROR] %s" % e)
    print("结论: %s" % ("不通过" if res["errors"] else "通过"))
    return 1 if res["errors"] else 0


if __name__ == "__main__":
    sys.exit(main())
