#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 3a 的【文字判定】与【渲染实物】放进同一份 deck，供人眼复核。

为什么需要这个：形态层（reference/transition-shapes.md）里有两类东西 ——

  * 数字（band_travel = -0.747、coverage = 0.957、span = 15 帧）—— 这些有
    合成标定样本 + 逐帧地面真值核对，可信。
  * 把数字翻译成中文（"前沿从右向左"、"整页滑走"）—— **这一层没有任何机器
    验证**。它是我写的，而它恰恰是读者唯一会用的东西。

两者必须能被分开检查，否则"数字对"会被误当成"描述对"。

这份 deck 的做法：每族一页，左边写我的判定（可被质疑的断言），右边贴
【真实渲染视频抽出的帧序列】（地面真值）。人眼扫一遍就知道哪边错了。

用法：
  python scripts/build_shape_evidence.py <video_dir> <out.pptx>

注意：这只回答"我的文字与渲染是否一致"，**不回答"这个效果好不好看"**，
也不回答"该不该用"（那是 3b 的事）。
"""
import io
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)


# 每族挑代表 + 写死"我的判定"（就是要被质疑的那些话）。
# 这些字符串必须与 reference/transition-shapes.md 里的一致 —— 不一致就是
# 文档与证据脱节，测试会抓（见 tests/test_transition_table.py 第 10 组）。
CLAIMS = [
    {"spec": "cut", "family": "硬切",
     "claim": "瞬间替换，无过程。活跃帧 = 0。",
     "pred": "抽帧应看到：一帧之内整幅换掉，中间没有过渡画面。"},
    {"spec": "push", "family": "整页平移",
     "claim": "整页滑入，默认方向【右→左】。变化重心全程停在 ~0.5，"
              "但逐帧相位相关给出稳定 −4~−5px/帧。",
     "pred": "抽帧应看到：整个画面内容整体横向平移，不是「某条边在动」。"},
    {"spec": "pan", "family": "整页平移",
     "claim": "同 push 类，位移方向【右→左】。",
     "pred": "抽帧应看到：整页连续平移。"},
    {"spec": "wipe", "family": "方向揭示",
     "claim": "一条前沿从【右】向【左】扫过。实测差异重心 0.87 → 0.425"
              "（0=最左，1=最右）。",
     "pred": "抽帧应看到：画面被一条推进的边界分割，亮区自右向左扩大。"},
    {"spec": "cover", "family": "方向揭示",
     "claim": "覆盖式推进，方向【右→左】。实测重心 1.0 → 0.51。",
     "pred": "抽帧应看到：新画面从右侧压过来。"},
    {"spec": "uncover", "family": "方向揭示",
     "claim": "揭开式，方向【右→左】。实测重心 0.47 → 0.04 —— "
              "起点在中间，终点在最左。",
     "pred": "抽帧应看到：新画面从中偏右揭开，向左侧扩张。"},
    {"spec": "fade", "family": "中心扩散",
     "claim": "整幅均匀淡变，无方向。径向剖面 [1.0, 0.15, ...] 与 wipe 雷同，"
              "靠第二探测的 band_travel 才分开（fade = 0.015）。",
     "pred": "抽帧应看到：没有可辨认的移动边界，整幅亮度一起变。"},
    {"spec": "split", "family": "中心扩散",
     "claim": "从中心向两侧打开。band_travel = −0.015（几乎不漂）。",
     "pred": "抽帧应看到：一条缝从中间张开，向左右两侧扩张。"},
    {"spec": "checkerboard", "family": "方向揭示",
     "claim": "棋盘格逐格翻转，判为【l→r】。",
     "pred": "抽帧应看到：格子分批变化，批次顺序自左向右。"},
    {"spec": "curtains", "family": "方向揭示",
     "claim": "帘式，判为【t→b】（自上而下）。",
     "pred": "抽帧应看到：上下开合式的揭示。"},
    {"spec": "honeycomb", "family": "方向揭示",
     "claim": "归入方向揭示族，但【本层给不出可靠方向】—— "
              "位移指标是噪声级，真实形态是内部纹理在变。这里标注而非硬填。",
     "pred": "抽帧应看到：蜂巢式局部变化，但它是否算「一条前沿」值得怀疑。"},
    {"spec": "flash", "family": "均匀淡变",
     "claim": "整幅同时变化、无方向。靠「覆盖率高 + 漂移≈0」与 fade 分开。",
     "pred": "抽帧应看到：整幅亮度一起闪，没有任何推进边界。"},
    {"spec": "random", "family": "随机",
     "claim": "⚠️ 本层【测不了】：每次播放随机挑一个效果，没有稳定形态。"
              "两个 mp4 字节数相同（186629）但 md5 不同即为证。",
     "pred": "这一页没有可比的形态 —— 你每次放映看到的可能都不一样。"},
]


def grab_frames(video, n=6, boundary_frame=60):
    """从过渡区间里均匀抽 n 帧。

    过渡发生在第 60 帧附近（每页停留 2s @30fps）。取 60-2 到 60+40 这一段，
    刚好覆盖过渡本身。
    """
    import cv2

    cap = cv2.VideoCapture(video)
    frames = []
    while True:
        ok, f = cap.read()
        if not ok:
            break
        frames.append(f)
    cap.release()
    if not frames:
        return []
    lo = max(0, boundary_frame - 3)
    hi = min(len(frames), boundary_frame + 45)
    seg = frames[lo:hi]
    if len(seg) <= n:
        return seg
    idx = [int(round(i * (len(seg) - 1) / (n - 1))) for i in range(n)]
    return [seg[i] for i in idx]


def main(video_dir, out_pptx, claims_path=None):
    from pptx import Presentation
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    claims = CLAIMS
    if claims_path:
        claims = json.load(io.open(claims_path, encoding="utf-8-sig"))

    prs = Presentation()
    prs.slide_width = Emu(12192000)      # 16:9 宽屏，横向放左右两栏
    prs.slide_height = Emu(6858000)
    blank = prs.slide_layouts[6]
    W, H = prs.slide_width, prs.slide_height

    def add_page(title, claim, pred, imgs, note=""):
        s = prs.slides.add_slide(blank)
        # 标题
        tb = s.shapes.add_textbox(Emu(300000), Emu(180000),
                                  W - Emu(600000), Emu(700000))
        p = tb.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(20, 20, 20)

        # 左栏：我的判定
        lw = int(W * 0.32)
        lt = s.shapes.add_textbox(Emu(300000), Emu(1000000),
                                  Emu(lw), H - Emu(1400000))
        tf = lt.text_frame
        tf.word_wrap = True
        head = tf.paragraphs[0]
        head.text = "我的判定（可质疑）"
        head.font.size = Pt(20)
        head.font.bold = True
        head.font.color.rgb = RGBColor(190, 40, 40)
        for txt, sz, col in ((claim, 16, (40, 40, 40)),
                             ("", 8, (0, 0, 0)),
                             ("一眼该看到什么", 16, (30, 30, 30)),
                             (pred, 14, (70, 70, 70))):
            para = tf.add_paragraph()
            para.text = txt
            para.font.size = Pt(sz)
            para.font.color.rgb = RGBColor(*col)

        # 右栏：真实渲染抽帧，横向排
        if imgs:
            rx = lw + Emu(500000)
            avail = W - rx - Emu(300000)
            n = len(imgs)
            gap = Emu(60000)
            fw = int((avail - gap * (n - 1)) / n)
            fh = int(fw * 9 / 16)
            top = Emu(1100000)
            from PIL import Image
            for i, img in enumerate(imgs):
                x = rx + i * (fw + gap)
                pic = io.BytesIO()
                Image.fromarray(img[:, :, ::-1]).save(pic, format="PNG")
                pic.seek(0)
                s.shapes.add_picture(pic, x, top, width=fw, height=fh)
            cap_tb = s.shapes.add_textbox(rx, Emu(1100000) + fh + Emu(80000),
                                          avail, Emu(500000))
            cp = cap_tb.text_frame.paragraphs[0]
            cp.text = ("← 真实渲染抽帧（地面真值）　"
                       "过渡区间的等间隔采样，左 → 右 = 时间先后")
            cp.font.size = Pt(13)
            cp.font.color.rgb = RGBColor(70, 70, 70)
            if note:
                np_ = cap_tb.text_frame.add_paragraph()
                np_.text = note
                np_.font.size = Pt(12)
                np_.font.color.rgb = RGBColor(120, 80, 20)
        return s

    # 封面
    s = prs.slides.add_slide(blank)
    tb = s.shapes.add_textbox(Emu(700000), int(H * 0.30),
                              W - Emu(1400000), Emu(1500000))
    tf = tb.text_frame
    tf.word_wrap = True
    t = tf.paragraphs[0]
    t.text = "切换形态层：我的判定 vs 渲染实物"
    t.font.size = Pt(40)
    t.font.bold = True
    p2 = tf.add_paragraph()
    p2.text = "左边是我写的中文结论，右边是 PowerPoint 真实渲染抽的帧。"
    p2.font.size = Pt(20)
    p3 = tf.add_paragraph()
    p3.text = ("如果我写反了方向、或把某种运动认错了族，你在这一页就能看出来。")
    p3.font.size = Pt(20)
    p4 = tf.add_paragraph()
    p4.text = "注：抽帧用的是【探测 deck】（网格+一个会移动变色的小圆），"
    p4.font.size = Pt(15)
    p4.font.color.rgb = RGBColor(110, 110, 110)
    p5 = tf.add_paragraph()
    p5.text = ("所以你会看到那个小圆的运动，而不是邮件/图表那种真实内容。")
    p5.font.size = Pt(15)
    p5.font.color.rgb = RGBColor(110, 110, 110)

    missing = []
    for c in claims:
        v = os.path.join(video_dir, c["spec"] + ".mp4")
        imgs = []
        if os.path.exists(v):
            imgs = grab_frames(v)
        else:
            missing.append(c["spec"])
        add_page("%s　—　%s" % (c["spec"], c["family"]),
                 c["claim"], c["pred"], imgs,
                 note=("⚠️ 该效果的 mp4 缺失，本页只有判定没有实物。"
                       if not imgs else ""))

    prs.save(out_pptx)
    print("-> %s" % out_pptx)
    print("   %d 页（封面 + %d 个效果）" % (len(claims) + 1, len(claims)))
    if missing:
        print("   ⚠️ 缺素材：%s" % missing)
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1], sys.argv[2],
                  sys.argv[3] if len(sys.argv) > 3 else None))
