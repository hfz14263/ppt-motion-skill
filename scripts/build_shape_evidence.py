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
【真实渲染视频剪出的循环动画】（地面真值）。人眼扫一遍就知道哪边错了。

⚠️ 第一版这里是"静态抽帧"，被用户当场指出没用 —— 抽帧只能证明"某一刻画面上
有什么"，证明不了"它是怎么动的"，而形态层判的恰恰是**运动**。
现在改成嵌入**循环 GIF**：pptx 里存的是 image.gif，PowerPoint 会真的播。

用法：
  python scripts/build_shape_evidence.py <video_dir> <out.pptx> [gif_dir]

注意：这只回答"我的文字与渲染是否一致"，**不回答"这个效果好不好看"**，
也不回答"该不该用"（那是 3b 的事）。
"""
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

# 过渡窗口（帧）。所有探测视频都是：第 1 页停 2s（60 帧 @30fps）后切到第 2 页。
# 实测峰值帧落在 59–65，取 55–95 让"切之前 / 运动中 / 切之后"三段都在画面里。
WIN_START = 55
WIN_END = 95
GIF_FPS = 15          # 15fps 足够看清运动，体积只有 30fps 的一半
GIF_WIDTH = 640       # 720p 源降到 640 宽，肉眼仍清晰

# ⚠️ 关键：不同效果要用不同的探测 deck 才看得见。
#
# deck1（网格 + 一个移动圆盘）两页的网格完全相同，只有一个圆盘动 —— 所以
# 「整页平移」类看得很清楚（整个网格在移），但「均匀淡变 / 扫描」类几乎不动：
# 实测 wipe/fade/split/random 在 deck1 上帧间差异只有 0.2（肉眼等于静止）。
#
# deck2（全帧随机噪声）每个像素都在变，专门用来分离这些效果：
# 同样四个效果帧间差异升到 3.2 —— 15 倍。
#
# 所以：能用 deck1 的用 deck1（画面干净、方向清楚），deck1 看不出的换 deck2。
# 名单不是猜的，是逐个量帧间差异量出来的。
DECK2_ONLY = ("wipe", "fade", "split", "random", "dissolve", "shape",
              "randombar", "box", "clock")

# 一条 GIF 要被判为"看得出运动"，平均帧间差异至少要这么大。
# 实测：deck1 的 wipe/fade/split/random ≈ 0.2（静止），
#       deck2 的同样效果 ≈ 3.2，deck1 的 push ≈ 38。
# 取 1.0 做分界 —— 低于它就该换 deck，而不是硬塞给用户。
MIN_VISIBLE = 1.0


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
     "claim": "揭开式，方向【r→l】。新画面自右向左铺开 —— "
              "起点在右缘，变化前结束于中间偏左。",
     "pred": "这里该看到：新画面从右边进来、向左侧展开。"},
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
     "claim": "帘式，方向【l↔r】（左右对开）—— 不是上下。"
              "实测变化重心全程钉在 0.50 附近（对称），说明两侧同时开。",
     "pred": "这里该看到：从中缝向左右两边同时拉开，左右对称。"},
    {"spec": "honeycomb", "family": "方向揭示",
     "claim": "归入方向揭示族，但【本层给不出可靠方向】—— "
              "位移指标是噪声级，真实形态是内部纹理在变。这里标注而非硬填。",
     "pred": "抽帧应看到：蜂巢式局部变化，但它是否算「一条前沿」值得怀疑。"},
    {"spec": "flash", "family": "均匀淡变",
     "claim": "整幅同时变化、无方向。靠「覆盖率高 + 漂移≈0」与 fade 分开。",
     "pred": "抽帧应看到：整幅亮度一起闪，没有任何推进边界。"},
    {"spec": "random", "family": "随机",
     "claim": "⚠️ 本层【测不了】：每次播放随机挑一个效果，没有稳定形态。"
              "两个 mp4 字节数相同（186629）但 md5 不同即为证。"
              "本页录到的这一次，它抽中了 wheel（即 clock 的形态）。",
     "pred": "这一页没有「它的」形态 —— 你每次放映看到的可能都不一样。"},
]


def _ffmpeg():
    """定位 ffmpeg。优先用 imageio_ffmpeg 自带的那份（不依赖系统安装）。"""
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        pass
    for cand in ("ffmpeg", r"C:\ffmpeg\bin\ffmpeg.exe"):
        try:
            subprocess.check_output([cand, "-version"],
                                    stderr=subprocess.DEVNULL)
            return cand
        except Exception:
            continue
    return None


def make_gif(video, out_gif):
    """把过渡窗口剪成一个无限循环的 GIF。

    为什么要循环：人眼判断"方向"需要看几遍。循环播放让用户不用反复点重放。

    用调色板两遍法（palettegen + paletteuse）—— 直接转 GIF 会有严重的
    色带和抖动，而探测 deck 是网格 + 移动圆盘，色带会把"运动"糊掉。
    两遍法先生成该片段的专用调色板，再套用，颜色损失小很多。
    """
    ff = _ffmpeg()
    if not ff:
        return None
    fps = 30.0  # 探测视频固定 30fps
    ss = WIN_START / fps
    dur = (WIN_END - WIN_START) / fps
    vf = ("fps=%d,scale=%d:-1:flags=lanczos,split[s0][s1];"
          "[s0]palettegen=max_colors=128[p];"
          "[s1][p]paletteuse=dither=bayer:bayer_scale=3" % (GIF_FPS, GIF_WIDTH))
    cmd = [ff, "-y", "-ss", "%.3f" % ss, "-t", "%.3f" % dur, "-i", video,
           "-vf", vf, "-loop", "0", out_gif]
    r = subprocess.run(cmd, capture_output=True)
    if r.returncode != 0 or not os.path.exists(out_gif):
        return None
    return out_gif


def gif_motion(gif_path):
    """量化这个 GIF 到底动得看不看得见 —— 返回平均帧间差异。

    存在的意义：第一版 deck 里 wipe/fade/split 的 GIF "有 21 帧"却几乎不动
    （差异 0.2），肉眼等于静止。**帧数多不等于看得见运动。** 这个函数让
    "这条 GIF 能不能用来判断方向"变成一个可以自动检查的数字。
    """
    from PIL import Image, ImageSequence
    import numpy as np
    im = Image.open(gif_path)
    fr = [np.array(f.convert("L"), dtype=int) for f in ImageSequence.Iterator(im)]
    if len(fr) < 2:
        return 0.0
    d = [float(np.abs(fr[i] - fr[i + 1]).mean()) for i in range(len(fr) - 1)]
    return sum(d) / len(d)


def main(video_dir, out_pptx, gif_dir=None, deck2_dir=None):
    from pptx import Presentation
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor

    claims = CLAIMS
    if gif_dir is None:
        gif_dir = os.path.join(os.path.dirname(os.path.abspath(out_pptx)),
                               "evidence_gifs")
    os.makedirs(gif_dir, exist_ok=True)
    if deck2_dir is None:
        # 约定：deck2 渲染产物默认放在 <video_dir 的兄弟目录>/shape3
        cand = os.path.join(os.path.dirname(os.path.abspath(video_dir)),
                            "shape3")
        deck2_dir = cand if os.path.isdir(cand) else None

    prs = Presentation()
    prs.slide_width = Emu(12192000)      # 16:9 宽屏，横向放左右两栏
    prs.slide_height = Emu(6858000)
    blank = prs.slide_layouts[6]
    W, H = prs.slide_width, prs.slide_height

    def add_page(title, claim, pred, gif_path, note="", deck_label=""):
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
        lw = int(W * 0.30)
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

        # 右栏：真实渲染动画（循环 GIF，PowerPoint 会播）
        rx = lw + Emu(500000)
        avail = W - rx - Emu(400000)
        top = Emu(1100000)
        if gif_path:
            fh = int(avail * 9 / 16)
            s.shapes.add_picture(gif_path, rx, top, width=avail, height=fh)
            cap_top = top + fh + Emu(90000)
        else:
            cap_top = top
        cap_tb = s.shapes.add_textbox(rx, cap_top, avail, Emu(700000))
        ct = cap_tb.text_frame
        ct.word_wrap = True
        cp = ct.paragraphs[0]
        cp.text = ("↑ 真实渲染动画（地面真值）　循环播放　%s"
                   % (("　" + deck_label) if deck_label else ""))
        cp.font.size = Pt(13)
        cp.font.color.rgb = RGBColor(70, 70, 70)
        cp2 = ct.add_paragraph()
        cp2.text = "这段是探测 deck 原样剪出来的，没有重绘或加速。"
        cp2.font.size = Pt(12)
        cp2.font.color.rgb = RGBColor(110, 110, 110)
        if note:
            np_ = ct.add_paragraph()
            np_.text = note
            np_.font.size = Pt(12)
            np_.font.color.rgb = RGBColor(120, 80, 20)
        return s

    # 封面
    s = prs.slides.add_slide(blank)
    tb = s.shapes.add_textbox(Emu(700000), int(H * 0.28),
                              W - Emu(1400000), Emu(1800000))
    tf = tb.text_frame
    tf.word_wrap = True
    t = tf.paragraphs[0]
    t.text = "切换形态层：我的判定 vs 渲染实物"
    t.font.size = Pt(40)
    t.font.bold = True
    p2 = tf.add_paragraph()
    p2.text = "左边是我写的中文结论，右边是 PowerPoint 真实渲染的【循环动画】。"
    p2.font.size = Pt(20)
    p3 = tf.add_paragraph()
    p3.text = "如果我写反了方向、或把某种运动认错了族，看这一眼就能发现。"
    p3.font.size = Pt(20)
    p4 = tf.add_paragraph()
    p4.text = ""
    p4.font.size = Pt(10)
    p5 = tf.add_paragraph()
    p5.text = "⚠️ 必须按 F5 全屏放映才看得到动 —— 编辑视图里 GIF 可能静止。"
    p5.font.size = Pt(18)
    p5.font.bold = True
    p5.font.color.rgb = RGBColor(190, 40, 40)
    p6 = tf.add_paragraph()
    p6.text = "注：动画用的是【探测 deck】（网格 + 一个会移动变色的小圆），"
    p6.font.size = Pt(15)
    p6.font.color.rgb = RGBColor(110, 110, 110)
    p7 = tf.add_paragraph()
    p7.text = "所以你会看到那个小圆的运动，而不是邮件/图表那种真实内容。"
    p7.font.size = Pt(15)
    p7.font.color.rgb = RGBColor(110, 110, 110)

    missing = []
    no_gif = []
    too_static = []
    for c in claims:
        spec = c["spec"]
        # 选 deck：deck1 看得见就用 deck1，看不见就换 deck2（见 DECK2_ONLY 注释）
        cands = []
        if spec in DECK2_ONLY and deck2_dir:
            cands.append(os.path.join(deck2_dir, spec + ".mp4"))
        cands.append(os.path.join(video_dir, spec + ".mp4"))
        if deck2_dir:
            cands.append(os.path.join(deck2_dir, spec + ".mp4"))
        src = next((p for p in cands if os.path.exists(p)), None)

        gif = None
        src_name = ""
        if src:
            out_gif = os.path.join(gif_dir, spec + ".gif")
            gif = make_gif(src, out_gif)
            src_name = os.path.basename(os.path.dirname(src))
            if not gif:
                no_gif.append(spec)
            else:
                # 看得见才算数 —— 帧数多不代表能用（见 gif_motion 注释）。
                # 例外：硬切本来就没有中间态，帧间差异天然接近 0 —— 低不是缺陷，
                # 恰恰是它的形态特征，所以不报警告。
                m = gif_motion(gif)
                if m < MIN_VISIBLE and c["family"] != "硬切":
                    too_static.append((spec, round(m, 2)))
        else:
            missing.append(spec)

        note = ""
        if spec in missing:
            note = "⚠️ 该效果的 mp4 缺失，本页只有判定没有实物。"
        elif spec in no_gif:
            note = "⚠️ GIF 生成失败，本页只有判定。"
        elif spec in [x[0] for x in too_static]:
            note = "⚠️ 警告：这条动画帧间差异极低（%.2f），可能看不出运动。" % \
                   dict(too_static)[spec]
        # 标注用的是哪个探测 deck —— 两个 deck 的画面完全不同，
        # 不写清楚的话用户会以为所有页都是同一套画面。
        label = ("探测 deck2（噪声纹理）" if src_name == "shape3"
                 else "探测 deck1（网格+圆盘）")
        add_page("%s　—　%s" % (spec, c["family"]),
                 c["claim"], c["pred"], gif, note=note, deck_label=label)

    prs.save(out_pptx)
    n_gif = len(claims) - len(missing) - len(no_gif)
    print("-> %s" % out_pptx)
    print("   %d 页（封面 + %d 个效果），其中 %d 页带动画"
          % (len(claims) + 1, len(claims), n_gif))
    print("   GIF 目录：%s" % gif_dir)
    if missing:
        print("   ⚠️ 缺 mp4：%s" % missing)
    if no_gif:
        print("   ⚠️ GIF 失败：%s" % no_gif)
    if too_static:
        print("   ⚠️ 动画几乎看不出（需换 deck）：%s" % too_static)
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1], sys.argv[2],
                  sys.argv[3] if len(sys.argv) > 3 else None,
                  sys.argv[4] if len(sys.argv) > 4 else None))
