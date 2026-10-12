#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""程序化生成太阳系贴图（纯 numpy + PIL，无外部素材、无网络调用）。

为什么不用图像生成 API：
- 泄漏面小（不发任何请求）
- **可复现**：同一份代码永远产出同一张图 —— 这才可能写"逐值核对"的判据。
  AI 生图做不到这点，它每次出的图都不一样。
- 行星是规则球体，程序化着色能精确控制明暗界线（terminator）、木星条带、
  土星环缝；生成模型反而经常画不出这些物理细节。

输出的图对应示例 deck「太阳系漫游指南」：
    background.png   银河星空背景（总览页与详情页共用底图）
    sun.png          太阳（日冕光晕）
    <planet>.png     8 行星，其中 saturn 带环

用法：
    python examples/material/solar/make_assets.py
"""
from __future__ import annotations

import pathlib

import numpy as np
from PIL import Image, ImageDraw, ImageFilter

OUT = pathlib.Path(__file__).resolve().parent

# 输出尺寸。行星贴图按各自大小给：详情页里单个行星要占画面主要面积，
# 所以给到 512~1024，总览页里会缩小使用，缩放由 PowerPoint 处理。
SIZE_STAR = 512
SIZE_SUN = 900
SIZE_BG = 1920


def _blur(a: np.ndarray, radius: int) -> np.ndarray:
    """盒式模糊（用累积和，O(n)），用作局部均值。只用 numpy，不依赖 scipy。"""
    k = radius * 2 + 1
    a32 = a.astype(np.float32)
    if a32.ndim == 2:
        a32 = a32[..., None]
    pad = np.pad(a32, ((radius, radius), (radius, radius), (0, 0)), mode="edge")
    cs = np.cumsum(np.cumsum(pad, axis=0), axis=1)
    cs = np.pad(cs, ((1, 0), (1, 0), (0, 0)))
    h, w = a32.shape[:2]
    total = (cs[k:k + h, k:k + w] - cs[0:h, k:k + w] - cs[k:k + h, 0:w] + cs[0:h, 0:w])
    total = total / (k * k)
    return total if a.ndim == 3 else total[..., 0]


def _sphere_shade(w: int, h: int, light_dir, terminator: float = 0.72):
    """生成球体的明暗贴图。

    关键点：明暗界线（terminator）要有**柔和过渡带**。硬边会让球看起来像
    贴纸，而不是被单一光源照亮的球体 —— 这也是纯程序化最容易露怯的地方。

    光源方向默认偏向**侧上方**而不是正上方：正上方光源会让亮度沿纵向单调
    衰减（顶部亮底部暗），而木星/土星的条带是横向的 —— 两者叠加会把条带
    对比度彻底吃掉，条带就"消失"了（第一版就是这样）。把主光压到接近
    水平，条带才看得见，这更接近真实插画的做法。
    """
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    cx, cy = w / 2.0, h / 2.0
    r = min(cx, cy) * 0.94
    # 归一化到单位圆
    nx = (xx - cx) / r
    ny = (yy - cy) / r
    nz = np.sqrt(np.clip(1.0 - nx * nx - ny * ny, 0.0, 1.0))

    # 朗伯 + 边缘暗化（rim）：单纯 Lambert 会让边缘看起来平
    lx, ly, lz = light_dir
    lam = np.clip(nx * lx + ny * ly + nz * lz, 0.0, 1.0)
    lam = lam ** terminator

    # 大气/边缘散射：靠近轮廓稍亮一点再压暗，避免"贴纸边"
    rim = (1.0 - nz) ** 3
    shade = lam * 0.92 + rim * 0.10 - rim * 0.14

    # 抗锯齿的圆盘遮罩
    alpha = np.clip((r - np.sqrt(nx * nx + ny * ny) * r) / 2.0, 0.0, 1.0)
    return np.clip(shade, 0.0, 1.3), alpha


def _bands(h: int, w: int, colors, n_bands: int, warp: float = 0.0, rng=None):
    """生成横向条带纹理（木星/土星）。

    用**平滑插值**而不是按行硬切分：硬切在放大后是一圈圈的台阶，
    一眼假（第一版就是这么废掉的）。做法是先在 v 上做一维平滑，
    再取色 —— 等价于"色带之间有过渡带"。
    """
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    v = yy / h
    if warp and rng is not None:
        n = (np.sin(xx / w * 6.0 + v * 9.0) * 0.6 +
             np.sin(xx / w * 11.0 - v * 5.0) * 0.4)
        v = v + n * warp
    t = np.clip(v, 0, 0.9999) * (len(colors) - 1)
    i0 = np.floor(t).astype(int)
    i1 = np.minimum(i0 + 1, len(colors) - 1)
    f = (t - i0)[..., None]
    pal = np.asarray(colors, np.float32)
    out = pal[i0] * (1 - f) + pal[i1] * f
    # 细颗粒：气体带不是纯色
    out *= (0.94 + 0.12 * _speckle(h, w, rng)[..., None])
    return out


def _speckle(h: int, w: int, rng, scale=0.35):
    """分形噪声（多倍频），用于岩质行星表面与云带细节。"""
    acc = np.zeros((h, w), np.float32)
    for octave, amp in ((4, 0.5), (8, 0.3), (16, 0.2), (32, 0.1)):
        small = rng.random((octave, octave)).astype(np.float32)
        ys = (np.arange(h) * octave // h).clip(0, octave - 1)
        xs = (np.arange(w) * octave // w).clip(0, octave - 1)
        acc += small[np.ix_(ys, xs)] * amp
    return acc / acc.max()


def make_planet(name, size, base, *, bands_colors=None, n_bands=0,
                warp=0.0, noise=0.0, light=(-0.62, -0.34, 0.52),
                terminator=0.72, seed=7):
    rng = np.random.default_rng(seed)
    h = w = size
    shade, alpha = _sphere_shade(w, h, light, terminator)

    if bands_colors:
        base_img = _bands(h, w, bands_colors, n_bands, warp, rng)
    else:
        base_img = np.zeros((h, w, 3), np.float32)
        base_img[:] = base
        if noise:
            base_img *= (0.88 + 0.24 * _speckle(h, w, rng)[..., None])

    # 贴图自身也做一点明暗（自阴影），与球面光照相乘
    #
    # 条带行星必须走【另一套合成】：直接 base*shade 会把条带吃掉 ——
    # shade 的纵向跨度约 10 倍，而条带的纵向对比只有约 1.2 倍，
    # 相乘之后条带幅度被压到 2% 以下，肉眼等于没有（第一版就是这样）。
    # 做法：先用球面光照的**局部均值**去归一化 base（把大梯度抽掉），
    # 再把条带作为乘性调制叠回去。这样条带只在"受光程度"上起伏。
    if bands_colors:
        lum = shade[..., None] + 1e-3
        # 抽掉球面光照的局部均值，只留条带；再乘回光照。
        # 四个坑，都踩过（记下来是因为每次的表现都不同）：
        #  1) lm 不归一化到均值≈1 → 整体压暗
        #  2) lm 之外又乘 gain → 衰减做两遍，球体变黑
        #  3) lm 的背光侧接近 0（最小 0.0008），除它等于除以 0 → 溢出成纯白/纯黑
        #  4) **除以 lm*255 而不除以 lm** —— 这是最隐蔽的一个：
        #     lm 是"相对亮度比"（≈1），base 是 0..255 量级，
        #     两者量纲差 255 倍，除完只剩 0.5，球体全黑。
        lm = _blur(lum, radius=max(8, w // 24))
        lm = lm / max(float(lm.mean()), 1e-3)
        lm = np.clip(lm, 0.30, None)          # 坑 3：钳下限，不在背光侧做除法
        flat = np.clip(base_img / lm, 0, 255) # 坑 4：只除 lm
        img = flat * lum                      # 乘回光照
        # 收尾归一化：把受光面的均值拉回"看得清"的水平（约 150）。
        # 不做这一步条带虽在，但整颗球偏暗 —— demo 缩到屏幕上就糊成一片。
        lit = img[shade > 0.35]
        if lit.size and lit.mean() > 1:
            img *= min(1.8, 150.0 / float(lit.mean()))
    else:
        img = base_img * shade[..., None]
    # 极轻微的高光提亮，让球有"光滑"感
    img = img + np.clip(shade[..., None] - 0.82, 0, None) * 90
    rgb = np.clip(img, 0, 255).astype(np.uint8)
    rgba = np.dstack([rgb, (alpha * 255).astype(np.uint8)])

    im = Image.fromarray(rgba, "RGBA")
    if name == "saturn":
        im = _add_ring(im)
    if bands_colors:
        # 条带边界软化：半径过小会在相邻色带交界留下可见的细线
        # （演示尺寸 320pt 上很明显）。0.4px 足够掩盖，又不糊掉条带。
        im = im.filter(ImageFilter.GaussianBlur(0.4))
    im.save(OUT / f"{name}.png")
    return name


def _add_ring(im: Image.Image) -> Image.Image:
    """给土星加环。

    三个必须做对的地方（第一版全错）：
    1. **反投影**：环是倾斜的圆，画在屏幕上是**椭圆**。按角度采样会把圆画成
       一串短线段 —— 那是虚线，不是环。正确做法是先在单位圆上取 3600 个点，
       再按 (cos·a, sin·a·ry) 投影成椭圆，一次性填多边形。
    2. **遮挡关系**：环必须分成"球后"与"球前"两半，后半画在球下面。
       只画前半会变成"环贴在球面上"；完全不遮挡也不对。
    3. **卡西尼缝**：土星环最容易被认出的特征，靠半径区间留空得到，
       不是靠画一条线。
    """
    w, h = im.size
    pad_x = int(w * 0.50)
    pad_y = int(h * 0.34)
    W, H = w + pad_x * 2, h + pad_y * 2
    cx, cy = W / 2.0, H / 2.0
    rx = w * 0.72                       # 环的横向半径
    ry = rx * 0.30                      # 透视压扁 → 屏幕上是椭圆

    N = 3600
    ta = np.linspace(0.0, 2 * np.pi, N, endpoint=False)
    xs = cx + np.cos(ta) * rx
    ys = cy + np.sin(ta) * ry

    def half(t_start, t_end):
        """返回某一半环（角度区间）的 RGBA 数组，形状 (H, W, 4)。"""
        m = (ta >= t_start) & (ta < t_end)
        if not m.any():
            return None
        # 反投影得到全量归一化半径，再做掩码 —— 不能先用 m 切数组
        # （布尔索引返回的是压缩过的数组，长度不等于 N，会和全量 rad 撞形状）
        u = (xs - cx) / rx
        v = (ys - cy) / ry
        rad_all = np.sqrt(u * u + v * v)

        from PIL import ImageDraw
        canvas = np.zeros((H, W, 4), np.float32)
        col = np.array([226.0, 203.0, 162.0])
        tmp = Image.fromarray((canvas * 0).astype(np.uint8), "RGBA")
        d = ImageDraw.Draw(tmp)
        # 内外边缘各做一次描边填充：draw.polygon 一次只能一个半径，
        # 所以按三段半径区间分别填。
        segs = [(0.60, 0.725, 140), (0.745, 0.94, 255), (0.940, 1.005, 190)]
        for lo, hi, A in segs:
            keep = m & (rad_all >= lo) & (rad_all < hi)
            if keep.sum() < 3:
                continue
            idx = np.nonzero(keep)[0]
            pts = [(float(xs[j]), float(ys[j])) for j in idx]
            pts.append(pts[0])          # 首尾相接，闭合多边形
            d.line(pts, fill=(int(col[0]), int(col[1]), int(col[2]), A),
                   width=max(2, int(h * 0.030)), joint="curve")
        arr = np.asarray(tmp).astype(np.float32)
        # 少量颗粒 + 边缘柔化
        rng = np.random.default_rng(31)
        a = arr[..., 3] / 255.0
        a *= (0.90 + 0.16 * _speckle(H, W, rng))
        arr[..., 3] = np.clip(a * 255, 0, 255)
        arr[..., 0:3] = col[None, None, :]
        return arr

    canvas = np.zeros((H, W, 4), np.float32)

    # ① 球后半的环（角度 180°..360°，即 sin<0 的上半）
    back = half(np.pi, 2 * np.pi)
    if back is not None:
        canvas[..., :] = back
    # ② 球体（覆盖环的后半）
    #    球体贴在 canvas 的中心，尺寸不等 —— 必须先把球体放进同尺寸画布，
    #    否则 numpy 会因形状不等直接报 broadcast 错（第一版就踩了）。
    ball = np.zeros((H, W, 4), np.float32)
    bw, bh = im.size
    ball[pad_y:pad_y + bh, pad_x:pad_x + bw] = np.asarray(im).astype(np.float32)
    ba = ball[..., 3:4] / 255.0
    ca = canvas[..., 3:4] / 255.0
    out_a = ca + ba * (1 - ca)
    safe = np.where(out_a > 1e-6, out_a, 1.0)
    canvas[..., 0:3] = (canvas[..., 0:3] * ca + ball[..., 0:3] * ba * (1 - ca)) / safe[..., 0][..., None]
    canvas[..., 3:4] = out_a * 255
    # ③ 球前半的环（sin>0 的下半，压在球上）
    front = half(0.0, np.pi)
    if front is not None:
        fa = front[..., 3:4] / 255.0
        ca = canvas[..., 3:4] / 255.0
        out_a2 = fa + ca * (1 - fa)
        safe2 = np.where(out_a2 > 1e-6, out_a2, 1.0)
        canvas[..., 0:3] = (front[..., 0:3] * fa + canvas[..., 0:3] * ca * (1 - fa)) / safe2[..., 0][..., None]
        canvas[..., 3:4] = out_a2 * 255

    out = Image.fromarray(np.clip(canvas, 0, 255).astype(np.uint8), "RGBA")
    return out.filter(ImageFilter.GaussianBlur(0.7))


def make_sun(size=SIZE_SUN):
    """太阳：本体 + 日冕光晕。光晕要多层叠加，单层渐变看起来像贴纸。"""
    h = w = size
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    cx = cy = w / 2
    rr = np.sqrt((xx - cx) ** 2 + (yy - cy) ** 2) / (w * 0.5)

    core = np.clip(1.0 - rr / 0.42, 0, 1)
    core = core ** 0.6
    # 米粒组织的粗噪声，避免纯色渐变
    rng = np.random.default_rng(11)
    core *= (0.90 + 0.20 * _speckle(h, w, rng))

    glow = np.clip(1.0 - rr, 0, 1) ** 2.4 * 0.85

    rgb = np.zeros((h, w, 3), np.float32)
    rgb[..., 0] = core * 255 + glow * 255
    rgb[..., 1] = core * 210 + glow * 170
    rgb[..., 2] = core * 110 + glow * 90
    alpha = np.clip(core + glow, 0, 1) * 255
    im = Image.fromarray(np.dstack([np.clip(rgb, 0, 255).astype(np.uint8),
                                    alpha.astype(np.uint8)]), "RGBA")
    im = im.filter(ImageFilter.GaussianBlur(1.2))
    im.save(OUT / "sun.png")
    return "sun"


def make_background(w=SIZE_BG, h=1080, seed=23):
    """银河星空背景。方向键翻页时它不动 —— 星空必须是"远景"，一动就露馅。"""
    rng = np.random.default_rng(seed)
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    base = np.zeros((h, w, 3), np.float32)
    base[..., 0], base[..., 1], base[..., 2] = 3, 4, 12

    # 银河：沿一条倾斜带的高斯亮雾
    ang = np.deg2rad(-24)
    d = (np.cos(ang) * (xx - w * 0.62) - np.sin(ang) * (yy - h * 0.46))
    band = np.exp(-(d / (h * 0.26)) ** 2)
    mott = _speckle(h, w, rng)
    mist = band * (0.30 + 0.70 * mott) * 0.30
    # 抖动只用来打散 uint8 量化台阶，幅度必须**远小于信号本身**。
    # 抖动一放大就变成可见的紫色噪点带（第一版就是这样）。
    # 判据：抖动幅度 ≤ 该处信号的 8%。
    jit = (rng.random((h, w)).astype(np.float32) - 0.5) * 2.0
    mist = np.clip(mist + jit * (band * 0.08 + 0.02), 0, None)
    base[..., 0] += mist * 132
    base[..., 1] += mist * 148
    base[..., 2] += mist * 205

    # 恒星：三层亮度分布，暗星远多于亮星（真实星等分布）
    for count, bright in ((2600, 0.42), (420, 0.75), (70, 1.0)):
        ys = rng.integers(0, h, count)
        xs = rng.integers(0, w, count)
        size = rng.random(count) * 1.5 + 0.5
        for y, x, s, b in zip(ys, xs, size, np.full(count, bright)):
            v = b * (0.55 + 0.45 * rng.random())
            tint = np.array([1.0, 0.94 + 0.06 * rng.random(), 0.86 + 0.14 * rng.random()])
            yy0, xx0 = int(y), int(x)
            for dy in range(0, 2):
                for dx in range(0, 2):
                    y2, x2 = yy0 + dy, xx0 + dx
                    if 0 <= y2 < h and 0 <= x2 < w:
                        k = v * (1.0 if (dy == 0 and dx == 0) else 0.42)
                        base[y2, x2] += tint * 235 * k * (1.0 if s > 1 else 0.6)

    im = Image.fromarray(np.clip(base, 0, 255).astype(np.uint8), "RGB")
    # 两道：先 0.35 消高频，再 1.8 的较大半径把银河的低频色阶台阶抹平。
    # 只用抖动是不够的 —— 抖动把台阶换成噪点，噪点同样会被量化放大；
    # 必须让银河本身足够平滑，抖动只做最后 1~2 个色阶的 dither。
    im = im.filter(ImageFilter.GaussianBlur(0.35))
    im = im.filter(ImageFilter.GaussianBlur(1.8))
    im.save(OUT / "background.png", optimize=True)
    return "background"


# --------------------------------------------------------------------------
# 8 行星参数。距离与直径比例刻意【不按真实比例】—— 真实比例下地球只有
# 1 个像素宽。这里用"视觉可读"的比例：太阳大、行星依次可辨、土星环完整。
# 这一点必须在 README 里说明，否则会被当成科普错误。
# --------------------------------------------------------------------------
PLANETS = [
    # name, size,  base rgb,           bands,                       warp, noise, seed
    ("mercury", 420, (150, 143, 130), None, 0, 0.30, 3),
    ("venus",   480, (222, 184, 110), None, 0, 0.22, 5),
    ("earth",   540, (58, 106, 168),  None, 0, 0.34, 7),
    ("mars",    460, (196, 84, 46),   None, 0, 0.28, 9),
    ("jupiter", 820, (206, 176, 126),
     [(226, 208, 178), (158, 116, 74), (120, 84, 52), (240, 228, 202)], 7, 0.055, 11),
    ("saturn",  740, (216, 194, 152),
     [(234, 218, 184), (188, 158, 112), (146, 118, 82), (226, 208, 176)], 6, 0.045, 13),
    ("uranus",  520, (128, 196, 202), None, 0, 0.16, 17),
    ("neptune", 500, (48, 82, 186),  None, 0, 0.20, 19),
]


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    print("背景...")
    make_background()
    print("太阳...")
    make_sun()
    for name, size, base, bands, nb, warp, seed in PLANETS:
        print(f"{name}...")
        make_planet(name, size, base,
                    bands_colors=bands, n_bands=nb, warp=warp,
                    noise=0.0 if bands else 1.0, seed=seed)
    for p in sorted(OUT.glob("*.png")):
        with Image.open(p) as im:
            print(f"  {p.name:18} {im.size[0]:5}x{im.size[1]:<5} {p.stat().st_size/1024:8.1f} KB")


if __name__ == "__main__":
    main()
