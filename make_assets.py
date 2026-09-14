#!/usr/bin/env python3
"""从 assets/crest/ 的立绘原图生成侧栏头像和正文立绘。

用法：
    pip install pillow numpy
    python3 make_assets.py

会做两件事：
  assets/avatar/<角色名>.png    76×76 的脸部方图（靠肤色自动找脸）
  assets/portrait/<角色名>.webp 300×360 统一画框的半身像，下缘柔化

角色名以 convert.py 里的 ROSTER 为准。原图文件名去掉空格后能对上就行，
对不上的（繁简、日文名）在下面的 ALIAS 里点名。
"""
import os
import sys

import numpy as np
from PIL import Image

from convert import ROSTER

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "assets", "crest")
AV = os.path.join(HERE, "assets", "avatar")
PT = os.path.join(HERE, "assets", "portrait")

# ROSTER 里的角色名 -> 立绘文件名（去空格后仍对不上的写在这里）
ALIAS = {
    "手塚国光": "手冢国光",
    "海堂熏": "海堂 薰",
    "赤澤吉朗": "赤泽𠮷朗",
    "観月はじめ": "观月 初",
    "神尾アキラ": "神尾 明",
}

AV_PX = 76                      # 头像边长
W, H = 300, 360                 # 立绘画框（网页上按 150×180 显示）
HEAD = 0.47 * W                 # 脸部方框统一缩到这么宽
HEAD_X, HEAD_Y = 0.50, 0.285    # 头心在画框里的位置


def index_sources():
    if not os.path.isdir(SRC):
        sys.exit(f"找不到 {SRC}，请先把立绘原图放进去")
    out = {}
    for f in os.listdir(SRC):
        if f.lower().endswith((".png", ".webp")):
            key = os.path.splitext(f)[0].replace(" ", "").replace("　", "")
            out[key] = os.path.join(SRC, f)
    return out


def find(name, table):
    for cand in (ALIAS.get(name, ""), name):
        if cand and cand.replace(" ", "") in table:
            return table[cand.replace(" ", "")]
    return None


def head_box(im):
    """靠肤色找脸。只在立绘上段找，并取肤色最密集的那一横带当脸，
    免得举起的手、交叉的手臂把中心拽下去。"""
    a = np.array(im)
    ys, xs = np.nonzero(a[..., 3] > 8)
    y0, y1, x0, x1 = ys.min(), ys.max(), xs.min(), xs.max()
    bw, bh = x1 - x0, y1 - y0

    top_h = int(bh * 0.45)
    top = a[y0:y0 + top_h, x0:x1 + 1]
    R, G, B, A = (top[..., i].astype(int) for i in range(4))
    skin = (A > 200) & (R > 185) & (G > 135) & (B > 110) & (R > G + 12) & (G > B + 2)

    rows = skin.sum(axis=1).astype(float)
    lo_y = 0
    if rows.sum() > 400 and rows.max() > 6:
        k = max(3, int(bh * 0.02)) | 1
        sm = np.convolve(rows, np.ones(k) / k, mode="same")
        peak = int(np.argmax(sm))
        band = int(bh * 0.085)
        lo_y, hi_y = max(0, peak - band), min(top_h, peak + band)
        sy, sx = np.nonzero(skin[lo_y:hi_y])
    else:
        sx = sy = np.array([])

    if len(sx) > 200:
        # 脸宽只取峰值那一段连续的肤色，举到头边的手就不会算进来
        cols = np.bincount(sx, minlength=top.shape[1]).astype(float)
        kx = max(3, int(bw * 0.02)) | 1
        cols = np.convolve(cols, np.ones(kx) / kx, mode="same")
        pk = int(np.argmax(cols))
        thr = cols[pk] * 0.22
        l = r = pk
        while l > 0 and cols[l - 1] > thr:
            l -= 1
        while r < len(cols) - 1 and cols[r + 1] > thr:
            r += 1
        cx = x0 + (l + r) // 2
        cy = y0 + lo_y + int(np.median(sy))
        size = int(max(r - l, 40) * 2.32)
        cy -= int(size * 0.14)
    else:                                        # 兜底：按立绘比例估
        size = int(bh * 0.42)
        cx, cy = (x0 + x1) // 2, y0 + size // 2

    size = int(np.clip(size, bh * 0.22, min(bh * 0.62, bw * 1.3)))
    half = size // 2
    cy = max(cy, y0 + int(half * 0.80))
    return (cx - half, cy - half, cx + half, cy + half)


def smoothstep(t):
    t = np.clip(t, 0, 1)
    return t * t * (3 - 2 * t)


def soft_mask():
    """下缘淡出、左右收一点边，让立绘融进页面底色。"""
    y = np.linspace(0, 1, H)[:, None]
    x = np.linspace(0, 1, W)[None, :]
    bottom = 1 - smoothstep((y - 0.66) / 0.34)
    side = 0.25 + 0.75 * smoothstep(x / 0.09) * smoothstep((1 - x) / 0.09)
    return np.clip(bottom * side * smoothstep(y / 0.05), 0, 1)


MASK = soft_mask()


def make_avatar(im, box):
    canvas = Image.new("RGBA", (box[2] - box[0], box[3] - box[1]), (0, 0, 0, 0))
    canvas.paste(im.crop(box), (0, 0))
    canvas = canvas.resize((AV_PX, AV_PX), Image.LANCZOS)
    bg = Image.new("RGBA", (AV_PX, AV_PX), (255, 255, 255, 255))
    bg.alpha_composite(canvas)
    return bg.convert("RGB")


def make_portrait(im, box):
    s = HEAD / (box[2] - box[0])
    big = im.resize((max(1, int(im.width * s)), max(1, int(im.height * s))), Image.LANCZOS)
    cx, cy = int((box[0] + box[2]) / 2 * s), int((box[1] + box[3]) / 2 * s)
    left, top = cx - int(W * HEAD_X), cy - int(H * HEAD_Y)

    canvas = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    canvas.paste(big.crop((left, top, left + W, top + H)), (0, 0))
    a = np.array(canvas)
    a[..., 3] = (a[..., 3] * MASK).astype(np.uint8)
    return Image.fromarray(a)


def main():
    os.makedirs(AV, exist_ok=True)
    os.makedirs(PT, exist_ok=True)
    table = index_sources()

    made, missing = 0, []
    for _sname, _crest, _color, names in ROSTER:
        for cname in names:
            path = find(cname, table)
            if not path:
                missing.append(cname)
                continue
            im = Image.open(path).convert("RGBA")
            box = head_box(im)
            make_avatar(im, box).save(os.path.join(AV, cname + ".png"), optimize=True)
            # 柔化后的 alpha 是渐变，调色板 PNG 存不了，所以立绘用 WebP
            make_portrait(im, box).save(os.path.join(PT, cname + ".webp"), quality=84, method=6)
            made += 1

    print(f"✓ 生成 {made} 套头像 + 立绘")
    if missing:
        print("  缺立绘（在 ALIAS 里补对应关系，或把原图放进 assets/crest/）：")
        print("   ", "、".join(missing))


if __name__ == "__main__":
    main()
