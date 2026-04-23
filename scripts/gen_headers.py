# -*- coding: utf-8 -*-
"""
生成 C4 楷体 + C6 仿宋 的报头 PNG 图片到 assets/

动机：iOS Gmail app 会把楷体（Kaiti）/ 仿宋（FangSong）fallback 成系统默认宋体，
      失去毛笔味和账本味。把最醒目的报头文字预渲染成 PNG，用 <img> 嵌入邮件，
      Gmail 无法修改图片内字体。

输出：
  assets/c4/month-01.png ~ month-12.png   (12 张，"X 月 記" 华文楷体)
  assets/c6/header-monthly.png            ("个人消费月报表" 华文仿宋)
  assets/c6/header-weekly.png             ("个人消费周报表")

尺寸：560 x 160 / 120（2x retina，邮件里按 280 或 200 px 显示）
"""

from PIL import Image, ImageDraw, ImageFont
from pathlib import Path
import os
import sys

ROOT = Path(__file__).parent.parent  # smbc-auto-ledger/

FONTS_DIR = r"C:\Windows\Fonts"
FONT_KAITI    = os.path.join(FONTS_DIR, "STKAITI.TTF")
FONT_FANGSONG = os.path.join(FONTS_DIR, "STFANGSO.TTF")

ASSETS_C4 = ROOT / "assets" / "c4"
ASSETS_C6 = ROOT / "assets" / "c6"
ASSETS_C4.mkdir(parents=True, exist_ok=True)
ASSETS_C6.mkdir(parents=True, exist_ok=True)

# 颜色匹配邮件模板背景
C4_FG = (42, 24, 16)       # #2a1810 深褐
C4_BG = (241, 229, 203)    # #f1e5cb 宣纸
C6_FG = (26, 20, 16)       # #1a1410
C6_BG = (243, 234, 208)    # #f3ead0

MONTH_ZH = ['一','二','三','四','五','六','七','八','九','十','十一','十二']


def render_centered(text, font, fg, bg, width, height, letter_spacing=0):
    """渲染单行文字，水平 + 垂直居中。letter_spacing 单位 px。"""
    img = Image.new("RGB", (width, height), bg)
    draw = ImageDraw.Draw(img)

    # 逐字符宽度（处理中文字符变宽情况）
    char_widths = [draw.textlength(ch, font=font) for ch in text]
    total_w = sum(char_widths) + letter_spacing * max(0, len(text) - 1)

    x = (width - total_w) / 2
    y = height / 2

    for i, ch in enumerate(text):
        cw = char_widths[i]
        draw.text((x + cw / 2, y), ch, fill=fg, font=font, anchor='mm')
        x += cw + letter_spacing

    return img


def main():
    saved = []

    # ===== C4: 12 张 "X 月 記" 楷体 =====
    font_c4 = ImageFont.truetype(FONT_KAITI, 96)
    for m in range(1, 13):
        title = f"{MONTH_ZH[m-1]} 月 記"
        img = render_centered(title, font_c4, C4_FG, C4_BG, 560, 160, letter_spacing=24)
        out = ASSETS_C4 / f"month-{m:02d}.png"
        img.save(out, "PNG", optimize=True)
        saved.append(out)
        print(f"  [ok] {out.relative_to(ROOT)}")

    # ===== C6: 月报 + 周报各一张 仿宋 =====
    font_c6 = ImageFont.truetype(FONT_FANGSONG, 64)
    for fname, text in [
        ("header-monthly.png", "个人消费月报表"),
        ("header-weekly.png",  "个人消费周报表"),
    ]:
        img = render_centered(text, font_c6, C6_FG, C6_BG, 560, 120, letter_spacing=16)
        out = ASSETS_C6 / fname
        img.save(out, "PNG", optimize=True)
        saved.append(out)
        print(f"  [ok] {out.relative_to(ROOT)}")

    print(f"\nGenerated {len(saved)} files under assets/")
    return saved


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
