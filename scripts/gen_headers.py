# -*- coding: utf-8 -*-
"""
生成全部 6 套邮件模板的报头 PNG 图片到 assets/

动机：iOS Gmail app 会对 font-family 做强制 fallback，非系统默认衬线的装饰字体
      （楷体 / 仿宋 / Playfair Display / 粗宋体）会被替换成系统默认 serif。
      把每个模板最醒目的报头文字预渲染成 PNG，用 <img> 通过 GitHub raw URL 嵌入，
      Gmail 无法干预图片内字体。

输出（18 张 PNG）：
  assets/a/header.png                     "The Monthly Ledger"   Georgia Bold Italic
  assets/c1/header.png                    "月度账本"              STZhongsong 粗宋
  assets/c2/header.png                    "月 计 报"              STZhongsong 特大
  assets/c4/month-01.png ~ month-12.png   "X 月 記"              STKaiti 楷体（12 张）
  assets/c6/header-monthly.png            "个人消费月报表"         STFangsong 仿宋
  assets/c6/header-weekly.png             "个人消费周报表"         STFangsong 仿宋
  assets/e7/header.png                    "The Monthly\nLedger"   Playfair Display Black（两行）

尺寸：2x retina（邮件里按 ~280px 宽度显示）
"""

from PIL import Image, ImageDraw, ImageFont
from pathlib import Path
import os
import sys

ROOT = Path(__file__).parent.parent  # smbc-auto-ledger/

# ===== 字体路径 =====
WIN_FONTS = r"C:\Windows\Fonts"
FONT_KAITI    = os.path.join(WIN_FONTS, "STKAITI.TTF")      # C4 楷体
FONT_FANGSONG = os.path.join(WIN_FONTS, "STFANGSO.TTF")     # C6 仿宋
FONT_ZHONGS   = os.path.join(WIN_FONTS, "STZHONGS.TTF")     # C1 / C2 粗宋
FONT_GEORGIA_BI = os.path.join(WIN_FONTS, "georgiaz.ttf")   # A  Georgia Bold Italic
FONT_PLAYFAIR = str(ROOT / "assets" / "fonts" / "PlayfairDisplay-VF.ttf")  # E7 维多利亚（variable font, set to Black 900）

# ===== 输出目录 =====
for sub in ("a", "c1", "c2", "c4", "c6", "e7"):
    (ROOT / "assets" / sub).mkdir(parents=True, exist_ok=True)

# ===== 配色（匹配邮件模板背景）=====
C_A    = dict(fg=(26, 26, 26),   bg=(255, 255, 255))    # A  白底黑字
C_C1   = dict(fg=(26, 26, 26),   bg=(244, 237, 224))    # C1 米底
C_C2   = dict(fg=(26, 20, 16),   bg=(244, 237, 224))    # C2 米底
C_C4   = dict(fg=(42, 24, 16),   bg=(241, 229, 203))    # C4 宣纸
C_C6   = dict(fg=(26, 20, 16),   bg=(243, 234, 208))    # C6 账本底
C_E7   = dict(fg=(26, 20, 16),   bg=(235, 224, 198))    # E7 羊皮底

MONTH_ZH = ['一','二','三','四','五','六','七','八','九','十','十一','十二']


# ===== 渲染 helper =====

def _draw_line(draw, text, font, x_center, y_center, color, letter_spacing):
    """在 (x_center, y_center) 居中绘制一行文字（逐字符 + 自定义字距）。"""
    char_widths = [draw.textlength(ch, font=font) for ch in text]
    total_w = sum(char_widths) + letter_spacing * max(0, len(text) - 1)
    x = x_center - total_w / 2
    for i, ch in enumerate(text):
        cw = char_widths[i]
        draw.text((x + cw / 2, y_center), ch, fill=color, font=font, anchor='mm')
        x += cw + letter_spacing


def render_centered(text, font, fg, bg, width, height, letter_spacing=0):
    """单行居中。"""
    img = Image.new("RGB", (width, height), bg)
    draw = ImageDraw.Draw(img)
    _draw_line(draw, text, font, width / 2, height / 2, fg, letter_spacing)
    return img


def render_two_lines(line1, line2, font, fg, bg, width, height, letter_spacing=0, line_height_ratio=0.95):
    """两行居中，行高按字号 * ratio。"""
    img = Image.new("RGB", (width, height), bg)
    draw = ImageDraw.Draw(img)
    # 字号（font.size）× ratio = 行间距
    line_gap = font.size * line_height_ratio
    y1 = height / 2 - line_gap / 2
    y2 = height / 2 + line_gap / 2
    _draw_line(draw, line1, font, width / 2, y1, fg, letter_spacing)
    _draw_line(draw, line2, font, width / 2, y2, fg, letter_spacing)
    return img


def save(img, rel_path):
    out = ROOT / rel_path
    img.save(out, "PNG", optimize=True)
    print(f"  [ok] {rel_path}")
    return out


# ===== 生成 =====

def main():
    count = 0

    # --- A: "The Monthly Ledger"  Georgia Bold Italic ---
    font_a = ImageFont.truetype(FONT_GEORGIA_BI, 56)
    img = render_centered("The Monthly Ledger", font_a, C_A['fg'], C_A['bg'], 560, 140, letter_spacing=-1)
    save(img, "assets/a/header.png"); count += 1

    # --- C1: "月度账本"  STZhongsong 粗宋 ---
    font_c1 = ImageFont.truetype(FONT_ZHONGS, 108)
    img = render_centered("月度账本", font_c1, C_C1['fg'], C_C1['bg'], 560, 180, letter_spacing=14)
    save(img, "assets/c1/header.png"); count += 1

    # --- C2: "月 计 报"  STZhongsong 超大 ---
    font_c2 = ImageFont.truetype(FONT_ZHONGS, 124)
    img = render_centered("月 计 报", font_c2, C_C2['fg'], C_C2['bg'], 560, 200, letter_spacing=20)
    save(img, "assets/c2/header.png"); count += 1

    # --- C4: "X 月 記" × 12  STKaiti 楷体 ---
    font_c4 = ImageFont.truetype(FONT_KAITI, 96)
    for m in range(1, 13):
        title = f"{MONTH_ZH[m-1]} 月 記"
        img = render_centered(title, font_c4, C_C4['fg'], C_C4['bg'], 560, 160, letter_spacing=24)
        save(img, f"assets/c4/month-{m:02d}.png"); count += 1

    # --- C6: 月报 + 周报  STFangsong 仿宋 ---
    font_c6 = ImageFont.truetype(FONT_FANGSONG, 64)
    for fname, text in [
        ("header-monthly.png", "个人消费月报表"),
        ("header-weekly.png",  "个人消费周报表"),
    ]:
        img = render_centered(text, font_c6, C_C6['fg'], C_C6['bg'], 560, 120, letter_spacing=16)
        save(img, f"assets/c6/{fname}"); count += 1

    # --- E7: "The Monthly / Ledger"  Playfair Display Black（variable font, wght=900，两行） ---
    font_e7 = ImageFont.truetype(FONT_PLAYFAIR, 72)
    try:
        font_e7.set_variation_by_axes([900])
    except Exception as e:
        print(f"  [warn] Playfair VF weight 900 failed: {e}, using default", file=sys.stderr)
    img = render_two_lines("The Monthly", "Ledger", font_e7, C_E7['fg'], C_E7['bg'],
                           560, 180, letter_spacing=-1, line_height_ratio=0.95)
    save(img, "assets/e7/header.png"); count += 1

    print(f"\nGenerated {count} files under assets/")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
