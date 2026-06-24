#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 PPT 用的示例界面截图（模拟通话复盘助手网页）。"""

import os
import textwrap

from PIL import Image, ImageDraw, ImageFont

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "assets", "screenshots")
W, H = 960, 540
BG = (245, 249, 253)
PRIMARY = (31, 78, 121)
ACCENT = (46, 134, 193)
WHITE = (255, 255, 255)
GRAY = (85, 95, 107)
RED_BG = (255, 244, 244)
GREEN_BG = (238, 250, 240)
WARN_BG = (255, 248, 225)


def _font(size, bold=False):
    paths = [
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/System/Library/Fonts/PingFang.ttc",
        "C:/Windows/Fonts/msyh.ttc",
    ]
    for p in paths:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()


def _card(draw, x, y, w, h, fill=WHITE, radius=12):
    draw.rounded_rectangle([x, y, x + w, y + h], radius=radius, fill=fill, outline=(220, 228, 236))


def gen_upload_page():
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)
    f_title = _font(28)
    f_body = _font(16)
    f_small = _font(13)

    # 顶栏
    draw.rectangle([0, 0, W, 56], fill=PRIMARY)
    draw.text((24, 14), "📞 通话复盘助手", fill=WHITE, font=f_title)

    _card(draw, 40, 80, W - 80, H - 120)
    draw.text((70, 100), "上传录音转写文件，一键生成销售复盘", fill=PRIMARY, font=f_body)

    # 拖拽区
    draw.rounded_rectangle([70, 140, W - 70, 280], radius=14, outline=ACCENT, width=2, fill=(232, 243, 252))
    draw.text((W // 2 - 120, 175), "📁 点击选择，或把文件拖到这里", fill=PRIMARY, font=f_body)
    draw.text((W // 2 - 150, 210), "支持 Excel(.xlsx) / CSV / txt", fill=GRAY, font=f_small)
    draw.text((W // 2 - 100, 240), "已选择：sample_call.xlsx", fill=ACCENT, font=f_small)

    # 按钮
    draw.rounded_rectangle([70, 310, W - 70, 360], radius=10, fill=PRIMARY)
    draw.text((W // 2 - 55, 325), "🚀 开始分析", fill=WHITE, font=f_body)

    draw.text((70, 380), "⚠️ 上传前请先脱敏（去掉真实姓名、电话、公司名）", fill=GRAY, font=f_small)

    path = os.path.join(OUT_DIR, "01_上传页面.png")
    img.save(path)
    return path


def gen_result_page():
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)
    f_title = _font(22)
    f_body = _font(14)
    f_small = _font(12)

    draw.rectangle([0, 0, W, 48], fill=PRIMARY)
    draw.text((20, 12), "📞 通话复盘结果", fill=WHITE, font=f_title)

    y = 60
    draw.text((24, y), "综合评分  64 / 100", fill=ACCENT, font=f_title)
    y += 40
    draw.text((24, y), "客户画像：X 先生 | 采购负责人 | 餐饮", fill=GRAY, font=f_small)
    y += 28

    sections = [
        ("客户需求", "30 家门店会员数据分散，老板看不到汇总", WHITE),
        ("客户异议", "【切换风险】担心影响营业 → 建议给试点方案", WARN_BG),
        ("待改进", "❌「保证三天上线」→ 改为「通常 3-5 个工作日」", RED_BG),
        ("可复用金句", "「一个后台看全部门店数据，相当于多请一个助理」", GREEN_BG),
    ]
    for title, text, bg in sections:
        _card(draw, 24, y, W - 48, 52, fill=bg)
        draw.text((36, y + 6), title, fill=PRIMARY, font=f_body)
        draw.text((36, y + 26), text[:52], fill=GRAY, font=f_small)
        y += 60

    path = os.path.join(OUT_DIR, "02_分析结果.png")
    img.save(path)
    return path


def gen_steps_page():
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)
    f_title = _font(24)
    f_body = _font(16)

    draw.text((24, 20), "使用步骤（3 步搞定）", fill=PRIMARY, font=f_title)
    steps = [
        ("1", "打开网站链接", "浏览器访问你部署好的网址"),
        ("2", "上传脱敏后的转写文件", "Excel / CSV / txt 均可"),
        ("3", "点击「开始分析」", "几秒后查看复盘结果，可下载保存"),
    ]
    y = 80
    for num, title, sub in steps:
        draw.ellipse([40, y, 76, y + 36], fill=ACCENT)
        draw.text((52, y + 6), num, fill=WHITE, font=f_body)
        draw.text((96, y + 2), title, fill=PRIMARY, font=f_body)
        draw.text((96, y + 26), sub, fill=GRAY, font=_font(13))
        y += 70

    path = os.path.join(OUT_DIR, "03_使用步骤.png")
    img.save(path)
    return path


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    for fn in (gen_upload_page, gen_result_page, gen_steps_page):
        p = fn()
        print(f"生成：{p}")


if __name__ == "__main__":
    main()
