#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把分享内容生成为真正的 .pptx 文件。

依赖：python-pptx
  pip install python-pptx

用法：
  python tools/build_pptx.py            # 生成 AI学习成果分享.pptx
  python tools/build_pptx.py --out 我的分享.pptx
"""

import argparse

from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


PRIMARY = RGBColor(0x1F, 0x4E, 0x79)   # 深蓝
ACCENT = RGBColor(0x2E, 0x86, 0xC1)    # 亮蓝
DARK = RGBColor(0x22, 0x29, 0x2F)
GRAY = RGBColor(0x55, 0x5F, 0x6B)

# 每页：标题 + 若干要点（支持二级缩进）+ 备注（演讲稿）
SLIDES = [
    {
        "title": "从“凭经验”到“靠工具”",
        "subtitle": "我用 Cursor + AI 提升电话销售效率的实践\n分享人：XXX",
        "cover": True,
        "notes": "大家好，我是做电话销售的。这次分享的不是高大上的 AI 理论，而是我用 AI 实实在在解决了三个每天都要干的麻烦事。全程只用一个工具——Cursor。",
    },
    {
        "title": "痛点：电话销售的“隐形时间黑洞”",
        "bullets": [
            ("打完电话要手动复盘、记笔记", 0),
            ("每个客户打电话前要翻半天资料、想话术", 0),
            ("客户问产品细节，知识库要连 VPN、搜半天", 0),
        ],
        "notes": "我们真正赚钱的动作是打电话，但很多时间花在了打电话之外：打完要复盘，打之前要准备话术，客户一问细节还得去翻知识库。这些事不做不行，做了又很耗时间。",
    },
    {
        "title": "我手上有什么“原料”",
        "bullets": [
            ("CRM 销售记录  →  时间该花在谁身上？", 0),
            ("知识库网站  →  客户问的我答得准吗？", 0),
            ("通话录音  →  我哪句话说错 / 说漏了？", 0),
            ("把这三样和 AI 结合，做成了三个小工具", 0),
        ],
        "notes": "我盘了一下手上的资源：CRM 销售记录、知识库网站、通话录音。这三样分别对应销售最关键的三个问题，我把它们和 AI 结合，做出了三个小工具。",
    },
    {
        "title": "我的 AI 学习路径",
        "bullets": [
            ("学会用 Cursor 跑脚本，用大白话让 AI 帮我做工具", 0),
            ("理解了关键概念 RAG：让 AI 先查资料再回答，不瞎编", 0),
            ("学会写“好提问”（Prompt），结果质量差别巨大", 0),
        ],
        "notes": "我主要学了三件事：把 Cursor 用起来，哪怕不懂代码也能让它帮我做工具；搞懂了 RAG，就是让 AI 先查我们自己的资料再回答；学会了怎么把问题问清楚。",
    },
    {
        "title": "成果一 · 通话复盘助手（现场演示）",
        "bullets": [
            ("把录音转写丢进去，2 秒生成复盘表", 0),
            ("自动提取：需求 / 预算 / 异议 / 下一步", 1),
            ("话术诊断：哪里没接住、哪里该追问", 1),
            ("合规检查 + 金句沉淀", 1),
            ("复盘时间：20 分钟 → 2 分钟", 0),
        ],
        "notes": "第一个也是我最想给大家看的——通话复盘助手。现场打开 Cursor 运行 call_review.py，把一段脱敏转写丢进去，它马上整理出客户需求、异议、我哪里做得不好、漏了哪些动作，甚至挑出违规话术。以前复盘一通要二十分钟，现在两分钟。",
    },
    {
        "title": "成果二 · 客户话术生成器",
        "bullets": [
            ("输入客户画像，输出整通电话的话术", 0),
            ("开场白 / 核心卖点 / 异议应对", 1),
            ("该问的关键问题 + 跟进文案", 1),
            ("打电话前准备：翻半天 → 1 分钟", 0),
        ],
        "notes": "第二个，话术生成器。打电话前把客户在 CRM 里的情况告诉 AI，它就生成一整套话术：开场怎么说、三个卖点怎么讲、客户挑刺怎么接、最后该约什么。相当于每次打电话前都有个老销售帮我对稿。",
    },
    {
        "title": "成果三 · 知识库即问即答（RAG）",
        "bullets": [
            ("用大白话问，AI 带着出处回答", 0),
            ("只答知识库里有的，查不到就说“建议核实”", 0),
            ("不用再连 VPN 翻文档", 0),
        ],
        "notes": "第三个解决知识库难搜的问题。把常用文档脱敏后放进文件夹，用大白话问，它基于我们自己的资料回答，还告诉我出处。关键是只答资料里有的，查不到就老实说建议核实，不会瞎编坑我。",
    },
    {
        "title": "量化成果（估算）",
        "bullets": [
            ("单通电话复盘：20 分钟 → 2 分钟", 0),
            ("单个客户话术准备：15 分钟 → 1 分钟", 0),
            ("知识库查询：5 分钟 → 30 秒", 0),
            ("省下的时间 = 每天多打 N 通有效电话", 0),
        ],
        "notes": "算笔账：这三件事加起来，我每天大概能省出一两个小时，拿去多打几通有效电话。而且复盘做细了话术也在变好，这是复利。",
    },
    {
        "title": "数据安全红线（重点）",
        "bullets": [
            ("客户信息、录音、知识库 → 先脱敏再用", 0),
            ("优先用公司批准的工具和接口", 0),
            ("不把真实客户数据传到公网", 0),
        ],
        "notes": "特别强调：这些数据涉及客户隐私和公司机密，我所有演示都是脱敏的，真实数据只在公司允许的环境里用。用 AI 提效的前提是安全合规，这条红线不能碰。",
    },
    {
        "title": "怎么复制给团队",
        "bullets": [
            ("工具、Prompt 模板都整理好了，复制即用", 0),
            ("不需要会编程，会用 Cursor 就行", 0),
            ("建议先从“通话复盘”开始，见效最快", 0),
        ],
        "notes": "这套东西我整理成了模板，复制就能用，不需要会编程。如果只想试一个，建议从通话复盘开始。需要的同事会后找我。谢谢大家！",
    },
]


def _set_text(frame, text, size, color, bold=False, align=PP_ALIGN.LEFT):
    frame.word_wrap = True
    p = frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Microsoft YaHei"


def add_cover(prs, slide_def):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版式
    # 顶部色块
    left, top, width, height = Inches(0), Inches(2.4), prs.slide_width, Inches(2.6)
    box = slide.shapes.add_textbox(Inches(0.8), Inches(2.2), prs.slide_width - Inches(1.6), Inches(1.6))
    _set_text(box.text_frame, slide_def["title"], 40, PRIMARY, bold=True, align=PP_ALIGN.CENTER)
    sub = slide.shapes.add_textbox(Inches(0.8), Inches(3.8), prs.slide_width - Inches(1.6), Inches(1.4))
    tf = sub.text_frame
    tf.word_wrap = True
    for i, line in enumerate(slide_def["subtitle"].split("\n")):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = PP_ALIGN.CENTER
        r = para.add_run()
        r.text = line
        r.font.size = Pt(20)
        r.font.color.rgb = GRAY
        r.font.name = "Microsoft YaHei"
    _add_notes(slide, slide_def.get("notes"))


def add_content(prs, slide_def):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # 标题条
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), prs.slide_width - Inches(1.2), Inches(1.0))
    _set_text(title_box.text_frame, slide_def["title"], 30, PRIMARY, bold=True)
    # 标题下划线
    line = slide.shapes.add_shape(1, Inches(0.65), Inches(1.35), Inches(2.2), Pt(3))
    line.fill.solid()
    line.fill.fore_color.rgb = ACCENT
    line.line.fill.background()

    body = slide.shapes.add_textbox(Inches(0.8), Inches(1.7), prs.slide_width - Inches(1.6), Inches(4.8))
    tf = body.text_frame
    tf.word_wrap = True
    for i, (text, level) in enumerate(slide_def["bullets"]):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.level = level
        para.space_after = Pt(10)
        run = para.add_run()
        bullet = "•  " if level == 0 else "–  "
        run.text = bullet + text
        run.font.size = Pt(22 if level == 0 else 18)
        run.font.bold = level == 0
        run.font.color.rgb = DARK if level == 0 else GRAY
        run.font.name = "Microsoft YaHei"
    _add_notes(slide, slide_def.get("notes"))


def _add_notes(slide, notes):
    if notes:
        slide.notes_slide.notes_text_frame.text = notes


def build(out_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)   # 16:9
    prs.slide_height = Inches(7.5)
    for sd in SLIDES:
        if sd.get("cover"):
            add_cover(prs, sd)
        else:
            add_content(prs, sd)
    prs.save(out_path)
    print(f"[完成] 已生成：{out_path}（共 {len(SLIDES)} 页，含演讲稿备注）")


def main():
    parser = argparse.ArgumentParser(description="生成 AI 学习成果分享 PPTX")
    parser.add_argument("--out", "-o", default="AI学习成果分享.pptx", help="输出文件名")
    args = parser.parse_args()
    build(args.out)


if __name__ == "__main__":
    main()
