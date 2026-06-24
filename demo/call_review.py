#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通话录音复盘助手（Demo）

用途：把一段「通话录音转写文字」自动整理成结构化的销售复盘表。
适合在分享会上现场演示：左边放转写文字，右边 2 秒生成复盘。

支持的输入文件：.txt / .md / .csv / .xlsx（Excel）
  - Excel/CSV：脚本会把所有单元格的文字按行读出来拼成转写文字，
    所以不管你的表格是「时间/角色/内容」三列，还是把对话写在一列里，都能用。
  - 解析 Excel 用的是 Python 标准库，无需安装任何依赖。

两种运行模式：
  1) 真·AI 模式：配置好兼容 OpenAI 的接口（公司内部接口 / 任意大模型），
     脚本会真的调用大模型来生成复盘。
  2) 离线 Demo 模式：没有配置任何接口时，自动用内置的示例结果演示效果，
     断网也能现场跑，不会报错、不会把客户数据传出去。

用法：
  # 离线演示（最稳，适合现场）
  python demo/call_review.py --input samples/sample_call.txt

  # 用 Excel 文件作为输入
  python demo/call_review.py --input 我的录音转写.xlsx

  # 接入真实大模型（OpenAI 兼容接口）
  export LLM_API_KEY="你的key"
  export LLM_BASE_URL="https://api.openai.com/v1"   # 或公司内部兼容地址
  export LLM_MODEL="gpt-4o-mini"
  python demo/call_review.py --input 我的录音转写.xlsx

  # 把结果同时导出成 markdown 文件
  python demo/call_review.py --input samples/sample_call.txt --out 复盘结果.md

注意：演示前请先对录音/转写做脱敏（去掉真实姓名、电话、公司名）。
"""

import argparse
import csv
import json
import os
import re
import sys
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
import zipfile


SYSTEM_PROMPT = """你是一名资深的电话销售教练，擅长复盘销售通话。
请阅读用户提供的「通话转写文字」，输出一份**详尽、可执行**的销售复盘。
要求：
- 每条「待改进」必须引用原话或具体场景，并给出「更好的说法」示例；
- 「客户异议」要分类（价格/信任/切换风险/决策权等）；
- 「下一步行动」要具体到时间点和交付物；
- 语气像资深销售教练，直接、实用。
只输出一个 JSON 对象，不要输出任何额外文字、解释或 markdown 代码块标记。
JSON 必须严格符合以下结构：

{
  "customer": {
    "name": "客户称呼（脱敏，如 X 先生 / A 公司）",
    "role": "客户身份/职位",
    "industry": "所在行业"
  },
  "demand": ["客户的核心需求/关注点，逐条列出"],
  "budget": "客户预算或采购意向（没有则写 未明确）",
  "objections": ["客户提出的异议或顾虑，逐条列出"],
  "did_well": ["销售在本通电话里做得好的点，逐条列出"],
  "to_improve": ["销售可以改进的点，要具体到话术，逐条列出"],
  "missed_actions": ["销售遗漏的关键动作，比如没要预算/没约下一步等"],
  "next_step": "建议的下一步行动（一句话，可执行）",
  "compliance": ["合规风险提示，比如夸大承诺/违规话术；没有则写 未发现明显问题"],
  "golden_lines": ["可沉淀复用的高质量话术/金句，逐条列出"],
  "score": "本通电话综合评分（0-100 的整数，字符串）"
}
"""


def read_input(path: str) -> str:
    """根据扩展名读取输入，统一返回转写文字。
    支持 .txt/.md（纯文本）、.csv、.xlsx（Excel）。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".txt", ".md", ""):
        with open(path, "r", encoding="utf-8") as f:
            return f.read().strip()
    if ext == ".csv":
        return _read_csv(path)
    if ext == ".xlsx":
        return _read_xlsx(path)
    if ext == ".xls":
        raise ValueError(
            "暂不支持老版 .xls 格式，请在 Excel 里「另存为」.xlsx 后再试。"
        )
    # 其它后缀，尝试按纯文本读
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        return f.read().strip()


def _read_csv(path: str) -> str:
    """把 CSV 每行的非空单元格用空格连起来，行与行换行。"""
    rows = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        for row in csv.reader(f):
            cells = [c.strip() for c in row if c and c.strip()]
            if cells:
                rows.append("  ".join(cells))
    return "\n".join(rows).strip()


def _read_xlsx(path: str) -> str:
    """仅用标准库解析 .xlsx：读取第一个工作表，按行拼出所有单元格文字。"""
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path) as z:
        # 共享字符串表（Excel 把文本统一存这里）
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("m:si", ns):
                shared.append("".join(t.text or "" for t in si.iter("{%s}t" % ns["m"])))

        # 找第一个工作表
        sheet_names = [n for n in z.namelist()
                       if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
        if not sheet_names:
            return ""
        sheet_names.sort()
        root = ET.fromstring(z.read(sheet_names[0]))

        lines = []
        for row in root.iter("{%s}row" % ns["m"]):
            cells = []
            for c in row.findall("m:c", ns):
                t = c.get("t")
                v = c.find("m:v", ns)
                if t == "s" and v is not None:  # 共享字符串
                    idx = int(v.text)
                    if 0 <= idx < len(shared):
                        cells.append(shared[idx])
                elif t == "inlineStr":
                    is_el = c.find("m:is", ns)
                    if is_el is not None:
                        cells.append("".join(x.text or "" for x in is_el.iter("{%s}t" % ns["m"])))
                elif v is not None and v.text:
                    cells.append(v.text)
            cells = [c.strip() for c in cells if c and c.strip()]
            if cells:
                lines.append("  ".join(cells))
        return "\n".join(lines).strip()


def build_user_prompt(transcript: str) -> str:
    return f"以下是一通销售电话的转写文字，请按要求复盘：\n\n---\n{transcript}\n---"


def call_llm(transcript: str, api_key: str, base_url: str, model: str) -> dict:
    """调用兼容 OpenAI 的 chat/completions 接口，返回解析后的 JSON。"""
    url = base_url.rstrip("/") + "/chat/completions"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": build_user_prompt(transcript)},
        ],
        "temperature": 0.3,
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        body = json.loads(resp.read().decode("utf-8"))
    content = body["choices"][0]["message"]["content"].strip()
    content = _strip_code_fence(content)
    return json.loads(content)


def _strip_code_fence(text: str) -> str:
    """去掉模型可能多包的 ```json ... ``` 包裹。"""
    if text.startswith("```"):
        lines = text.splitlines()
        lines = [ln for ln in lines if not ln.strip().startswith("```")]
        return "\n".join(lines).strip()
    return text


def mock_result() -> dict:
    """离线演示用的内置示例结果（基于 samples/sample_call.txt）。"""
    return {
        "customer": {
            "name": "X 先生",
            "role": "采购负责人",
            "industry": "连锁餐饮",
        },
        "demand": [
            "想要一套能管理 30 家门店的会员系统",
            "关心能否和现有的收银系统打通",
            "希望能看到各门店的销售数据汇总",
        ],
        "budget": "未明确，但提到'预算这块还要内部商量'",
        "objections": [
            "担心系统切换会影响门店正常营业",
            "之前用过别家产品，售后跟不上",
            "觉得价格比预期高",
        ],
        "did_well": [
            "开场很快说明来意，没有绕圈子",
            "主动询问了客户门店数量，做了需求确认",
        ],
        "to_improve": [
            "客户提到'担心切换影响营业'时，没有给出具体的过渡方案，只是说'放心'，说服力不够",
            "客户说价格高，直接降价应对，应先讲价值再谈价格",
            "全程没有为下一步动作设定明确时间点",
        ],
        "missed_actions": [
            "没有问清楚客户的决策流程和拍板人",
            "没有约定下一次沟通的具体时间",
            "没有索要客户邮箱以便发送资料",
        ],
        "next_step": "本周五前发送'门店平滑切换方案+成功案例'，并预约下周二上午电话演示。",
        "compliance": [
            "提到'保证三天上线'属于过度承诺，建议改为'通常 3-5 个工作日，视门店数量而定'",
        ],
        "golden_lines": [
            "'您现在 30 家店的数据是分开看的，我们能让您在一个后台看到全部，相当于多请了一个数据分析助理。'",
        ],
        "score": "72",
    }


def _split_turns(transcript: str):
    """把转写按行拆成 (说话人, 内容)。支持 Excel 三列格式：时间  角色  内容"""
    turns = []
    skip_headers = {"时间", "角色", "内容", "说话人", "speaker", "time", "text"}
    for raw in transcript.splitlines():
        line = raw.strip()
        if not line or line in skip_headers:
            continue
        line = re.sub(r"^\d{1,2}:\d{2}(:\d{2})?\s+", "", line)
        speaker = "未知"
        # Excel 三列：销售  您好… / 00:01  销售  您好…（时间已去掉）
        m = re.match(r"^(销售|业务|顾问|客服|我方|我|sales)\s{2,}(.+)$", line, re.I)
        if m:
            speaker, line = "销售", m.group(2).strip()
        else:
            m = re.match(r"^(客户|顾客|对方|客|买家|customer)\s{2,}(.+)$", line, re.I)
            if m:
                speaker, line = "客户", m.group(2).strip()
            else:
                m = re.match(r"^(销售|业务|顾问|客服|我方|我|sales)\s*[:：、\-]\s*(.+)$", line, re.I)
                if m:
                    speaker, line = "销售", m.group(2).strip()
                else:
                    m = re.match(r"^(客户|顾客|对方|客|买家|customer)\s*[:：、\-]\s*(.+)$", line, re.I)
                    if m:
                        speaker, line = "客户", m.group(2).strip()
        if line:
            turns.append((speaker, line))
    return turns


def _dedup(seq, limit=8):
    out = []
    seen = set()
    for x in seq:
        x = x.strip("　 ,，。.!！?？")
        key = x[:40]
        if x and key not in seen:
            seen.add(key)
            out.append(x)
        if len(out) >= limit:
            break
    return out


def _short(text, n=48):
    t = text.strip()
    return t if len(t) <= n else t[:n] + "…"


def _infer_customer(turns, cust_lines):
    """从对话推断客户画像。"""
    text = " ".join(cust_lines)
    text_all = " ".join(t for _, t in turns)
    name = "客户"
    for pat in [r"[A-Z]\s*先生", r"[A-Z]\s*女士", r"[A-Z]\s*总", r"[\u4e00-\u9fff]{1,2}\s*先生", r"[\u4e00-\u9fff]{1,2}\s*女士"]:
        m = re.search(pat, text_all)
        if m:
            name = m.group(0)
            break
    industry_map = [
        ("餐饮", ["餐饮", "门店", "连锁", "店", "会员", "收银"]),
        ("零售", ["零售", "商场", "超市", "门店"]),
        ("制造", ["工厂", "生产", "制造", "产线"]),
        ("教育", ["学校", "培训", "教育", "校区"]),
        ("医疗", ["医院", "诊所", "医疗"]),
    ]
    industry = "—"
    for ind, kws in industry_map:
        if sum(1 for k in kws if k in text) >= 2:
            industry = ind
            break
    role = "采购/业务负责人" if any(k in text for k in ["采购", "负责人", "老板", "商量"]) else "—"
    return {"name": name, "role": role, "industry": industry}


# 异议类型 → 标准应对话术模板
_OBJ_TEMPLATES = {
    "切换/实施风险": {
        "kws": ["切换", "上线", "折腾", "影响", "营业", "开单", "迁移", "实施"],
        "better": "完全理解您的顾虑。我们建议先选 1 家店试点，老店照常开单，验证后再逐店推开。我可以把《平滑切换方案》发您过目。",
    },
    "价格/预算": {
        "kws": ["贵", "价格", "预算", "预期", "优惠", "便宜", "多少钱"],
        "better": "我先帮您算笔账：现在人工汇总/跟进的隐性成本是多少？系统省下的时间和避免的流失，往往比价格更值。您方便说下大概门店规模吗？",
    },
    "信任/售后": {
        "kws": ["售后", "找不到人", "别家", "之前", "坑", "不信", "担心"],
        "better": "您这个顾虑特别对。我们每个客户有专属对接人，工作时间内 30 分钟响应。我把直线电话给您，有事直接找我，不绕客服。",
    },
    "决策权/流程": {
        "kws": ["老板", "商量", "内部", "汇报", "决策", "拍板"],
        "better": "明白，这类采购通常要老板一起定。您看方便的话，我整理一份一页纸方案，您转给老板过目？或者约个三方短会？",
    },
}


def heuristic_result(transcript: str) -> dict:
    """不调用 AI，用增强规则分析真实上传内容，输出更详尽的复盘。"""
    turns = _split_turns(transcript)
    cust = [t for s, t in turns if s == "客户"]
    sales = [t for s, t in turns if s == "销售"]
    all_lines = [t for _, t in turns] or [l.strip() for l in transcript.splitlines() if l.strip()]
    cust_lines = cust or all_lines
    sales_lines = sales or []

    def hits(lines, kws):
        return [l for l in lines if any(k in l for k in kws)]

    # ---------- 需求：提炼为要点句 ----------
    NEED_KWS = ["想要", "需要", "希望", "想", "头疼", "最", "能不能", "可不可以", "有没有", "解决", "管理", "打通", "对接", "汇总", "数据"]
    demand_raw = hits(cust_lines, NEED_KWS)
    demand = []
    for line in _dedup(demand_raw, 6):
        # 把长句提炼成「客户需要…」格式
        if len(line) > 60:
            demand.append(f"客户关注：{_short(line, 55)}")
        else:
            demand.append(line)
    if not demand:
        demand = ["客户尚未明确表达核心需求，建议下通电话重点追问痛点"]

    # ---------- 异议：分类 + 原话 + 建议应对 ----------
    objections = []
    for cat, cfg in _OBJ_TEMPLATES.items():
        matched = hits(cust_lines, cfg["kws"])
        if matched:
            quote = _short(matched[0], 42)
            objections.append(f"【{cat}】客户原话：「{quote}」→ 建议应对：{cfg['better']}")

    # 兜底：其它异议关键词
    OBJ_EXTRA = ["担心", "顾虑", "怕", "不敢", "犹豫", "再看看", "再考虑", "麻烦", "不放心"]
    extra = hits(cust_lines, OBJ_EXTRA)
    for line in _dedup(extra, 3):
        if not any(_short(line, 20) in o for o in objections):
            objections.append(f"【其它顾虑】「{_short(line, 42)}」→ 先认同情绪，再追问具体担心点，给可验证的案例或方案")

    if not objections:
        objections = ["客户暂未提出明显异议，处于信息了解阶段"]

    # ---------- 预算 ----------
    BUDGET_KWS = ["预算", "价格", "报价", "多少钱", "费用", "成本", "优惠", "便宜", "预期"]
    budget_hits = hits(all_lines, BUDGET_KWS)
    if budget_hits:
        budget = f"涉及价格/预算讨论。客户原话摘录：「{_short(budget_hits[0], 50)}」"
        if any(k in " ".join(cust_lines) for k in ["商量", "老板", "内部"]):
            budget += "；决策需内部商量，预算尚未拍板。"
    else:
        budget = "未明确（通话中未深入讨论预算或报价）"

    # ---------- 合规 ----------
    OVER_KWS = ["保证", "绝对", "百分百", "100%", "一定", "肯定", "立刻", "马上", "稳赚", "永久", "没问题", "放心"]
    over_hits = hits(sales_lines or all_lines, OVER_KWS)
    compliance = []
    for line in _dedup(over_hits, 4):
        if any(k in line for k in ["保证", "绝对", "百分百", "一定", "肯定", "永久"]):
            compliance.append(
                f"⚠️ 过度承诺风险：「{_short(line, 40)}」\n"
                f"   建议改为：「视门店数量和对接情况，通常需要 X 个工作日，我们先评估再给您准确时间。」"
            )
        elif "放心" in line or "没问题" in line:
            compliance.append(
                f"⚠️ 空泛安抚：「{_short(line, 40)}」\n"
                f"   建议改为：给出具体过渡方案或案例，而不是只说「放心」。"
            )
    if not compliance:
        compliance = ["✅ 未发现明显违规话术"]

    # ---------- 做得好的地方（引用原话）----------
    did_well = []
    for line in sales_lines[:8]:
        if any(k in line for k in ["请问", "了解", "门店", "几家", "规模", "目前"]):
            did_well.append(f"主动探需：「{_short(line, 45)}」")
            break
    if sales_lines and any(k in sales_lines[0] for k in ["您好", "你好", "我是"]):
        did_well.append(f"开场规范：「{_short(sales_lines[0], 45)}」")
    for line in sales_lines:
        if any(k in line for k in ["汇总", "后台", "数据", "相当于", "帮您", "节省"]):
            did_well.append(f"价值表达：「{_short(line, 45)}」")
            break
    if not did_well:
        did_well.append("通话整体节奏平稳，可结合录音进一步标注亮点")

    # ---------- 待改进（原话 + 问题 + 改法）----------
    to_improve = []
    weak_phrases = [("放心", "对客户顾虑只回复「放心/没问题」，缺乏具体方案",
                     "应给出试点方案、时间线或同行案例，让客户看到可控路径"),
                    ("便宜", "客户嫌贵时直接谈优惠/降价", "先讲价值与 ROI，再问预算范围，最后再谈方案匹配"),
                    ("考虑", "通话以「您先考虑」结束，未锁定下一步", "应争取具体时间：「我周五发方案，下周二上午 10 点跟您确认，可以吗？」")]
    for kw, problem, fix in weak_phrases:
        for line in sales_lines:
            if kw in line:
                to_improve.append(f"❌ 原话：「{_short(line, 38)}」\n   问题：{problem}\n   ✅ 改法：{fix}")
                break

    if over_hits:
        to_improve.append(
            f"❌ 原话：「{_short(over_hits[0], 38)}」\n"
            "   问题：过度承诺会降低信任，后续做不到反而丢单\n"
            "   ✅ 改法：用「通常…视情况而定…我们先评估」替代绝对化表述"
        )
    if objections and not any("试点" in l or "方案" in l or "案例" in l for l in sales_lines):
        to_improve.append(
            "❌ 客户提出顾虑后，销售未给出可落地的解决方案（方案/案例/试点）\n"
            "   ✅ 改法：每个异议对应一个「证据」：案例名、试点步骤、或书面方案"
        )
    if not hits(all_lines, ["下周", "明天", "周一", "周二", "周三", "周四", "周五", "几点", "约", "演示", "发您"]):
        to_improve.append(
            "❌ 未约定明确的下一步时间与交付物\n"
            "   ✅ 改法：「我今天下班前把方案发您邮箱，我们约周四下午 3 点电话过一遍，您看可以吗？」"
        )
    if not to_improve:
        to_improve.append("整体表现不错，建议把本通中的金句和异议应对沉淀到团队话术库")

    # ---------- 遗漏动作 ----------
    missed = []
    checks = [
        (["决策", "拍板", "老板", "负责人", "谁定", "定夺"], "确认决策人：「这事最终是您拍板，还是需要老板一起定？」"),
        (["邮箱", "微信", "电话", "号码", "联系方式", "发您", "发你"], "索要联系方式：「方便留个邮箱/微信吗？我把方案和案例发您。」"),
        (["下周", "明天", "周二", "约", "几点", "演示", "下次见"], "约定下次沟通：「我们约下周二上午做一次 15 分钟演示，您看方便吗？」"),
        (["竞品", "别家", "对比", "用过"], "了解竞品经历：「您之前那套主要卡在哪？我们重点帮您避开。」"),
    ]
    for kws, action in checks:
        if not hits(all_lines, kws):
            missed.append(f"遗漏：{action}")
    if not missed:
        missed.append("跟进动作覆盖较完整")

    # ---------- 下一步（结合内容生成）----------
    parts = ["48 小时内完成跟进："]
    if any("切换" in l or "营业" in l for l in cust_lines):
        parts.append("1）发送《平滑切换方案+同行业案例》；")
    if any(k in " ".join(cust_lines) for k in ["价格", "贵", "预算"]):
        parts.append("2）准备一页纸 ROI 说明（省人力/提复购）；")
    parts.append("3）主动约下一次沟通的具体时间（建议电话或 15 分钟演示）。")
    next_step = " ".join(parts)

    # ---------- 金句 ----------
    golden = []
    for line in sales_lines:
        if any(k in line for k in ["相当于", "一个后台", "汇总", "数据", "帮您", "节省"]) and len(line) > 15:
            golden.append(f"「{_short(line, 55)}」")
    golden = _dedup(golden, 3) or ["可从本通通话中进一步提炼价值表达话术"]

    # ---------- 评分 ----------
    score = 78
    score -= 12 * min(len([c for c in compliance if "⚠️" in c]), 2)
    score -= 4 * len([m for m in missed if m.startswith("遗漏")])
    score -= 3 if hits(sales_lines, ["便宜", "优惠", "降价"]) else 0
    score += 3 if did_well and "价值表达" in did_well[-1] else 0
    score = max(35, min(92, score))

    customer = _infer_customer(turns, cust_lines)

    return {
        "customer": customer,
        "demand": demand,
        "budget": budget,
        "objections": objections,
        "did_well": did_well,
        "to_improve": to_improve,
        "missed_actions": missed,
        "next_step": next_step,
        "compliance": compliance,
        "golden_lines": golden,
        "score": str(score),
        "_engine": "offline",
    }


def render_markdown(data: dict) -> str:
    """把复盘 JSON 渲染成易读的 markdown 表格。"""
    c = data.get("customer", {})

    def bullets(items):
        if not items:
            return "—"
        if isinstance(items, str):
            items = [items]
        return "\n".join(f"- {x}" for x in items)

    md = []
    md.append("# 通话复盘表\n")
    md.append(f"**综合评分：{data.get('score', '—')} / 100**\n")
    md.append("## 一、客户画像")
    md.append(f"- **称呼**：{c.get('name', '—')}")
    md.append(f"- **身份**：{c.get('role', '—')}")
    md.append(f"- **行业**：{c.get('industry', '—')}")
    md.append("")
    md.append("## 二、客户需求")
    md.append(bullets(data.get("demand")))
    md.append("")
    md.append(f"## 三、预算/意向\n{data.get('budget', '—')}")
    md.append("")
    md.append("## 四、客户异议")
    md.append(bullets(data.get("objections")))
    md.append("")
    md.append("## 五、我做得好的地方")
    md.append(bullets(data.get("did_well")))
    md.append("")
    md.append("## 六、待改进（重点看这里）")
    md.append(bullets(data.get("to_improve")))
    md.append("")
    md.append("## 七、遗漏的关键动作")
    md.append(bullets(data.get("missed_actions")))
    md.append("")
    md.append(f"## 八、下一步行动\n{data.get('next_step', '—')}")
    md.append("")
    md.append("## 九、合规检查")
    md.append(bullets(data.get("compliance")))
    md.append("")
    md.append("## 十、可复用金句")
    md.append(bullets(data.get("golden_lines")))
    md.append("")
    return "\n".join(md)


def main() -> int:
    parser = argparse.ArgumentParser(description="通话录音复盘助手 Demo")
    parser.add_argument("--input", "-i", required=True,
                        help="通话转写文件路径，支持 .txt/.md/.csv/.xlsx")
    parser.add_argument("--out", "-o", help="可选：把复盘结果导出为 markdown 文件")
    parser.add_argument("--mock", action="store_true", help="强制使用离线示例结果")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"[错误] 找不到输入文件：{args.input}", file=sys.stderr)
        return 1

    try:
        transcript = read_input(args.input)
    except ValueError as e:
        print(f"[错误] {e}", file=sys.stderr)
        return 1
    except Exception as e:  # noqa: BLE001
        print(f"[错误] 读取输入文件失败：{e}", file=sys.stderr)
        return 1

    if not transcript:
        print("[错误] 没从输入文件里读到任何文字（可能是空文件，或 Excel 第一个表是空的）。", file=sys.stderr)
        return 1

    api_key = os.environ.get("LLM_API_KEY")
    base_url = os.environ.get("LLM_BASE_URL", "https://api.openai.com/v1")
    model = os.environ.get("LLM_MODEL", "gpt-4o-mini")

    if args.mock:
        data = mock_result()
    elif not api_key:
        print("[提示] 未检测到 LLM_API_KEY，使用【离线规则分析】（基于你上传的真实内容，不调用 AI）。\n"
              "       想要更细致的分析，可设置 LLM_API_KEY / LLM_BASE_URL / LLM_MODEL。\n",
              file=sys.stderr)
        data = heuristic_result(transcript)
    else:
        print(f"[提示] 正在调用大模型（{model}）生成复盘……\n", file=sys.stderr)
        try:
            data = call_llm(transcript, api_key, base_url, model)
        except urllib.error.HTTPError as e:
            print(f"[错误] 接口返回 {e.code}：{e.read().decode('utf-8', 'ignore')}", file=sys.stderr)
            return 2
        except Exception as e:  # noqa: BLE001
            print(f"[错误] 调用大模型失败：{e}", file=sys.stderr)
            return 2

    md = render_markdown(data)
    print(md)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(md)
        print(f"\n[已保存] 复盘结果写入：{args.out}", file=sys.stderr)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
