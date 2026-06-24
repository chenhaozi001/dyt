"""通话复盘的核心分析逻辑。

包含两条路径：
1. 在线分析：调用 Groq（OpenAI 兼容接口）的大模型生成复盘。
2. 离线分析：纯规则、不联网，根据关键词与对话结构生成复盘。
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

import pandas as pd
import requests

# Groq 提供的 OpenAI 兼容接口
GROQ_ENDPOINT = "https://api.groq.com/openai/v1/chat/completions"
DEFAULT_GROQ_MODEL = "llama-3.3-70b-versatile"

# 关键词词典：用于离线规则分析
KEYWORDS = {
    "价格/报价": ["价格", "报价", "多少钱", "费用", "预算", "成本", "优惠", "折扣", "性价比"],
    "需求/痛点": ["需要", "想要", "希望", "问题", "困难", "痛点", "麻烦", "解决", "提升", "效率"],
    "异议/顾虑": ["但是", "不过", "太贵", "考虑一下", "再看看", "担心", "不确定", "比较一下", "犹豫", "贵了"],
    "竞品对比": ["竞品", "对手", "别家", "其他家", "友商", "对比"],
    "下一步/承诺": ["跟进", "回头", "下周", "明天", "联系", "发资料", "安排", "demo", "试用", "方案", "合同", "签约"],
}

# 常见的销售方称呼，用于区分说话人角色
SALES_TOKENS = ["销售", "客服", "顾问", "坐席", "客户经理", "sales", "agent", "rep"]
CUSTOMER_TOKENS = ["客户", "顾客", "用户", "对方", "customer", "client", "买家"]


@dataclass
class Turn:
    """一轮发言。"""

    speaker: str  # 原始说话人标签
    role: str  # 归一化角色：sales / customer / unknown
    text: str


@dataclass
class ParsedCall:
    """解析后的整通通话。"""

    raw_text: str
    turns: list[Turn] = field(default_factory=list)


def _decode_bytes(data: bytes) -> str:
    """尽量稳妥地把字节解码成文本（兼容 utf-8 / gbk）。"""
    for enc in ("utf-8", "utf-8-sig", "gb18030", "gbk", "latin-1"):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return data.decode("utf-8", errors="ignore")


def _classify_role(speaker: str) -> str:
    s = (speaker or "").lower()
    for token in SALES_TOKENS:
        if token in s:
            return "sales"
    for token in CUSTOMER_TOKENS:
        if token in s:
            return "customer"
    return "unknown"


def _split_speaker_line(line: str) -> tuple[str, str] | None:
    """把形如 "销售：你好" / "客户: ..." / "A | ..." 的行拆成 (说话人, 内容)。"""
    m = re.match(r"^\s*([^:：\t|]{1,20})\s*[:：\t|]\s*(.+)$", line)
    if m:
        speaker, text = m.group(1).strip(), m.group(2).strip()
        # 避免把时间戳（如 00:12）误判为说话人
        if re.fullmatch(r"[\d\s.\-/]+", speaker):
            return None
        return speaker, text
    return None


def _turns_from_text(text: str) -> list[Turn]:
    turns: list[Turn] = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        parsed = _split_speaker_line(line)
        if parsed:
            speaker, content = parsed
            turns.append(Turn(speaker=speaker, role=_classify_role(speaker), text=content))
        else:
            # 没有说话人前缀的行，归到上一轮或标记未知
            if turns:
                turns[-1].text += " " + line
            else:
                turns.append(Turn(speaker="未知", role="unknown", text=line))
    return turns


def _turns_from_dataframe(df: pd.DataFrame) -> tuple[str, list[Turn]]:
    """从表格中提取说话人与文本列，兼容常见列名。"""
    cols = {str(c).strip().lower(): c for c in df.columns}

    def find_col(candidates: list[str]):
        for cand in candidates:
            for low, original in cols.items():
                if cand in low:
                    return original
        return None

    speaker_col = find_col(["说话人", "角色", "speaker", "role", "name", "姓名"])
    text_col = find_col(["内容", "文本", "对话", "text", "content", "utterance", "转写", "句子"])

    turns: list[Turn] = []
    lines: list[str] = []
    if text_col is not None:
        for _, row in df.iterrows():
            content = str(row[text_col]).strip()
            if not content or content.lower() == "nan":
                continue
            speaker = str(row[speaker_col]).strip() if speaker_col is not None else "未知"
            turns.append(Turn(speaker=speaker, role=_classify_role(speaker), text=content))
            lines.append(f"{speaker}：{content}")
    else:
        # 没识别出文本列：把每行所有单元格拼起来当作文本
        for _, row in df.iterrows():
            content = " ".join(str(v) for v in row.values if str(v).strip() and str(v).lower() != "nan")
            if content:
                turns.extend(_turns_from_text(content))
                lines.append(content)
    return "\n".join(lines), turns


def parse_uploaded_file(filename: str, data: bytes) -> ParsedCall:
    """根据文件扩展名解析上传的通话文件。"""
    name = (filename or "").lower()
    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(data))
        raw_text, turns = _turns_from_dataframe(df)
    elif name.endswith(".csv"):
        text = _decode_bytes(data)
        try:
            df = pd.read_csv(io.StringIO(text))
            raw_text, turns = _turns_from_dataframe(df)
        except Exception:
            raw_text = text
            turns = _turns_from_text(text)
    else:  # txt / md / 其它纯文本
        raw_text = _decode_bytes(data)
        turns = _turns_from_text(raw_text)

    return ParsedCall(raw_text=raw_text, turns=turns)


# --------------------------------------------------------------------------- #
# 离线规则分析
# --------------------------------------------------------------------------- #
def _count_keywords(text: str) -> dict[str, list[str]]:
    hits: dict[str, list[str]] = {}
    for label, words in KEYWORDS.items():
        found = sorted({w for w in words if w.lower() in text.lower()})
        if found:
            hits[label] = found
    return hits


def offline_review(call: ParsedCall) -> str:
    """不联网，基于规则生成一份结构化复盘报告（Markdown）。"""
    text = call.raw_text
    turns = call.turns
    char_total = len(re.sub(r"\s", "", text))

    sales_turns = [t for t in turns if t.role == "sales"]
    customer_turns = [t for t in turns if t.role == "customer"]
    sales_chars = sum(len(t.text) for t in sales_turns)
    customer_chars = sum(len(t.text) for t in customer_turns)

    questions = [t for t in turns if ("？" in t.text or "?" in t.text or t.text.endswith("吗"))]
    sales_questions = [t for t in questions if t.role == "sales"]

    hits = _count_keywords(text)

    # 简单评分（满分 5）
    score_components = []
    talk_ratio_note = "无法区分说话人"
    if sales_chars + customer_chars > 0:
        sales_ratio = sales_chars / (sales_chars + customer_chars)
        talk_ratio_note = f"销售约 {sales_ratio:.0%}，客户约 {1 - sales_ratio:.0%}"
        # 健康的销售对话客户应有足够表达空间
        score_components.append(5 if 0.35 <= sales_ratio <= 0.6 else 3)
    if sales_questions:
        score_components.append(min(5, 2 + len(sales_questions)))
    if "下一步/承诺" in hits:
        score_components.append(5)
    else:
        score_components.append(2)
    if "需求/痛点" in hits:
        score_components.append(4)
    overall = round(sum(score_components) / len(score_components), 1) if score_components else 3.0

    lines: list[str] = []
    lines.append("## 📋 通话复盘报告（离线规则版）")
    lines.append("> 未填写 AI Key，以下为基于关键词与对话结构的离线分析。填入免费 Groq Key 可获得更细致的 AI 复盘。")
    lines.append("")
    lines.append("### 一、对话基本统计")
    lines.append(f"- 总发言轮数：**{len(turns)}**")
    lines.append(f"- 有效字数：**{char_total}**")
    lines.append(f"- 说话占比：**{talk_ratio_note}**")
    lines.append(f"- 销售主动提问次数：**{len(sales_questions)}**")
    lines.append("")

    lines.append("### 二、客户需求与痛点")
    if "需求/痛点" in hits:
        lines.append(f"- 检测到需求/痛点相关表达：{('、'.join(hits['需求/痛点']))}")
        for t in customer_turns:
            if any(w in t.text for w in KEYWORDS["需求/痛点"]):
                lines.append(f"  - 客户原话：「{t.text[:80]}」")
    else:
        lines.append("- ⚠️ 未明显捕捉到客户需求/痛点表达，建议复盘是否做了足够的需求挖掘。")
    lines.append("")

    lines.append("### 三、异议与风险")
    if "异议/顾虑" in hits:
        lines.append(f"- 检测到异议/顾虑信号：{('、'.join(hits['异议/顾虑']))}")
        lines.append("- 建议针对这些顾虑准备标准话术，下次主动化解。")
    else:
        lines.append("- 未检测到明显异议词。注意：客户没说不代表没有顾虑，可主动确认。")
    if "竞品对比" in hits:
        lines.append(f"- 出现竞品/对比相关内容：{('、'.join(hits['竞品对比']))}，需强化差异化价值。")
    if "价格/报价" in hits:
        lines.append(f"- 涉及价格讨论：{('、'.join(hits['价格/报价']))}，确认是否在价值铺垫充分后再谈价。")
    lines.append("")

    lines.append("### 四、销售表现")
    if score_components and (sales_chars / max(sales_chars + customer_chars, 1)) > 0.65:
        lines.append("- ⚠️ 销售说得偏多，客户表达空间不足，建议多用开放式提问、多倾听。")
    if sales_questions:
        lines.append(f"- ✅ 有主动提问（{len(sales_questions)} 次），有助于挖掘需求。")
    else:
        lines.append("- ⚠️ 几乎没有主动提问，建议增加开放式问题。")
    lines.append("")

    lines.append("### 五、下一步行动建议")
    if "下一步/承诺" in hits:
        lines.append(f"- ✅ 通话中明确了后续动作：{('、'.join(hits['下一步/承诺']))}。请确保按时跟进。")
    else:
        lines.append("- ⚠️ 未明确下一步动作。每通电话结束前都应约定下一步（发资料/约 demo/再次通话时间）。")
    lines.append("- 复盘三件事：本次目标是否达成？客户的真实顾虑是什么？下一步谁、在何时、做什么？")
    lines.append("")

    lines.append(f"### 六、综合评分：**{overall} / 5**")
    lines.append("> 评分基于说话占比、提问情况、需求挖掘与下一步明确度的启发式估算，仅供参考。")

    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# 在线 AI 分析（Groq）
# --------------------------------------------------------------------------- #
SYSTEM_PROMPT = (
    "你是一位资深的销售教练。请基于用户提供的通话转写记录，输出一份结构化的销售复盘报告，"
    "使用中文与 Markdown。报告需包含以下部分：\n"
    "1. 通话概览（双方、主题、阶段）\n"
    "2. 客户需求与痛点\n"
    "3. 客户异议与顾虑及其根因\n"
    "4. 销售表现：亮点与不足（含说话占比、提问质量、倾听）\n"
    "5. 关键节点与转折\n"
    "6. 下一步行动建议（具体、可执行）\n"
    "7. 综合评分（满分 5 分）并给出理由。\n"
    "请客观、具体，多引用对话中的原话作为依据。"
)


def groq_review(call: ParsedCall, api_key: str, model: str = DEFAULT_GROQ_MODEL, timeout: int = 60) -> str:
    """调用 Groq 接口生成复盘报告。失败时抛出异常，由上层兜底到离线分析。"""
    transcript = call.raw_text.strip()
    # 控制长度，避免超出上下文
    max_chars = 24000
    if len(transcript) > max_chars:
        transcript = transcript[:max_chars] + "\n…（转写过长，已截断）"

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"以下是通话转写记录，请生成复盘报告：\n\n{transcript}"},
        ],
        "temperature": 0.3,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    resp = requests.post(GROQ_ENDPOINT, json=payload, headers=headers, timeout=timeout)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"]
