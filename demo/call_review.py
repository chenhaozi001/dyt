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
请阅读用户提供的「通话转写文字」，输出一份结构化的销售复盘。
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

    if args.mock or not api_key:
        if not args.mock:
            print("[提示] 未检测到 LLM_API_KEY，使用离线示例结果演示。\n"
                  "       如需接入真实大模型，请设置 LLM_API_KEY / LLM_BASE_URL / LLM_MODEL。\n",
                  file=sys.stderr)
        data = mock_result()
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
