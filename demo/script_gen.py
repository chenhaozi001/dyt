#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
客户专属话术生成器（Demo）

用途：输入「客户画像（来自 CRM）」+「我们的产品要点（来自知识库）」，
自动生成下一通电话的开场白、重点话术、异议应对、跟进微信/邮件文案。

运行模式同 call_review.py：
  - 配了 LLM_API_KEY 就调用真实大模型；
  - 没配就用内置离线示例，断网也能现场演示。

用法：
  # 离线演示
  python demo/script_gen.py --customer samples/customer.json

  # 接真实大模型
  export LLM_API_KEY="你的key"
  export LLM_BASE_URL="https://api.openai.com/v1"
  export LLM_MODEL="gpt-4o-mini"
  python demo/script_gen.py --customer samples/customer.json --out 话术.md

customer.json 字段说明见 samples/customer.json。
注意：客户信息请先脱敏。
"""

import argparse
import json
import os
import sys
import urllib.request
import urllib.error


SYSTEM_PROMPT = """你是一名顶尖的电话销售话术教练。
根据用户提供的「客户画像」和「产品要点」，生成一套用于下一通电话的实战话术。
要求：口语化、自然、有人情味，避免书面腔和过度承诺；先讲价值再谈价格。
只输出一个 JSON 对象，不要任何额外文字或 markdown 代码块标记。结构如下：

{
  "opening": "30 秒内的开场白，要能勾起客户兴趣",
  "key_points": ["本通电话要传达的 3 个核心卖点（结合客户痛点）"],
  "objection_handling": [
    {"objection": "预计客户会提的异议", "response": "应对话术"}
  ],
  "questions_to_ask": ["需要向客户确认的关键问题，比如决策流程、预算、时间"],
  "next_step_ask": "本通电话要争取的明确下一步（约时间/发资料等）",
  "follow_up_message": "通话后发给客户的微信/短信文案（简短、有钩子）"
}
"""


def build_user_prompt(customer: dict) -> str:
    return (
        "客户画像（来自 CRM，已脱敏）：\n"
        + json.dumps(customer, ensure_ascii=False, indent=2)
        + "\n\n请据此生成下一通电话的实战话术。"
    )


def call_llm(customer: dict, api_key: str, base_url: str, model: str) -> dict:
    url = base_url.rstrip("/") + "/chat/completions"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": build_user_prompt(customer)},
        ],
        "temperature": 0.5,
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
    if content.startswith("```"):
        lines = [ln for ln in content.splitlines() if not ln.strip().startswith("```")]
        content = "\n".join(lines).strip()
    return json.loads(content)


def mock_result(customer: dict) -> dict:
    name = customer.get("name", "X 先生")
    return {
        "opening": (
            f"{name}您好，我是 XX 公司的小李。上次您提到 30 家门店的会员数据各管各的，"
            "老板想看个总数据很费劲，我这边整理了一个专门解决这个问题的方案，"
            "想花两分钟跟您对一下，看是不是您要的，方便吗？"
        ),
        "key_points": [
            "一个后台看全部门店数据：相当于多请一个数据分析助理，老板随时掌握全局。",
            "和现有收银无缝对接：不用换掉现在的系统，降低切换成本。",
            "平滑切换不影响营业：分门店灰度上线，老店照常开单，新功能逐步切。",
        ],
        "objection_handling": [
            {
                "objection": "担心切换影响门店正常营业",
                "response": (
                    "完全理解，这正是您最该关心的。我们不是一刀切，而是先挑 1 家店试点，"
                    "跑通了再逐店推开，整个过程老店照常开单。我可以把'平滑切换方案'发您看看。"
                ),
            },
            {
                "objection": "之前用别家系统售后差",
                "response": (
                    "您这个顾虑太对了。我们每个客户都有专属对接人，承诺工作时间 30 分钟内响应，"
                    "我可以把我的直线电话给您，有事直接找我。"
                ),
            },
            {
                "objection": "价格比预期高",
                "response": (
                    "我先不谈价格，您算一笔账：现在每月人工汇总数据要花多少时间？"
                    "系统帮您省下来的时间和避免的会员流失，其实早就把成本赚回来了。"
                ),
            },
        ],
        "questions_to_ask": [
            "这个事最终是您拍板，还是要老板一起定？",
            "如果方案合适，大概希望什么时候能用上？",
            "预算上有没有一个大致的范围，我好给您配最合适的版本。",
        ],
        "next_step_ask": "争取约定下周二上午做一次 15 分钟的线上演示，并先要到客户邮箱发方案。",
        "follow_up_message": (
            f"{name}您好，我是 XX 公司小李。这是刚跟您说的'门店平滑切换方案+同行案例'，"
            "您先过目～方便的话下周二上午我给您做个 15 分钟演示，您看可以吗？"
        ),
    }


def render_markdown(data: dict) -> str:
    md = ["# 客户专属话术\n"]
    md.append("## 一、开场白（30 秒）")
    md.append(data.get("opening", "—"))
    md.append("\n## 二、核心卖点")
    for i, p in enumerate(data.get("key_points", []), 1):
        md.append(f"{i}. {p}")
    md.append("\n## 三、异议应对")
    for item in data.get("objection_handling", []):
        md.append(f"- **客户可能说：** {item.get('objection', '')}")
        md.append(f"  - **你可以答：** {item.get('response', '')}")
    md.append("\n## 四、要向客户确认的关键问题")
    for q in data.get("questions_to_ask", []):
        md.append(f"- {q}")
    md.append(f"\n## 五、本通电话要争取的下一步\n{data.get('next_step_ask', '—')}")
    md.append(f"\n## 六、通话后跟进文案（微信/短信）\n> {data.get('follow_up_message', '—')}")
    md.append("")
    return "\n".join(md)


def main() -> int:
    parser = argparse.ArgumentParser(description="客户专属话术生成器 Demo")
    parser.add_argument("--customer", "-c", required=True, help="客户画像 json 文件路径")
    parser.add_argument("--out", "-o", help="可选：导出 markdown 文件")
    parser.add_argument("--mock", action="store_true", help="强制使用离线示例")
    args = parser.parse_args()

    if not os.path.exists(args.customer):
        print(f"[错误] 找不到文件：{args.customer}", file=sys.stderr)
        return 1

    with open(args.customer, "r", encoding="utf-8") as f:
        customer = json.load(f)

    api_key = os.environ.get("LLM_API_KEY")
    base_url = os.environ.get("LLM_BASE_URL", "https://api.openai.com/v1")
    model = os.environ.get("LLM_MODEL", "gpt-4o-mini")

    if args.mock or not api_key:
        if not args.mock:
            print("[提示] 未检测到 LLM_API_KEY，使用离线示例结果演示。\n", file=sys.stderr)
        data = mock_result(customer)
    else:
        print(f"[提示] 正在调用大模型（{model}）生成话术……\n", file=sys.stderr)
        try:
            data = call_llm(customer, api_key, base_url, model)
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
        print(f"\n[已保存] {args.out}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
