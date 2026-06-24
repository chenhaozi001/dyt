#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
知识库即问即答（RAG Demo）

用途：把知识库文档放到一个文件夹里（脱敏的 .md / .txt），
你用大白话提问，脚本先「检索」出最相关的片段，再让大模型「带出处」回答。
这就是 RAG（检索增强生成）——能有效减少大模型乱编。

特点：只用 Python 标准库，无需安装任何依赖，Cursor 里直接跑。
检索用的是轻量关键词打分（中文按字切分 + 词频），够 Demo 用；
真要上生产可换成向量检索，这里先把原理讲清楚。

运行模式同其他脚本：配 LLM_API_KEY 走真实大模型；不配则用离线示例答案演示。

用法：
  # 离线演示
  python demo/kb_qa.py --kb knowledge_base --q "你们支持和收银系统对接吗？"

  # 接真实大模型
  export LLM_API_KEY="你的key"
  python demo/kb_qa.py --kb knowledge_base --q "切换系统会影响门店营业吗？"
"""

import argparse
import glob
import json
import os
import re
import sys
import urllib.request
import urllib.error
from collections import Counter


def load_docs(kb_dir: str):
    """读取知识库目录下所有 .md / .txt，按段落切成 chunk。"""
    chunks = []
    paths = glob.glob(os.path.join(kb_dir, "**", "*.md"), recursive=True)
    paths += glob.glob(os.path.join(kb_dir, "**", "*.txt"), recursive=True)
    for path in sorted(set(paths)):
        with open(path, "r", encoding="utf-8") as f:
            text = f.read()
        for block in re.split(r"\n\s*\n", text):
            block = block.strip()
            if len(block) >= 8:
                chunks.append({"source": os.path.basename(path), "text": block})
    return chunks


def tokenize(text: str):
    """简单分词：英文/数字按词，中文按单字。够关键词打分用。"""
    text = text.lower()
    tokens = re.findall(r"[a-z0-9]+", text)
    tokens += re.findall(r"[\u4e00-\u9fff]", text)
    return tokens


def retrieve(chunks, query, top_k=3):
    """对每个 chunk 用「query 词在其中出现的次数」打分，取最相关的 top_k 个。"""
    q_tokens = set(tokenize(query))
    scored = []
    for ch in chunks:
        counts = Counter(tokenize(ch["text"]))
        score = sum(counts[t] for t in q_tokens)
        if score > 0:
            scored.append((score, ch))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [ch for _, ch in scored[:top_k]]


SYSTEM_PROMPT = """你是公司内部知识库问答助手，服务于电话销售同事。
只能依据用户提供的「知识库片段」回答问题，不能编造。
如果片段里没有答案，就明确说「知识库里没有查到，建议核实」。
回答要简洁、口语化，方便销售直接对客户讲。
最后用一行注明引用了哪些来源文件。"""


def build_user_prompt(query, contexts):
    ctx = "\n\n".join(
        f"[来源：{c['source']}]\n{c['text']}" for c in contexts
    )
    return f"知识库片段如下：\n\n{ctx}\n\n---\n客户/我的问题：{query}\n请据此回答。"


def call_llm(query, contexts, api_key, base_url, model):
    url = base_url.rstrip("/") + "/chat/completions"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": build_user_prompt(query, contexts)},
        ],
        "temperature": 0.2,
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url, data=data,
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        body = json.loads(resp.read().decode("utf-8"))
    return body["choices"][0]["message"]["content"].strip()


def mock_answer(query, contexts):
    """离线演示：直接把检索到的片段拼成一个'看起来像答案'的回复。"""
    if not contexts:
        return "知识库里没有查到相关内容，建议核实后再答复客户。"
    snippet = contexts[0]["text"].replace("\n", " ")
    sources = "、".join(sorted({c["source"] for c in contexts}))
    return (
        f"（离线示例答案，基于检索到的知识库片段）\n"
        f"针对「{query}」，知识库里的相关说明是：{snippet}\n\n"
        f"引用来源：{sources}"
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="知识库即问即答 RAG Demo")
    parser.add_argument("--kb", required=True, help="知识库文件夹路径（放 .md/.txt）")
    parser.add_argument("--q", required=True, help="你的问题")
    parser.add_argument("--top-k", type=int, default=3, help="检索片段数量")
    parser.add_argument("--mock", action="store_true", help="强制离线示例")
    args = parser.parse_args()

    if not os.path.isdir(args.kb):
        print(f"[错误] 知识库目录不存在：{args.kb}", file=sys.stderr)
        return 1

    chunks = load_docs(args.kb)
    if not chunks:
        print(f"[错误] {args.kb} 里没有找到 .md/.txt 文档。", file=sys.stderr)
        return 1

    contexts = retrieve(chunks, args.q, top_k=args.top_k)

    print(f"问题：{args.q}\n")
    print(f"检索到 {len(contexts)} 个相关片段：")
    for i, c in enumerate(contexts, 1):
        preview = c["text"].replace("\n", " ")[:50]
        print(f"  {i}. [{c['source']}] {preview}…")
    print()

    api_key = os.environ.get("LLM_API_KEY")
    base_url = os.environ.get("LLM_BASE_URL", "https://api.openai.com/v1")
    model = os.environ.get("LLM_MODEL", "gpt-4o-mini")

    if args.mock or not api_key:
        if not args.mock:
            print("[提示] 未检测到 LLM_API_KEY，使用离线示例答案。\n", file=sys.stderr)
        answer = mock_answer(args.q, contexts)
    else:
        print(f"[提示] 正在调用大模型（{model}）生成答案……\n", file=sys.stderr)
        try:
            answer = call_llm(args.q, contexts, api_key, base_url, model)
        except urllib.error.HTTPError as e:
            print(f"[错误] 接口返回 {e.code}：{e.read().decode('utf-8', 'ignore')}", file=sys.stderr)
            return 2
        except Exception as e:  # noqa: BLE001
            print(f"[错误] 调用大模型失败：{e}", file=sys.stderr)
            return 2

    print("=" * 40)
    print(answer)
    print("=" * 40)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
