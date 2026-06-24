# -*- coding: utf-8 -*-
"""
通话复盘助手 · 在线版（Streamlit）

这个版本是为「免费部署到网上、点链接就能用」准备的。
推荐部署到 Streamlit Community Cloud（免费、用 GitHub 登录、不用信用卡）。
详细部署步骤见仓库里的「部署指南.md」。

本地预览（可选）：
  pip install -r requirements.txt
  streamlit run app.py
"""

import os
import sys
import tempfile

import streamlit as st

# 复用 demo/call_review.py 里的解析与复盘逻辑
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "demo"))
import call_review  # noqa: E402

ALLOWED_EXT = ("txt", "md", "csv", "xlsx")


def get_secret(name, default=""):
    """优先读 Streamlit secrets，其次读环境变量。"""
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.environ.get(name, default)


st.set_page_config(page_title="通话复盘助手", page_icon="📞", layout="centered")

st.markdown(
    """
    <style>
      .block-container { max-width: 820px; }
      h1 { color: #1f4e79; }
      .stAlert { border-radius: 10px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📞 通话复盘助手")
st.caption("上传录音转写文件（Excel / CSV / txt），一键生成销售复盘 · 请先脱敏后再上传")

# ---------- 侧边栏：AI 接口设置 ----------
with st.sidebar:
    st.header("⚙️ AI 接口设置")
    st.caption(
        "要分析真实内容需要一个 AI 接口 Key。\n\n"
        "推荐用**免费**的 Groq（不用信用卡），\n"
        "领 Key：https://console.groq.com/keys"
    )
    default_key = get_secret("LLM_API_KEY", "")
    default_base = get_secret("LLM_BASE_URL", "https://api.groq.com/openai/v1")
    default_model = get_secret("LLM_MODEL", "llama-3.3-70b-versatile")

    api_key = st.text_input("API Key", value=default_key, type="password",
                            help="不会被保存，仅本次分析使用")
    base_url = st.text_input("接口地址 Base URL", value=default_base)
    model = st.text_input("模型名称", value=default_model)

    if not api_key:
        st.info("没填 Key 也能用：会用**离线规则**分析你上传的真实文件。填了免费 Key 分析会更细致。")

# ---------- 主体 ----------
uploaded = st.file_uploader(
    "选择或拖入文件",
    type=list(ALLOWED_EXT),
    help="支持 Excel(.xlsx) / CSV / 文本(.txt)，请使用脱敏后的转写文件",
)

go = st.button("🚀 开始分析", type="primary", use_container_width=True)

st.warning(
    "⚠️ 数据安全：录音/转写涉及客户隐私，上传前请去掉真实姓名、电话、公司名。"
    "若部署在公开链接上，任何拿到链接的人都能访问，请勿上传未脱敏的真实数据。",
    icon="🔒",
)


def render_result(data, demo_mode):
    if demo_mode:
        st.info("💡 当前是**离线分析**（没填 AI Key 也能用）：以下结果是根据你上传的真实文件、"
                "用关键词规则分析得出的，可直接参考。想要更细致的分析，可在左侧填入免费 AI Key 升级。")

    c = data.get("customer", {})
    score = data.get("score", "—")
    st.metric("综合评分", f"{score} / 100")

    st.subheader("客户画像")
    st.write(f"**称呼：**{c.get('name', '—')}　|　**身份：**{c.get('role', '—')}　|　**行业：**{c.get('industry', '—')}")

    def section(title, items, kind="normal"):
        st.subheader(title)
        if not items:
            st.write("—")
            return
        if isinstance(items, str):
            items = [items]
        text = "\n".join(f"- {x}" for x in items)
        if kind == "warn":
            st.warning(text)
        elif kind == "error":
            st.error(text)
        elif kind == "success":
            st.success(text)
        else:
            st.markdown(text)

    section("客户需求", data.get("demand"))
    st.subheader("预算 / 意向")
    st.write(data.get("budget", "—"))
    section("客户异议", data.get("objections"))
    section("我做得好的地方", data.get("did_well"), "success")
    section("待改进（重点）", data.get("to_improve"), "error")
    section("遗漏的关键动作", data.get("missed_actions"))
    st.subheader("下一步行动")
    st.write(data.get("next_step", "—"))
    section("合规检查", data.get("compliance"), "warn")
    section("可复用金句", data.get("golden_lines"), "success")


if go:
    if uploaded is None:
        st.error("请先选择一个文件。")
        st.stop()

    ext = "." + uploaded.name.rsplit(".", 1)[-1].lower()
    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tf:
        tf.write(uploaded.getbuffer())
        tmp_path = tf.name
    try:
        transcript = call_review.read_input(tmp_path)
    except Exception as e:  # noqa: BLE001
        st.error(f"读取文件失败：{e}")
        os.unlink(tmp_path)
        st.stop()
    os.unlink(tmp_path)

    if not transcript:
        st.error("没从文件里读到文字（可能是空文件，或 Excel 第一个表是空的）。")
        st.stop()

    if api_key:
        with st.spinner("正在调用 AI 生成复盘…"):
            try:
                data = call_review.call_llm(transcript, api_key, base_url, model)
            except Exception as e:  # noqa: BLE001
                st.error(f"调用 AI 失败：{e}\n\n请检查左侧的 Key / 接口地址 / 模型名是否正确。")
                st.stop()
        render_result(data, demo_mode=False)
    else:
        with st.spinner("正在离线分析你的文件…"):
            data = call_review.heuristic_result(transcript)
        render_result(data, demo_mode=True)
