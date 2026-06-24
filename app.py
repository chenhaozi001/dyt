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
st.caption("上传录音转写文件（Excel / CSV / txt），智能生成结构化销售复盘")

# API 配置仅从后台 Secrets / 环境变量读取，页面上不展示（演示更简洁）
api_key = get_secret("LLM_API_KEY", "")
base_url = get_secret("LLM_BASE_URL", "https://api.groq.com/openai/v1")
model = get_secret("LLM_MODEL", "llama-3.3-70b-versatile")

# ---------- 主体 ----------
uploaded = st.file_uploader(
    "选择或拖入文件",
    type=list(ALLOWED_EXT),
    help="支持 Excel(.xlsx) / CSV / 文本(.txt)，请使用脱敏后的转写文件",
)

go = st.button("🚀 开始分析", type="primary", use_container_width=True)

st.warning(
    "⚠️ 数据安全：上传前请对录音转写做脱敏处理（去掉真实姓名、电话、公司名）。",
    icon="🔒",
)


def render_result(data):
    st.success("✅ 分析完成")
    c = data.get("customer", {})
    score = data.get("score", "—")
    col1, col2, col3 = st.columns(3)
    col1.metric("综合评分", f"{score} / 100")
    col2.metric("客户行业", c.get("industry", "—"))
    col3.metric("通话阶段", "意向沟通" if data.get("objections") else "—")

    st.subheader("👤 客户画像")
    st.markdown(
        f"| 称呼 | 身份 | 行业 |\n|------|------|------|\n"
        f"| {c.get('name', '—')} | {c.get('role', '—')} | {c.get('industry', '—')} |"
    )

    def section(title, items, kind="normal"):
        st.subheader(title)
        if not items:
            st.write("—")
            return
        if isinstance(items, str):
            items = [items]
        for item in items:
            if kind == "warn":
                st.warning(item)
            elif kind == "error":
                st.error(item)
            elif kind == "success":
                st.success(item)
            else:
                st.markdown(f"- {item}")

    section("📋 客户需求", data.get("demand"))
    st.subheader("💰 预算 / 意向")
    st.write(data.get("budget", "—"))
    section("🚧 客户异议（含建议应对）", data.get("objections"))
    section("✅ 我做得好的地方", data.get("did_well"), "success")
    section("🔧 待改进（原话 + 问题 + 改法）", data.get("to_improve"), "error")
    section("⏭️ 遗漏的关键动作", data.get("missed_actions"))
    st.subheader("📌 下一步行动")
    st.info(data.get("next_step", "—"))
    section("⚖️ 合规检查", data.get("compliance"), "warn")
    section("💬 可复用金句", data.get("golden_lines"), "success")

    with st.expander("📥 导出复盘结果"):
        md = call_review.render_markdown(data)
        st.download_button("下载 Markdown 文件", md, file_name="通话复盘.md", mime="text/markdown")
        st.code(md, language="markdown")


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
        with st.spinner("正在智能分析通话内容，请稍候…"):
            try:
                data = call_review.call_llm(transcript, api_key, base_url, model)
            except Exception as e:  # noqa: BLE001
                st.error(f"分析失败：{e}")
                st.stop()
    else:
        with st.spinner("正在智能分析通话内容，请稍候…"):
            data = call_review.heuristic_result(transcript)
    render_result(data)
