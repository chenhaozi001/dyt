"""通话复盘助手 · Streamlit 应用入口。

上传录音转写文件（Excel / CSV / txt / md），一键生成销售复盘。
- 填写免费的 Groq API Key 时，使用大模型生成细致复盘。
- 不填 Key 时，使用离线规则分析（不联网）。
"""

from __future__ import annotations

import streamlit as st

from analysis import (
    DEFAULT_GROQ_MODEL,
    groq_review,
    offline_review,
    parse_uploaded_file,
)

st.set_page_config(page_title="通话复盘助手", page_icon="📞", layout="centered")

# --------------------------------------------------------------------------- #
# 侧边栏：AI 接口设置
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("⚙️ AI 接口设置")
    st.write("要分析真实内容需要一个 AI 接口 Key。")
    st.markdown(
        "推荐用免费的 **Groq**（不用信用卡），领 Key："
        "[console.groq.com/keys](https://console.groq.com/keys)"
    )
    st.caption("没填 Key 也能用：会用离线规则分析你上传的真实文件。填了免费 Key 分析会更细致。")

    api_key = st.text_input("Groq API Key", type="password", placeholder="gsk_...")
    model = st.text_input("模型", value=DEFAULT_GROQ_MODEL)


# --------------------------------------------------------------------------- #
# 主区域
# --------------------------------------------------------------------------- #
st.title("📞 通话复盘助手")
st.write("上传录音转写文件（Excel / CSV / txt），一键生成销售复盘 · 请先脱敏后再上传")

uploaded = st.file_uploader(
    "Upload",
    type=["txt", "md", "csv", "xlsx", "xls"],
    help="200MB per file • TXT, MD, CSV, XLSX",
)

start = st.button("🚀 开始分析", type="primary", use_container_width=True)

if start:
    if uploaded is None:
        st.warning("请先上传一个转写文件再开始分析。")
        st.stop()

    with st.spinner("正在解析文件…"):
        try:
            call = parse_uploaded_file(uploaded.name, uploaded.getvalue())
        except Exception as exc:  # noqa: BLE001
            st.error(f"文件解析失败：{exc}")
            st.stop()

    if not call.raw_text.strip():
        st.error("没有从文件中读到任何文本内容，请检查文件。")
        st.stop()

    with st.expander("查看解析出的转写内容", expanded=False):
        st.text(call.raw_text[:5000] + ("…" if len(call.raw_text) > 5000 else ""))

    report = None
    if api_key.strip():
        with st.spinner("正在用 AI 生成复盘…"):
            try:
                report = groq_review(call, api_key.strip(), model=model.strip() or DEFAULT_GROQ_MODEL)
                st.success("AI 复盘完成 ✅")
            except Exception as exc:  # noqa: BLE001
                st.warning(f"AI 接口调用失败（{exc}），已自动切换到离线规则分析。")

    if report is None:
        with st.spinner("正在进行离线规则分析…"):
            report = offline_review(call)

    st.markdown(report)
    st.download_button(
        "⬇️ 下载复盘报告 (Markdown)",
        data=report,
        file_name="通话复盘报告.md",
        mime="text/markdown",
        use_container_width=True,
    )

st.divider()
st.caption(
    "⚠️ 数据安全：录音/转写涉及客户隐私，上传前请去掉真实姓名、电话、公司名。"
    "若部署在公开链接上，任何拿到链接的人都能访问，请勿上传未脱敏的真实数据。"
)
