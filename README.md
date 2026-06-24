# 📞 通话复盘助手 (dyt)

一个基于 [Streamlit](https://streamlit.io/) 的销售通话复盘小工具：上传录音转写文件（Excel / CSV / txt / md），一键生成结构化的销售复盘报告。

- **填写免费的 Groq API Key**：调用大模型生成细致的复盘报告。
- **不填 Key 也能用**：使用离线规则分析（不联网），根据关键词与对话结构生成复盘。

> ⚠️ **数据安全**：录音/转写涉及客户隐私，上传前请去掉真实姓名、电话、公司名。若部署在公开链接上，任何拿到链接的人都能访问，请勿上传未脱敏的真实数据。

## 本地运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

浏览器会自动打开 `http://localhost:8501`。

## 部署到 Streamlit Community Cloud

1. 把本仓库推送到 GitHub。
2. 打开 [share.streamlit.io](https://share.streamlit.io/)，用 GitHub 账号登录。
3. 点击 **New app**，选择本仓库、分支以及主文件 `app.py`，点击 **Deploy**。
4. （可选）在应用的 **Settings → Secrets** 里可以预置 `GROQ_API_KEY`，否则在侧边栏手动填入即可。

> 注意：Streamlit Community Cloud 的 App 在长时间无人访问后会进入“休眠（sleep）”状态，重新访问时会从 GitHub 仓库重新拉取代码并重启。请务必保证仓库里始终保留这套源码，否则重启会失败。

## 获取免费的 Groq API Key

访问 [console.groq.com/keys](https://console.groq.com/keys)，注册后即可免费领取（无需信用卡）。

## 支持的文件格式

| 格式 | 说明 |
| --- | --- |
| `.txt` / `.md` | 纯文本转写，支持 `说话人：内容` 形式分行 |
| `.csv` | 自动识别 `说话人/角色` 与 `内容/文本` 等常见列名 |
| `.xlsx` / `.xls` | 同上，使用 openpyxl 读取 |

## 项目结构

```
.
├── app.py            # Streamlit 应用入口（界面）
├── analysis.py       # 文件解析 + 在线/离线复盘分析逻辑
├── requirements.txt  # 依赖
└── README.md
```
