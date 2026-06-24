# AI 电话销售提效工具包（纯 Cursor 版）

为「软件电话销售」的 **AI 学习成果分享** 准备的一套可演示、可复用的工具包。
基于你手上的三类数据，做了三个小工具，**全程只用 Cursor，不依赖其它平台**。

> 🌟 **不懂代码？看这里 →** 直接用「[网页版通话复盘助手](#-网页版通话复盘助手不用懂代码)」，
> 打开网页、上传文件、点一下就出结果，**不用敲任何命令**。

| 数据源 | 工具 | 作用 |
|---|---|---|
| 通话录音（转写文字） | `demo/call_review.py` | 自动生成结构化复盘表 |
| CRM 销售记录 | `demo/script_gen.py` | 生成客户专属电话话术 |
| 知识库网站 | `demo/kb_qa.py` | 知识库即问即答（RAG，带出处） |

> 配套还有：`分享PPT文案.md`（逐页文案+演讲稿）、`AI学习成果分享.pptx`（已生成好的 PPT）、`prompts/`（可直接粘进 Cursor 对话框的 Prompt 模板）。

---

## 🌟 网页版通话复盘助手（不用懂代码）

如果你不想敲命令，用这个：在网页上传录音转写文件，点一下就出复盘结果。

### 启动（任选一种，只需一次）

- **方式 A（最简单）**：在 Cursor 里打开 `webapp/server.py`，点右上角的「运行 ▶」按钮。
- **方式 B（双击启动）**：
  - Mac：双击 `webapp/启动网站.command`
  - Windows：双击 `webapp/启动网站.bat`

启动后浏览器会**自动打开** `http://localhost:8000`（没自动打开就手动在浏览器输入这个地址）。

### 使用

1. 把录音转写文件（Excel/CSV/txt，**先脱敏**）拖进网页，或点击选择。
2. 点「开始分析」，几秒后就能看到复盘结果，可以「打印 / 存为 PDF」。
3. **想分析真实内容**：展开网页里的「⚙️ 高级设置」，把公司给的大模型
   `API Key`（必要时还有接口地址、模型名）填进去再分析即可——**不用碰命令行**。
   不填则显示示例结果，方便先看效果。

> 这个网站只在你自己电脑上运行，数据不会上传到任何服务器；填的 Key 仅本次使用、用完即弃。

---

## 快速开始（在 Cursor 里，命令行方式）

无需安装任何依赖，三个脚本都只用 Python 标准库。**没配 API Key 时会自动用离线示例演示，断网也能跑、不会泄露数据。**

```bash
# 1) 通话复盘（支持 .txt/.md/.csv/.xlsx）
python demo/call_review.py --input samples/sample_call.txt
python demo/call_review.py --input 我的录音转写.xlsx     # 用 Excel 文件

# 2) 客户话术生成
python demo/script_gen.py --customer samples/customer.json

# 3) 知识库问答
python demo/kb_qa.py --kb knowledge_base --q "切换系统会影响门店营业吗？"
```

### 接入真实大模型（可选）

把环境变量配好，脚本就会真的调用大模型（任意 OpenAI 兼容接口，包括公司内部接口）：

```bash
export LLM_API_KEY="你的key"
export LLM_BASE_URL="https://api.openai.com/v1"   # 或公司内部兼容地址
export LLM_MODEL="gpt-4o-mini"
```

---

## 两种使用方式

1. **跑脚本**（适合现场演示、批量处理）：见上面的命令。
2. **直接在 Cursor 对话框用 Prompt**（适合日常、零门槛）：
   打开 `prompts/` 里的模板，复制粘贴到 Cursor 对话框，把【】里的内容换成你脱敏后的数据即可。

---

## 目录结构

```
.
├── README.md                      本说明
├── 分享PPT文案.md                 逐页 PPT 文案 + 演讲稿
├── AI学习成果分享.pptx            已生成好的 PPT（含演讲稿备注）
├── requirements.txt               仅生成 PPT 时需要的依赖
├── webapp/                        网页版通话复盘助手（不用懂代码）
│   ├── server.py                  网站程序（纯标准库，无需依赖）
│   ├── 启动网站.command           Mac 双击启动
│   └── 启动网站.bat               Windows 双击启动
├── tools/
│   └── build_pptx.py              把分享内容生成为 .pptx 的脚本
├── demo/
│   ├── call_review.py             通话复盘助手（支持 txt/csv/xlsx）
│   ├── script_gen.py              客户话术生成器
│   └── kb_qa.py                   知识库问答（RAG）
├── prompts/
│   ├── 01_通话复盘_Prompt.md
│   ├── 02_客户话术生成_Prompt.md
│   └── 03_知识库问答_Prompt.md
├── samples/
│   ├── sample_call.txt            脱敏通话转写示例
│   └── customer.json              脱敏客户画像示例
└── knowledge_base/                脱敏知识库示例文档（.md）
    ├── 产品功能.md
    ├── 实施与售后.md
    └── 价格与套餐.md
```

---

## ⚠️ 数据安全红线（务必遵守）

- 通话录音、CRM、知识库都涉及**客户隐私 / 公司机密**，使用前请**脱敏**（去掉真实姓名、电话、公司名）。
- 优先使用**公司批准的 AI 工具和接口**；用公网模型时**绝不上传真实客户数据**。
- 本仓库所有示例数据均为**虚构脱敏**，仅供演示。

---

## 只想演示「通话复盘助手」？（最简流程）

1. 把你的录音转写文件准备好（支持 `.xlsx` Excel / `.csv` / `.txt`），**先脱敏**。
   - Excel 里不管是「时间 / 角色 / 内容」三列，还是把对话写在一列里，脚本都能读。
2. 在 Cursor 里打开本项目，运行：

```bash
python demo/call_review.py --input 你的文件.xlsx
```

   - 没配 API Key 时会用离线示例演示（最稳，断网也能跑）。
   - 想让它真分析你的内容，先配好 `LLM_API_KEY` 等环境变量（见上文），再运行同一条命令。
3. 想把结果存成文件，加 `--out 复盘结果.md`。

> 提示：Excel 解析用的是 Python 标准库，**不用安装任何东西**。

---

## 生成 / 重新生成 PPT

`AI学习成果分享.pptx` 已经生成好，可直接用。若想改内容后重新生成：

```bash
pip install python-pptx          # 仅此一项依赖
python tools/build_pptx.py       # 生成 AI学习成果分享.pptx
```

PPT 共 10 页、16:9，每页**备注里都带演讲稿**（在 PowerPoint/WPS 的"备注"区查看）。
改文字直接编辑 `tools/build_pptx.py` 顶部的 `SLIDES` 列表即可。

---

## 把录音变成文字怎么做？

本工具包以「转写后的文字」为输入。录音转文字可用公司批准的语音转写工具，
或 Cursor 里调用的转写接口。转写完成后做脱敏，再喂给上面的脚本/Prompt。
