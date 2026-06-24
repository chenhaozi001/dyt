#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通话复盘助手 · 网页版

一个不需要懂代码的网站：在浏览器里上传录音转写文件（Excel/CSV/txt），
点一下就能看到结构化的销售复盘结果。

特点：
- 只用 Python 标准库，无需安装任何依赖。
- 不配大模型也能用（显示示例结果，方便先看效果）；
  想分析真实内容，可以直接在网页上填大模型接口信息，或设置环境变量。

启动方式（任选其一）：
  1) 在 Cursor 里打开本文件，点右上角的「运行」按钮。
  2) 双击 webapp 目录下的「启动网站.command」(Mac) / 「启动网站.bat」(Windows)。
  3) 命令行：python webapp/server.py

启动后浏览器会自动打开 http://localhost:8000
"""

import html
import json
import os
import re
import sys
import tempfile
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

# 复用 demo/call_review.py 里的解析与复盘逻辑
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "demo"))
import call_review  # noqa: E402

PORT = int(os.environ.get("PORT", "8000"))
ALLOWED_EXT = (".txt", ".md", ".csv", ".xlsx")


# ----------------------------- 页面模板 -----------------------------

PAGE_HEAD = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>通话复盘助手</title>
<style>
  * { box-sizing: border-box; }
  body {
    margin: 0; font-family: "Microsoft YaHei", -apple-system, "PingFang SC", sans-serif;
    background: linear-gradient(135deg, #1f4e79 0%, #2e86c1 100%);
    color: #222; min-height: 100vh;
  }
  .wrap { max-width: 880px; margin: 0 auto; padding: 32px 20px 60px; }
  .card {
    background: #fff; border-radius: 16px; padding: 32px;
    box-shadow: 0 12px 40px rgba(0,0,0,.18);
  }
  h1 { color: #fff; text-align: center; margin: 8px 0 4px; font-size: 30px; }
  .sub { color: #e8f1fb; text-align: center; margin-bottom: 28px; font-size: 15px; }
  h2 { color: #1f4e79; font-size: 18px; margin: 26px 0 8px; border-left: 4px solid #2e86c1; padding-left: 10px; }
  .drop {
    border: 2px dashed #9cc3e6; border-radius: 14px; padding: 40px 20px; text-align: center;
    background: #f5f9fd; cursor: pointer; transition: .2s;
  }
  .drop.hover { background: #e8f3fc; border-color: #2e86c1; }
  .drop .big { font-size: 18px; color: #1f4e79; font-weight: bold; }
  .drop .small { color: #777; margin-top: 8px; font-size: 13px; }
  #fileName { margin-top: 14px; color: #1f4e79; font-weight: bold; }
  .adv { margin-top: 18px; }
  .adv summary { cursor: pointer; color: #2e86c1; font-size: 14px; }
  .adv input {
    width: 100%; padding: 10px 12px; margin-top: 8px; border: 1px solid #cfd8e3;
    border-radius: 8px; font-size: 14px;
  }
  .adv label { font-size: 13px; color: #555; display:block; margin-top:10px; }
  .btn {
    display: block; width: 100%; margin-top: 22px; padding: 15px;
    background: #1f4e79; color: #fff; border: none; border-radius: 10px;
    font-size: 17px; font-weight: bold; cursor: pointer; transition: .2s;
  }
  .btn:hover { background: #163a5c; }
  .btn:disabled { background: #9aa7b4; cursor: not-allowed; }
  .note { background: #fff8e1; border: 1px solid #ffe08a; color: #8a6d00;
    padding: 12px 14px; border-radius: 10px; font-size: 13px; margin-top: 18px; }
  .demo-banner { background:#e8f3fc; border:1px solid #9cc3e6; color:#1f4e79;
    padding:12px 14px; border-radius:10px; font-size:14px; margin-bottom:18px; }
  ul { margin: 6px 0; padding-left: 22px; }
  li { margin: 5px 0; line-height: 1.6; }
  .score { font-size: 22px; font-weight: bold; color: #2e86c1; }
  .improve { background:#fff4f4; border:1px solid #f3c2c2; padding:12px 16px; border-radius:10px; }
  .compliance { background:#fff8e1; border:1px solid #ffe08a; padding:12px 16px; border-radius:10px; }
  .golden { background:#eefaf0; border:1px solid #b6e2c1; padding:12px 16px; border-radius:10px; }
  .actions { display:flex; gap:12px; margin-top:24px; flex-wrap: wrap;}
  .actions a { flex:1; text-align:center; text-decoration:none; padding:13px;
    border-radius:10px; font-weight:bold; }
  .a-primary { background:#1f4e79; color:#fff; }
  .a-second { background:#eef2f6; color:#1f4e79; }
  .spinner { display:none; text-align:center; color:#1f4e79; margin-top:16px; }
  table.kv { width:100%; border-collapse: collapse; }
  table.kv td { padding:7px 4px; border-bottom:1px solid #eef0f3; }
  table.kv td:first-child { color:#777; width:90px; }
</style>
</head>
<body>
<div class="wrap">
<h1>📞 通话复盘助手</h1>
<div class="sub">上传录音转写文件，一键生成销售复盘 · 仅供内部使用，请先脱敏</div>
"""

PAGE_FOOT = """
</div>
<script>
  const drop = document.getElementById('drop');
  const fileInput = document.getElementById('file');
  const fileName = document.getElementById('fileName');
  const form = document.getElementById('form');
  const btn = document.getElementById('btn');
  const spinner = document.getElementById('spinner');
  if (drop) {
    drop.addEventListener('click', () => fileInput.click());
    ['dragenter','dragover'].forEach(e => drop.addEventListener(e, ev => {
      ev.preventDefault(); drop.classList.add('hover');
    }));
    ['dragleave','drop'].forEach(e => drop.addEventListener(e, ev => {
      ev.preventDefault(); drop.classList.remove('hover');
    }));
    drop.addEventListener('drop', ev => {
      fileInput.files = ev.dataTransfer.files;
      showName();
    });
    fileInput.addEventListener('change', showName);
    function showName(){
      if (fileInput.files.length) fileName.textContent = '已选择：' + fileInput.files[0].name;
    }
    form.addEventListener('submit', () => {
      if (!fileInput.files.length){ alert('请先选择一个文件'); return false; }
      btn.disabled = true; btn.textContent = '正在分析…'; spinner.style.display = 'block';
    });
  }
</script>
</body>
</html>
"""


def index_page():
    return PAGE_HEAD + """
<div class="card">
  <form id="form" method="POST" action="/analyze" enctype="multipart/form-data">
    <div class="drop" id="drop">
      <div class="big">📁 点击选择，或把文件拖到这里</div>
      <div class="small">支持 Excel(.xlsx) / CSV / 文本(.txt)</div>
      <div id="fileName"></div>
    </div>
    <input type="file" id="file" name="file" accept=".xlsx,.csv,.txt,.md" style="display:none">

    <button class="btn" id="btn" type="submit">开始分析</button>
    <div class="spinner" id="spinner">⏳ 正在智能分析通话内容，请稍候…</div>
  </form>
  <div class="note">⚠️ 数据安全：上传前请对录音转写做脱敏处理（去掉真实姓名、电话、公司名）。</div>
</div>
""" + PAGE_FOOT


def _ul(items):
    if not items:
        return "<ul><li>—</li></ul>"
    if isinstance(items, str):
        items = [items]
    return "<ul>" + "".join(f"<li>{html.escape(str(x))}</li>" for x in items) + "</ul>"


def result_page(data, error=None):
    if error:
        body = f"""<div class="card">
        <h2>分析出错了</h2>
        <p>{html.escape(error)}</p>
        <div class="actions"><a class="a-primary" href="/">← 返回重新上传</a></div>
        </div>"""
        return PAGE_HEAD + body + PAGE_FOOT

    c = data.get("customer", {})
    body = f"""<div class="card">
  <p style="color:#1a7f37;font-weight:bold;margin-bottom:12px;">✅ 分析完成</p>
  <div style="text-align:center;margin-bottom:10px;">
    <span class="score">综合评分 {html.escape(str(data.get('score','—')))} / 100</span>
  </div>

  <h2>客户画像</h2>
  <table class="kv">
    <tr><td>称呼</td><td>{html.escape(str(c.get('name','—')))}</td></tr>
    <tr><td>身份</td><td>{html.escape(str(c.get('role','—')))}</td></tr>
    <tr><td>行业</td><td>{html.escape(str(c.get('industry','—')))}</td></tr>
  </table>

  <h2>客户需求</h2>{_ul(data.get('demand'))}
  <h2>预算 / 意向</h2><p>{html.escape(str(data.get('budget','—')))}</p>
  <h2>客户异议</h2>{_ul(data.get('objections'))}
  <h2>我做得好的地方</h2>{_ul(data.get('did_well'))}
  <h2>待改进（重点）</h2><div class="improve">{_ul(data.get('to_improve'))}</div>
  <h2>遗漏的关键动作</h2>{_ul(data.get('missed_actions'))}
  <h2>下一步行动</h2><p>{html.escape(str(data.get('next_step','—')))}</p>
  <h2>合规检查</h2><div class="compliance">{_ul(data.get('compliance'))}</div>
  <h2>可复用金句</h2><div class="golden">{_ul(data.get('golden_lines'))}</div>

  <div class="actions">
    <a class="a-primary" href="/">分析下一个 ↻</a>
    <a class="a-second" href="javascript:window.print()">打印 / 存为 PDF</a>
  </div>
</div>"""
    return PAGE_HEAD + body + PAGE_FOOT


# ----------------------------- multipart 解析 -----------------------------

def parse_multipart(body: bytes, boundary: bytes):
    """极简 multipart/form-data 解析，返回 (fields, files)。
    files: {name: (filename, bytes)}"""
    fields, files = {}, {}
    delim = b"--" + boundary
    for seg in body.split(delim):
        if seg in (b"", b"--", b"--\r\n", b"\r\n"):
            continue
        if seg.startswith(b"\r\n"):
            seg = seg[2:]
        if seg.endswith(b"\r\n"):
            seg = seg[:-2]
        if b"\r\n\r\n" not in seg:
            continue
        header_blob, content = seg.split(b"\r\n\r\n", 1)
        headers = header_blob.decode("utf-8", "ignore")
        name_m = re.search(r'name="([^"]*)"', headers)
        if not name_m:
            continue
        name = name_m.group(1)
        file_m = re.search(r'filename="([^"]*)"', headers)
        if file_m and file_m.group(1):
            files[name] = (file_m.group(1), content)
        else:
            fields[name] = content.decode("utf-8", "ignore").strip()
    return fields, files


# ----------------------------- HTTP Handler -----------------------------

class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):  # 静音默认日志
        pass

    def _send_html(self, content, code=200):
        data = content.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self._send_html(index_page())
        else:
            self._send_html("<h1>404</h1><a href='/'>返回首页</a>", 404)

    def do_POST(self):
        if self.path != "/analyze":
            self._send_html("<h1>404</h1>", 404)
            return
        try:
            ctype = self.headers.get("Content-Type", "")
            m = re.search(r"boundary=([^;]+)", ctype)
            if not m:
                self._send_html(result_page(None, error="上传格式不正确。"))
                return
            boundary = m.group(1).strip('"').encode("utf-8")
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length)
            fields, files = parse_multipart(body, boundary)

            if "file" not in files or not files["file"][1]:
                self._send_html(result_page(None, error="没有收到文件，请重新选择。"))
                return

            filename, content = files["file"]
            ext = os.path.splitext(filename)[1].lower()
            if ext not in ALLOWED_EXT:
                self._send_html(result_page(None,
                    error=f"暂不支持「{ext or '该'}」格式，请上传 Excel(.xlsx)、CSV 或 txt 文件。"))
                return

            # 写到临时文件，复用 call_review 的解析逻辑
            with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tf:
                tf.write(content)
                tmp_path = tf.name
            try:
                transcript = call_review.read_input(tmp_path)
            finally:
                os.unlink(tmp_path)

            if not transcript:
                self._send_html(result_page(None,
                    error="没从文件里读到文字（可能是空文件，或 Excel 第一个表是空的）。"))
                return

            # 决定用真实大模型还是示例
            api_key = fields.get("api_key") or os.environ.get("LLM_API_KEY")
            base_url = fields.get("base_url") or os.environ.get("LLM_BASE_URL", "https://api.openai.com/v1")
            model = fields.get("model") or os.environ.get("LLM_MODEL", "gpt-4o-mini")

            if api_key:
                try:
                    data = call_review.call_llm(transcript, api_key, base_url, model)
                    self._send_html(result_page(data))
                    return
                except Exception as e:  # noqa: BLE001
                    self._send_html(result_page(None,
                        error=f"分析失败：{e}"))
                    return
            else:
                data = call_review.heuristic_result(transcript)
                self._send_html(result_page(data))
        except Exception as e:  # noqa: BLE001
            self._send_html(result_page(None, error=f"服务器出错：{e}"))


def main():
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    url = f"http://localhost:{PORT}"
    print("=" * 50)
    print("  通话复盘助手已启动！")
    print(f"  请在浏览器打开： {url}")
    print("  (关闭这个窗口即可停止网站)")
    print("=" * 50)
    sys.stdout.flush()
    # 仅在有图形界面的本机电脑上自动打开浏览器；
    # 服务器/无界面环境（如云端）跳过，避免报错和无谓的日志。
    want_browser = os.environ.get("OPEN_BROWSER", "auto")
    has_display = bool(os.environ.get("DISPLAY") or sys.platform in ("darwin", "win32"))
    if want_browser == "1" or (want_browser == "auto" and has_display):
        try:
            threading.Timer(1.0, lambda: webbrowser.open(url)).start()
        except Exception:
            pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")
        server.shutdown()


if __name__ == "__main__":
    main()
