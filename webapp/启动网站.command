#!/bin/bash
# Mac 用户：双击本文件即可启动「通话复盘助手」网站。
# 如果双击没反应，右键 → 打开；或第一次需在「访达」里右键 → 打开 以授权。
cd "$(dirname "$0")/.."
echo "正在启动通话复盘助手……"
python3 webapp/server.py || python webapp/server.py
echo ""
echo "网站已停止，可以关闭本窗口。"
read -n 1 -s -r -p "按任意键关闭"
