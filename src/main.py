# -*- coding: utf-8 -*-
"""PDF 批量识别重命名工具 - 程序入口

说明：
- ORT/OMP 线程数环境变量必须在 import onnxruntime 之前设置（置顶处理），
  避免多线程并发 OCR 时线程爆炸导致栈溢出（沿用原脚本的修复）
- 高 DPI 适配：Windows 下声明进程 DPI 感知，避免高分屏界面模糊
"""

import os
import sys

# 限制 onnxruntime/OpenCV 内部线程数（必须在任何 OCR 相关 import 之前）
os.environ.setdefault("ORT_NUM_THREADS", "1")
os.environ.setdefault("OMP_NUM_THREADS", "1")

# 开发模式运行（python src/main.py）时，确保本目录在模块搜索路径中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Windows 高 DPI 适配（需在创建任何窗口之前调用）
if sys.platform == "win32":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass  # 旧系统无该 API 时忽略


def main():
    from app import run
    run()


if __name__ == "__main__":
    main()
