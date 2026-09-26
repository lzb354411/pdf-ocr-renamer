# -*- coding: utf-8 -*-
"""应用数据路径管理（独立 EXE 版）

数据全部存放在用户目录 %APPDATA% 下 PdfOcrRenamer 文件夹，
与程序安装目录（Program Files）分离：
- 升级覆盖安装时不碰用户数据（规则、撤销日志）
- 程序卸载重装后规则仍在
"""

import os
import sys

# 应用数据目录（用户区）
_APP_DIR_NAME = "PdfOcrRenamer"

# 规则文件名（沿用原脚本的文件名，便于老用户迁移）
RULES_FILENAME = "PDFocr 映射规则.xlsx"


def get_app_dir():
    """获取应用数据目录 %APPDATA%\\PdfOcrRenamer（不存在则创建）"""
    base = os.environ.get("APPDATA", os.path.expanduser("~"))
    app_dir = os.path.join(base, _APP_DIR_NAME)
    os.makedirs(app_dir, exist_ok=True)
    return app_dir


def get_rules_dir():
    """获取规则文件夹（不存在则创建）"""
    rules_dir = os.path.join(get_app_dir(), "rules")
    os.makedirs(rules_dir, exist_ok=True)
    return rules_dir


def get_rules_file():
    """获取规则 Excel 文件完整路径"""
    return os.path.join(get_rules_dir(), RULES_FILENAME)


def get_rename_log_file():
    """获取撤销日志文件路径（记录每次重命名操作，用于一键回滚）"""
    return os.path.join(get_app_dir(), "rename_log.json")


def is_frozen():
    """当前是否以 PyInstaller 打包后的 EXE 方式运行"""
    return getattr(sys, "frozen", False)
