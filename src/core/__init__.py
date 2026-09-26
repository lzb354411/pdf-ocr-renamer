# -*- coding: utf-8 -*-
"""核心逻辑层（自 v2.0 单文件脚本迁移并模块化）

模块划分：
- paths:  应用数据路径（%APPDATA% 用户区，与程序目录分离）
- rules:  规则 Excel 读写 / 文本清洗 / 规则匹配 / 文件名工具
- ocr:    OCR 引擎 / 方向检测 / 标题提取（顶部 + 图纸右下角）
- batch:  扫描 / 分析预览 / 执行重命名 / 撤销（含拆分撤销日志）
- splitter: 整份扫描 PDF 拆分（分析 + 执行 + 项目边界判定）
- project_config / name_parser: 工程名称解析配置与解析器

历史脚本已归档至 `legacy/`（见该目录 README），本包不引用它。
"""
