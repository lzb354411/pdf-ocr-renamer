# -*- coding: utf-8 -*-
"""批量处理核心：扫描 → 分析预览 → 执行重命名 → 撤销

与原脚本的区别：原脚本「识别到匹配立即重命名」，本模块拆分为两步以支持
GUI 预览确认模式：
1. analyze_pdf：只识别和匹配，不改动任何文件（预览阶段）
2. execute_renames：用户确认后，才真正执行重命名并记录撤销日志

撤销机制：每次执行重命名前，把 (旧路径 → 新路径) 清单写入
%APPDATA% 下 PdfOcrRenamer 文件夹的 rename_log.json，撤销时逆序恢复。
"""

import json
import os
import threading
import concurrent.futures
from datetime import datetime

from . import paths
from .rules import match_rule, build_new_name, resolve_conflict, sanitize_filename

# 最大并发线程数（PDF 渲染并行，OCR 推理在 ocr.py 内受锁保护串行执行）
MAX_WORKERS = min(2, os.cpu_count() or 1)

# 文件操作锁（重命名/撤销时保护并发冲突）
_FILE_OP_LOCK = threading.Lock()


# ============================================================
# 扫描
# ============================================================

def scan_pdfs(folder):
    """递归扫描文件夹下所有 PDF（含子文件夹），返回路径列表"""
    pdf_files = []
    for root, _dirs, files in os.walk(folder):
        for fn in files:
            if fn.lower().endswith(".pdf"):
                pdf_files.append(os.path.join(root, fn))
    return pdf_files


# ============================================================
# 单文件分析（预览阶段，不改动文件）
# ============================================================

def analyze_pdf(pdf_path, rules):
    """识别单个 PDF 并匹配规则，返回分析结果（不重命名）

    两阶段识别（与原脚本一致）：
    1. 顶部标题识别（普通文档）
    2. 未匹配时，图纸右下角标题识别（竣工图纸）

    返回 dict：
        status: "match"（匹配）/ "no_match"（未匹配）/ "failed"（识别失败）
        title:  识别到的标题（用于展示与排查）
        mapping: 映射内容
        new_name: 预计新文件名（执行时若有冲突会自动加序号）
    """
    base_name = os.path.basename(pdf_path)
    result = {"pdf_path": pdf_path, "base_name": base_name,
              "status": "no_match", "title": "", "mapping": "",
              "new_name": "", "error": ""}
    try:
        from .ocr import extract_pdf_title, extract_drawing_title

        # 第一阶段：顶部标题识别
        title = extract_pdf_title(pdf_path, rules)
        mapping = match_rule(title, rules)

        # 第二阶段：未匹配时尝试图纸右下角识别
        if not mapping:
            title2 = extract_drawing_title(pdf_path, rules)
            mapping2 = match_rule(title2, rules)
            if mapping2:
                title, mapping = title2, mapping2
            elif len(title2) > len(title):
                # 保留字符更多的识别结果供展示
                title = title2

        result["title"] = title
        if mapping:
            result["status"] = "match"
            result["mapping"] = mapping
            result["new_name"] = ""  # 由 analyze_batch 统一计算（需要 by_folder）
        elif mapping == "":
            result["status"] = "no_match"
    except MemoryError:
        result["status"] = "failed"
        result["error"] = "内存不足（MemoryError）"
    except Exception as e:
        result["status"] = "failed"
        result["error"] = str(e)
    return result


def analyze_batch(pdf_files, rules, by_folder, progress_cb=None, stop_event=None):
    """批量分析（预览阶段）：并发识别 + 匹配，不改动任何文件

    参数：
        pdf_files: PDF 路径列表
        rules: 规则列表
        by_folder: 命名方式（True=文件夹名+映射内容）
        progress_cb: 进度回调 callback(done_count, total, result_dict)
        stop_event: threading.Event，置位后停止处理后续文件
    返回分析结果列表。
    """
    total = len(pdf_files)
    results = [None] * total

    def _work(args):
        idx, pdf_path = args
        if stop_event is not None and stop_event.is_set():
            return idx, {"pdf_path": pdf_path,
                         "base_name": os.path.basename(pdf_path),
                         "status": "cancelled", "title": "", "mapping": "",
                         "new_name": "", "error": "已停止"}
        return idx, analyze_pdf(pdf_path, rules)

    if total <= 1:
        for idx, pdf_path in enumerate(pdf_files):
            i, res = _work((idx, pdf_path))
            if res["status"] != "cancelled":
                if res["status"] == "match":
                    res["new_name"] = build_new_name(
                        res["mapping"], res["pdf_path"], by_folder)
            results[i] = res
            if progress_cb:
                progress_cb(i + 1, total, res)
        return results

    done = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(_work, (idx, p)): idx
                   for idx, p in enumerate(pdf_files)}
        for future in concurrent.futures.as_completed(futures):
            try:
                i, res = future.result()
            except Exception as e:
                i = futures[future]
                res = {"status": "failed", "title": "", "mapping": "",
                       "new_name": "", "error": str(e),
                       "pdf_path": pdf_files[i],
                       "base_name": os.path.basename(pdf_files[i])}
            if res["status"] != "cancelled":
                if res["status"] == "match":
                    res["new_name"] = build_new_name(
                        res["mapping"], res["pdf_path"], by_folder)
            results[i] = res
            done += 1
            if progress_cb:
                progress_cb(done, total, res)
    return results


# ============================================================
# 执行重命名 + 撤销日志
# ============================================================
#
# 日志结构（阶段三 3.1 扩展）：
#   {"history": [
#      {"time": "...", "type": "rename", "items": [{"old":..., "new":...}]},
#      {"time": "...", "type": "split",  "items": [{"dst":...}]},
#   ]}
#
# **向后兼容**：阶段三之前的记录**没有 `type` 字段**，一律视为 "rename"。
# 读取端必须用 `_record_type()` 归一化，不得直接 `rec["type"]`。

# 记录类型常量
REC_RENAME = "rename"
REC_SPLIT = "split"


def _record_type(record):
    """归一化读取记录类型（缺失/异常一律视为 rename，兼容旧格式）"""
    if not isinstance(record, dict):
        return REC_RENAME
    t = record.get("type")
    if t in (REC_RENAME, REC_SPLIT):
        return t
    return REC_RENAME


def _load_log():
    """读取撤销日志，返回 {"history": [...]}"""
    log_file = paths.get_rename_log_file()
    if not os.path.exists(log_file):
        return {"history": []}
    try:
        with open(log_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and isinstance(data.get("history"), list):
            return data
    except Exception:
        pass
    return {"history": []}


def _save_log(data):
    """保存撤销日志"""
    log_file = paths.get_rename_log_file()
    with open(log_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _append_record(record_type, items, keep=20):
    """向日志追加一笔记录（统一入口，自动裁剪历史长度）

    参数：
        record_type: REC_RENAME / REC_SPLIT
        items:       记录条目列表
        keep:        最多保留多少笔（防无限增长）
    """
    if not items:
        return
    data = _load_log()
    data.setdefault("history", []).append({
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "type": record_type,
        "items": items,
    })
    data["history"] = data["history"][-keep:]
    _save_log(data)


def execute_renames(match_results, by_folder, progress_cb=None, stop_event=None,
                    log_undo=True):
    """执行重命名（用户在预览表确认后调用）

    参数：
        match_results: analyze_batch 返回中 status=="match" 的结果列表
        by_folder: 命名方式（True=文件夹名+映射内容，与预览时保持一致）
        progress_cb: callback(done, total, item_result)
        stop_event: 置位后停止处理后续文件
        log_undo:   是否写入撤销日志（默认 True，GUI 依赖）。
                    **测试/一次性脚本应传 False**：撤销日志位于 %APPDATA%，
                    是全局共享状态，写入会污染用户数据与其它套件的基线断言
                    （见阶段文档 §5.5 第 1 条）。
                    阶段四补齐：与 `splitter.execute_split(log_undo=...)` 对称。
    返回 (renamed_items, skipped_info)：
        renamed_items: [{"old": 旧路径, "new": 新路径}, ...] 实际改名成功的
        skipped_info: [{"base_name": ..., "info": ...}, ...] 跳过的及原因
    """
    renamed_items = []
    skipped_info = []
    total = len(match_results)
    done = 0

    for res in match_results:
        done += 1
        pdf_path = res["pdf_path"]
        base_name = res["base_name"]
        if stop_event is not None and stop_event.is_set():
            skipped_info.append({"base_name": base_name, "info": "已停止，未处理"})
            if progress_cb:
                progress_cb(done, total, None)
            continue

        try:
            if not os.path.exists(pdf_path):
                skipped_info.append({"base_name": base_name,
                                     "info": "文件已不存在（可能已被移动或改名）"})
                if progress_cb:
                    progress_cb(done, total, None)
                continue

            new_name = build_new_name(res["mapping"], pdf_path, by_folder)
            if not new_name:
                skipped_info.append({"base_name": base_name, "info": "映射内容为空"})
                if progress_cb:
                    progress_cb(done, total, None)
                continue

            with _FILE_OP_LOCK:
                new_path = os.path.join(os.path.dirname(pdf_path), new_name)
                new_path = resolve_conflict(new_path)
                if os.path.normcase(new_path) == os.path.normcase(pdf_path):
                    skipped_info.append({"base_name": base_name, "info": "名称未变化"})
                    if progress_cb:
                        progress_cb(done, total, None)
                    continue
                os.rename(pdf_path, new_path)

            renamed_items.append({"old": pdf_path, "new": new_path})
            if progress_cb:
                progress_cb(done, total,
                            {"base_name": base_name, "new_name": os.path.basename(new_path)})
        except Exception as e:
            skipped_info.append({"base_name": base_name, "info": str(e)})
            if progress_cb:
                progress_cb(done, total, None)

    # 记录撤销日志（统一入口，自动裁剪历史长度防无限增长）
    # log_undo=False 时跳过（测试/脚本用，避免污染全局撤销日志）
    if renamed_items and log_undo:
        _append_record(REC_RENAME, renamed_items)

    return renamed_items, skipped_info


# ============================================================
# 撤销
# ============================================================

def _last_record(record_type):
    """取最近一笔指定类型的记录；无则返回 None

    阶段三 3.1：撤销重命名与撤销拆分共用同一日志文件，必须**按类型**取记录，
    避免「撤销重命名」误读拆分的记录（反之亦然）。
    """
    data = _load_log()
    for rec in reversed(data.get("history", [])):
        if _record_type(rec) == record_type:
            return rec
    return None


def get_last_rename_count():
    """查询最近一笔可撤销的**重命名**数量（无记录返回 0）

    只统计 type 为 rename 的记录（含旧格式无 type 的记录），
    不受拆分记录影响。
    """
    rec = _last_record(REC_RENAME)
    if not rec:
        return 0
    return len(rec.get("items", []))


def get_last_split_count():
    """查询最近一笔可撤销的**拆分**文件数量（无记录返回 0）"""
    rec = _last_record(REC_SPLIT)
    if not rec:
        return 0
    return len(rec.get("items", []))


def undo_last():
    """撤销最近一次批量重命名（逆序恢复旧文件名）

    阶段三：仅处理 type=="rename" 的记录（旧格式无 type 亦视为 rename），
    不会被拆分记录干扰。

    返回 (restored, failed_info)：
        restored: 恢复成功的数量
        failed_info: [{"file": ..., "info": ...}, ...] 恢复失败的及原因
    """
    data = _load_log()
    history = data.get("history", [])

    # 定位最近一笔 rename 记录的下标（从后往前找）
    idx = -1
    for i in range(len(history) - 1, -1, -1):
        if _record_type(history[i]) == REC_RENAME:
            idx = i
            break
    if idx < 0:
        return 0, [{"file": "", "info": "没有可撤销的重命名记录"}]

    last = history.pop(idx)
    restored = 0
    failed_info = []

    # 逆序恢复：后改名的先恢复，最大程度还原处理前状态
    for item in reversed(last.get("items", [])):
        new_path = item.get("new", "")
        old_path = item.get("old", "")
        base = os.path.basename(new_path) if new_path else "(未知)"
        try:
            if not os.path.exists(new_path):
                failed_info.append({"file": base,
                                    "info": "改名后的文件不存在，无法恢复"})
                continue
            with _FILE_OP_LOCK:
                target = old_path
                if os.path.exists(target):
                    target = resolve_conflict(target)
                os.rename(new_path, target)
            restored += 1
        except Exception as e:
            failed_info.append({"file": base, "info": str(e)})

    _save_log(data)
    return restored, failed_info


# ============================================================
# 拆分撤销（阶段三 R-5）
# ============================================================

def log_split(written_items):
    """记录一次拆分的输出文件（供「撤销上次拆分」使用）

    参数 written_items: splitter.execute_split 返回的 written 列表
                        （元素含 "dst" 键）。也接受 [{"dst": ...}] 形式。

    返回写入日志的记录条数。
    """
    items = []
    for it in (written_items or []):
        dst = ""
        if isinstance(it, dict):
            dst = it.get("dst") or it.get("new") or ""
        if dst:
            items.append({"dst": dst})
    if not items:
        return 0
    _append_record(REC_SPLIT, items)
    return len(items)


def undo_last_split(delete_empty_dirs=True):
    """撤销最近一次拆分：删除**本次拆分新写出的**文件

    **安全保证（务必遵守）**：
        1. 只删除日志中记录的、本次拆分**新建**的文件（`dst` 列表），
           **绝不删除**原整份 PDF，也不删除任何未出现在清单中的用户文件；
        2. 删除前校验路径存在性；文件已被用户移动/改名 → **跳过并报告**，
           不报错中断；
        3. 只删除文件，**不递归删除目录**；仅当某目录因本次撤销而变空时，
           可选地移除该空目录（`delete_empty_dirs=False` 可关闭）；
        4. 若清单中的路径已被用户替换为**同名新文件**，本函数无法区分，
           因此调用方（GUI）**必须在删除前弹确认框**。

    返回 (deleted, skipped_info)：
        deleted:      成功删除的数量
        skipped_info: [{"file": ..., "info": ...}, ...] 未删除的及原因
    """
    data = _load_log()
    history = data.get("history", [])

    idx = -1
    for i in range(len(history) - 1, -1, -1):
        if _record_type(history[i]) == REC_SPLIT:
            idx = i
            break
    if idx < 0:
        return 0, [{"file": "", "info": "没有可撤销的拆分记录"}]

    last = history.pop(idx)
    deleted = 0
    skipped_info = []
    touched_dirs = set()

    # 逆序删除（后写出的先删）
    for item in reversed(last.get("items", [])):
        dst = item.get("dst", "") if isinstance(item, dict) else ""
        base = os.path.basename(dst) if dst else "(未知)"
        if not dst:
            skipped_info.append({"file": base, "info": "记录缺少路径"})
            continue
        try:
            if not os.path.exists(dst):
                skipped_info.append(
                    {"file": base, "info": "文件不存在（可能已被移动或改名），已跳过"})
                continue
            if os.path.isdir(dst):
                # 防御：清单里若混入目录，绝不删除
                skipped_info.append({"file": base, "info": "路径是目录，已跳过"})
                continue
            with _FILE_OP_LOCK:
                os.remove(dst)
            deleted += 1
            touched_dirs.add(os.path.dirname(dst))
        except Exception as e:
            skipped_info.append({"file": base, "info": str(e)})

    # 清理因撤销而变空的目录（仅限本次涉及、且确实为空的目录；不递归）
    if delete_empty_dirs:
        for d in touched_dirs:
            try:
                if d and os.path.isdir(d) and not os.listdir(d):
                    os.rmdir(d)
            except Exception:
                pass  # 目录清理失败不影响撤销结果

    _save_log(data)
    return deleted, skipped_info
