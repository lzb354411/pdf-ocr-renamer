# -*- coding: utf-8 -*-
"""页级 OCR 结果缓存（阶段四 R-7）

背景：
    `splitter._classify_page()` 对单页最多执行 4 角度 × 4 级裁剪 + 3 个右下角
    区域 ≈ 28 次 OCR 推理。30 页扫描件最坏情况 800+ 次推理，实测单份 30 页
    分析约 310–335 秒。同一份 PDF 二次分析（反复调参、重复预览）代价高昂。

设计原则（严格遵守阶段文档「缓存正确性优先于性能」）：
    1. **缓存键必须覆盖全部影响结果的输入** —— 文件身份（路径+大小+mtime）、
       页码、渲染参数（scale）。任一变化即视为不同缓存条目；
    2. **缓存的是「页级最终结果」**：`(mapping, rows, angle)`，
       而不是单次裁剪的 OCR 结果。理由见阶段文档阶段四风险提示第 3 条：
       `_classify_page` 是「按角度顺序命中即返回」的状态机，
       只缓存某次裁剪会掩盖角度差异，导致 A4-2 无法真正验证；
    3. **失效由键保证**：文件 mtime/大小变化 → 键变化 → 自动 miss，
       不做额外的"猜测式"失效判断；
    4. **失败降级**：缓存读写任何异常都不得影响分析结果（退化为无缓存）。

存储：
    %APPDATA%\\PdfOcrRenamer\\cache\\pages\\<key>.json
    单文件单条目，便于按 mtime 做 LRU 淘汰，避免整包重写。
"""

import hashlib
import json
import os
import shutil
import threading
import time

from . import paths

# 缓存格式版本：键结构或值结构变更时必须递增，使旧缓存全部失效
CACHE_VERSION = 1

# 默认容量上限（条目数）与单条体积上限
DEFAULT_MAX_ENTRIES = 2000
MAX_ENTRY_BYTES = 2 * 1024 * 1024      # 单条 >2MB 不缓存（异常页）

# 内存级开关与统计（进程内）
_CACHE_LOCK = threading.Lock()
_STATS = {"hit": 0, "miss": 0, "write": 0, "error": 0}

# 是否启用缓存（可通过 enable_cache 切换；测试可临时关闭）
_ENABLED = True


# ============================================================
# 路径与开关
# ============================================================

def get_cache_dir():
    """获取缓存目录（不存在则创建）"""
    d = os.path.join(paths.get_app_dir(), "cache", "pages")
    os.makedirs(d, exist_ok=True)
    return d


def is_enabled():
    """当前是否启用缓存"""
    return _ENABLED


def enable_cache(flag=True):
    """启用/禁用缓存（禁用时读写全部短路，行为等价于无缓存）"""
    global _ENABLED
    _ENABLED = bool(flag)
    return _ENABLED


def reset_stats():
    """重置命中统计（便于测试对比）"""
    with _CACHE_LOCK:
        for k in _STATS:
            _STATS[k] = 0


def get_stats():
    """返回统计副本：{"hit","miss","write","error"}"""
    with _CACHE_LOCK:
        return dict(_STATS)


def _bump(key, n=1):
    with _CACHE_LOCK:
        _STATS[key] = _STATS.get(key, 0) + n


# ============================================================
# 缓存键
# ============================================================

def make_page_key(pdf_path, page_index, render_scale,
                  extra=None, file_stat=None):
    """构造页级缓存键（sha1 十六进制串）

    参与哈希的输入（全部必须覆盖，否则会串味）：
        - 缓存格式版本 CACHE_VERSION
        - PDF 绝对路径（normcase，避免大小写差异造成重复条目）
        - 文件大小 + mtime_ns（**失效依据**：文件一改，键就变）
        - 页码（0 基）
        - 渲染参数 render_scale（scale 影响位图，进而影响 OCR 结果）
        - extra：算法相关参数（如裁剪比例元组、角度顺序），
          这些同样影响最终 (mapping, rows)，必须纳入键

    返回 str；任何异常返回 None（调用方据此跳过缓存）。
    """
    try:
        if not pdf_path:
            return None
        ap = os.path.abspath(pdf_path)
        st = file_stat or os.stat(ap)
        parts = [
            "v%d" % CACHE_VERSION,
            os.path.normcase(ap),
            str(st.st_size),
            str(getattr(st, "st_mtime_ns", int(st.st_mtime * 1e9))),
            "p%d" % int(page_index),
            "s%s" % _fmt_scale(render_scale),
        ]
        if extra:
            # 稳定序列化：sort_keys 保证同一逻辑参数得到同一键
            try:
                parts.append(json.dumps(extra, sort_keys=True,
                                        ensure_ascii=False, default=str))
            except Exception:
                parts.append(repr(extra))
        raw = "|".join(parts).encode("utf-8", "replace")
        return hashlib.sha1(raw).hexdigest()
    except Exception:
        _bump("error")
        return None


def _fmt_scale(scale):
    """渲染参数的稳定字符串形式（避免 1 与 1.0 生成不同键）"""
    try:
        f = float(scale)
        return ("%.6f" % f).rstrip("0").rstrip(".") or "0"
    except (TypeError, ValueError):
        return str(scale)


# ============================================================
# 读写
# ============================================================

def _entry_path(key):
    return os.path.join(get_cache_dir(), key + ".json")


def load_page(key):
    """读取页级缓存，返回 dict 或 None（未命中/损坏/禁用）"""
    if not _ENABLED or not key:
        _bump("miss")
        return None
    p = _entry_path(key)
    try:
        if not os.path.exists(p):
            _bump("miss")
            return None
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict) or data.get("v") != CACHE_VERSION:
            _bump("miss")
            return None
        # 命中时刷新访问时间（供 LRU 淘汰）
        try:
            os.utime(p, None)
        except OSError:
            pass
        _bump("hit")
        return data
    except Exception:
        # 损坏缓存视为 miss，并尽力清理，避免反复失败
        _bump("miss")
        _bump("error")
        try:
            os.remove(p)
        except OSError:
            pass
        return None


def save_page(key, mapping, rows, angle, meta=None):
    """写入页级缓存

    参数：
        key:     make_page_key 的返回值
        mapping: 命中的映射内容（可为 None）
        rows:    [(y, text), ...]（供工程名称提取）
        angle:   命中角度
        meta:    附加信息（如渲染参数、耗时），仅作诊断，不参与判定

    返回 True/False。任何失败都只记录统计，不抛异常。
    """
    if not _ENABLED or not key:
        return False
    payload = {
        "v": CACHE_VERSION,
        "mapping": mapping,
        "rows": [[float(y), str(t)] for y, t in (rows or [])],
        "angle": angle,
        "meta": meta or {},
        "saved_at": time.time(),
    }
    try:
        blob = json.dumps(payload, ensure_ascii=False)
        if len(blob.encode("utf-8")) > MAX_ENTRY_BYTES:
            return False
        p = _entry_path(key)
        tmp = p + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(blob)
        os.replace(tmp, p)          # 原子替换，避免读到半截 JSON
        _bump("write")
        return True
    except Exception:
        _bump("error")
        return False


def rows_from_entry(entry):
    """把缓存条目里的 rows 还原为 [(y, text), ...]"""
    out = []
    for item in (entry or {}).get("rows") or []:
        try:
            out.append((float(item[0]), str(item[1])))
        except (TypeError, ValueError, IndexError):
            continue
    return out


# ============================================================
# 容量管理
# ============================================================

def _list_entries():
    """列出缓存条目 [(path, mtime, size), ...]"""
    out = []
    try:
        d = get_cache_dir()
        for fn in os.listdir(d):
            if not fn.endswith(".json"):
                continue
            p = os.path.join(d, fn)
            try:
                st = os.stat(p)
                out.append((p, st.st_mtime, st.st_size))
            except OSError:
                continue
    except OSError:
        pass
    return out


def cache_size():
    """返回 (条目数, 总字节数)"""
    items = _list_entries()
    return len(items), sum(s for _p, _m, s in items)


def enforce_limit(max_entries=None, max_bytes=None):
    """按容量上限淘汰最久未使用的条目（LRU）

    返回被删除的条目数。
    """
    try:
        limit = int(max_entries if max_entries is not None
                    else DEFAULT_MAX_ENTRIES)
    except (TypeError, ValueError):
        limit = DEFAULT_MAX_ENTRIES
    items = _list_entries()
    removed = 0

    # 先按总字节数淘汰（若有上限）
    if max_bytes:
        total = sum(s for _p, _m, s in items)
        if total > max_bytes:
            for p, _m, s in sorted(items, key=lambda t: t[1]):
                if total <= max_bytes:
                    break
                try:
                    os.remove(p)
                    total -= s
                    removed += 1
                except OSError:
                    pass
            items = _list_entries()

    # 再按条目数淘汰（最旧优先）
    if len(items) > limit:
        for p, _m, _s in sorted(items, key=lambda t: t[1])[:len(items) - limit]:
            try:
                os.remove(p)
                removed += 1
            except OSError:
                pass
    return removed


def clear_cache():
    """清空缓存目录

    返回 (删除条目数, 释放字节数)。目录本身保留，便于后续继续缓存。
    """
    items = _list_entries()
    n = 0
    freed = 0
    d = get_cache_dir()
    for p, _m, s in items:
        try:
            os.remove(p)
            n += 1
            freed += s
        except OSError:
            pass
    # 顺带清理残留的 .tmp（写中断留下的）
    try:
        for fn in os.listdir(d):
            if fn.endswith(".tmp"):
                try:
                    os.remove(os.path.join(d, fn))
                except OSError:
                    pass
    except OSError:
        pass
    return n, freed


def clear_all():
    """彻底删除 cache 根目录（含其它子目录）"""
    base = os.path.join(paths.get_app_dir(), "cache")
    if os.path.isdir(base):
        shutil.rmtree(base, ignore_errors=True)
    return True
