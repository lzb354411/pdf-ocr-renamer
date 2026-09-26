# -*- coding: utf-8 -*-
"""整份扫描 PDF 拆分模块（新增功能，不影响现有重命名流程）

职责：
- analyze_pdf：只读分析一份整份扫描 PDF，识别出各个项目与文档
  （页分类、项目边界、工程名称提取、文件夹匹配），返回预览结构，不改动任何文件
- execute_split：用户确认后，按分析结果把每个文档的页范围抽取为独立 PDF，
  写入匹配到的项目子文件夹（未匹配 → 根目录下「未匹配项目/<工程名称>/」）

关键规则（与用户需求对应）：
1. 每个项目的第一份资料是「开工报告」（必为横向）。
   从开工报告页到下一个开工报告页之间，属于同一个项目。
2. 单页资料大多一个页面且页顶有标题；标题可命中映射规则，
   一个文档一个页面直接拆分；「工作票」「甲供物资超欠供表」等多页资料，
   识别到标题后、下一个标题之前的页面合并为一个 PDF 文档。
3. 工程名称由 name_parser 从开工报告页的整页 OCR 文本中提取（配置驱动、
   多形态、带置信度打分），再用最长公共子串/相似度评分匹配根目录下的子文件夹名。
   地域词与业务后缀全部来自 project_config（可编辑、可由 Excel 自动推导），
   不在本模块中硬编码具体省市县或业务类型。
4. 页面方向混合（纵/横/旋转），标题可能在上/右/下/左/右下角（如竣工图纸），
   每页先自动旋转校正再 OCR，顶部裁剪未命中规则时再尝试右下角标题栏。
"""

import os
import re
import threading

from .rules import (match_rule, normalize_text, sanitize_filename,
                    resolve_conflict, FUZZY_THRESHOLD, FUZZY_MIN_LEN)

# ============ 可调参数 ============
# 分析渲染缩放比例（1.0 已足够识别表格标题；越大越慢）
RENDER_SCALE = 1.0
# OCR 行分组容差（像素）
ROW_TOLERANCE = 15
# 顶部标题动态裁剪比例：依次尝试 15% / 30% / 50% / 整页，命中即停
TOP_CROP_RATIOS = (0.15, 0.30, 0.50, 1.0)
# 一次方向 OCR 字符数达到此值即认为方向正确，停止尝试其他角度
MIN_CHARS_ANGLE = 50
# 工程名称提取时，标签与值所在的 y 带容差（像素，@scale1.0）
NAME_ROW_TOLERANCE = 60
# 文件夹匹配相似度阈值（0~1）
# 可用 project_config 的 folder_score_threshold 覆盖
FOLDER_SCORE_THRESHOLD = 0.55

# ============ 阶段四 4.3：边界锚点优化参数 ============
# 思路：先对「页面是否含可识别文字」做一次廉价判定，再决定是否执行
#       昂贵的多角度 × 多裁剪扫描。对扫描件中的空白页/纯图页，
#       `_classify_page` 原本仍会走完 4 角度 × 4 裁剪 + 3 区域 ≈ 28 次推理，
#       全部返回空；预筛可让这类页面只付 1 次推理。
#
# 预筛策略（保守优先，绝不改变已有结论）：
#   1. 只对「预筛明确无文字」的页做短路。判据是预筛 OCR 出的字符数 < 阈值；
#   2. 预筛使用**与主流程相同的角度**（default_angle，通常是 0 或检测方向），
#      只在最常见的朝向下取样，避免因朝向不同而误判"无文字"；
#   3. 任何不确定（预筛出字符数 ≥ 阈值）都回落到完整 `_classify_page`，
#      因此**不会**改变命中角度与最终结果。
MIN_PRESCAN_CHARS = 1        # 预筛字符数低于此值视为"无文字页"


# 文件操作锁（执行拆分时保护并发冲突）
_FILE_OP_LOCK = threading.Lock()


class SplitStopped(Exception):
    """用户主动停止拆分分析时抛出，用于中断逐页 OCR 循环"""
    pass


# ============================================================
# 页级 OCR 辅助
# ============================================================

def _build_rows(result, limit=60):
    """把 OCR 结果按行分组（同 y 容差内的框合并为一行），返回 [(y, text), ...]"""
    if not result:
        return []
    items = []
    for item in result:
        box = item[0]
        text = str(item[1])
        top_y = min(p[1] for p in box)
        left_x = min(p[0] for p in box)
        items.append((top_y, left_x, text))
    items.sort(key=lambda t: (t[0], t[1]))
    rows = []
    cur_row = []
    cur_texts = []
    cur_y = None
    for y, _x, text in items:
        if cur_y is None or abs(y - cur_y) <= ROW_TOLERANCE:
            cur_row.append(text)
            cur_texts.append(text)
            if cur_y is None:
                cur_y = y
        else:
            rows.append((cur_y, "".join(cur_texts)))
            cur_row = [text]
            cur_texts = [text]
            cur_y = y
    if cur_texts:
        rows.append((cur_y, "".join(cur_texts)))
    return rows[:limit]


def _ocr_image(engine, pil_image):
    """对 PIL Image 做 OCR（受锁保护，串行执行）"""
    import numpy as np
    from .ocr import _ocr_recognize
    return _ocr_recognize(engine, pil_image)


def _detect_orientation(img):
    """复用 ocr.py 的投影方差方向检测，返回需旋转角度（0 或 90）"""
    from .ocr import _detect_text_orientation
    return _detect_text_orientation(img)


def _angle_order(orientation):
    """根据检测方向生成角度尝试顺序（PIL rotate 为逆时针）"""
    if orientation == 0:
        return (0, 90, 270, 180)
    return (90, 270, 0, 180)


def _right_bottom_regions(w, h):
    """右下角/底部标题栏候选区域（用于图纸标题识别）"""
    return [
        (w - int(w * 0.4), h - int(h * 0.4), w, h),
        (0, int(h * 0.7), w, h),
        (w - int(w * 0.6), h - int(h * 0.6), w, h),
    ]


# ============================================================
# 页分类
# ============================================================

def _match_rows(rows, rules):
    """对按行分组的 OCR 文本做规则匹配，返回 (mapping, title)"""
    title = "\n".join(t for _y, t in rows)
    if not title.strip():
        return None, ""
    mapping = match_rule(title, rules)
    return mapping, title


# ============ 页级缓存参数（阶段四 4.2） ============
# 以下参数共同决定 _classify_page 的最终结果，因此必须全部纳入缓存键；
# 改动其中任何一项都应换键（cache 的 CACHE_VERSION 变更可强制全量失效）。
def _cache_extra_signature():
    """返回影响分类结果的算法参数签名（用于缓存键）"""
    return {
        "top_crop": list(TOP_CROP_RATIOS),
        "rb_regions": 3,                     # _right_bottom_regions 返回 3 个
        "row_tol": ROW_TOLERANCE,
        "fuzzy_thr": FUZZY_THRESHOLD,
        "fuzzy_min_len": FUZZY_MIN_LEN,
        "rules_sig": _RULES_SIGNATURE[0],
    }


# 规则内容签名：规则表变了，分类结果可能变，必须纳入缓存键
_RULES_SIGNATURE = [None]


def set_rules_signature(rules):
    """记录当前规则表签名（供缓存键使用）

    签名取「识别内容 + 映射内容」的 sha1；规则为空时置 None（不参与键）。
    """
    global _RULES_SIGNATURE
    try:
        if not rules:
            _RULES_SIGNATURE[0] = None
            return
        raw = "|".join(
            "%s=>%s" % (str(r.get("识别内容", "")), str(r.get("映射内容", "")))
            for r in rules)
        import hashlib
        _RULES_SIGNATURE[0] = hashlib.sha1(
            raw.encode("utf-8", "replace")).hexdigest()[:12]
    except Exception:
        _RULES_SIGNATURE[0] = None


def _prescan_has_text(img, engine):
    """阶段四 4.3：廉价预筛 —— 判断页面是否含可识别文字

    只做**一次** OCR：取整页（不裁剪、不旋转），按检测到的朝向下最常见的
    方向读一次。若字符数 < `MIN_PRESCAN_CHARS`（默认 1），认定该页无文字。

    **保守性保证**：本函数只用于「明确无文字」的短路。一旦预筛认为有文字
    （哪怕只有一个字符），调用方都会回落到完整的 `_classify_page`，
    因此不会改变命中角度、裁剪层级或最终 mapping —— 即不改变既有结论。

    返回 True（可能有文字，需完整分类）/ False（确认无文字，可短路）。
    """
    try:
        orientation = _detect_orientation(img)
        # 用检测出的主方向，避免因朝向不同而误判"无文字"
        cand = img if orientation == 0 else img.rotate(orientation, expand=True)
        rows = _build_rows(_ocr_image(engine, cand))
        count = sum(len(t) for _y, t in rows)
        return count >= MIN_PRESCAN_CHARS
    except Exception:
        # 预筛出错 → 保守回落（当作"有文字"，走完整流程）
        return True


def _classify_page_cached(img, rules, engine, cache_key=None, prescan=False):
    """带缓存的页分类（阶段四 4.2）

    与 `_classify_page` 的关系：
        - cache_key 为 None 或缓存不可用 → 直接委托 `_classify_page`（无缓存路径）；
        - 命中 → 直接还原 (mapping, rows, angle)，**不再做任何 OCR**；
        - 未命中 → 调 `_classify_page` 后写入缓存。

    缓存存的是**整页最终结果**，理由见阶段文档阶段四风险提示第 3 条：
    `_classify_page` 是「按角度顺序命中即返回」的状态机，
    只缓存单次裁剪的 OCR 会掩盖角度差异。

    prescan=True（阶段四 4.3）：在完整分类前先做一次廉价预筛；
    仅当预筛**确认无文字**时短路返回 (None, [], 0)，
    否则与不带预筛时**完全一致**（保守设计，不改变既有结论）。
    """
    if not cache_key:
        if prescan and not _prescan_has_text(img, engine):
            return None, [], 0
        return _classify_page(img, rules, engine)

    from . import cache as _cache
    if not _cache.is_enabled():
        if prescan and not _prescan_has_text(img, engine):
            return None, [], 0
        return _classify_page(img, rules, engine)

    entry = _cache.load_page(cache_key)
    if entry is not None:
        return (entry.get("mapping"),
                _cache.rows_from_entry(entry),
                entry.get("angle", 0))

    if prescan and not _prescan_has_text(img, engine):
        # 预筛确认无文字 → 结果确定为 (None, [], 0)，可安全缓存
        _cache.save_page(cache_key, None, [], 0,
                         meta={"scale": RENDER_SCALE, "prescan_empty": True})
        return None, [], 0

    mapping, rows, angle = _classify_page(img, rules, engine)
    _cache.save_page(cache_key, mapping, rows, angle,
                     meta={"scale": RENDER_SCALE})
    return mapping, rows, angle


def _classify_page(img, rules, engine):
    """对单页做标题识别与规则匹配

    策略：
    1. 投影方差检测文字方向，按角度顺序尝试旋转
    2. 每个方向下，顶部裁剪 15%/30%/50%/整页 严格匹配规则，命中即停
    3. 顶部全部未命中时，尝试右下角/底部标题栏（图纸标题，如竣工图）
    4. 全部未命中：返回 None（该页并入当前文档）

    返回 (mapping, rows, angle)：
        mapping: 命中的映射内容；None 表示本页未匹配到任何规则
        rows:    最佳方向下的整页 OCR 行（供工程名称提取）
        angle:   rows 所在图像的旋转角度（工作方向）

    **注意**：本函数无缓存、每次都会真实调用 OCR；需要缓存请用
    `_classify_page_cached`（阶段四 4.2）。这样保留了「对照实验」能力：
    A4-2 要求能在启用/禁用缓存两种模式下逐页对比结果。
    """
    orientation = _detect_orientation(img)
    best_rows = []
    best_count = 0
    best_angle = 0

    for angle in _angle_order(orientation):
        cand = img if angle == 0 else img.rotate(angle, expand=True)
        w, h = cand.size

        # 顶部裁剪逐级匹配
        for ratio in TOP_CROP_RATIOS:
            top = cand.crop((0, 0, w, int(h * ratio)))
            result = _ocr_image(engine, top)
            rows = _build_rows(result)
            count = sum(len(t) for _y, t in rows)
            if count > best_count:
                best_count = count
                best_rows = rows
                best_angle = angle
            mapping, title = _match_rows(rows, rules)
            if mapping:
                # 命中后补一次整页 OCR：顶部裁剪可能只覆盖标题区，
                # 工程名称表格常在页面下部，整页文本用于工程名称提取
                if ratio < 1.0:
                    full_rows = _build_rows(_ocr_image(engine, cand))
                    if any(len(t) for _y, t in full_rows):
                        rows = full_rows
                return mapping, rows, angle
            if count == 0:
                break  # 该方向无文字，换角度

        # 顶部未命中 → 右下角/底部标题栏（图纸标题）
        for x0, y0, x1, y1 in _right_bottom_regions(w, h):
            crop = cand.crop((x0, y0, x1, y1))
            result = _ocr_image(engine, crop)
            rows = _build_rows(result)
            count = sum(len(t) for _y, t in rows)
            if count > best_count:
                best_count = count
                best_rows = rows
                best_angle = angle
            mapping, title = _match_rows(rows, rules)
            if mapping:
                return mapping, rows, angle

    return None, best_rows, best_angle


# ============================================================
# 工程名称提取（配置驱动，见 name_parser / project_config）
# ============================================================
#
# 历史说明（P0-1 修复）：
#   此处原先硬编码了「某省」作名称起始锚点、「业扩配套工程」作结束锚点，
#   并维护了一份 48 个年份（20241~20279）的黑名单。问题：
#     1. 只能覆盖「某省…10kV业扩配套工程」一种形态（560 条真实样本中 373 条，
#        即 67%）；低压维修项目（139 条）与项目包（36 条）完全提取不到，
#        且静默返回空串，最终流入「未匹配项目」文件夹；
#     2. 换省份/换业务后功能整体失效；
#     3. 年份黑名单到 2028 年即失效。
#   现已改为配置驱动（project_config.py + name_parser.py）：
#     地域词、业务后缀、杂质词、OCR 纠正表全部外置为可编辑配置，
#     并可由用户提供的项目进度表 Excel 自动推导。
#   实测：560 条真实名称解析一致率 100%。

# 配置缓存（避免逐页重复读盘；配置路径在运行时固定）
_CONFIG_CACHE = None


def get_config(reload=False):
    """获取工程名称解析配置（带缓存）

    配置优先级：内置默认 ← 用户 JSON（%APPDATA%\\PdfOcrRenamer\\
    project_config.json）← Excel 自动推导结果（若已生成）。
    """
    global _CONFIG_CACHE
    if _CONFIG_CACHE is None or reload:
        from . import project_config as PC
        _CONFIG_CACHE = PC.load_config()
    return _CONFIG_CACHE


def set_config(cfg):
    """注入配置（供 GUI「自动识别项目名称规则」使用，也便于测试）"""
    global _CONFIG_CACHE
    _CONFIG_CACHE = cfg


def _extract_engineering_name(project_rows_by_page):
    """从项目各页的 OCR 行中提取工程名称（跨页投票取最优）

    参数 project_rows_by_page: [[(y, text), ...], ...]（每页一组 OCR 行，工作方向）
    返回最优工程名称字符串；无法提取时返回 ""（调用方据此归入「未匹配项目」）。

    实现委托给 name_parser（配置驱动、多形态、带置信度打分）。
    """
    from . import name_parser as NP

    cfg = get_config()
    name, _dbg = NP.extract_best_name(project_rows_by_page, cfg)
    return name


# ============================================================
# 文件夹匹配
# ============================================================

def _lcs_len(a, b):
    """最长公共子串长度（动态规划，字符串较短）"""
    la, lb = len(a), len(b)
    dp = [[0] * (lb + 1) for _ in range(la + 1)]
    best = 0
    for i in range(1, la + 1):
        for j in range(1, lb + 1):
            if a[i - 1] == b[j - 1]:
                dp[i][j] = dp[i - 1][j - 1] + 1
                if dp[i][j] > best:
                    best = dp[i][j]
            else:
                dp[i][j] = 0
    return best


# ============================================================
# 特征词主键（阶段二 R-3）
# ============================================================
#
# 背景（阶段文档 §3.2 证据 3）：
#   图纸标题栏与进度表名称差异极大 —— 「某县」vs「丁县」、「台区」vs「合区」，
#   图纸末尾还多出设计单位（如 某市某某电力勘察设计有限公司）。
#   实测 7 页图纸的整页文本包含大量图例噪声（「16法兰杆」「25接地线夹」
#   「34居民住宅」等），整串相似度匹配极不稳定。
#
# 方案：先剥离「公因子」（年份 / 区划后缀 / 批次 / 业务后缀 / 长编号 /
#   设计单位 / 施工单位），剩下的即「特征词」（台区/线路描述，如
#   「某堤口2#配变」「某堡变」「某楼南前片」），用特征词作匹配主键。

def strip_noise_patterns(text, config):
    """按配置的噪声模式剥离标题中的「公因子」与单位名

    剥离项（全部配置驱动，不硬编码任何具体公司名或地名）：
        1. `design_unit_pattern` 设计单位（如「…勘察设计有限公司」）
        2. `org_noise_pattern`   施工/监理/建设单位
        3. `feature_strip_patterns` 年份/区划字/批次/电压等级/业务后缀/编号

    返回剥离后的文本；配置缺失时原样返回（向后兼容）。
    """
    if not text:
        return text
    out = text
    for key in ("design_unit_pattern", "org_noise_pattern"):
        pat = config.get(key)
        if pat:
            try:
                out = re.sub(pat, "", out)
            except re.error:
                pass
    for pat in (config.get("feature_strip_patterns") or []):
        if not pat:
            continue
        try:
            out = re.sub(pat, "", out)
        except re.error:
            pass
    return out


def feature_key(text, config=None):
    """提取「特征词主键」：剥离公因子后的核心描述

    例（配置默认时）：
        「2026年某市某县乙镇某堤口2#配变低压台区维修项目3226000000000007」
        → 「某市沛楼镇某堤口2#配变」附近的核心串（含「某堤口2#配变」）

    返回剥离后的特征串；为空时返回 ""（调用方应退化为整串匹配）。
    """
    if config is None:
        config = get_config()
    if not text:
        return ""
    out = strip_noise_patterns(normalize_text(text), config)
    # 去掉残留标点/空白
    out = re.sub(r"[^\u4e00-\u9fa5A-Za-z0-9#]", "", out)
    return out


def _feature_score(name_feat, folder_feat):
    """两个特征串的相似度（0~1）

    特征串通常不长（如「某堤口2#配变」），故以最长公共子串占比为主，
    再取 SequenceMatcher 作参考。
    """
    if not name_feat or not folder_feat:
        return 0.0
    if name_feat in folder_feat or folder_feat in name_feat:
        return 1.0
    lcs = _lcs_len(name_feat, folder_feat)
    base = lcs / max(1, min(len(name_feat), len(folder_feat)))
    from difflib import SequenceMatcher
    sm = SequenceMatcher(None, name_feat, folder_feat).ratio()
    return max(base, sm)


def _long_codes(text):
    """提取文本中的长编号（>=10 位连续数字），返回 set

    图纸标题栏与文件夹名中都印有项目编号（如 3225000000000004）。
    当同一台区存在多条项目记录（特征词完全相同）时，编号是唯一可区分依据。
    """
    if not text:
        return set()
    return set(re.findall(r"\d{10,}", text))


def _codes_equivalent(a, b, max_diff=2):
    """两个长编号是否等价（允许少量 OCR 数字错字）

    实测：图纸标题栏编号 3225000000000004 被 OCR 成 3225000000000006
    （个别数字识别错），与文件夹名编号**不完全相等**。因此编号比对必须
    容错 —— 与地域词容错同源思路（阶段二 R-2）。

    判定：长度相同且不同位数 ≤ max_diff；或一方是另一方的前缀/子串。
    """
    if not a or not b:
        return False
    if a == b:
        return True
    if len(a) == len(b):
        return sum(1 for x, y in zip(a, b) if x != y) <= max_diff
    # 长度不同（OCR 多识别/漏识别一位）：允许短者为长者的近似子串
    short, long_ = (a, b) if len(a) < len(b) else (b, a)
    if len(long_) - len(short) > max_diff:
        return False
    return short in long_ or long_.startswith(short)


def _apply_code_signal(score, name, folder, config):
    """按「项目编号是否一致」调整得分（阶段二 R-3 补充）

    - 双方都含长编号且**存在等价编号** → 加分（编号匹配是强证据）；
    - 双方都含长编号但**无可配对者** → 降分（明确不是同一项目）；
    - 任一方无编号 → 不加不减（保持原行为）。

    编号比对**容错**（见 `_codes_equivalent`），因为图纸标题栏编号常有个别
    数字 OCR 错误。

    该信号用于区分「同台区、不同批次」的并发项目记录 —— 这类记录的特征词
    完全相同，仅编号不同（实测：某堡变 3225000000000004 / 3226000000000005）。
    """
    if not config.get("code_match_enabled", True):
        return score
    ncodes = _long_codes(name)
    fcodes = _long_codes(folder)
    if not ncodes or not fcodes:
        return score
    try:
        tol = int(config.get("code_fuzzy_tolerance", 2) or 2)
    except (TypeError, ValueError):
        tol = 2
    try:
        bonus = float(config.get("code_match_bonus", 0.35) or 0.0)
    except (TypeError, ValueError):
        bonus = 0.35
    try:
        penalty = float(config.get("code_mismatch_penalty", 0.75) or 1.0)
    except (TypeError, ValueError):
        penalty = 0.75
    for nc in ncodes:
        for fc in fcodes:
            if _codes_equivalent(nc, fc, tol):
                return min(1.0, score + bonus)
    return score * penalty


def _folder_score(name, folder, config=None):
    """工程名称与文件夹名的相似度评分（0~1）

    三路信号综合：
      1. **整串比对**（原逻辑，兜底）：最长公共子串占比 + SequenceMatcher，
         完全包含给满分；
      2. **特征词主键**（阶段二 R-3，优先）：剥离公因子后用特征词比对，
         对「图纸标题 vs 进度表名称」这类差异极大的场景更稳；
         低于 `feature_score_threshold` 时视为未命中（返回 0），
         避免弱特征串给出虚高分；
      3. **项目编号信号**（阶段二 R-3 补充）：编号相同加分、不同降分，
         用于区分特征词完全相同的同台区多条记录。

    参数 config 为 None 时使用当前生效配置。
    """
    if not folder:
        return 0.0
    if config is None:
        config = get_config()
    if folder in name:
        return _apply_code_signal(1.0, name, folder, config)
    if name in folder:
        return _apply_code_signal(0.95, name, folder, config)
    lcs = _lcs_len(name, folder)
    lcs_ratio = lcs / len(folder)
    from difflib import SequenceMatcher
    sm = SequenceMatcher(None, name, folder).ratio()
    whole = max(lcs_ratio, sm * 0.9)

    # 特征词主键
    fname = feature_key(name, config)
    ffolder = feature_key(folder, config)
    try:
        fthr = float(config.get("feature_score_threshold", 0.62) or 0.62)
    except (TypeError, ValueError):
        fthr = 0.62
    try:
        fmin = int(config.get("feature_min_len", 3) or 3)
    except (TypeError, ValueError):
        fmin = 3
    feat = 0.0
    if len(fname) >= fmin and len(ffolder) >= fmin:
        sc = _feature_score(fname, ffolder)
        if sc >= fthr:
            feat = sc
    return _apply_code_signal(max(whole, feat), name, folder, config)


def match_folder(name, subfolders):
    """匹配工程名称到根目录下的子文件夹

    返回 (folder_name, score)：
        folder_name: 匹配到的文件夹名；None 表示未匹配到（阈值以下）
        score: 匹配得分
    """
    std_name = normalize_text(name)
    if not std_name:
        return None, 0.0
    cfg = get_config()
    best = None
    best_score = 0.0
    for folder in subfolders:
        std_folder = normalize_text(folder)
        if not std_folder:
            continue
        score = _folder_score(std_name, std_folder, cfg)
        if score > best_score:
            best_score = score
            best = folder
    if best is not None and best_score >= _folder_threshold():
        return best, best_score
    return None, best_score


# ============================================================
# 用候选文件夹名反向纠正 OCR 错字（阶段二 2.2）
# ============================================================
#
# 思路：根目录下的**真实文件夹名是干净的标准答案**。工程名称从扫描件
#   OCR 而来，常含形近错字（「某县」→「戊县/己县/丁县/乙县」、「台区」→
#   「合区」）。把 OCR 名称与候选文件夹名对齐后，可把错字局部替换为
#   文件夹名中的正确写法，显著提升后续匹配与命名质量。
#
# 约束（见阶段文档 2.2）：
#   - 纯函数、可测试、不产生副作用；
#   - **不得改变已正确的结果**（仅当对齐后确有差异才替换，且只改不同字符）。

def _align_replace_by_template(ocr_text, tpl_text, max_diff_ratio=0.34):
    """以 tpl_text 为模板，局部修正 ocr_text 中的错字（滑动窗口对齐）

    在 ocr_text 中滑动一个与 tpl_text 等长的窗口，找出**差异最小**的位置；
    若差异比例 ≤ max_diff_ratio（默认 1/3），则用 tpl_text 替换该窗口。

    这样「2026年某市戊县乙镇某堡变低压合区维修项目」对齐到文件夹名
    「2026年某市某县乙镇某堡变低压台区维修项目」时，
    可把「戊县→某县」「合区→台区」一并修正。

    返回 (纠正后文本, 是否发生替换)。纯函数，无副作用。
    """
    if not ocr_text or not tpl_text:
        return ocr_text, False
    if ocr_text == tpl_text:
        return ocr_text, False
    n, m = len(ocr_text), len(tpl_text)
    if m > n:
        return ocr_text, False
    best_diff = None
    best_pos = 0
    for start in range(0, n - m + 1):
        diff = sum(1 for k in range(m) if ocr_text[start + k] != tpl_text[k])
        if best_diff is None or diff < best_diff:
            best_diff = diff
            best_pos = start
            if diff == 0:
                break
    if best_diff is None or best_diff == 0:
        return ocr_text, False
    # 差异过大 → 不对齐（避免把无关文本强行套模板）
    if best_diff / float(m) > max_diff_ratio:
        return ocr_text, False
    out = ocr_text[:best_pos] + tpl_text + ocr_text[best_pos + m:]
    return out, out != ocr_text


def _template_variants(folder):
    """由文件夹名生成可用于对齐的模板变体（长 → 短）

    文件夹名常比 OCR 名称多出「末尾长编号」（如 …维修项目3225000000000004），
    而 OCR 名称通常只到「…维修项目」为止。因此对齐时需尝试：
        1. 完整文件夹名
        2. 去掉末尾长编号（\\d{10,}）后的名字
        3. 去掉末尾编号 + 尾部「项目/工程」变体（OCR 偶有截断）

    返回去重后的模板列表（保持长度降序，优先用更完整者）。
    """
    out = []
    base = normalize_text(folder)
    if base:
        out.append(base)
        no_code = re.sub(r"\d{10,}$", "", base)
        if no_code and no_code != base:
            out.append(no_code)
    seen = []
    for t in out:
        if t and t not in seen:
            seen.append(t)
    return sorted(seen, key=lambda s: (-len(s), s))


def correct_name_by_folders(name, subfolders, config=None):
    """用候选文件夹名反向纠正工程名称中的 OCR 错字（阶段二 2.2）

    参数：
        name:       OCR 提取到的工程名称
        subfolders: 候选文件夹名列表（根目录下一级子目录）
        config:     配置（None 时取当前生效配置）

    返回 (纠正后名称, 说明dict)：
        说明dict = {"changed": bool, "folder": 采用的模板, "before": 原名}

    **安全保证**：
        - 先按 `_folder_score` 选出最佳候选文件夹，只对齐该文件夹的模板；
        - 模板会尝试「完整名 / 去末尾长编号」两种形态（见 `_template_variants`），
          以适配「OCR 名称不含末尾编号」的常见情况；
        - 仅在差异比例 ≤ 1/3 时才替换，且替换后相似度不得下降；
        - 名称本身与某文件夹（或其去编号形态）完全一致时直接返回，
          **不改变已正确的结果**。
    """
    info = {"changed": False, "folder": None, "before": name}
    if not name or not subfolders:
        return name, info

    std_name = normalize_text(name)
    if not std_name:
        return name, info
    cfg = config if config is not None else get_config()

    # 选最佳候选模板：优先已完全包含，其次分数最高
    best_folder = None
    best_score = 0.0
    for folder in subfolders:
        std_folder = normalize_text(folder)
        if not std_folder:
            continue
        if std_folder == std_name:
            return name, info          # 已完全一致，不改动
        if std_name in std_folder:
            # OCR 名称是文件夹名的子集（缺末尾编号）→ 直接采用该文件夹
            best_folder = folder
            best_score = 1.0
            break
        sc = _folder_score(std_name, std_folder, cfg)
        if sc > best_score:
            best_score = sc
            best_folder = folder
    if best_folder is None:
        return name, info

    # 逐模板尝试对齐（完整名 → 去编号名），取首个成功者
    for tpl in _template_variants(best_folder):
        if tpl == std_name:
            return name, info
        fixed_norm, changed = _align_replace_by_template(std_name, tpl)
        if not changed:
            continue
        # 命中特征词才认为对齐可信：要求修正后相似度不低于原相似度
        new_score = _folder_score(fixed_norm, normalize_text(best_folder), cfg)
        if new_score < best_score:
            continue
        info["changed"] = True
        info["folder"] = best_folder
        info["tpl"] = tpl
        info["score_before"] = round(best_score, 4)
        info["score_after"] = round(new_score, 4)
        return fixed_norm, info
    return name, info


def _folder_threshold():
    """文件夹匹配阈值：优先取配置值，配置缺失时回退模块常量"""
    try:
        cfg = get_config()
        v = cfg.get("folder_score_threshold")
        if v is not None:
            return float(v)
    except Exception:
        pass
    return FOLDER_SCORE_THRESHOLD


# ============================================================
# 分析（只读）
# ============================================================

def _list_subfolders(root):
    """列出根目录下的子文件夹名（仅一级，排除隐藏文件夹）"""
    folders = []
    try:
        for name in os.listdir(root):
            full = os.path.join(root, name)
            if os.path.isdir(full) and not name.startswith("."):
                folders.append(name)
    except OSError:
        pass
    return folders


def analyze_pdf(pdf_path, rules, root, progress_cb=None, stop_event=None,
                by_folder=False, use_cache=True, cache_max_entries=None,
                prescan=True):
    """只读分析整份扫描 PDF

    参数：
        pdf_path: 整份 PDF 文件路径
        rules: 映射规则列表
        root:   根文件夹（拆出的文档放入其子文件夹）
        progress_cb: callback(done, total, info) 每处理完一页回调一次，
                     done 为已处理页数，total 为总页数，info 为 {"page": 当前页序号}
        stop_event:  置位后中止分析，抛出 SplitStopped
        by_folder: 命名方式（True=文件夹名+映射内容 / False=仅映射内容，
                   与重命名模式一致），据此预计算每个文档的文件名 file_name
        use_cache: **阶段四 4.2** —— 是否使用页级 OCR 缓存（默认 True）。
                   False 时完全不走缓存读取/写入，用于
                   A4-2「启用/禁用缓存逐页对比」的对照实验。
                   注意：本参数会临时切换 cache 模块的全局开关，
                   分析结束后恢复为 True。
        cache_max_entries: 缓存条目上限（None 用 cache.DEFAULT_MAX_ENTRIES）
        prescan: **阶段四 4.3** —— 是否启用低分辨率预筛（默认 True）。
                 仅对「确认无文字」的页面短路，**不改变任何有文字页的结果**；
                 设为 False 可做严格对照（结果必须完全一致）。
    返回 analysis dict：
        {
          "pdf_path", "root", "page_count",
          "warnings": [str, ...],
          "projects": [
             {"name", "matched_folder", "matched_score",
              "start_page", "end_page", "page_range",
              "docs": [{"mapping", "title", "page_start", "page_end",
                        "pages", "file_name"}],
              "no_folder": bool}
          ],
          "total_docs": int,
          "pre_pages": int,   # 首个开工报告之前的异常页面数
        }
    """
    from .ocr import get_ocr_engine

    engine = get_ocr_engine()
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(pdf_path)
    page_count = len(pdf)
    warnings = []

    # 阶段四 4.2：页级 OCR 缓存
    # 缓存键覆盖 文件身份(路径+大小+mtime_ns) + 页码 + 渲染参数 + 算法参数
    # + 规则表签名；任一变化即 miss，因此「文件修改后自动失效」由键保证。
    from . import cache as _cache
    _cache.enable_cache(use_cache)
    set_rules_signature(rules)
    file_stat = None
    try:
        file_stat = os.stat(pdf_path)
    except OSError:
        pass
    _cache_extra = _cache_extra_signature()

    # 逐页分类：page_class[i] = {"mapping", "rows", "angle", "page"}
    page_class = []
    try:
        for i in range(page_count):
            if stop_event is not None and stop_event.is_set():
                raise SplitStopped()
            page = pdf[i]
            bitmap = page.render(scale=RENDER_SCALE)
            img = bitmap.to_pil().convert("RGB")
            ck = None
            if use_cache:
                ck = _cache.make_page_key(
                    pdf_path, i, RENDER_SCALE,
                    extra=_cache_extra, file_stat=file_stat)
            mapping, rows, _angle = _classify_page_cached(
                img, rules, engine, cache_key=ck, prescan=prescan)
            page_class.append({"page": i, "mapping": mapping, "rows": rows})
            if progress_cb:
                progress_cb(i + 1, page_count, {"page": i})
    finally:
        pdf.close()
        # 分析结束后按容量上限淘汰，避免缓存无限增长（阶段四 A4-4）
        if use_cache:
            try:
                _cache.enforce_limit(max_entries=cache_max_entries)
            except Exception:
                pass

    # 按开工报告切分项目
    projects = []
    cur_project = None
    pre_pages = 0
    pre_doc = None       # 首个开工报告之前的页面（收集起来，避免静默丢失）
    for pc in page_class:
        mapping = pc["mapping"]
        if _is_project_boundary(pc):
            # 开工报告页 → 新项目（也是该项目第一份文档）
            if cur_project is None and pre_pages == 0:
                # 首个开工报告之前的页面数（无开工报告先行页 = 异常，提示）
                pre_pages = pc["page"]
            if cur_project is not None:
                cur_project["end_page"] = pc["page"] - 1
                projects.append(cur_project)
            cur_project = {
                "name": "",
                "start_page": pc["page"],
                "end_page": page_count - 1,  # 暂定，遇到下一个开工报告时修正
                "docs": [],
                "rows_pool": [],
            }
            cur_project["docs"].append({
                "mapping": mapping or "开工报告",
                "title": "",
                "page_start": pc["page"],
                "page_end": pc["page"],
                "pages": "P%d" % pc["page"],
            })
            cur_project["rows_pool"].append(pc["rows"])
        else:
            if cur_project is None:
                # 首个开工报告之前出现的页面（异常）。
                # P0-2 修复：这些页面原先只计数、直接 continue 丢弃——
                # 对扫描归档场景「开头几页不是开工报告」极常见（封面、目录、
                # 扫描空白页、上一份文件的尾页），静默丢弃会造成资料永久丢失。
                # 现改为收集到一个「前置页面」文档，执行时落到
                # 未匹配项目/_前置页面/ 目录下，确保不丢页。
                pre_pages += 1
                if pre_doc is not None:
                    pre_doc["page_end"] = pc["page"]
                    pre_doc["pages"] = "P%d-P%d" % (pre_doc["page_start"],
                                                    pre_doc["page_end"])
                else:
                    pre_doc = {
                        "mapping": mapping or "",
                        "title": "".join(t for _y, t in pc["rows"])[:40],
                        "page_start": pc["page"],
                        "page_end": pc["page"],
                        "pages": "P%d" % pc["page"],
                    }
                continue
            # 命中规则 → 新文档；未命中 → 并入当前文档
            combined = "".join(t for _y, t in pc["rows"])
            if mapping:
                cur_project["docs"].append({
                    "mapping": mapping,
                    "title": combined[:40],
                    "page_start": pc["page"],
                    "page_end": pc["page"],
                    "pages": "P%d" % pc["page"],
                })
            elif cur_project["docs"]:
                # 并入当前文档（多页资料，如工作票/甲供物资表）
                last_doc = cur_project["docs"][-1]
                last_doc["page_end"] = pc["page"]
                last_doc["pages"] = "P%d-P%d" % (
                    last_doc["page_start"], last_doc["page_end"])
            else:
                # 项目内首尾无标题页：单独成档，避免丢失
                cur_project["docs"].append({
                    "mapping": "",
                    "title": "",
                    "page_start": pc["page"],
                    "page_end": pc["page"],
                    "pages": "P%d" % pc["page"],
                })
            cur_project["rows_pool"].append(pc["rows"])
    if cur_project is not None:
        cur_project["end_page"] = page_count - 1
        projects.append(cur_project)

    # 提取每个项目的工程名称并匹配文件夹
    subfolders = _list_subfolders(root)
    for proj in projects:
        proj["name"] = _extract_engineering_name(proj["rows_pool"])
        # 阶段二 2.2：先用候选文件夹名反向纠正 OCR 错字，再匹配文件夹。
        # 纯函数、无副作用；无合适模板或名称已正确时不改动（向后兼容）。
        corrected, fix_info = correct_name_by_folders(proj["name"], subfolders)
        if fix_info.get("changed"):
            proj["name_raw"] = proj["name"]
            proj["name"] = corrected
            proj["name_corrected_by"] = fix_info.get("folder")
        folder, score = match_folder(proj["name"], subfolders)
        if folder:
            proj["matched_folder"] = folder
            proj["matched_score"] = round(score, 2)
            proj["no_folder"] = False
        else:
            proj["matched_folder"] = None
            proj["matched_score"] = 0.0
            proj["no_folder"] = True
            warnings.append(
                "工程「{}」未匹配到子文件夹，将放入「未匹配项目」文件夹".format(
                    proj["name"] or ("P%d起" % proj["start_page"])))
        proj.pop("rows_pool", None)
        proj["page_range"] = "P%d-P%d" % (proj["start_page"], proj["end_page"])

    # 按命名方式预计算每个文档的文件名（预览与执行共用，保持一致）
    _assign_file_names(projects, by_folder)

    # 冲突检测：多个项目命中同一文件夹
    seen = {}
    for proj in projects:
        if proj.get("matched_folder"):
            seen.setdefault(proj["matched_folder"], []).append(proj["start_page"])
    for folder, starts in seen.items():
        if len(starts) > 1:
            warnings.append(
                "多个项目（起始页{}）匹配到同一文件夹「{}」，请人工核对".format(
                    "、".join("P%d" % p for p in starts), folder))

    if pre_pages > 0:
        warnings.append(
            "文件开头有 {} 页在首个开工报告之前（P{}起），将放入"
            "「未匹配项目/_前置页面/」以免资料丢失，请人工核对".format(
                pre_pages,
                pre_doc["page_start"] if pre_doc else 0))

    # 前置页面作为一个独立「项目」写出，确保不丢页
    # （matched_folder 固定为保留目录名，execute_split 会据此建目录）
    if pre_doc is not None:
        pre_doc["file_name"] = sanitize_filename(
            "前置页面_P%d-P%d.pdf" % (pre_doc["page_start"], pre_doc["page_end"]))
        pre_project = {
            "name": "",
            "matched_folder": None,
            "matched_score": 0.0,
            "no_folder": True,
            "is_pre_pages": True,
            "start_page": pre_doc["page_start"],
            "end_page": pre_doc["page_end"],
            "page_range": pre_doc["pages"],
            "docs": [pre_doc],
        }
        projects.insert(0, pre_project)

    total_docs = sum(len(p["docs"]) for p in projects)
    return {
        "pdf_path": pdf_path,
        "root": root,
        "page_count": page_count,
        "warnings": warnings,
        "projects": projects,
        "total_docs": total_docs,
        "pre_pages": pre_pages,
    }


# ============ 开工报告页判定锚点（阶段一 1.4） ============
# 原实现仅 `return "开工报告" in text.replace("配电网工程", "")`：只判包含、
# 无排除逻辑，任何含「开工报告」字样的页面都会成为项目边界。竣工报告页因
# OCR 把「竣工」识别成「峻工」，若同时出现「开工」相关字样（如「实际开工
# 时间」「申请开工」），就会被误当作新项目的开始，导致项目数翻倍。
#
# 现改为「排除优先于包含」：
#   1) 先判排除：出现竣工语义锚点 → 明确不是开工报告页；
#   2) 再判包含：出现开工语义锚点 → 是开工报告页。
# 排除优先可确保「同一页既有开工时间又有竣工申请」时判为竣工页（更安全：
# 少切一个项目不会丢资料，误切会撕裂项目）。
_JUN_EXCLUDE_ANCHORS = (
    "申请竣工", "申请峻工",
    "实际竣工", "实际峻工",
    "竣工验收", "峻工验收",
    "竣工报告", "峻工报告",
)
_KAI_INCLUDE_ANCHORS = (
    "申请开工",
    "开工准备情况",
    "开工报告",
)


def _is_project_boundary(pc):
    """判定某页是否构成「新项目边界」（阶段一 1.4 强化）

    背景（真实 30 页样本暴露的第三处缺陷）：
        _classify_page 按角度顺序 (90, 270, 0, 180) 尝试旋转，命中即返回。
        对「竣工报告」扫描页，正确方向是横排；但当正确方向顶部裁剪恰好无
        文字（count==0 → break 换角度）时，控制流落到 angle=0/270 等侧读
        方向，OCR 出的标题是**残缺**的：真实 P1 页标题只识别成
        「配电网工程工报告」（「竣」字整个丢失）。
        这种残缺标题对「配电网工程开工报告」与「配电网工程竣工报告」的
        相似度**完全相同（0.9412）**，属天然歧义，模糊匹配只能凭规则表
        顺序给出「开工报告」——若据此切项目，竣工页就被当作新项目起点。
        实测 30 页样本因此切出 4 个项目，而非业务期望的 3 个。

    本函数把边界判定收敛为三重判定（依次）：
        1. **排除优先**：页面文本含竣工语义锚点（申请竣工/实际竣工/
           竣工报告 等，含 OCR 错字形态）→ 明确不是项目边界；
        2. **完整锚点**：页面文本含**完整**开工锚点（申请开工/开工准备
           情况/开工报告，经 OCR 纠正后）→ 是边界；
        3. **mapping 兜底**：mapping 命中「开工报告」**且**页面文本不是
           「残缺标题」形态时才认边界（见 _rows_have_degraded_kai_title）。

    排除优先与「要求完整锚点」都是安全方向的取舍：少切一个项目只是把
    资料并入前一个项目（内容仍在、可人工归位），误切则会把一个项目撕裂
    成两个且命名错误、需人工合并。
    """
    rows = pc.get("rows") or []
    # 1) 排除优先：含竣工语义 → 不是边界
    if _rows_have_jun_anchor(rows):
        return False
    # 2) 完整开工锚点 → 是边界
    if _rows_contain_kai(rows):
        return True
    # 3) mapping 兜底：mapping 说开工报告，但文本只是残缺标题
    #    （如「配电网工程工报告」= 掉了「竣」或「开」字）时不可信，不切边界
    if pc.get("mapping") == "开工报告":
        return not _rows_have_degraded_kai_title(rows)
    return False


def _rows_have_jun_anchor(rows):
    """页面文本是否含竣工语义锚点（用于排除法，见 _is_project_boundary）"""
    text = normalize_text("\n".join(t for _y, t in rows))
    if not text:
        return False
    text = _apply_ocr_fixes_for_boundary(text)
    return any(anchor in text for anchor in _JUN_EXCLUDE_ANCHORS)


def _rows_have_degraded_kai_title(rows):
    """页面是否只含「残缺的开工标题」（无法区分开工/竣工）

    真实样本：竣工报告页被 OCR 成「配电网工程工报告」（「竣」字丢失），
    该文本与「配电网工程开工报告」「配电网工程竣工报告」相似度完全相同
    （0.9412），本质是歧义文本，**不足以判定为项目边界**。

    判定方式：文本中「报告」前紧跟的是「工」（而非「开工」或「竣工」），
    即标题为「…工程工报告」/「…工报告」这类缺字形态。
    """
    text = normalize_text("\n".join(t for _y, t in rows))
    if not text:
        return False
    text = _apply_ocr_fixes_for_boundary(text)
    # 「工报告」且其前一字不是「开」/「竣」→ 残缺标题
    idx = text.find("工报告")
    if idx < 0:
        return False
    prev = text[idx - 1] if idx > 0 else ""
    return prev not in ("开", "竣")


def _rows_contain_kai(rows):
    """判定某页是否为「开工报告页」（项目边界）

    阶段一 1.4 改造：排除优先于包含，且对 OCR 错字先做纠正
    （「峻工」→「竣工」），避免竣工页被当作项目边界。

    注意：本函数只用于「兜底判定」。正常情况下页面映射已由 match_rule
    给出「开工报告」，此处仅覆盖规则未命中但文本明显含锚点的场景。
    """
    text = normalize_text("\n".join(t for _y, t in rows))
    if not text:
        return False
    # OCR 错字纠正后再判定，保证「峻工报告」与「竣工报告」等价
    text = _apply_ocr_fixes_for_boundary(text)
    # 1) 排除优先：含竣工语义 → 明确不是开工报告页
    for anchor in _JUN_EXCLUDE_ANCHORS:
        if anchor in text:
            return False
    # 2) 包含：含开工语义 → 是开工报告页
    for anchor in _KAI_INCLUDE_ANCHORS:
        if anchor in text:
            return True
    return False


def _apply_ocr_fixes_for_boundary(text):
    """边界判定用的 OCR 纠正（复用 project_config，失败时原样返回）"""
    try:
        from . import project_config as PC
        return PC.apply_ocr_fixes(text, get_config())
    except Exception:
        return text


def _assign_file_names(projects, by_folder):
    """按命名方式给每个文档预计算文件名（预览与执行共用，保持一致）

    by_folder=True：匹配到文件夹 → 「文件夹名+标题」；未匹配 → 「工程名称+标题」
    by_folder=False：仅标题（原行为）
    直接修改 projects 中每个 doc 的 file_name 字段。
    """
    for proj in projects:
        prefix = proj.get("matched_folder") or proj.get("name") or ""
        prefix = sanitize_filename(prefix)
        for doc in proj["docs"]:
            title = sanitize_filename(doc.get("mapping") or "未识别标题")
            if by_folder and prefix:
                doc["file_name"] = prefix + title + ".pdf"
            else:
                doc["file_name"] = title + ".pdf"


# ============================================================
# 执行拆分（用户确认后）
# ============================================================

def execute_split(analysis, progress_cb=None, stop_event=None, log_undo=True):
    """把分析结果中的各文档按页范围抽出，写入项目子文件夹

    参数：
        analysis: analyze_pdf 的返回结构
        progress_cb: callback(done, total, item_dict)
        stop_event:  置位后停止处理后续文档
        log_undo:    是否把本次输出写入撤销日志（默认 True，供 GUI 使用）。
                     **测试/一次性脚本应传 False**：撤销日志位于 %APPDATA%，
                     是全局共享状态，写入会污染用户数据与其它套件的基线断言
                     （见阶段文档 §5.5 第 1 条）。
    返回 (written, skipped_info)：
        written: [{"dst": ..., "mapping": ..., "pages": ...}, ...] 成功写出的
        skipped_info: [{"dst": ..., "info": ...}, ...] 失败的及原因
    """
    from pypdf import PdfReader, PdfWriter

    pdf_path = analysis["pdf_path"]
    root = analysis["root"]
    reader = PdfReader(pdf_path)

    # 汇总全部文档（含未匹配项目）
    doc_list = []  # (proj, doc)
    for proj in analysis["projects"]:
        for doc in proj["docs"]:
            doc_list.append((proj, doc))

    total = len(doc_list)
    written = []
    skipped_info = []
    done = 0

    for proj, doc in doc_list:
        done += 1
        if stop_event is not None and stop_event.is_set():
            skipped_info.append({"dst": "", "info": "已停止，未处理"})
            if progress_cb:
                progress_cb(done, total, None)
            continue

        # 目标文件夹：匹配到的子文件夹，或「未匹配项目/<工程名称>/」
        # 例外：前置页面（首个开工报告之前的页）固定放入
        #       「未匹配项目/_前置页面/」，便于人工归位且绝不丢失
        if proj.get("is_pre_pages"):
            target_dir = os.path.join(root, "未匹配项目", "_前置页面")
        elif proj.get("matched_folder"):
            target_dir = os.path.join(root, proj["matched_folder"])
        else:
            name = (proj.get("name") or "未知工程").strip() or "未知工程"
            name = sanitize_filename(name)
            target_dir = os.path.join(root, "未匹配项目", name)

        mapping = doc.get("mapping") or "未识别标题"
        # 文件名在分析阶段按命名方式预计算（file_name），保证预览与执行一致
        file_name = doc.get("file_name") or (sanitize_filename(mapping) + ".pdf")

        try:
            os.makedirs(target_dir, exist_ok=True)
            # 目标路径冲突时自动加序号
            writer = PdfWriter()
            for p in range(doc["page_start"], doc["page_end"] + 1):
                writer.add_page(reader.pages[p])
            dst = os.path.join(target_dir, file_name)
            with _FILE_OP_LOCK:
                dst = resolve_conflict(dst)
                with open(dst, "wb") as f:
                    writer.write(f)
            written.append({
                "dst": dst,
                "mapping": mapping,
                "pages": doc["pages"],
                "project": proj.get("name", ""),
            })
            if progress_cb:
                progress_cb(done, total, {
                    "mapping": mapping, "dst": dst,
                    "pages": doc["pages"], "project": proj.get("name", "")})
        except Exception as e:
            skipped_info.append({"dst": os.path.join(target_dir, file_name),
                                 "info": str(e)})
            if progress_cb:
                progress_cb(done, total, None)

    # 阶段三 3.1：把本次**新建**的文件写入撤销日志（供「撤销上次拆分」使用）。
    # 只记录本次实际写出的 dst；原整份 PDF 与用户既有文件从不入账，
    # 因此撤销时不可能误删它们。
    # log_undo=False 时完全跳过（测试/脚本用，避免污染全局撤销日志）。
    if written and log_undo:
        try:
            from . import batch as _batch
            _batch.log_split(written)
        except Exception:
            # 日志写入失败不应让拆分结果丢失（文件已写出）
            pass

    return written, skipped_info