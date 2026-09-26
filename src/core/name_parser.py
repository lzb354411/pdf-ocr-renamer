# -*- coding: utf-8 -*-
"""工程名称解析器（配置驱动，替代 splitter.py 中的硬编码逻辑）

背景（P0-1）：
    原实现用写死的「某省」作起始锚点、「业扩配套工程」作结束锚点，
    只能覆盖 2026年业扩项目进度表.xlsx 中 373/560（67%）的名称形态；
    形态 C/D（低压维修项目 / 项目包N，共 187 条，33%）完全提取不到名称，
    且失效时静默返回空串，最终流入「未匹配项目」文件夹。

本模块实现「配置驱动 + 多形态解析」：
    1. 所有地域词 / 业务后缀 / 杂质词来自 project_config，不在代码中写死；
    2. 支持 4 种已归纳的名称形态（见 project_config 模块文档）；
    3. 每种形态给出候选与置信度，取最高分；
    4. 解析失败时返回空串 + 原因，便于界面提示（而非静默）。

对外主接口：
    extract_name(text, config)  ->  (name, debug_info)
    score_fix(name, config)     ->  应用 OCR 错字纠正
"""

import re

from . import project_config as PC


# ============================================================
# 工具
# ============================================================

def _norm(text):
    """工程名称提取用的轻度归一化

    与 rules.normalize_text 不同：这里保留中文与数字的整体结构，
    只去掉空白与全角标点，避免把「1#变」的 # 与数字粘连关系破坏掉。

    特别注意：小数点必须保留（「0.4kV」拆成「04kV」会导致后缀匹配失败）。
    """
    if not text:
        return ""
    text = str(text)
    # 全角数字/字母转半角
    out = []
    for ch in text:
        o = ord(ch)
        if 0xFF10 <= o <= 0xFF19:
            out.append(chr(o - 0xFF10 + ord("0")))
        elif 0xFF21 <= o <= 0xFF3A:
            out.append(chr(o - 0xFF21 + ord("A")))
        elif 0xFF41 <= o <= 0xFF5A:
            out.append(chr(o - 0xFF41 + ord("a")))
        elif ch in "，。、；：？！「」『』《》【】":
            # 中文标点：剔除（但保留全角括号，见下）
            continue
        elif ch in "（）":
            # 全角括号转半角并保留：企业名中「（个体工商户）」是名称组成部分，
            # 直接剔除会让提取结果与原始名称不一致。
            out.append("(" if ch == "（" else ")")
        else:
            out.append(ch)
    text = "".join(out)
    # 去掉空白
    text = "".join(text.split())
    # 电压等级写法规范化：0.4KV → 0.4kV；10KV → 10kV
    text = re.sub(r"(\d+(?:\.\d+)?)\s*[kK][vV]", lambda m: m.group(1) + "kV", text)
    return text


def _apply_ocr_fixes(text, config):
    """应用配置中的 OCR 错字纠正表

    实现在 project_config.apply_ocr_fixes（与 rules.match_rule 共用一套逻辑：
    长键优先、空表向后兼容），此处仅做薄封装以保持既有调用点不变。
    """
    return PC.apply_ocr_fixes(text, config)


def _strip_noise(text, config):
    """剔除配置中的杂质字段（标签词）"""
    if not text:
        return text
    for w in (config.get("noise_words") or []):
        if w and w in text:
            text = text.replace(w, "")
    # 剔除「工程名称」「工程编号」等标签的 OCR 变体
    for key in ("name_label_pattern", "code_label_pattern"):
        pat = config.get(key)
        if pat:
            try:
                text = re.sub(pat, "", text)
            except re.error:
                pass
    return text


def _apply_symbol_fixes(text, config):
    """应用配置中的「编号符号等价替换」（阶段二）

    真实样本：图纸标题「某堤口2#配变」被 OCR 识别为「某堤口2*配变」，
    `#` 变成 `*`。`*` 不在名称合法字符集内，会导致整个名称被判非法、
    静默返回空串。此处按 `symbol_fixes` 配置做等价替换（配置驱动）。
    """
    if not text:
        return text
    fixes = config.get("symbol_fixes")
    if not isinstance(fixes, dict):
        return text
    for bad, good in fixes.items():
        if isinstance(bad, str) and bad and isinstance(good, str) and bad in text:
            text = text.replace(bad, good)
    return text


def _max_digit_run(text):
    """最长连续数字长度"""
    runs = re.findall(r"\d+", text or "")
    return max((len(r) for r in runs), default=0)


def _expanded_suffixes(config):
    """返回「配置后缀 + 其宽松形态」的完整候选集

    宽松形态包括：
      - 去掉尾部「工程/项目」的形式（「业扩配套工程」→「业扩配套」），
        用于兼容真实数据中「…0.4kV业扩配套」这类截断式名称；
      - 加电压等级前缀的形式（「业扩配套工程」→「10kV业扩配套工程」）。
    """
    base = [s for s in (config.get("business_suffixes") or []) if s]
    out = list(base)
    for s in base:
        for tail in ("工程", "项目"):
            if s.endswith(tail) and len(s) > len(tail) + 2:
                out.append(s[:-len(tail)])
    for s in list(out):
        if s.startswith(("业扩", "配套")) and "kV" not in s:
            for v in ("10kV", "0.4kV"):
                out.append(v + s)
    return sorted(set(out), key=lambda s: (-len(s), s))


def _is_valid_name(cand, config):
    """按配置判定候选是否为合法工程名称

    重要设计变更（对比旧实现）：
        旧实现要求名称必须含「业扩」或「配套工程」这类业务关键词，
        导致用电主体是个人姓名时（如「某省某市某县…某某某0.4kV业扩配套工程」）
        被判为非法。实际上名称的合法性由「结构」决定——
        含业务后缀 + 长度合理 + 无杂质，而非必须含某个业务词。

        因此这里把「有效关键词」从**必要条件**降级为**打分项**
        （见 extract_name 的 score 计算），只保留结构性硬校验。
    """
    if not cand:
        return False
    cand = cand.strip()
    min_len = int(config.get("min_name_len", 8) or 8)
    max_len = int(config.get("max_name_len", 60) or 60)
    if not (min_len <= len(cand) <= max_len):
        return False
    # 必须整体是中文/数字/#/字母/括号组合，不含残留标点
    # 字符集由配置 `name_charset` 提供（默认含 # 与 OCR 变体 * ※）
    charset = config.get("name_charset") or r"\u4e00-\u9fa5A-Za-z0-9#（）()\-\."
    try:
        if not re.fullmatch("[" + charset + "]+", cand):
            return False
    except re.error:
        # 配置写坏的字符集不应让解析崩溃：退化为默认字符集
        if not re.fullmatch(r"[\u4e00-\u9fa5A-Za-z0-9#（）()\-\.]+", cand):
            return False
    # 连续数字过长 → 把编号卷进来了
    max_run = int(config.get("max_digit_run", 4) or 4)
    if _max_digit_run(cand) > max_run:
        return False
    # 杂质词
    for w in (config.get("noise_words") or []):
        if w and w in cand:
            return False
    # 结构性要求：必须含业务后缀之一（含宽松形态），证明它是工程名而非纯地名
    suffixes = _expanded_suffixes(config)
    if suffixes and not any(s in cand for s in suffixes):
        return False
    return True


# ============================================================
# 后缀 / 前缀定位（配置驱动）
# ============================================================

def _find_business_suffix(text, start, config):
    """在 text[start:] 中寻找业务后缀，返回 (end_index, suffix)

    取「最早出现」的后缀；同一起点取最长者。
    返回 (None, None) 表示未找到。

    补充：真实数据中存在「…0.4kV业扩配套」（结尾无「工程」二字）的打包类名称，
    因此在标准后缀之外，额外尝试「后缀去掉尾部工程/项目」形态，
    以及「kV+业务词」的组合形态。
    """
    suffixes = config.get("business_suffixes") or []
    # 长后缀优先，避免「维修项目」抢先于「低压台区维修项目」；
    # _expanded_suffixes 已包含宽松形态（去尾部工程/项目、加电压等级前缀），
    # 用于兼容真实数据中「…0.4kV业扩配套」这类截断式名称。
    expanded = _expanded_suffixes(config) or sorted(
        [s for s in suffixes if s], key=lambda s: (-len(s), s))

    best_end = None
    best_suffix = None
    for suf in expanded:
        idx = text.find(suf, start)
        if idx < 0:
            continue
        end = idx + len(suf)
        if best_end is None or idx < best_end - len(best_suffix or "") or \
                (idx == best_end - len(best_suffix) and len(suf) > len(best_suffix)):
            if best_end is None or idx < best_end:
                best_end = end
                best_suffix = suf
    if best_end is None:
        return None, None
    return best_end, best_suffix


def _region_tolerance(word, config):
    """计算某个地域词允许的容错字符数（配置驱动）

    - `region_fuzzy_tolerance` == 0 或配置缺失 → 0（关闭容错，等价于精确查找）
    - 整数 → 直接取用
    - "auto" → ⌊len(word)/3⌋ 自适应（长词放宽、短词严格）

    最终受 `region_fuzzy_max_tolerance` 上限约束，且不超过 len(word)-1
    （否则任意同长度词都会命中，等于失去判别力）。
    """
    raw = config.get("region_fuzzy_tolerance", 0)
    if raw is None:
        return 0
    if isinstance(raw, str):
        if raw.strip().lower() != "auto":
            try:
                raw = int(raw)
            except (TypeError, ValueError):
                return 0
        else:
            raw = len(word) // 3
    try:
        tol = int(raw)
    except (TypeError, ValueError):
        return 0
    if tol <= 0:
        return 0
    # 上限约束
    try:
        cap = int(config.get("region_fuzzy_max_tolerance", tol))
    except (TypeError, ValueError):
        cap = tol
    tol = min(tol, max(0, cap))
    # 不能超过 len-1，否则全同长度词都命中
    return max(0, min(tol, max(0, len(word) - 1)))


def _char_mismatch(a, b):
    """两个等长字符串的不同字符数；长度不等返回一个大值"""
    if len(a) != len(b):
        return 999
    return sum(1 for x, y in zip(a, b) if x != y)


def find_region_word(text, config, words=None):
    """在 text 中容错查找地域词，返回 (index, word, mismatches) 或 None

    容错规则（阶段二 R-2，配置驱动）：
        1. 只在**等长**窗口上比较，允许至多 `tol` 个字符不同
           （禁止增删字符 —— 否则「某省」可能匹配到长度不同的片段）；
        2. tol 由 `_region_tolerance(word, config)` 决定，可配置为 0 关闭；
        3. 若 `region_fuzzy_require_suffix_char` 为真，则要求窗口的**最后一字**
           等于地域词的最后一字（如都须以「县」结尾），
           这样「某县→甲县/乙县/丙县」可命中，而「某县→沛丰」不会；
        4. 多个候选时取**不匹配数最少**者，其次取**位置最靠前**者（确定性）。

    返回的 index 是窗口起始位置，与 `text.find(word)` 语义一致，可直接作为
    名称起点使用。

    **设计原则（宁可漏也不能错）**：默认容错度为 1 且要求区划后缀字一致，
    避免把无关文本误判为地域词起点。
    """
    if not text or not words:
        return None
    min_len = 2
    try:
        min_len = int(config.get("region_fuzzy_min_len", 2) or 2)
    except (TypeError, ValueError):
        min_len = 2
    require_suffix = bool(config.get("region_fuzzy_require_suffix_char", True))

    best = None  # (mismatches, index, word)
    for word in words:
        if not word or len(word) < min_len:
            continue
        tol = _region_tolerance(word, config)
        wlen = len(word)
        if wlen > len(text):
            continue
        if tol <= 0:
            # 容错关闭：退化为精确查找
            idx = text.find(word)
            if idx >= 0 and (best is None or idx < best[1]):
                best = (0, idx, word)
            continue
        for start in range(0, len(text) - wlen + 1):
            window = text[start:start + wlen]
            if require_suffix and window[-1] != word[-1]:
                continue
            mm = _char_mismatch(window, word)
            if mm > tol:
                continue
            # 同分时取位置最靠前；mismatch 更少者优先
            if best is None or (mm, start) < (best[0], best[1]):
                best = (mm, start, word)
    if best is None:
        return None
    return best[1], best[2], best[0]


def _find_start(text, config):
    """定位名称起点：优先地域词（容错），其次长数字编号，再次年份

    返回 (start_index, how)：
        how ∈ {"region", "region_fuzzy", "code_prefix", "code", "year", "head"}

    阶段二 R-2：地域词查找改为**容错**（见 `find_region_word`），
    容错度由 `region_fuzzy_tolerance` 控制；为 0 时退化为精确查找，
    行为与改动前完全一致（向后兼容）。
    """
    region_words = sorted([w for w in (config.get("region_words") or []) if w],
                          key=lambda w: (-len(w), w))
    # 1) 地域词（精确优先，避免容错抢走本来能精确命中的位置）
    best = None
    for w in region_words:
        idx = text.find(w)
        if idx >= 0 and (best is None or idx < best):
            best = idx
    if best is not None:
        return best, "region"

    # 1b) 精确失败 → 容错查找
    hit = find_region_word(text, config, region_words)
    if hit is not None:
        idx, _word, mm = hit
        return idx, "region_fuzzy" if mm > 0 else "region"

    # 2) 长数字编号（>=10 位）：名称常在编号之后的主体名
    m = re.search(r"\d{10,}", text)
    if m:
        # 起点取编号「之前」的地域性前缀（若有），否则取编号之后
        # 形态 A：某省某市某县 + 编号 + 主体 + 后缀
        #         → 起点应在编号之前（保留区划前缀）
        prefix = text[:m.start()]
        if len(prefix) >= 2 and re.fullmatch(r"[\u4e00-\u9fa5]+", prefix):
            return 0, "code_prefix"
        return m.end(), "code"

    # 3) 年份开头（形态 C/D：2026年...）
    m = re.match(r"^(20\d{2})年?", text)
    if m:
        return 0, "year"

    return 0, "head"


# ============================================================
# 多形态解析
# ============================================================

def _candidate_from_form_a(text, config):
    """形态 A：某省某市某县<长编号><用电主体><kV业扩配套工程>

    名称整体就是「区划前缀 + 编号 + 主体 + 后缀」，直接取整个片段。
    """
    out = []
    for m in re.finditer(r"\d{10,}", text):
        seg = text[:m.end()]           # 含编号及其前缀
        # 找业务后缀作为终点
        end, suf = _find_business_suffix(text, m.end(), config)
        if end is None:
            continue
        cand = text[:end]
        cand = re.sub(r"\d{10,}", "", cand)     # 去掉编号本身
        out.append((cand, 1.0, "formA"))
    return out


def _candidate_from_year(text, config):
    """形态 C/D：<年份><单位/区划><台区描述><业务后缀>[编号|包N]

    以年份为起点，业务后缀为终点。
    同时兼容「...低压维护项目包1」形态（后缀后带 包N）。
    """
    out = []
    for m in re.finditer(r"20\d{2}", text):
        start = m.start()
        end, suf = _find_business_suffix(text, start, config)
        if end is None:
            continue
        # 后缀后可能还有「包N」
        tail = text[end:]
        m2 = re.match(r"(包\d+)", tail)
        if m2:
            end += m2.end()
        cand = text[start:end]
        # 形态 C 的结尾编号已在 _find_business_suffix 之前被截断，
        # 但若编号夹杂在描述与后缀之间（如 ...变台区维修项目3225...），
        # 后缀定位已排除；此处再兜底去掉尾部编号
        cand = re.sub(r"\d{11,}$", "", cand)
        out.append((cand, 0.95, "formCD"))
    return out


def _candidate_from_region(text, config):
    """兜底：地域词起 → 业务后缀止（形态 B 及未知变体）

    阶段二 R-2：除精确命中外，额外尝试**容错命中**的位置
    （如「甲县…低压台区维修项目」中的「甲县」应作为「某县」的容错起点）。
    """
    out = []
    region_words = sorted([w for w in (config.get("region_words") or []) if w],
                          key=lambda w: (-len(w), w))

    def _emit(start):
        end, suf = _find_business_suffix(text, start, config)
        if end is None:
            return
        tail = text[end:]
        m2 = re.match(r"(包\d+)", tail)
        if m2:
            end += m2.end()
        cand = text[start:end]
        cand = re.sub(r"\d{10,}", "", cand)
        out.append((cand, 0.9, "formB"))

    for w in region_words:
        for m in re.finditer(re.escape(w), text):
            _emit(m.start())

    # 容错命中：仅在精确命中位置之外补充（避免重复候选）
    exact_starts = {m.start() for w in region_words
                    for m in re.finditer(re.escape(w), text)}
    hit = find_region_word(text, config, region_words)
    if hit is not None and hit[0] not in exact_starts and hit[2] > 0:
        _emit(hit[0])
    return out


# ============================================================
# 对外主接口
# ============================================================

def extract_name(text, config=None):
    """从一段 OCR 文本中提取工程名称

    参数：
        text:   原始文本（可以是整页 OCR 的拼接结果）
        config: 配置 dict；None 时读取默认/用户配置

    返回 (name, debug)：
        name:  提取到的工程名称（已应用 OCR 纠正与杂质剔除）；未提取到为 ""
        debug: {"how": 命中形态, "score": 置信度, "candidates": [...]}
    """
    cfg = config if config is not None else PC.load_config()
    raw = _norm(text)
    if not raw:
        return "", {"how": "empty", "score": 0.0, "candidates": []}

    cleaned = _apply_ocr_fixes(_strip_noise(raw, cfg), cfg)
    # 编号符号等价替换（如 OCR 把「2#配变」识别成「2*配变」）
    cleaned = _apply_symbol_fixes(cleaned, cfg)

    # 收集各形态候选
    cands = []
    cands.extend(_candidate_from_form_a(cleaned, cfg))
    cands.extend(_candidate_from_year(cleaned, cfg))
    cands.extend(_candidate_from_region(cleaned, cfg))

    # 兜底：整段直接作为候选（当文本本身就是一行名称时）
    start, how = _find_start(cleaned, cfg)
    end, suf = _find_business_suffix(cleaned, start, cfg)
    if end is not None:
        tail = cleaned[end:]
        m2 = re.match(r"(包\d+)", tail)
        if m2:
            end += m2.end()
        cand = re.sub(r"\d{10,}", "", cleaned[start:end])
        cands.append((cand, 0.8, "fallback_%s" % how))

    # 清洗 + 过滤
    scored = []
    for cand, base_score, how_tag in cands:
        c = _apply_ocr_fixes(cand, cfg)
        c = c.strip("：:号-—_ ")
        c = re.sub(r"\d{10,}", "", c)      # 残留长编号
        c = c.strip()
        if not _is_valid_name(c, cfg):
            continue
        # 打分：形态基础分 + 完整度加成（含后缀更可信、更长更完整）
        score = base_score
        if suf and suf in c:
            score += 0.3
        for kw in (cfg.get("valid_keywords") or []):
            if kw and kw in c:
                score += 0.15
                break
        score += min(0.2, len(c) * 0.004)
        scored.append((score, c, how_tag))

    if not scored:
        return "", {"how": "no_match", "score": 0.0, "candidates": []}

    scored.sort(key=lambda t: (-t[0], -len(t[1])))
    best_score, best_name, best_how = scored[0]
    return best_name, {
        "how": best_how,
        "score": round(best_score, 3),
        "candidates": [{"name": c, "score": round(s, 3), "how": h}
                       for s, c, h in scored[:5]],
    }


def extract_name_from_rows(rows, config=None):
    """从「一页/一项目的 OCR 行列表」中提取工程名称

    参数 rows: [text, ...] 或 [(y, text), ...]，两种都兼容。
    多行会拼接后解析（工程名称常跨行/跨单元格被 OCR 断开）。
    """
    texts = []
    for r in rows or []:
        if isinstance(r, (tuple, list)) and len(r) >= 2:
            texts.append(str(r[1]))
        else:
            texts.append(str(r))
    return extract_name("".join(texts), config)


def extract_best_name(rows_pool, config=None):
    """从项目多页文本中投票选出最优工程名称

    参数 rows_pool: [[rows_page1], [rows_page2], ...]
    返回 (name, debug)
    """
    cfg = config if config is not None else PC.load_config()
    all_cands = []
    for rows in rows_pool or []:
        name, dbg = extract_name_from_rows(rows, cfg)
        if name:
            all_cands.append((name, dbg.get("score", 0.0), dbg))
        for c in dbg.get("candidates", []):
            all_cands.append((c["name"], c["score"], dbg))

    if not all_cands:
        return "", {"how": "no_match", "score": 0.0, "candidates": []}

    # 同名合并取最高分
    best = {}
    for name, score, dbg in all_cands:
        if name not in best or score > best[name][0]:
            best[name] = (score, dbg)
    # 选分最高；同分取更长（更完整）
    ranked = sorted(best.items(), key=lambda kv: (-kv[1][0], -len(kv[0])))
    top_name, (top_score, top_dbg) = ranked[0]
    return top_name, {
        "how": top_dbg.get("how", ""),
        "score": round(top_score, 3),
        "candidates": [{"name": n, "score": round(s, 3)}
                       for n, (s, _d) in ranked[:5]],
    }
