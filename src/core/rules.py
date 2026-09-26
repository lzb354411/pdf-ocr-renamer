# -*- coding: utf-8 -*-
"""映射规则管理（读写 Excel）与规则匹配（自 v2.0 单文件脚本迁移）

职责：
- 规则 Excel 的读取 / 保存 / 导入（带回滚保护）
- OCR 文本清洗与规则匹配（严格包含 → 模糊匹配）
- 文件名构建与冲突处理
"""

import os
import shutil

from . import paths

# ============ 可调参数（与原脚本一致） ============
# 模糊匹配相似度阈值（严格包含匹配失败时启用）
FUZZY_THRESHOLD = 0.7
# 标题长度超过此值时才参与模糊匹配（避免短词误匹配）
FUZZY_MIN_LEN = 4

# 默认映射规则（首次运行自动生成）
DEFAULT_RULES = [
    {"识别内容": "配电网工程开工报告", "映射内容": "开工报告"},
    {"识别内容": "配电网工程竣工报告", "映射内容": "竣工报告"},
    {"识别内容": "架空配电线路安装施工记录", "映射内容": "架空配电线路安装施工记录"},
    {"识别内容": "接地装置接地电阻测试报告", "映射内容": "接地电阻测试报告"},
    {"识别内容": "竣工图", "映射内容": "竣工图纸"},
]


# ============================================================
# 规则文件读写
# ============================================================

def create_default_rules(path):
    """首次运行时创建包含默认规则的 Excel 文件"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "映射规则"
    ws.append(["识别内容", "映射内容"])
    for rule in DEFAULT_RULES:
        ws.append([rule["识别内容"], rule["映射内容"]])
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 32
    wb.save(path)


def read_rules(path):
    """读取映射规则 Excel，返回 [{"识别内容": str, "映射内容": str}, ...]

    智能识别策略（兼容各种格式的用户 Excel）：
    1. 遍历所有 sheet，找到包含「识别内容」和「映射内容」表头的 sheet
    2. 表头可在前 10 行任意位置（兼容首行有标题/说明的情况）
    3. 表头列可在任意位置（不依赖固定列序）
    4. 找到表头后，从下一行开始读取数据，跳过空行
    """
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    rules = []
    try:
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            # 在前 10 行中查找包含「识别内容」和「映射内容」的表头行
            header_row_idx = -1
            key_idx = -1
            val_idx = -1
            for i, row in enumerate(rows[:10]):
                if row is None:
                    continue
                header = [str(c or "").strip() for c in row]
                k_idx = -1
                v_idx = -1
                for j, h in enumerate(header):
                    if h == "识别内容" and k_idx < 0:
                        k_idx = j
                    elif h == "映射内容" and v_idx < 0:
                        v_idx = j
                if k_idx >= 0 and v_idx >= 0:
                    header_row_idx = i
                    key_idx = k_idx
                    val_idx = v_idx
                    break
            if header_row_idx < 0:
                # 当前 sheet 未找到标准表头，跳过尝试下一个 sheet
                continue
            # 从表头下一行开始读取数据
            for row in rows[header_row_idx + 1:]:
                if row is None:
                    continue
                key_val = row[key_idx] if key_idx < len(row) else None
                map_val = row[val_idx] if val_idx < len(row) else None
                if key_val is None and map_val is None:
                    continue
                key_str = str(key_val).strip() if key_val is not None else ""
                map_str = str(map_val).strip() if map_val is not None else ""
                if key_str:
                    rules.append({"识别内容": key_str, "映射内容": map_str})
            if rules:
                # 已找到规则数据，不再遍历其他 sheet
                break
    finally:
        wb.close()
    return rules


def save_rules(path, rules):
    """将规则列表保存为标准格式的 Excel（规则管理界面保存时调用）"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "映射规则"
    ws.append(["识别内容", "映射内容"])
    for rule in rules:
        ws.append([str(rule.get("识别内容", "")).strip(),
                   str(rule.get("映射内容", "")).strip()])
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 32
    wb.save(path)


def load_rules_safe():
    """加载当前规则文件；不存在则生成默认规则。

    返回 (rules, message)：message 描述加载情况（用于界面提示）。
    """
    rules_file = paths.get_rules_file()
    if not os.path.exists(rules_file):
        create_default_rules(rules_file)
        rules = list(DEFAULT_RULES)
        return rules, "未发现规则文件，已生成默认规则（{} 条）".format(len(rules))
    rules = read_rules(rules_file)
    return rules, "已加载规则 {} 条".format(len(rules))


def import_rules(src_path):
    """导入外部规则 Excel 到规则文件夹（覆盖现有规则）

    安全策略：
    1. 先备份现有规则文件（导入失败时自动回滚）
    2. 复制新文件后立即读取验证
    3. 验证失败（文件不可读或规则为空）则回滚到备份
    返回读取到的规则列表。
    """
    dst_path = paths.get_rules_file()
    # 备份现有规则文件
    backup_path = dst_path + ".bak"
    has_backup = os.path.exists(dst_path)
    if has_backup:
        shutil.copy2(dst_path, backup_path)
    try:
        shutil.copy2(src_path, dst_path)
        # 立即读取验证：确保新文件可读且包含有效规则
        rules = read_rules(dst_path)
        if not rules:
            # 新文件无规则，回滚
            if has_backup:
                shutil.copy2(backup_path, dst_path)
            raise ValueError(
                "导入的文件未读取到有效规则，请确认 Excel 中包含"
                "「识别内容」和「映射内容」两列")
        # 导入成功，清理备份
        if os.path.exists(backup_path):
            os.remove(backup_path)
        return rules
    except Exception:
        # 导入过程异常，回滚保护现有规则
        if has_backup and os.path.exists(backup_path):
            shutil.copy2(backup_path, dst_path)
            if os.path.exists(backup_path):
                os.remove(backup_path)
        raise


# ============================================================
# 文本清洗与规则匹配
# ============================================================

def normalize_text(text):
    """清洗 OCR 文本，提高匹配稳定性

    - 全角字符转半角（字母/数字/标点）
    - 去除所有空白、常见标点符号
    - 仅保留中文、字母、数字
    """
    if not text:
        return ""
    # 全角转半角
    table = {0xFF01: "!", 0xFF02: '"', 0xFF03: "#", 0xFF04: "$",
             0xFF05: "%", 0xFF06: "&", 0xFF07: "'", 0xFF08: "(",
             0xFF09: ")", 0xFF0A: "*", 0xFF0B: "+", 0xFF0C: ",",
             0xFF0D: "-", 0xFF0E: ".", 0xFF0F: "/",
             0xFF1A: ":", 0xFF1B: ";", 0xFF1C: "<", 0xFF1D: "=",
             0xFF1E: ">", 0xFF1F: "?", 0xFF20: "@",
             0xFF3B: "[", 0xFF3C: "\\", 0xFF3D: "]", 0xFF3E: "^",
             0xFF3F: "_", 0xFF40: "`", 0xFF5B: "{", 0xFF5C: "|",
             0xFF5D: "}", 0xFF5E: "~"}
    # 全角数字 0xFF10-0xFF19 → 0-9
    for i in range(10):
        table[0xFF10 + i] = str(i)
    # 全角大写字母 0xFF21-0xFF3A → A-Z
    for i in range(26):
        table[0xFF21 + i] = chr(65 + i)
    # 全角小写字母 0xFF41-0xFF5A → a-z
    for i in range(26):
        table[0xFF41 + i] = chr(97 + i)
    text = text.translate(table)
    # 去除所有空白
    text = "".join(text.split())
    # 去除常见中英文标点
    punct = "，。、；：？！「」『』《》（）【】<>,.;:!?'\"()[]{}-_/\\|*#@`~"
    for ch in punct:
        text = text.replace(ch, "")
    return text


def fuzzy_similarity(a, b):
    """计算两个字符串的相似度（0~1），基于最长公共子序列"""
    if not a or not b:
        return 0.0
    from difflib import SequenceMatcher
    return SequenceMatcher(None, a, b).ratio()


def _fuzzy_threshold(key_len):
    """按关键词长度调整模糊匹配阈值

    短关键词（如 4 字的「开工报告」）只凭 4 个字的相似度极易误配：
    「开工时间」「竣工时间」与「开工报告」的相似度就有 0.75。
    因此关键词越短，要求相似度越高；长关键词维持宽松阈值容错 OCR 错字。
    """
    if key_len < 5:
        return 0.85
    if key_len < 7:
        return 0.78
    return FUZZY_THRESHOLD


# ============================================================
# OCR 错字纠正（标题匹配链路）
# ============================================================
#
# 背景（阶段一 R-1 第一层根因）：
#   project_config.DEFAULT_CONFIG["ocr_fixes"] 中已有「峻工→竣工」，
#   但本模块原先完全没有引用 project_config，纠正表只被 name_parser
#   （工程名称提取阶段）使用。因此标题→文件名的匹配阶段不做任何纠错，
#   「配电网工程峻工报告」原样进入匹配，最终被误判为「开工报告」。
#
#   此处把纠正接入 match_rule 入口。为避免与 name_parser 逻辑分叉，
#   实现委托给 project_config.apply_ocr_fixes（长键优先、空表兼容）。
#
# 依赖方向：rules → project_config → paths，无反向导入，不存在循环导入。

def _apply_ocr_fixes(text):
    """按 project_config.ocr_fixes 纠正标题文本中的 OCR 错字

    纠正表不可用（缺失/为空/异常）时原样返回，行为与接入前完全一致。
    配置带模块级缓存，避免逐页/逐规则重复读盘。
    """
    if not text:
        return text
    try:
        from . import project_config as PC
        return PC.apply_ocr_fixes(text, _ocr_fix_config())
    except Exception:
        # 配置读取失败绝不能让匹配链路崩溃：退化为不纠正
        return text


# OCR 纠正表缓存：(配置来源标识, {错误写法: 正确写法})
_OCR_FIX_CACHE = None


def _ocr_fix_config():
    """读取一次配置并缓存（同一进程内配置路径固定）"""
    global _OCR_FIX_CACHE
    if _OCR_FIX_CACHE is None:
        from . import project_config as PC
        _OCR_FIX_CACHE = PC.load_config()
    return _OCR_FIX_CACHE


def clear_ocr_fix_cache():
    """清空 OCR 纠正表缓存（配置变更或测试注入时调用）"""
    global _OCR_FIX_CACHE
    _OCR_FIX_CACHE = None


def _pick_better(score, key_len, best_score, best_key_len):
    """模糊匹配平局消歧：判断当前 rule 是否应取代当前最优

    背景（阶段一 R-1 第二层根因）：
        原实现用 `score >= thr and score > best_score`，平局时先遍历者胜出。
        「配电网工程峻工报告」对「配电网工程开工报告」与「配电网工程竣工报告」
        的相似度完全相同，因规则表中开工报告在前，竣工页被误判为开工报告。

    确定性优先级链（不再依赖规则表遍历顺序）：
        1. 分数更高者胜；
        2. 分数相同时，**关键词更长者胜**（更长 = 更具体、更少歧义）；
        3. 仍相同时，视为等价 —— 保留先遍历者（规则表顺序），结果同样确定；
           映射内容相同的重复规则不会来回替换。

    说明：长度优先只作用于「同分」这一狭窄场景，不会让长关键词抢走
    本应命中短关键词的场景（分数更高者始终优先），因此不破坏既有
    「短关键词误配防护」用例（见 test_core.py 短关键词用例）。
    """
    if score != best_score:
        return score > best_score
    # 同分：更长更具体者胜；同长则等价，保持规则表顺序
    return key_len > best_key_len


def match_rule(title_text, rules):
    """根据标题文本匹配映射规则，返回映射内容（未匹配返回 None）

    匹配策略（依次尝试）：
    0. OCR 错字纠正：入口处先按 project_config.ocr_fixes 纠正标题文本，
       再做归一化与匹配。纠正表为空/缺失/异常时行为与未接入时完全一致
       （向后兼容）。**注意**：纠正必须在 normalize_text 之前做，且不能
       在 normalize_text 内部做——后者是纯文本清洗，不应耦合配置。
    1. 逐行严格包含：OCR 标题可能是多行拼接（第一页顶部多行文字），
       按行从上到下匹配——页面最上方一行最可能是真正的标题，
       避免下部行里的干扰文字（如表格中其他报告的标题）抢先命中；
       同一行内多条规则命中时，取「识别内容」更长（更具体）的规则
    2. 模糊匹配：整段标题与识别内容相似度 ≥ 阈值（容错 OCR 错字），
       短关键词阈值更高（见 _fuzzy_threshold），防止常见词组误匹配；
       分数相同时按 _pick_better 的确定性规则消歧（不再「先遍历者胜」）
    """
    if not title_text:
        return None
    title_text = _apply_ocr_fixes(title_text)
    title_clean = normalize_text(title_text)
    if not title_clean:
        return None

    # 第一轮：逐行严格包含匹配（最快、最准）
    for raw_line in title_text.splitlines():
        line_clean = normalize_text(raw_line)
        if not line_clean:
            continue
        line_match = None
        best_key_len = -1
        for rule in rules:
            key = normalize_text(_apply_ocr_fixes(
                str(rule.get("识别内容", ""))))
            if key and key in line_clean and len(key) > best_key_len:
                best_key_len = len(key)
                line_match = rule
        if line_match is not None:
            return str(line_match.get("映射内容", "")).strip()

    # 第二轮：模糊匹配（容错 OCR 错字/漏字）
    best_match = None
    best_score = 0.0
    best_key_len = 0
    for rule in rules:
        key = normalize_text(_apply_ocr_fixes(
            str(rule.get("识别内容", ""))))
        if not key or len(key) < FUZZY_MIN_LEN:
            continue
        thr = _fuzzy_threshold(len(key))
        # 标题中可能包含规则关键词的子串，滑动窗口找最高相似度
        if len(title_clean) >= len(key):
            max_sub_score = 0.0
            window = len(key)
            # 步长优化：避免全量滑动
            step = max(1, len(key) // 4)
            for start in range(0, len(title_clean) - window + 1, step):
                sub = title_clean[start:start + window]
                score = fuzzy_similarity(sub, key)
                if score > max_sub_score:
                    max_sub_score = score
                    if max_sub_score >= 1.0:
                        break
            score = max_sub_score
        else:
            score = fuzzy_similarity(title_clean, key)
        if score < thr:
            continue
        # 平局消歧：分数相同不再「先遍历者胜」，改用确定性规则（见 _pick_better）
        if best_match is None or _pick_better(score, len(key),
                                              best_score, best_key_len):
            best_score = score
            best_key_len = len(key)
            best_match = rule
    if best_match is not None:
        return str(best_match.get("映射内容", "")).strip()
    return None


# ============================================================
# 文件名工具
# ============================================================

def sanitize_filename(name):
    """清理文件名中的非法字符"""
    invalid = '<>:"/\\|?*'
    for ch in invalid:
        name = name.replace(ch, "_")
    return name.strip().rstrip(".")


def resolve_conflict(path):
    """目标路径已存在时，添加序号后缀避免覆盖"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists("{}_{}{}".format(base, i, ext)):
        i += 1
    return "{}_{}{}".format(base, i, ext)


def build_new_name(mapping_content, pdf_path, by_folder):
    """根据命名规则构建新文件名

    by_folder=True：文件夹名 + 映射内容（如：模板开工报告.pdf）
    by_folder=False：仅映射内容（如：开工报告.pdf）
    """
    mapping = sanitize_filename(mapping_content)
    if not mapping:
        return ""
    if by_folder:
        folder_name = os.path.basename(os.path.dirname(pdf_path))
        folder_name = sanitize_filename(folder_name)
        return "{}{}.pdf".format(folder_name, mapping)
    return "{}.pdf".format(mapping)
