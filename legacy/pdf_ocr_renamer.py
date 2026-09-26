# -*- coding: utf-8 -*-
"""
PDF 批量识别重命名脚本（热拔插脚本）

功能：
- 递归扫描选定文件夹下的所有 PDF 文件（含子文件夹）
- 用 OCR 识别每份 PDF 第一页顶部文字作为标题候选（支持扫描件）
- 按映射规则（Excel）匹配标题，得到映射内容
- 根据命名规则重命名文件：文件夹名+映射内容 / 仅映射内容
- 未匹配到规则的文件保持原名，并在日志中汇报

规则文件存放于 rules/ 文件夹（与平台 templates/ 同级），
文件名为「PDFocr 映射规则.xlsx」，两列：识别内容 / 映射内容。

依赖：openpyxl（规则读写）、pypdfium2（PDF 渲染为图片）、
      rapidocr_onnxruntime（OCR 文字识别）
"""

import os
import sys
import shutil
import time
import threading
import concurrent.futures

# 限制 onnxruntime/OpenCV 内部线程数，避免多线程并发时线程爆炸导致栈溢出
os.environ.setdefault("ORT_NUM_THREADS", "1")
os.environ.setdefault("OMP_NUM_THREADS", "1")

# ============ 脚本元信息 ============
TOOL_NAME = "PDF批量识别重命名"
TOOL_DESC = "OCR 识别 PDF 标题，按映射规则重命名文件（支持扫描件）"
SCRIPT_VERSION = "2.0.0"
CATEGORY = "单独功能"
AUTHOR = ""
UPDATE_DATE = "2026-07-29"

# AI 调用提示（用于 CLI/MCP Server 自描述）
PARAMS_HINT = "pdf_folder=PDF目录 import_excel_rules=规则Excel路径 rename_by_folder=true|false"
DESTRUCTIVE = True  # 会重命名文件，需用户确认

UI_SCHEMA = [
    {
        "key": "pdf_folder",
        "label": "PDF 文件夹",
        "type": "folder_picker",
        "required": True,
        "placeholder": "点击右侧按钮选择包含 PDF 的文件夹",
        "help": "将递归扫描该文件夹及所有子文件夹下的 PDF 文件",
    },
    {
        "key": "import_excel_rules",
        "label": "导入映射规则（可选）",
        "type": "file_picker",
        "required": False,
        "placeholder": "选择要导入的规则 Excel 文件（支持任意文件名）",
        "help": "支持任意文件名的 Excel；需包含「识别内容」和「映射内容」两列（列位置不限）。导入后保存到规则文件夹，重启后仍可用",
    },
    {
        "key": "export_rules",
        "label": "执行时导出规则到桌面",
        "type": "checkbox",
        "default": False,
        "help": "将当前映射规则复制一份到桌面",
    },
    {
        "key": "open_rules_folder",
        "label": "执行后打开规则文件夹",
        "type": "checkbox",
        "default": False,
    },
    {
        "key": "rename_by_folder",
        "label": "按文件夹名称命名（文件夹名+映射内容）",
        "type": "checkbox",
        "default": False,
        "help": "如：位于「模板」文件夹的开工报告 → 模板开工报告.pdf",
    },
    {
        "key": "rename_by_mapping_only",
        "label": "仅按映射内容命名",
        "type": "checkbox",
        "default": True,
        "help": "如：开工报告 → 开工报告.pdf（两种命名规则只能选其一）",
    },
]
# ====================================

# 规则文件名
RULES_FILENAME = "PDFocr 映射规则.xlsx"

# 默认映射规则（首次运行自动生成）
DEFAULT_RULES = [
    ("配电网工程开工报告", "开工报告"),
    ("配电网工程竣工报告", "竣工报告"),
    ("架空配电线路安装施工记录", "架空配电线路安装施工记录"),
    ("接地装置接地电阻测试报告", "接地电阻测试报告"),
    ("竣工图", "竣工图纸"),
]

# 标题提取时取第一页前 N 行作为候选区域
_TITLE_LINES = 12
# OCR 行分组容差（像素）：y 坐标差距小于此值视为同一行
_ROW_TOLERANCE = 15
# PDF 渲染缩放比例（越大越清晰但越慢）
_RENDER_SCALE = 2.0
# 动态裁剪比例：依次尝试顶部 15% / 30% / 50% / 整页
_CROP_RATIOS = (0.15, 0.30, 0.50, 1.0)
# 识别文字少于此字符数时，认为该方向未识别到有效内容
_MIN_TITLE_CHARS = 4
# 模糊匹配相似度阈值（严格包含匹配失败时启用）
_FUZZY_THRESHOLD = 0.7
# 标题长度超过此值时才参与模糊匹配（避免短词误匹配）
_FUZZY_MIN_LEN = 4

# OCR 引擎单例（首次调用时加载模型，后续复用）
_OCR_ENGINE = None

# ============ 并发控制 ============
# 线程安全锁
_OCR_INIT_LOCK = threading.Lock()
_OCR_CALL_LOCK = threading.Lock()  # 保护 OCR 推理调用（RapidOCR 非线程安全）
_FILE_OP_LOCK = threading.Lock()
# 最大并发线程数（PDF 渲染/匹配并行，OCR 推理受锁保护串行执行）
_MAX_WORKERS = min(2, os.cpu_count() or 1)
# 单个 PDF 处理超时（秒）
_PDF_TIMEOUT = 120
# 总处理超时（秒），超时后强制结束未完成任务
_TOTAL_TIMEOUT = 600


# ============================================================
# 路径与规则文件管理
# ============================================================

def _get_rules_dir():
    """获取规则文件夹路径（与平台 templates/ 同级的 rules/）

    优先使用平台 API 确保 rules/ 与 templates/ 在同一基目录下；
    脚本独立运行时回退到自行计算。
    """
    try:
        from src.platform_api import PlatformAPI
        tpl_dir = PlatformAPI.get_templates_dir()
        if tpl_dir:
            base = os.path.dirname(tpl_dir)
            rules_dir = os.path.join(base, "rules")
            os.makedirs(rules_dir, exist_ok=True)
            return rules_dir
    except Exception:
        pass
    if getattr(sys, "frozen", False):
        local_appdata = os.environ.get(
            "LOCALAPPDATA", os.path.expanduser("~\\AppData\\Local"))
        base = os.path.join(local_appdata, "某工程管理软件")
    else:
        scripts_dir = os.path.dirname(os.path.abspath(__file__))
        base = os.path.dirname(scripts_dir)
    rules_dir = os.path.join(base, "rules")
    os.makedirs(rules_dir, exist_ok=True)
    return rules_dir


def _get_rules_file():
    """获取规则 Excel 文件完整路径"""
    return os.path.join(_get_rules_dir(), RULES_FILENAME)


def _create_default_rules(path):
    """首次运行时创建包含默认规则的 Excel 文件"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "映射规则"
    ws.append(["识别内容", "映射内容"])
    for key, val in DEFAULT_RULES:
        ws.append([key, val])
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 32
    wb.save(path)


def _read_rules(path):
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


def _import_rules(src_path, dst_path):
    """导入外部规则 Excel 到规则文件夹（覆盖现有规则）

    安全策略：
    1. 先备份现有规则文件（导入失败时自动回滚）
    2. 复制新文件后立即读取验证
    3. 验证失败（文件不可读或规则为空）则回滚到备份
    返回读取到的规则数量。
    """
    # 备份现有规则文件
    backup_path = dst_path + ".bak"
    has_backup = os.path.exists(dst_path)
    if has_backup:
        shutil.copy2(dst_path, backup_path)
    try:
        shutil.copy2(src_path, dst_path)
        # 立即读取验证：确保新文件可读且包含有效规则
        rules = _read_rules(dst_path)
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
        return len(rules)
    except Exception:
        # 导入过程异常，回滚保护现有规则
        if has_backup and os.path.exists(backup_path):
            shutil.copy2(backup_path, dst_path)
            if os.path.exists(backup_path):
                os.remove(backup_path)
        raise


def _export_rules_to_desktop(src_path):
    """将规则文件复制到桌面"""
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.isdir(desktop):
        desktop = os.path.expanduser("~")
    export_path = os.path.join(desktop, RULES_FILENAME)
    if os.path.exists(export_path):
        base, ext = os.path.splitext(export_path)
        i = 1
        while os.path.exists("{}_{}{}".format(base, i, ext)):
            i += 1
        export_path = "{}_{}{}".format(base, i, ext)
    shutil.copy2(src_path, export_path)
    return export_path


# ============================================================
# OCR 引擎与 PDF 标题提取
# ============================================================

def _get_ocr_engine():
    """获取 OCR 引擎单例（首次调用时加载模型，后续复用）

    RapidOCR 模型加载约 2-3 秒，复用避免重复开销。
    使用双重检查锁定保证线程安全。
    """
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        with _OCR_INIT_LOCK:
            if _OCR_ENGINE is None:
                from rapidocr_onnxruntime import RapidOCR
                _OCR_ENGINE = RapidOCR()
    return _OCR_ENGINE


def _ocr_recognize(engine, pil_image):
    """对 PIL Image 做 OCR，返回 result 列表（失败返回空列表）

    使用 _OCR_CALL_LOCK 保护调用，因为 RapidOCR/onnxruntime 非线程安全，
    多线程并发调用会导致栈溢出崩溃。
    """
    import numpy as np
    with _OCR_CALL_LOCK:
        result, _elapse = engine(np.array(pil_image))
    return result or []


def _count_chars(result):
    """统计 OCR 结果中的总字符数"""
    if not result:
        return 0
    return sum(len(str(item[1])) for item in result)


def _build_title_from_result(result):
    """从 OCR 结果构建标题文本（按行分组排序，取前 N 行）"""
    if not result:
        return ""
    items = []
    for item in result:
        box = item[0]
        text = item[1]
        top_y = min(p[1] for p in box)
        left_x = min(p[0] for p in box)
        items.append((top_y, left_x, text))
    items.sort(key=lambda t: (t[0], t[1]))
    rows = []
    cur_row = []
    cur_y = None
    for y, _x, text in items:
        if cur_y is None or abs(y - cur_y) <= _ROW_TOLERANCE:
            cur_row.append(text)
            if cur_y is None:
                cur_y = y
        else:
            rows.append("".join(cur_row))
            cur_row = [text]
            cur_y = y
    if cur_row:
        rows.append("".join(cur_row))
    return "\n".join(rows[:_TITLE_LINES])


def _detect_text_orientation(pil_image):
    """通过图像投影方差检测文字方向，返回需要旋转的角度（0 或 90）

    原理：正向文字的水平投影（每行深色像素统计）有明显的波峰波谷
    （文字行和空白行交替），方差大于垂直投影。侧倒文字则相反。

    - 水平投影方差 >= 垂直投影方差：文字行水平 → 返回 0（正向）
    - 垂直投影方差 > 水平投影方差：文字行垂直（侧倒）→ 返回 90（向左旋转90度）

    倾斜容差：文字倾斜 < 30° 时，水平投影方差仍然大于垂直投影，不会误旋转。
    （扫描件稍有倾斜不影响水平投影的波峰波谷特征）

    注意：投影方差法无法区分正向/倒置、向左/向右侧倒。
    倒置和向右侧倒由 _extract_pdf_title 的 fallback 逻辑覆盖：
    0° 未匹配规则时会依次尝试 90°/270°/180°。
    """
    import numpy as np
    w, h = pil_image.size
    # 取中部 40% 区域做检测（减少计算量，避免边缘噪声）
    cx0, cy0 = int(w * 0.30), int(h * 0.30)
    cx1, cy1 = int(w * 0.70), int(h * 0.70)
    sample = pil_image.crop((cx0, cy0, cx1, cy1))

    gray = np.array(sample.convert("L"))
    # 自适应阈值二值化：低于均值为文字（深色）
    threshold = gray.mean()
    binary = (gray < threshold).astype(np.float32)

    # 水平投影方差（文字行水平时方差大：文字行和空白行交替）
    h_var = float(binary.sum(axis=1).var())
    # 垂直投影方差（文字行垂直时方差大）
    v_var = float(binary.sum(axis=0).var())

    if h_var >= v_var:
        return 0  # 文字行水平：正向
    return 90  # 文字行垂直：侧倒，优先向左旋转90度


def _extract_pdf_title(pdf_path, rules=None):
    """用 OCR 提取 PDF 第一页顶部文字（支持横向/倒置扫描件）

    优化策略：
    1. 方向检测：通过 OCR 识别文字框角度，判断文字方向（不依赖宽高比）
    2. 多级裁剪：依次尝试顶部 15% / 30% / 50% / 整页
    3. 多方向：以检测到的文字方向优先，未匹配时 fallback 到其他方向
    4. 命中即停：传入 rules 时，识别到匹配规则的标题立即返回
    5. 兜底：所有方向都未匹配时，返回字符最多的识别结果
    """
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(pdf_path)
    try:
        if len(pdf) == 0:
            return ""
        page = pdf[0]
        bitmap = page.render(scale=_RENDER_SCALE)
        pil_image = bitmap.to_pil().convert("RGB")
    finally:
        pdf.close()

    engine = _get_ocr_engine()
    best_title = ""
    best_count = 0

    # 通过图像投影方差检测文字方向（不依赖宽高比，避免误判）
    # 检测到文字侧倒时，优先向左/向右旋转90度
    orientation = _detect_text_orientation(pil_image)

    # 根据检测到的文字方向决定旋转顺序（PIL rotate 为逆时针）
    if orientation == 0:
        # 文字正向：0° 优先，未匹配时向左旋转90度次之
        angles = (0, 90, 270, 180)
    elif orientation == 90:
        # 文字向左侧倒：向左旋转90度优先
        angles = (90, 0, 270, 180)
    elif orientation == 270:
        # 文字向右侧倒：向右旋转90度优先
        angles = (270, 0, 90, 180)
    else:
        # 文字倒置：180度优先
        angles = (180, 0, 90, 270)

    for angle in angles:
        if angle == 0:
            img = pil_image
        else:
            # expand=True 保证旋转后不裁切
            img = pil_image.rotate(angle, expand=True)
        w, h = img.size
        # 多级动态裁剪：先小后大，命中即停
        for ratio in _CROP_RATIOS:
            top_img = img.crop((0, 0, w, int(h * ratio)))
            result = _ocr_recognize(engine, top_img)
            title = _build_title_from_result(result)
            count = _count_chars(result)

            # 传入规则时：匹配到标题立即返回（最快路径）
            if rules is not None and _match_rule(title, rules):
                return title

            # 记录字符最多的结果作为备选
            if count > best_count:
                best_title = title
                best_count = count

            # count == 0 表示当前方向完全没文字，换方向
            if count == 0:
                break
            # 传入 rules 时：不因识别到文字就 break，继续扩大裁剪区域
            # 尝试匹配规则（标题可能在页面更靠下的位置，被表格内容遮挡）
            # 未传 rules 时：识别到足够文字就 break（只需标题候选）
            if rules is None and count >= _MIN_TITLE_CHARS:
                break

    return best_title


def _extract_drawing_title(pdf_path, rules):
    """专门识别图纸类 PDF 的标题（标题在右下角标题栏）

    竣工图纸特点：
    - 一定是横向扫描的文件（文字方向为横向）
    - 标题在右下角的标题栏中
    - 页面大部分是图形图像，顶部裁剪无法识别到标题

    策略：
    1. 检测文字方向，旋转到正向后再识别
    2. 依次裁剪右下角 40%、底部 30%、右下角 60% 区域做 OCR
    3. 匹配到规则立即返回，未匹配返回字符最多的结果
    """
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(pdf_path)
    try:
        if len(pdf) == 0:
            return ""
        page = pdf[0]
        bitmap = page.render(scale=_RENDER_SCALE)
        pil_image = bitmap.to_pil().convert("RGB")
    finally:
        pdf.close()

    engine = _get_ocr_engine()

    # 图纸一定是横向扫描，先检测方向并旋转到正向
    orientation = _detect_text_orientation(pil_image)
    if orientation != 0:
        pil_image = pil_image.rotate(orientation, expand=True)

    w, h = pil_image.size
    best_title = ""
    best_count = 0

    # 依次尝试多个右下角/底部区域
    regions = [
        ("右下角40%", w - int(w * 0.4), h - int(h * 0.4), w, h),
        ("底部30%", 0, int(h * 0.7), w, h),
        ("右下角60%", w - int(w * 0.6), h - int(h * 0.6), w, h),
    ]
    for _label, x0, y0, x1, y1 in regions:
        crop = pil_image.crop((x0, y0, x1, y1))
        result = _ocr_recognize(engine, crop)
        title = _build_title_from_result(result)
        count = _count_chars(result)

        # 匹配到规则立即返回
        if _match_rule(title, rules):
            return title

        if count > best_count:
            best_title = title
            best_count = count

    return best_title


def _normalize_text(text):
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


def _fuzzy_similarity(a, b):
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
    return _FUZZY_THRESHOLD


def _match_rule(title_text, rules):
    """根据标题文本匹配映射规则，返回映射内容（未匹配返回 None）

    匹配策略（依次尝试）：
    1. 逐行严格包含：OCR 标题可能是多行拼接（第一页顶部多行文字），
       按行从上到下匹配——页面最上方一行最可能是真正的标题，
       避免下部行里的干扰文字（如表格中其他报告的标题）抢先命中；
       同一行内多条规则命中时，取「识别内容」更长（更具体）的规则
    2. 模糊匹配：整段标题与识别内容相似度 ≥ 阈值（容错 OCR 错字），
       短关键词阈值更高（见 _fuzzy_threshold），防止常见词组误匹配
    """
    if not title_text:
        return None
    title_clean = _normalize_text(title_text)
    if not title_clean:
        return None

    # 第一轮：逐行严格包含匹配（最快、最准）
    for raw_line in title_text.splitlines():
        line_clean = _normalize_text(raw_line)
        if not line_clean:
            continue
        line_match = None
        best_key_len = -1
        for rule in rules:
            key = _normalize_text(str(rule.get("识别内容", "")))
            if key and key in line_clean and len(key) > best_key_len:
                best_key_len = len(key)
                line_match = rule
        if line_match is not None:
            return str(line_match.get("映射内容", "")).strip()

    # 第二轮：模糊匹配（容错 OCR 错字/漏字）
    best_match = None
    best_score = 0.0
    for rule in rules:
        key = _normalize_text(str(rule.get("识别内容", "")))
        if not key or len(key) < _FUZZY_MIN_LEN:
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
                score = _fuzzy_similarity(sub, key)
                if score > max_sub_score:
                    max_sub_score = score
                    if max_sub_score >= 1.0:
                        break
            score = max_sub_score
        else:
            score = _fuzzy_similarity(title_clean, key)
        if score >= thr and score > best_score:
            best_score = score
            best_match = rule
    if best_match is not None:
        return str(best_match.get("映射内容", "")).strip()
    return None


# ============================================================
# 并发处理
# ============================================================

def _process_single_pdf(pdf_path, rules, by_folder, idx, total):
    """处理单个 PDF（线程安全）

    包含：OCR 识别 → 规则匹配 → 重命名
    文件操作受 _FILE_OP_LOCK 保护，避免并发冲突。

    返回 dict:
        status: "renamed" / "pending" / "skip" / "failed"
        pdf_path: 原始路径（用于第二阶段）
        base_name: 文件名
        info: 详情信息
        title: 识别到的标题（pending 时用于日志）
        idx: 原始序号
    """
    base_name = os.path.basename(pdf_path)
    try:
        title = _extract_pdf_title(pdf_path, rules)
        mapping = _match_rule(title, rules)

        if not mapping:
            return {"status": "pending", "pdf_path": pdf_path,
                    "base_name": base_name, "title": title, "idx": idx}

        if not mapping.strip():
            return {"status": "skip", "pdf_path": pdf_path,
                    "base_name": base_name, "info": "映射内容为空", "idx": idx}

        new_name = _build_new_name(mapping, pdf_path, by_folder)
        new_path = os.path.join(os.path.dirname(pdf_path), new_name)

        with _FILE_OP_LOCK:
            new_path = _resolve_conflict(new_path)
            if os.path.normcase(new_path) == os.path.normcase(pdf_path):
                return {"status": "skip", "pdf_path": pdf_path,
                        "base_name": base_name, "info": "名称未变化", "idx": idx}
            os.rename(pdf_path, new_path)

        return {"status": "renamed", "pdf_path": pdf_path,
                "base_name": base_name, "info": "→ {}".format(new_name),
                "title": title, "idx": idx}
    except MemoryError:
        return {"status": "failed", "pdf_path": pdf_path,
                "base_name": base_name, "info": "内存不足（MemoryError）", "idx": idx}
    except Exception as e:
        return {"status": "failed", "pdf_path": pdf_path,
                "base_name": base_name, "info": str(e), "idx": idx}


# ============================================================
# 文件名工具
# ============================================================

def _sanitize_filename(name):
    """清理文件名中的非法字符"""
    invalid = '<>:"/\\|?*'
    for ch in invalid:
        name = name.replace(ch, "_")
    return name.strip().rstrip(".")


def _resolve_conflict(path):
    """目标路径已存在时，添加序号后缀避免覆盖"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists("{}_{}{}".format(base, i, ext)):
        i += 1
    return "{}_{}{}".format(base, i, ext)


def _build_new_name(mapping_content, pdf_path, by_folder):
    """根据命名规则构建新文件名"""
    mapping = _sanitize_filename(mapping_content)
    if by_folder:
        folder_name = os.path.basename(os.path.dirname(pdf_path))
        folder_name = _sanitize_filename(folder_name)
        return "{}{}.pdf".format(folder_name, mapping)
    return "{}.pdf".format(mapping)


# ============================================================
# 主入口
# ============================================================

def _fail(msg, logs=None):
    """构建失败响应"""
    result = {"success": False, "message": msg, "data": None}
    if logs:
        result["logs"] = logs
    return result


def run(params):
    """脚本主入口

    参数:
        params: dict - 用户填写的表单参数

    返回:
        dict - {success, message, data:{total,renamed,skipped,failed,details}, logs}
    """
    logs = []

    # ---------- 1. 校验命名规则（互斥） ----------
    by_folder = bool(params.get("rename_by_folder", False))
    by_mapping = bool(params.get("rename_by_mapping_only", True))
    if by_folder and by_mapping:
        return _fail("命名规则冲突：请只选择一种命名方式（两种均已勾选）", logs)
    if not by_folder and not by_mapping:
        return _fail("请选择一种命名规则（按文件夹名称 或 仅映射内容）", logs)

    # ---------- 2. 校验 PDF 文件夹 ----------
    pdf_folder = (params.get("pdf_folder") or "").strip()
    if not pdf_folder:
        return _fail("请选择 PDF 所在文件夹", logs)
    if not os.path.isdir(pdf_folder):
        return _fail("所选文件夹不存在或不是目录：{}".format(pdf_folder), logs)

    # ---------- 3. 检查依赖库 ----------
    try:
        import pypdfium2  # noqa: F401
    except ImportError:
        return _fail("未安装 PDF 渲染库 pypdfium2，请执行：pip install pypdfium2", logs)
    try:
        from rapidocr_onnxruntime import RapidOCR  # noqa: F401
    except ImportError:
        return _fail(
            "未安装 OCR 识别库 rapidocr_onnxruntime，请执行：pip install rapidocr_onnxruntime",
            logs)

    # ---------- 4. 确保规则文件存在 ----------
    rules_dir = _get_rules_dir()
    rules_file = _get_rules_file()
    logs.append("[规则] 规则文件夹：{}".format(rules_dir))
    if not os.path.exists(rules_file):
        _create_default_rules(rules_file)
        logs.append("[规则] 未发现规则文件，已生成默认规则：{}".format(RULES_FILENAME))

    # ---------- 5. 导入规则（如用户选择了文件） ----------
    import_path = (params.get("import_excel_rules") or "").strip()
    if import_path:
        if not os.path.isfile(import_path):
            return _fail("导入的规则文件不存在：{}".format(import_path), logs)
        try:
            imported_count = _import_rules(import_path, rules_file)
            logs.append("[规则] 已导入并更新规则文件：{}（来源：{}）".format(
                RULES_FILENAME, os.path.basename(import_path)))
            logs.append("[规则] 导入后验证通过，共读取到 {} 条规则".format(
                imported_count))
        except Exception as e:
            return _fail("导入规则文件失败：{}".format(str(e)), logs)

    # ---------- 6. 读取规则 ----------
    try:
        rules = _read_rules(rules_file)
    except Exception as e:
        return _fail("读取规则文件失败：{}".format(str(e)), logs)
    logs.append("[规则] 共加载 {} 条映射规则".format(len(rules)))
    if not rules:
        return _fail(
            "映射规则为空，请先导入包含「识别内容」和「映射内容」两列的 Excel 文件",
            logs)
    # 输出前 3 条规则预览，便于用户确认规则是否正确加载
    for i, rule in enumerate(rules[:3], 1):
        logs.append("[规则]  {}. {} → {}".format(
            i, rule.get("识别内容", ""), rule.get("映射内容", "")))
    if len(rules) > 3:
        logs.append("[规则]  ...（共 {} 条，仅显示前 3 条）".format(len(rules)))

    # ---------- 7. 递归扫描 PDF ----------
    pdf_files = []
    for root, _dirs, files in os.walk(pdf_folder):
        for fn in files:
            if fn.lower().endswith(".pdf"):
                pdf_files.append(os.path.join(root, fn))
    logs.append("[扫描] 共发现 {} 份 PDF 文件（含子文件夹）".format(len(pdf_files)))

    if not pdf_files:
        return {
            "success": True,
            "message": "选定文件夹下未发现任何 PDF 文件",
            "data": {"total": 0, "renamed": 0, "skipped": 0, "failed": 0,
                     "scanned_count": 0, "details": []},
            "logs": logs,
        }

    # ---------- 8. 初始化 OCR 引擎（提前加载模型） ----------
    try:
        _get_ocr_engine()
        logs.append("[OCR] 引擎已就绪")
    except Exception as e:
        return _fail("OCR 引擎初始化失败：{}".format(str(e)), logs)

    # ---------- 9. 第一阶段：并行处理（顶部标题识别） ----------
    details = []
    renamed = 0
    skipped = 0
    failed = 0
    total = len(pdf_files)
    pending_pdfs = []  # 第一阶段未匹配的 PDF，待第二阶段图纸识别

    if total <= 1:
        # 单个文件无需并行，直接串行处理
        for idx, pdf_path in enumerate(pdf_files, 1):
            result = _process_single_pdf(pdf_path, rules, by_folder, idx, total)
            status = result["status"]
            base_name = result["base_name"]
            if status == "renamed":
                renamed += 1
                details.append(["renamed", base_name, result["info"]])
                logs.append("[{}/{}] 成功 {} {}".format(idx, total, base_name, result["info"]))
            elif status == "pending":
                pending_pdfs.append((result["pdf_path"], idx))
                preview = result.get("title", "").replace("\n", " ")[:30] if result.get("title") else "(未识别到文字)"
                logs.append("[{}/{}] 待定 {} — 顶部未匹配，标题：「{}」".format(idx, total, base_name, preview))
            elif status == "skip":
                skipped += 1
                details.append(["skip", base_name, result.get("info", "")])
                logs.append("[{}/{}] 跳过 {} — {}".format(idx, total, base_name, result.get("info", "")))
            else:
                failed += 1
                details.append(["failed", base_name, result.get("info", "")])
                logs.append("[{}/{}] 失败 {} — {}".format(idx, total, base_name, result.get("info", "")))
    else:
        # 多文件并行处理
        workers = min(_MAX_WORKERS, total)
        logs.append("[处理] 使用 {} 个线程并行识别（总超时 {}秒）".format(workers, _TOTAL_TIMEOUT))
        t_start = time.time()

        executor = concurrent.futures.ThreadPoolExecutor(max_workers=workers)
        future_to_pdf = {}
        for idx, pdf_path in enumerate(pdf_files, 1):
            future = executor.submit(_process_single_pdf, pdf_path, rules, by_folder, idx, total)
            future_to_pdf[future] = (pdf_path, idx)

        done_count = 0
        try:
            for future in concurrent.futures.as_completed(future_to_pdf, timeout=_TOTAL_TIMEOUT):
                pdf_path, idx = future_to_pdf[future]
                done_count += 1
                try:
                    result = future.result(timeout=5)
                except Exception as e:
                    result = {"status": "failed", "pdf_path": pdf_path,
                              "base_name": os.path.basename(pdf_path),
                              "info": str(e), "idx": idx}

                status = result["status"]
                base_name = result["base_name"]
                if status == "renamed":
                    renamed += 1
                    details.append(["renamed", base_name, result["info"]])
                    logs.append("[{}/{}] 成功 {} {}".format(idx, total, base_name, result["info"]))
                elif status == "pending":
                    pending_pdfs.append((result["pdf_path"], idx))
                    preview = result.get("title", "").replace("\n", " ")[:30] if result.get("title") else "(未识别到文字)"
                    logs.append("[{}/{}] 待定 {} — 顶部未匹配，标题：「{}」".format(idx, total, base_name, preview))
                elif status == "skip":
                    skipped += 1
                    details.append(["skip", base_name, result.get("info", "")])
                    logs.append("[{}/{}] 跳过 {} — {}".format(idx, total, base_name, result.get("info", "")))
                else:
                    failed += 1
                    details.append(["failed", base_name, result.get("info", "")])
                    logs.append("[{}/{}] 失败 {} — {}".format(idx, total, base_name, result.get("info", "")))
        except concurrent.futures.TimeoutError:
            elapsed = int(time.time() - t_start)
            logs.append("[警告] 总处理超时（{}秒），已完成 {}/{}，未完成的标记为失败".format(
                elapsed, done_count, total))

        # 处理未完成的任务（超时或被取消）
        for future, (pdf_path, idx) in future_to_pdf.items():
            if not future.done():
                base_name = os.path.basename(pdf_path)
                future.cancel()
                failed += 1
                details.append(["failed", base_name, "处理超时未完成"])
                logs.append("[{}/{}] 失败 {} — 处理超时未完成".format(idx, total, base_name))

        executor.shutdown(wait=False)
        elapsed = int(time.time() - t_start)
        logs.append("[处理] 第一阶段完成，耗时 {}秒".format(elapsed))

    # ---------- 9b. 第二阶段：对未匹配的 PDF 进行图纸标题识别 ----------
    if pending_pdfs:
        logs.append("[图纸识别] 对 {} 份未匹配的 PDF 进行图纸标题识别（右下角）".format(
            len(pending_pdfs)))
        for pdf_path, idx in pending_pdfs:
            base_name = os.path.basename(pdf_path)
            try:
                title = _extract_drawing_title(pdf_path, rules)
                mapping = _match_rule(title, rules)

                if not mapping:
                    skipped += 1
                    preview = title.replace("\n", " ")[:30] if title else "(未识别到文字)"
                    details.append(["skip", base_name, "未匹配到映射规则，保持原名"])
                    logs.append("[{}/{}] 跳过 {} — 图纸识别仍未匹配，标题：「{}」".format(
                        idx, total, base_name, preview))
                    continue

                if not mapping.strip():
                    skipped += 1
                    details.append(["skip", base_name, "映射内容为空，保持原名"])
                    logs.append("[{}/{}] 跳过 {} — 映射内容为空".format(
                        idx, total, base_name))
                    continue

                new_name = _build_new_name(mapping, pdf_path, by_folder)
                new_path = os.path.join(os.path.dirname(pdf_path), new_name)
                new_path = _resolve_conflict(new_path)

                if os.path.normcase(new_path) == os.path.normcase(pdf_path):
                    skipped += 1
                    details.append(["skip", base_name, "名称未变化，保持原名"])
                    logs.append("[{}/{}] 跳过 {} — 名称未变化".format(
                        idx, total, base_name))
                    continue

                os.rename(pdf_path, new_path)
                renamed += 1
                details.append(["renamed", base_name, "→ {}（图纸识别）".format(new_name)])
                logs.append("[{}/{}] 成功（图纸识别）{} → {}".format(
                    idx, total, base_name, new_name))
            except Exception as e:
                failed += 1
                details.append(["failed", base_name, str(e)])
                logs.append("[{}/{}] 失败 {} — {}".format(
                    idx, total, base_name, str(e)))

    # ---------- 10. 导出规则（如勾选） ----------
    if bool(params.get("export_rules", False)):
        try:
            export_path = _export_rules_to_desktop(rules_file)
            logs.append("[规则] 已导出到桌面：{}".format(export_path))
        except Exception as e:
            logs.append("[规则] 导出失败：{}".format(str(e)))

    # ---------- 11. 打开规则文件夹（如勾选） ----------
    if bool(params.get("open_rules_folder", False)):
        try:
            os.startfile(rules_dir)  # Windows 专用
        except Exception as e:
            logs.append("[规则] 打开规则文件夹失败：{}".format(str(e)))

    # ---------- 12. 汇总 ----------
    summary = "共扫描 {total} 份 PDF：重命名 {renamed}，跳过 {skipped}，失败 {failed}".format(
        total=total, renamed=renamed, skipped=skipped, failed=failed)
    logs.append("[完成] {}".format(summary))

    return {
        "success": failed == 0,
        "message": summary,
        "data": {
            "total": total,
            "renamed": renamed,
            "skipped": skipped,
            "failed": failed,
            "scanned_count": total,
            "details": details,
        },
        "logs": logs,
    }
