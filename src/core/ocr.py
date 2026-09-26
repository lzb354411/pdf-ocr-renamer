# -*- coding: utf-8 -*-
"""OCR 引擎与 PDF 标题提取（自 v2.0 单文件脚本迁移，逻辑保持一致）

职责：
- OCR 引擎单例管理（RapidOCR，线程安全）
- PDF 第一页渲染为图片（pypdfium2）
- 文字方向检测（投影方差法，支持横向/倒置扫描件）
- 标题提取：普通文档（顶部多级裁剪）与图纸（右下角标题栏）两种策略
"""

import threading

# ============ 可调参数（与原脚本一致） ============
# 标题提取时取第一页前 N 行作为候选区域
TITLE_LINES = 12
# OCR 行分组容差（像素）：y 坐标差距小于此值视为同一行
ROW_TOLERANCE = 15
# PDF 渲染缩放比例（越大越清晰但越慢）
RENDER_SCALE = 2.0
# 动态裁剪比例：依次尝试顶部 15% / 30% / 50% / 整页
CROP_RATIOS = (0.15, 0.30, 0.50, 1.0)
# 识别文字少于此字符数时，认为该方向未识别到有效内容
MIN_TITLE_CHARS = 4

# OCR 引擎单例（首次调用时加载模型，后续复用）
_OCR_ENGINE = None
# 线程安全锁：初始化锁 + 推理调用锁（RapidOCR/onnxruntime 非线程安全）
_OCR_INIT_LOCK = threading.Lock()
_OCR_CALL_LOCK = threading.Lock()


def get_ocr_engine():
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
        if cur_y is None or abs(y - cur_y) <= ROW_TOLERANCE:
            cur_row.append(text)
            if cur_y is None:
                cur_y = y
        else:
            rows.append("".join(cur_row))
            cur_row = [text]
            cur_y = y
    if cur_row:
        rows.append("".join(cur_row))
    return "\n".join(rows[:TITLE_LINES])


def _detect_text_orientation(pil_image):
    """通过图像投影方差检测文字方向，返回需要旋转的角度（0 或 90）

    原理：正向文字的水平投影（每行深色像素统计）有明显的波峰波谷
    （文字行和空白行交替），方差大于垂直投影。侧倒文字则相反。

    - 水平投影方差 >= 垂直投影方差：文字行水平 → 返回 0（正向）
    - 垂直投影方差 > 水平投影方差：文字行垂直（侧倒）→ 返回 90（向左旋转90度）

    倾斜容差：文字倾斜 < 30° 时，水平投影方差仍然大于垂直投影，不会误旋转。
    （扫描件稍有倾斜不影响水平投影的波峰波谷特征）

    注意：投影方差法无法区分正向/倒置、向左/向右侧倒。
    倒置和向右侧倒由 extract_pdf_title 的 fallback 逻辑覆盖：
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


def _render_first_page(pdf_path):
    """将 PDF 第一页渲染为 PIL Image（失败返回 None）"""
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(pdf_path)
    try:
        if len(pdf) == 0:
            return None
        page = pdf[0]
        bitmap = page.render(scale=RENDER_SCALE)
        return bitmap.to_pil().convert("RGB")
    finally:
        pdf.close()


def extract_pdf_title(pdf_path, rules=None):
    """用 OCR 提取 PDF 第一页顶部文字（支持横向/倒置扫描件）

    优化策略：
    1. 方向检测：通过图像投影方差判断文字方向（不依赖宽高比）
    2. 多级裁剪：依次尝试顶部 15% / 30% / 50% / 整页
    3. 多方向：以检测到的文字方向优先，未匹配时 fallback 到其他方向
    4. 命中即停：传入 rules 时，识别到匹配规则的标题立即返回
    5. 兜底：所有方向都未匹配时，返回字符最多的识别结果
    """
    from .rules import match_rule

    pil_image = _render_first_page(pdf_path)
    if pil_image is None:
        return ""

    engine = get_ocr_engine()
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
        for ratio in CROP_RATIOS:
            top_img = img.crop((0, 0, w, int(h * ratio)))
            result = _ocr_recognize(engine, top_img)
            title = _build_title_from_result(result)
            count = _count_chars(result)

            # 传入规则时：匹配到标题立即返回（最快路径）
            if rules is not None and match_rule(title, rules):
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
            if rules is None and count >= MIN_TITLE_CHARS:
                break

    return best_title


def extract_drawing_title(pdf_path, rules):
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
    from .rules import match_rule

    pil_image = _render_first_page(pdf_path)
    if pil_image is None:
        return ""

    engine = get_ocr_engine()

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
        if match_rule(title, rules):
            return title

        if count > best_count:
            best_title = title
            best_count = count

    return best_title
