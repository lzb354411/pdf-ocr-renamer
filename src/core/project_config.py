# -*- coding: utf-8 -*-
"""工程名称解析配置（地域 / 业务可配置化）

背景（P0-1 问题）：
    原 splitter.py 把「某省」「某市」「某县」「业扩配套工程」以及 48 个年份
    黑名单硬编码在源码里，导致：
    1. 换省份 / 换业务类型（如光伏并网、充电桩）后拆分功能整体失效；
    2. 失效方式是静默返回空名称，再流入「未匹配项目」文件夹，用户不易察觉。

本模块把上述知识全部外置为配置，来源优先级（高 → 低）：
    1. 用户配置：%APPDATA%\\PdfOcrRenamer\\project_config.json
    2. 数据推导：用户提供的项目进度表 Excel（自动学习地域词与业务后缀）
    3. 内置默认：DEFAULT_CONFIG（通用电力工程语法，不含任何具体区县）

设计原则：
    - 代码只实现「解析语法」，不承载「业务词典」；
    - 所有具体地名 / 业务词都可由用户覆盖；
    - 推导失败时退化为通用语法，绝不硬失败。

名称语法（由 2026年业扩项目进度表.xlsx 560 条真实数据归纳）：

    形态 A（373 条，中压业扩 / 单体变压器）：
        某省某市某县 3225000000000001 某县某某建设发展有限公司 10kV业扩配套工程
        = [行政区划前缀] [长数字编号] [用电主体] [电压等级业务后缀]

    形态 B（12 条，资本打包）：
        某省某市某县2026年第一批包十0.4kV业扩配套
        = [行政区划前缀] [年份] 第N批 包X [电压等级业务后缀]

    形态 C（139 条，低压成本单体）：
        2026年某县供电公司某某北北变台区低压维修项目3225000000000002
        2026年某市某县丙镇某辛庄村某合村变低压台区维修项目3226000000000003
        = [年份] [供电单位 | 行政区划+乡镇] [台区/线路描述] [业务后缀] [编号]

    形态 D（36 条，低压维护/台区维修打包）：
        2026年某市某县低压维护项目包1
        2026年某市某县低压台区维修项目包31
        = [年份] [行政区划] [业务后缀] 包N
"""

import json
import os
import re

# ============================================================
# 配置文件名与版本
# ============================================================

CONFIG_FILENAME = "project_config.json"
CONFIG_VERSION = 1


# ============================================================
# 内置默认配置（通用电力工程语法，不含具体区县）
# ============================================================

DEFAULT_CONFIG = {
    "version": CONFIG_VERSION,
    "_comment": (
        "工程名称解析配置。region_words=地域词（用于定位名称起点/校验），"
        "business_suffixes=业务后缀（用于定位名称终点），"
        "org_markers=单位主体标记，noise_words=杂质字段，"
        "ocr_fixes=OCR 常见错字纠正表。可按需增删。"
    ),

    # ---------- 地域词 ----------
    # 名称起点锚点：名称通常以「省/市/县/区」等区划词开头或包含区划词
    # （如「某省」「某市」「某县」）。留空则退化为「以长数字编号或年份为起点」。
    "region_words": [],

    # 区划层级后缀字符（判断一段文字是否像行政区划名）
    "region_suffix_chars": "省市县区镇乡村",

    # ---------- 地域词容错（阶段二 R-2）----------
    # 背景：真实扫描件中「某县」被 OCR 识别为 甲县/乙县/丙县/丁县/壬路 等
    # （见阶段文档 §3.2 证据 2），基于 text.find(地域词) 的精确查找在 OCR
    # 场景下不可靠，导致地域词起点定位失败、工程名称提取不出。
    #
    # region_fuzzy_tolerance：允许地域词中不匹配的字符数上限。
    #   0 = 关闭容错（退化为精确查找，与改动前行为一致）；
    #   1 = 允许 1 字不同（推荐，可覆盖「某县→甲县」等单字错字）；
    #   也可为 "auto"：按 ⌊len/3⌋ 自适应（长词放宽、短词严格）。
    # 注意：容错与误匹配是一对矛盾，原则是「宁可漏也不能错」。
    #   - 只允许「同长度替换」，不允许增删字符（避免把长文本片段误当地域词）；
    #   - 容忍度受 region_fuzzy_max_tolerance 上限约束；
    #   - 容错命中时返回的起点必须仍落在合理位置。
    "region_fuzzy_tolerance": 1,
    # 容错度硬上限（即使配置写成 auto，也不超过此值）
    "region_fuzzy_max_tolerance": 1,
    # 容错匹配的最短地域词长度（短于此时不做容错，避免 2 字词误命中）
    "region_fuzzy_min_len": 2,
    # 是否要求容错命中的词仍以区划后缀字结尾（如 县/市/区）
    # 开启后「甲县」「乙县」可命中「某县」，而「沙漠」不会。
    "region_fuzzy_require_suffix_char": True,

    # ---------- 业务后缀（名称终点锚点）----------
    # 由真实数据归纳：名称几乎一定以这些词结尾。
    # 顺序无关，匹配时取最长匹配。
    "business_suffixes": [
        "业扩配套工程",
        "低压台区维修项目",
        "低压维修项目",
        "低压维护项目",
        "台区维修项目",
        "线路维修项目",
        "维修项目",
        "配套工程",
        "业扩工程",
        "改造项目",
        "新建工程",
        "项目",
    ],

    # 打包类名称的「包N」形态（形态 B / D）：...项目包1 / ...第一包
    "package_patterns": [
        r"包\s*\d+",
        r"第[一二三四五六七八九十百]+批",
    ],

    # ---------- 单位主体标记 ----------
    # 这些词出现时，说明当前片段是「单位」而非「工程名称」
    "org_markers": [
        "供电公司", "供电所", "电力公司", "电力局",
    ],

    # ---------- 杂质字段（名称中不应出现）----------
    "noise_words": [
        "建设单位", "施工单位", "监理单位", "设计单位",
        "项目负责人", "施工负责人", "填报人",
        "计划开工", "计划竣工", "开工时间", "竣工时间",
        "工程地点", "工程编号", "线路名称", "工程量", "项目名称",
    ],

    # ---------- 有效名称判定 ----------
    # 名称必须含其中至少一个词（留空则不校验）
    "valid_keywords": [],
    # 名称最小长度
    "min_name_len": 8,
    # 名称最大长度（超过视为把多个字段卷进来了）
    "max_name_len": 60,
    # 连续数字超过此长度视为编号，不属于名称
    "max_digit_run": 4,
    # 名称允许的字符集（正则字符类内容，配置驱动）。
    # 「#」是台区/配变编号的常用写法（某堤口2#配变），必须保留；
    # 另容错 OCR 常见的「*」「※」（真实样本把「2#」识别成「2*」）。
    "name_charset": r"\u4e00-\u9fa5A-Za-z0-9#*※（）()\-\.·",
    # OCR 编号符号的等价替换（应用于名称提取的归一化阶段）
    # 真实样本：图纸标题里「某堤口2#配变」被识别为「某堤口2*配变」
    "symbol_fixes": {
        "*": "#",
        "※": "#",
    },

    # ---------- OCR 错字纠正 ----------
    # 通用性纠正（业务词被识别错）；地域性纠正由 Excel 推导或用户补充
    #
    # 语义：键 = 错误写法，值 = 正确写法（错误 → 正确）。
    # 该表同时被 name_parser（工程名称提取）与 rules.match_rule（标题匹配）
    # 使用，两处共用 apply_ocr_fixes()，避免逻辑分叉。
    #
    # 注意（阶段一 1.3）：只允许加入「具体、无歧义」的词对。
    # 切勿加入「开工→竣工」这类宽泛替换，会破坏「开工报告」的正常识别。
    # 长锚点（如「峻工报告」）优先于短锚点（「峻工」）生效——见
    # apply_ocr_fixes()，长键先替换，避免长错词被短键切碎后残留。
    "ocr_fixes": {
        "业扩配食工程": "业扩配套工程",
        "业扩配门工程": "业扩配套工程",
        "业扩配套工稈": "业扩配套工程",
        "低压台区维休项目": "低压台区维修项目",
        "低压维休项目": "低压维修项目",
        # 「竣工」在真实扫描件中高频被识别为「峻工」（见阶段文档证据 1）
        "峻工报告": "竣工报告",
        "峻工": "竣工",
        # 「台区」在真实竣工图纸标题栏中 7/7 被识别为「合区」（证据 3）
        "合区": "台区",
    },

    # ---------- 地域词内部的 OCR 错字纠正 ----------
    # 用于修正「(子市)」被识别成「(某市)」这类区划名错字。
    # 由 infer_region_ocr_fixes 从 Excel 自动学习，也可手工补充。
    "region_ocr_fixes": {},

    # ---------- 文件夹匹配 ----------
    "folder_score_threshold": 0.55,

    # ---------- 文件夹匹配：特征词主键（阶段二 R-3）----------
    # 背景（阶段文档 §3.2 证据 3）：图纸标题栏与进度表名称差异极大
    #   「某县」vs「丁县」、「台区」vs「合区」，图纸末尾还多出设计单位，
    #   整串相似度匹配不稳定。改用「台区/线路特征词」（如 某堤口2配变）
    #   作匹配主键更可靠。
    #
    # feature_strip_patterns：匹配前先剥离的「公因子」正则（配置驱动，
    #   不含任何具体公司名/地名）。剥离后剩下的即特征词。
    "feature_strip_patterns": [
        r"20\d{2}年?",                    # 年份
        r"[省市区县镇乡村]",               # 区划后缀字（保留地名主体另处理）
        r"第[一二三四五六七八九十百]+批",   # 批次
        r"低压|高压|中压",                 # 电压等级业务前缀
        r"(?:台区|线路)?(?:维修|维护|改造|新建|配套)?(?:项目|工程)",  # 业务后缀
        r"包\d+",                          # 项目包编号
        r"\d{10,}",                        # 长编号
    ],
    # 设计单位噪声模式（图纸标题栏末尾常带设计/勘察/施工单位名）
    # 配置驱动，不硬编码具体公司名。
    "design_unit_pattern": (
        r"[\u4e00-\u9fa5]{2,20}"
        r"(?:电力勘察设计|勘察设计|电力设计|设计|勘测)"
        r"(?:院|有限公司|有限责任公司|公司|分公司)?"
    ),
    # 其它需从标题中剥离的噪声（施工单位/监理单位等整段）
    "org_noise_pattern": (
        r"[\u4e00-\u9fa5]{2,20}"
        r"(?:工程安装|安装工程|建设|建工|施工|监理)"
        r"(?:有限公司|有限责任公司|公司|分公司)"
    ),
    # 特征词最小长度（短于此不参与主键匹配，避免噪声词当主键）
    "feature_min_len": 3,
    # 特征词匹配的最低相似度（特征词比对低于此值不算命中）
    "feature_score_threshold": 0.62,
    # 项目编号参与匹配（阶段二 R-3 补充）
    # 背景：同一台区在「第七批7项」中有两条记录（如某堡变 3225000000000004
    # 与 3226000000000005），其特征词完全相同，仅末尾编号不同。此时**编号**
    # 是唯一可区分的依据，且图纸标题栏也印有该编号。
    # 开启后：名称与文件夹中若都含长编号，编号相同则显著加分、不同则降分。
    "code_match_enabled": True,
    # 编号完全相同时的加分权重（叠加到 feature/whole 得分上，上限 1.0）
    "code_match_bonus": 0.35,
    # 编号不同但双方都有编号时的惩罚系数（乘到得分上）
    "code_mismatch_penalty": 0.75,
    # 编号比对容错位数（图纸标题栏编号常有 1~2 位数字 OCR 错误，
    # 实测 3225000000000004 被识别为 3225000000000006）
    "code_fuzzy_tolerance": 2,

    # ---------- 标签正则（工程名称/编号标签的 OCR 变体）----------
    "name_label_pattern": r"工程(?:名称|名栋|名[厦回因]?|标[名]?)",
    "code_label_pattern": r"工程编号[:：号]?\d{9,}|\d{11,}\s*$|工程编号",
}


# ============================================================
# 配置加载 / 保存
# ============================================================

def _merge(base, override):
    """把 override 合并进 base 的副本（浅合并 dict，列表整体替换）

    列表整体替换而非追加：用户配置应当能"完全接管"某类词表，
    否则无法删除内置项。
    """
    if not isinstance(override, dict):
        return base
    out = dict(base)
    for k, v in override.items():
        if k.startswith("_"):
            continue
        if v is None:
            continue
        out[k] = v
    return out


def default_config():
    """返回内置默认配置的深拷贝友好副本"""
    cfg = dict(DEFAULT_CONFIG)
    # 列表 / dict 需要复制，避免调用方修改污染内置常量
    for k, v in cfg.items():
        if isinstance(v, (list, dict)):
            cfg[k] = v.copy() if isinstance(v, dict) else list(v)
    return cfg


def config_path():
    """用户配置文件路径（%APPDATA%\\PdfOcrRenamer\\project_config.json）"""
    from . import paths
    return os.path.join(paths.get_app_dir(), CONFIG_FILENAME)


def load_config(path=None):
    """加载配置：内置默认 ← 用户 JSON 覆盖

    参数 path 为空时使用 config_path()。文件不存在或损坏时返回默认配置
    （并附带 load_error 字段说明原因，便于界面上报，但不抛异常）。
    """
    cfg = default_config()
    p = path or config_path()
    if not p or not os.path.exists(p):
        return cfg
    try:
        with open(p, "r", encoding="utf-8") as f:
            user = json.load(f)
        cfg = _merge(cfg, user)
    except Exception as e:
        cfg["load_error"] = str(e)
    return cfg


def save_config(cfg, path=None):
    """保存配置到用户目录（只写入与默认值不同的键，保持文件精简可读）"""
    p = path or config_path()
    out = {k: v for k, v in cfg.items()
           if not k.startswith("_") and k in DEFAULT_CONFIG
           and DEFAULT_CONFIG.get(k) != v}
    out["_comment"] = DEFAULT_CONFIG["_comment"]
    out["version"] = CONFIG_VERSION
    with open(p, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    return p


def ensure_config_file(path=None):
    """确保用户配置文件存在（首次运行时生成一份可编辑的模板）

    返回 (路径, 是否新建)。
    """
    p = path or config_path()
    if os.path.exists(p):
        return p, False
    save_config(load_config(p), p)
    return p, True


# ============================================================
# OCR 错字纠正（rules 与 name_parser 共用，避免逻辑分叉）
# ============================================================

def get_ocr_fixes(cfg=None):
    """取出生效的 OCR 纠正表（dict：错误写法 → 正确写法）

    合并两处配置（阶段二扩展）：
        1. `ocr_fixes`        通用业务词纠错（内置 + 用户覆盖）
        2. `region_ocr_fixes` 地域词内部纠错（由 Excel 自动学习或手工补充）

    两者语义一致（错误 → 正确）；冲突时以 `ocr_fixes` 为准（更通用、更靠前）。
    配置缺失、为空或类型异常时返回 {}，调用方据此退化为「不做纠正」。
    """
    if cfg is None:
        cfg = load_config()
    cfg = cfg or {}
    out = {}
    # 先合并地域纠错，再合并通用纠错（通用优先覆盖）
    for key in ("region_ocr_fixes", "ocr_fixes"):
        table = cfg.get(key)
        if isinstance(table, dict):
            for k, v in table.items():
                if isinstance(k, str) and k and isinstance(v, str):
                    out[k] = v
    return out


def apply_ocr_fixes(text, cfg=None):
    """按配置的 ocr_fixes 纠正 OCR 错字，返回纠正后的文本

    设计要点：
        1. **长键优先**：先替换更长的错误词（「峻工报告」先于「峻工」），
           否则长错词会被短键切碎，产生「竣工报告」之外的残留写法；
           同长度按键名排序，保证结果确定（不依赖 dict 插入顺序）。
        2. **向后兼容**：纠正表为空/缺失/类型异常时原样返回输入，
           行为与「未接入纠正表」完全一致。
        3. 纯函数、不修改入参，可被 rules.match_rule 与 name_parser 复用。
    """
    if not text:
        return text
    fixes = get_ocr_fixes(cfg)
    if not fixes:
        return text
    # 长键优先 + 同长度按字典序，保证替换顺序确定
    for bad in sorted(fixes, key=lambda k: (-len(k), k)):
        good = fixes[bad]
        if bad in text:
            text = text.replace(bad, good)
    return text


# ============================================================
# 从项目进度表 Excel 推导配置
# ============================================================

def _looks_like_region(word):
    """判断一个词是否像行政区划名（如「某省」「某市」「某县」「丙镇」）"""
    if not word or len(word) < 2 or len(word) > 12:
        return False
    if not re.fullmatch(r"[\u4e00-\u9fa5]+", word):
        return False
    return word[-1] in "省市县区镇乡村"


def infer_from_names(names):
    """从一批真实「项目名称」中推导配置片段

    返回 dict（可直接 _merge 进配置）：
        region_words:       高频行政区划前缀（如 某省/某市/某县/某市）
        business_suffixes:  观察到的名称结尾形态
        valid_keywords:     高频业务关键词（如 业扩/维修/配套工程）

    推导策略：
        - 地域词：统计每条名称的前 12 字中出现的区划词，按频次取高频者；
        - 业务后缀：统计名称结尾形态，取出现 ≥2 次者；
        - 有效关键词：从后缀中切出 2-4 字核心词。
    """
    names = [str(n).strip() for n in names if n and str(n).strip()]
    if not names:
        return {}

    # ---------- 1. 地域词 ----------
    # 从每条名称开头扫描连续区划词：某省 / 某市 / 某县 / 某市
    # 关键：只从「名称最开头」连续切分，避免把「...镇某集西2#变」里的
    # 「某集西」之类描述词误当作区划词。
    region_counter = {}

    def _split_region_chain(head):
        """从头部连续切出区划链，返回 [词, ...]；遇到非区划词即停止"""
        chain = []
        i = 0
        # 允许跳过开头的 4 位年份
        m = re.match(r"20\d{2}年?", head)
        if m:
            i = m.end()
        while i < len(head):
            m = re.match(r"[\u4e00-\u9fa5]{1,5}?[省市县区镇乡村]", head[i:])
            if not m:
                break
            word = m.group()
            # 「年某市某县」这类残留要剔除：词首不得含「年」
            if word.startswith("年"):
                break
            chain.append(word)
            i += m.end()
            if len(chain) >= 4:
                break
        return chain

    for name in names:
        head = name[:18]
        chain = _split_region_chain(head)
        for word in chain:
            region_counter[word] = region_counter.get(word, 0) + 1
        # 同时记录整条链（如「某省某市某县」），长链是极强的起点锚点
        if len(chain) >= 2:
            joined = "".join(chain)
            region_counter[joined] = region_counter.get(joined, 0) + 1

    # 取出现频次 ≥ 总样本 5% 且 ≥2 次的，按长度降序（长词更具体，优先匹配）
    threshold = max(2, int(len(names) * 0.05))
    region_words = sorted(
        [w for w, c in region_counter.items() if c >= threshold],
        key=lambda w: (-len(w), w))

    # ---------- 2. 业务后缀 ----------
    # 观察名称结尾形态。难点：真实名称尾部常带「变/台区/配1主变」等业务对象词
    # （如「...某合村变低压台区维修项目」），若直接截尾会把对象词一起学成后缀。
    # 策略：先用「业务词库」找到结尾的业务词组，再向左吸收紧邻的修饰，
    #       但只在「业务词开头」处切断——即后缀必须以业务词打头。
    BIZ_HEADS = ("业扩", "配套", "低压", "高压", "台区", "线路", "维修",
                 "维护", "改造", "新建", "配电", "电缆", "箱变", "工程", "项目")

    # 电压等级前缀（10kV / 0.4kV / 0.4KV 等）常出现在业务词之前，
    # 需一并吸收进后缀，否则「0.4kV业扩配套工程」会被截成「业扩配套工程」，
    # 导致含该后缀的名称无法精确匹配。统一规范成 <N>kV 形式。
    _VOLT_RE = re.compile(r"^(?:\d+(?:\.\d+)?)\s*[kK][vV]")

    def _canon_suffix(seg):
        """把后缀中的电压等级写法规范化（0.4KV → 0.4kV）"""
        m = _VOLT_RE.match(seg)
        if not m:
            return seg
        volt = m.group()
        canon = re.sub(r"\s*[kK][vV]$", "kV", volt)
        return canon + seg[m.end():]

    suffix_counter = {}
    for name in names:
        tail = re.sub(r"\d+$", "", name)          # 去掉结尾编号
        tail = re.sub(r"包\d+$", "包N", tail)      # ...项目包12 → ...项目包N
        # 从左到右找最后一个「业务词开头」的位置，取其到结尾作为后缀
        best_pos = None
        for i in range(len(tail)):
            seg = tail[i:]
            # 允许以电压等级开头（0.4kV业扩配套工程）
            mv = _VOLT_RE.match(seg)
            body = seg[mv.end():] if mv else seg
            if len(body) < 3 or len(body) > 14:
                continue
            if not body.endswith(("工程", "项目", "包N")):
                continue
            if not body.startswith(BIZ_HEADS):
                continue
            # 后缀内部不应含长数字编号
            if re.search(r"\d{5,}", body):
                continue
            best_pos = _canon_suffix(seg)          # 取最长（最靠左）的合法后缀
            break
        if best_pos:
            suffix_counter[best_pos] = suffix_counter.get(best_pos, 0) + 1

    # 频次 ≥2 的保留；按长度降序（长词优先匹配）
    suffixes = sorted(
        [s for s, c in suffix_counter.items() if c >= 2],
        key=lambda s: (-len(s), s))

    # 保证内置通用后缀一定在列（即使本批数据未出现，如「改造项目」），
    # 否则换业务类型后配置反而变窄。
    merged = list(dict.fromkeys(suffixes + list(DEFAULT_CONFIG["business_suffixes"])))
    suffixes = sorted(merged, key=lambda s: (-len(s), s))

    # ---------- 3. 有效关键词 ----------
    kw_counter = {}
    for name in names:
        for kw in ("业扩", "配套工程", "维修", "维护", "改造", "新建", "台区"):
            if kw in name:
                kw_counter[kw] = kw_counter.get(kw, 0) + 1
    keywords = sorted([k for k, c in kw_counter.items() if c >= max(2, len(names) * 0.03)],
                      key=lambda k: (-kw_counter[k], k))

    inferred = {}
    if region_words:
        inferred["region_words"] = region_words
    if suffixes:
        # 与内置后缀合并去重，保持长词优先
        merged = list(dict.fromkeys(suffixes + DEFAULT_CONFIG["business_suffixes"]))
        inferred["business_suffixes"] = sorted(merged, key=lambda s: (-len(s), s))
    if keywords:
        inferred["valid_keywords"] = keywords
    return inferred


def infer_from_excel(xlsx_path, name_column="项目名称", max_rows=2000):
    """从项目进度表 Excel 推导配置片段

    自动定位包含「项目名称」表头的 sheet 与列（支持表头不在首行）。
    读取失败时返回 {}（调用方退化为默认配置）。
    """
    try:
        from openpyxl import load_workbook
    except ImportError:
        return {}
    try:
        wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    except Exception:
        return {}

    names = []
    try:
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True, max_row=max_rows))
            if not rows:
                continue
            # 在前 10 行找「项目名称」表头
            col = -1
            header_idx = -1
            for i, row in enumerate(rows[:10]):
                if row is None:
                    continue
                for j, cell in enumerate(row):
                    if cell is not None and str(cell).strip() == name_column:
                        col = j
                        header_idx = i
                        break
                if col >= 0:
                    break
            if col < 0:
                continue
            for row in rows[header_idx + 1:]:
                if row is None or col >= len(row):
                    continue
                v = row[col]
                if v is not None and str(v).strip():
                    names.append(str(v).strip())
    finally:
        try:
            wb.close()
        except Exception:
            pass

    return infer_from_names(names)


def auto_configure(excel_path=None, save=True):
    """一站式自动配置：读 Excel 推导 → 合并进配置 → 可选持久化

    返回 (config_dict, info_dict)：
        config_dict: 最终生效的配置
        info_dict:   {"source": ..., "inferred": {...}, "saved": 路径或None}
    """
    cfg = load_config()
    info = {"source": None, "inferred": {}, "saved": None}

    inferred = {}
    if excel_path and os.path.isfile(excel_path):
        inferred = infer_from_excel(excel_path)
        if inferred:
            info["source"] = excel_path

    if inferred:
        cfg = _merge(cfg, inferred)
        info["inferred"] = inferred
        if save:
            try:
                info["saved"] = save_config(cfg)
            except Exception as e:
                cfg["save_error"] = str(e)
    return cfg, info
