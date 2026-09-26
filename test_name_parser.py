# -*- coding: utf-8 -*-
"""工程名称解析回归测试（配置驱动，替代 splitter 硬编码）

覆盖：
1. 从真实项目进度表 Excel 推导配置（地域词 / 业务后缀 / 关键词）
2. 560 条真实「项目名称」的解析正确率
3. 四种名称形态（中压业扩 / 资本打包 / 低压维修 / 项目包）
4. OCR 噪声鲁棒性（换行、错字、标签干扰）
5. 跨地域 / 跨业务泛化（换区县仍可用；换业务不误提取）
6. 配置的读写与持久化

数据文件缺失时自动跳过依赖真实数据的用例（打印 SKIP），
保证在无样本环境下也能跑通结构性测试。
"""
import os
import re
import sys
import json
import shutil
import tempfile
import unicodedata

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from core import project_config as PC
from core import name_parser as NP

ROOT = os.path.dirname(os.path.abspath(__file__))
SAMPLE_XLSX = os.path.join(ROOT, "2026年业扩项目进度表.xlsx")

failures = []
skips = []


def check(name, cond, detail=""):
    if cond:
        print("  [PASS] {}".format(name))
    else:
        failures.append(name)
        print("  [FAIL] {}  {}".format(name, detail))


def skip(name):
    skips.append(name)
    print("  [SKIP] {}".format(name))


def read_sample_names(path):
    """从进度表读出全部「项目名称」"""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    names = []
    try:
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            col = -1
            header_idx = -1
            for i, row in enumerate(rows[:10]):
                if row is None:
                    continue
                for j, cell in enumerate(row):
                    if cell is not None and str(cell).strip() == "项目名称":
                        col, header_idx = j, i
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
        wb.close()
    return names


def canon(s):
    """规范化比较（全角/半角统一，去掉数字编号）"""
    return unicodedata.normalize("NFKC", re.sub(r"\d+", "", s or ""))


# ============================================================
print("== 1. 默认配置为通用语法（不含具体区县） ==")

cfg0 = PC.default_config()
check("默认配置无硬编码区县名", not cfg0.get("region_words"),
      str(cfg0.get("region_words")))
hardcoded = ["某省", "某市", "某县", "癸市", "庚县", "辛县"]
src_files = ["project_config.py", "name_parser.py"]
leaked = []
for fn in src_files:
    p = os.path.join(ROOT, "src", "core", fn)
    with open(p, "r", encoding="utf-8") as f:
        body = f.read()
    # 允许出现在注释/文档字符串中作为示例，但不允许作为代码字面量
    for word in hardcoded:
        for m in re.finditer(r'.*["\']' + word + r'.*', body):
            line = m.group().strip()
            if line.startswith("#"):
                continue
            # 文档字符串中的示例行以说明性文字结尾，此处仅标记疑似字面量
            if re.search(r'["\']' + word + r'["\']', line):
                leaked.append("{}: {}".format(fn, line[:80]))
check("无地域词作为代码字面量", not leaked, str(leaked[:3]))

check("默认配置含业务后缀表", len(cfg0.get("business_suffixes", [])) > 0)
check("默认配置无年份黑名单", not any(
    re.fullmatch(r"20\d{3}", str(w)) for w in cfg0.get("noise_words", [])),
    "疑似年份黑名单残留")

print("== 2. 配置读写与持久化 ==")
tmp = tempfile.mkdtemp()
try:
    p = os.path.join(tmp, "cfg.json")
    cfg = PC.default_config()
    cfg["region_words"] = ["测试省", "测试市"]
    cfg["business_suffixes"] = ["测试工程"]
    PC.save_config(cfg, p)
    check("配置文件已生成", os.path.exists(p))
    loaded = PC.load_config(p)
    check("地域词持久化", loaded["region_words"] == ["测试省", "测试市"],
          str(loaded.get("region_words")))
    check("业务后缀持久化", loaded["business_suffixes"] == ["测试工程"],
          str(loaded.get("business_suffixes")))
    # 损坏文件不应抛异常
    with open(p, "w", encoding="utf-8") as f:
        f.write("{ not valid json ")
    broken = PC.load_config(p)
    check("损坏配置回退默认且不抛异常",
          broken["business_suffixes"] == PC.DEFAULT_CONFIG["business_suffixes"]
          and "load_error" in broken, str(broken.get("load_error")))
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# ============================================================
if not os.path.exists(SAMPLE_XLSX):
    skip("真实数据用例（未找到 {}）".format(os.path.basename(SAMPLE_XLSX)))
else:
    names = read_sample_names(SAMPLE_XLSX)
    inferred = PC.infer_from_excel(SAMPLE_XLSX)
    cfg = PC.default_config()
    cfg.update(inferred)

    print("== 3. 从真实 Excel 推导配置 ==")
    check("推导出地域词", bool(inferred.get("region_words")),
          str(inferred.get("region_words")))
    check("地域词无「年」字污染",
          all(not w.startswith("年") for w in inferred.get("region_words", [])),
          str(inferred.get("region_words")))
    check("推导出业务后缀", bool(inferred.get("business_suffixes")))
    check("后缀含低压维修类", any("低压" in s or "维修" in s
                                  for s in inferred.get("business_suffixes", [])),
          str(inferred.get("business_suffixes", [])[:6]))
    check("推导出有效关键词", bool(inferred.get("valid_keywords")),
          str(inferred.get("valid_keywords")))
    check("样本量 > 500 条", len(names) > 500, str(len(names)))

    print("== 4. 全部 {} 条真实名称解析 ===".format(len(names)))
    ok = 0
    empty = 0
    mismatch = []
    for n in names:
        got, dbg = NP.extract_name(n, cfg)
        if not got:
            empty += 1
            mismatch.append((n, "", dbg.get("how")))
            continue
        a, b = canon(n), canon(got)
        if b and (b in a or a in b):
            ok += 1
        else:
            mismatch.append((n, got, dbg.get("how")))
    rate = 100.0 * ok / len(names) if names else 0.0
    check("解析一致率 100%", rate == 100.0,
          "{}% ({}/{}), 空={}, 不一致={}".format(
              round(rate, 2), ok, len(names), empty, len(mismatch) - empty))
    check("无静默失败（提取不到）", empty == 0, "空结果 {} 条".format(empty))
    for n, got, how in mismatch[:5]:
        print("        不一致: {}  ->  [{}] {}".format(n[:40], how, got[:40]))

    print("== 5. 名称形态识别 ==")
    cases = [
        ("形态A 中压业扩",
         "某省某市某县3225000000000001某县某某建设发展有限公司10kV业扩配套工程",
         "业扩配套工程"),
        ("形态A 个人主体（旧实现会失败）",
         "某省某市某县3226000000000008某某某0.4kV业扩配套工程", "某某某"),
        ("形态B 资本打包",
         "某省某市某县2026年第一批包十0.4kV业扩配套", "包十"),
        ("形态C 低压维修（旧实现会失败）",
         "2026年某县供电公司某某北北变台区低压维修项目3225000000000002",
         "低压维修项目"),
        ("形态C 乡镇台区",
         "2026年某市某县丙镇某辛庄村某合村变低压台区维修项目3226000000000003",
         "低压台区维修项目"),
        ("形态D 项目包（旧实现会失败）",
         "2026年某市某县低压台区维修项目包31", "包31"),
    ]
    for desc, text, expect in cases:
        got, dbg = NP.extract_name(text, cfg)
        check(desc, bool(got) and expect in got,
              "got={!r} how={}".format(got, dbg.get("how")))

    print("== 6. OCR 噪声鲁棒性 ==")
    noisy = [
        ("换行断开的名称",
         "某省某市某县3225000000000001\n某县某某建设发展有限公司\n10kV业扩配套工程",
         "业扩配套工程"),
        ("标签干扰（工程名称/工程编号）",
         "工程名称：某省某市某县3225000000000001某县某某建设发展有限公司10kV业扩配套工程"
         "\n工程编号：181000000009", "业扩配套工程"),
        ("OCR 错字 业扩配食工程",
         "某省某市某县3225000000000001某县某某建设发展有限公司10kV业扩配食工程",
         "业扩配套工程"),
        ("全角数字与标点",
         "某省某市某县３２２５000000000010某县某某建设发展有限公司１０ｋＶ业扩配套工程",
         "业扩配套工程"),
        ("干扰时间字段存在",
         "配电网工程开工报告\n工程名称2026年某市某县丙镇某合村变低压台区维修项目"
         "3226000000000003\n计划开工时间20260408", "低压台区维修项目"),
    ]
    for desc, text, expect in noisy:
        got, dbg = NP.extract_name(text, cfg)
        check(desc, bool(got) and expect in got,
              "got={!r} how={}".format(got, dbg.get("how")))

    print("== 7. 跨地域 / 跨业务泛化 ==")
    other_region = [
        ("换区县（无锡江阴）",
         "某省某市某区3225000000000001江阴某某科技有限公司10kV业扩配套工程",
         "业扩配套工程"),
        ("换省份（浙江）",
         "某省某市某区3225000000000001杭州某某科技有限公司10kV业扩配套工程",
         "业扩配套工程"),
        ("换省份（山东）",
         "某省某市某区3225000000000001青岛某某机械有限公司10kV业扩配套工程",
         "业扩配套工程"),
    ]
    for desc, text, expect in other_region:
        got, dbg = NP.extract_name(text, cfg)
        check(desc, bool(got) and expect in got,
              "got={!r} how={}".format(got, dbg.get("how")))

    print("== 8. 不应误提取 ==")
    negatives = [
        ("纯杂质字段", "工程名称\n工程编号\n建设单位\n施工单位\n计划开工时间20260101"),
        ("空文本", ""),
        ("只有编号", "3225000000000001"),
        ("普通表格文字", "序号 项目属性 乡镇 客户联系人 勘察交底情况"),
    ]
    for desc, text in negatives:
        got, dbg = NP.extract_name(text, cfg)
        check("正确拒绝：" + desc, not got, "误提取为 {!r}".format(got))

print("== 9. 配置驱动：改配置即换业务 ==")
tmp2 = tempfile.mkdtemp()
try:
    cfg2 = PC.default_config()
    cfg2["region_words"] = ["测试省", "测试市"]
    cfg2["business_suffixes"] = ["光伏并网工程", "并网工程"]
    cfg2["valid_keywords"] = ["光伏"]
    txt = "测试省测试市3225000000000001某某新能源有限公司10kV光伏并网工程"
    got, dbg = NP.extract_name(txt, cfg2)
    check("仅改配置即可解析新业务类型", bool(got) and "光伏并网工程" in got,
          "got={!r} how={}".format(got, dbg.get("how")))
    # 旧业务后缀在新配置下不生效（证明配置真正生效）
    getattr(NP, "_expanded_suffixes")
    cfg2["business_suffixes"] = ["光伏并网工程"]
    txt2 = "测试省测试市3225000000000001某某新能源有限公司10kV业扩配套工程"
    got2, _ = NP.extract_name(txt2, cfg2)
    check("配置未覆盖的业务不再误匹配", not got2, "got={!r}".format(got2))
finally:
    shutil.rmtree(tmp2, ignore_errors=True)

print("== 10. 多行/多页投票（extract_best_name） ==")
rows_pool = [
    ["工程名称", "某省某市某县32250000000011", "73某县某某建设发展有限公司"],
    ["10kV业扩配套工程", "工程编号181000000009"],
]
got, dbg = NP.extract_best_name(rows_pool, cfg if os.path.exists(SAMPLE_XLSX)
                                else PC.default_config())
check("跨页拼接后仍可提取", bool(got) and "业扩配套工程" in got,
      "got={!r} how={}".format(got, dbg.get("how")))

# ============================================================
print("== 11. 地域词容错匹配（阶段二 R-2，A2-3） ==")
# 背景（§3.2 证据 2）：真实扫描件中「某县」被 OCR 识别为 甲县/乙县/丙县/丁县/
# 壬路，基于 text.find(地域词) 的精确查找在 OCR 场景下不可靠。
cfg_r = PC.default_config()
cfg_r["region_words"] = ["某省", "某市", "某县"]
_words = sorted(cfg_r["region_words"], key=lambda w: (-len(w), w))

# A2-3：「某县」的 5 种 OCR 写法均能被识别为同一地域
_ok_variants = []
for bad in ("某县", "甲县", "乙县", "丙县", "丁县"):
    t = "2026年%s供电公司某某北北变台区低压维修项目3225000000000002" % bad
    hit = NP.find_region_word(t, cfg_r, _words)
    if hit is not None and hit[1] == "某县":
        _ok_variants.append(bad)
check("地域词容错：某县 5 种 OCR 写法均识别为「某县」（%s）"
      % "/".join(_ok_variants),
      len(_ok_variants) == 5, "仅识别 {}".format(_ok_variants))

# 容错度可配置：写 0 时应退化为精确查找（向后兼容）
cfg_t0 = PC.default_config()
cfg_t0["region_words"] = ["某县"]
cfg_t0["region_fuzzy_tolerance"] = 0
_hit0 = NP.find_region_word("2026年甲县供电公司低压维修项目", cfg_t0, ["某县"])
check("容错度为 0 时退化为精确查找（不命中错字）", _hit0 is None, str(_hit0))
# 容错度 1 时应命中
cfg_t1 = PC.default_config()
cfg_t1["region_fuzzy_tolerance"] = 1
_hit1 = NP.find_region_word("2026年甲县供电公司低压维修项目", cfg_t1, ["某县"])
check("容错度为 1 时命中单字错字", _hit1 is not None and _hit1[1] == "某县",
      str(_hit1))
# 容错放宽到 auto（长词 ⌊len/3⌋）
check("容错度 auto 自适应长词", NP._region_tolerance("某某省", {
    "region_fuzzy_tolerance": "auto", "region_fuzzy_max_tolerance": 3}) >= 1,
    str(NP._region_tolerance("某某省", {
        "region_fuzzy_tolerance": "auto",
        "region_fuzzy_max_tolerance": 3})))
# 上限约束：不得超过 len-1（否则失去判别力）
check("容错度受上限约束（不超过 len-1）",
      NP._region_tolerance("某县", {"region_fuzzy_tolerance": 9,
                                    "region_fuzzy_max_tolerance": 9}) == 1,
      str(NP._region_tolerance("某县", {"region_fuzzy_tolerance": 9,
                                        "region_fuzzy_max_tolerance": 9})))

# A2-6：容错不得引入误匹配（5 个无关文本反例）
cfg_r["region_fuzzy_tolerance"] = 1
_negatives = [
    ("无关中文文本", "完全无关的一段中文文本内容"),
    ("沙漠（末字非区划字）", "2026年沙漠地区低压台区维修项目"),
    ("普通表格文字", "序号项目属性乡镇客户联系人"),
    ("只有编号", "3225000000000001"),
    ("空文本", ""),
]
_neg_hits = []
for desc, t in _negatives:
    hit = NP.find_region_word(t, cfg_r, _words)
    if hit is not None and hit[2] > 0:      # 仅统计「容错命中」
        _neg_hits.append((desc, hit))
check("容错不引入误匹配：5 个反例均未容错命中", not _neg_hits, str(_neg_hits))

# 要求区划后缀字一致：末字不是区划字的不得容错命中
# 说明：用「某丰」（末字非县）验证 —— 若关闭末字校验，「某丰」会以 1 字之差
# 命中「某县」；开启后必须不命中。注意「沛丰县」这类输入本身含「丰县」子串
# （末字是县、与某县差 1 字），属**合法**的 OCR 变体，不应作为反例。
cfg_ns = PC.default_config()
cfg_ns["region_words"] = ["某县"]
cfg_ns["region_fuzzy_tolerance"] = 1
_hit_ns = NP.find_region_word("2026年某丰低压维修项目3225000000000002", cfg_ns, ["某县"])
check("容错要求末字为区划字（沛丰不命中某县）", _hit_ns is None, str(_hit_ns))
# 反向确认：关闭末字校验后该词会被命中（证明该校验确实在起作用）
cfg_ns2 = PC.default_config()
cfg_ns2["region_words"] = ["某县"]
cfg_ns2["region_fuzzy_tolerance"] = 1
cfg_ns2["region_fuzzy_require_suffix_char"] = False
_hit_ns2 = NP.find_region_word("2026年某丰低压维修项目3225000000000002", cfg_ns2, ["某县"])
check("关闭末字校验后「某丰」会被容错命中（证明校验生效）",
      _hit_ns2 is not None and _hit_ns2[2] == 1, str(_hit_ns2))
# 合法的「县」结尾变体仍应命中（丰县 是 某县 的单字变体）
check("同为区划后缀的变体仍可命中（丰县→某县）",
      NP.find_region_word("2026年丰县低压维修项目3225000000000002", cfg_ns,
                          ["某县"]) is not None)

print("== 12. region_ocr_fixes 并入公共纠正表（阶段二） ==")
_cfg_fx = PC.default_config()
check("get_ocr_fixes 合并 ocr_fixes",
      PC.get_ocr_fixes(_cfg_fx).get("峻工") == "竣工")
_cfg_fx["region_ocr_fixes"] = {"乙县": "某县"}
check("get_ocr_fixes 合并 region_ocr_fixes",
      PC.get_ocr_fixes(_cfg_fx).get("乙县") == "某县")
_cfg_fx["ocr_fixes"] = {"乙县": "覆盖值"}
check("两表冲突时以 ocr_fixes 为准",
      PC.get_ocr_fixes(_cfg_fx).get("乙县") == "覆盖值")
check("空/异常配置返回空表不抛异常",
      PC.get_ocr_fixes({}) == {} and PC.get_ocr_fixes({"ocr_fixes": None}) == {}
      and PC.get_ocr_fixes({"ocr_fixes": "bad"}) == {})

print()
if skips:
    print("跳过 {} 项（缺少样本数据）".format(len(skips)))
if failures:
    print("结果：{} 项失败 -> {}".format(len(failures), failures))
    sys.exit(1)
print("结果：全部通过")
