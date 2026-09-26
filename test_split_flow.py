# -*- coding: utf-8 -*-
"""拆分流程回归测试：前置页面保全（P0-2）与配置驱动集成

背景：
    原实现在「首个开工报告之前的页面」上只做计数后 continue，页面被静默丢弃
    （仅写进 warnings，界面只显示第一条，用户极易忽略），导致扫描归档时
    封面/目录/上一份文件尾页等资料永久不在任何输出中。

本测试构造一个 5 页整份 PDF：
    P0 封面（非开工报告，属前置页面）
    P1 开工报告（项目1开始）
    P2 竣工报告（项目1第二份资料）
    P3 开工报告（项目2开始）
    P4 无标题页（并入项目2的开工报告）
验证：
    1. 前置页面 P0 被保全到「未匹配项目/_前置页面/」
    2. 两个项目被正确切分
    3. 全部页面都被输出覆盖，无丢失（总页数守恒）
"""
import os
import sys
import re
import shutil
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from core import splitter as S
from core import project_config as PC
from core import name_parser as NP

failures = []


def check(name, cond, detail=""):
    if cond:
        print("  [PASS] {}".format(name))
    else:
        failures.append(name)
        print("  [FAIL] {}  {}".format(name, detail))


print("== 1. 前置页面保全（P0-2 修复验证） ==")
tmp = tempfile.mkdtemp()
try:
    from pypdf import PdfWriter, PdfReader

    src_pdf = os.path.join(tmp, "整份扫描.pdf")
    w = PdfWriter()
    for _ in range(5):
        w.add_blank_page(width=595, height=842)
    with open(src_pdf, "wb") as f:
        w.write(f)

    # 直接构造 analysis（跳过 OCR）：模拟 analyze_pdf 的输出结构
    # 其中 projects[0] 为前置页面项目（is_pre_pages=True）
    sub = os.path.join(tmp, "项目甲")
    os.makedirs(sub)

    analysis = {
        "pdf_path": src_pdf,
        "root": tmp,
        "page_count": 5,
        "pre_pages": 1,
        "warnings": ["文件开头有 1 页在首个开工报告之前"],
        "projects": [
            {"name": "", "matched_folder": None, "no_folder": True,
             "is_pre_pages": True, "start_page": 0, "end_page": 0,
             "page_range": "P0",
             "docs": [{"mapping": "", "title": "封面", "page_start": 0,
                       "page_end": 0, "pages": "P0",
                       "file_name": "前置页面_P0-P0.pdf"}]},
            {"name": "某省某县10kV业扩配套工程", "matched_folder": "项目甲",
             "no_folder": False, "start_page": 1, "end_page": 2,
             "page_range": "P1-P2",
             "docs": [
                 {"mapping": "开工报告", "page_start": 1, "page_end": 1,
                  "pages": "P1", "file_name": "项目甲开工报告.pdf"},
                 {"mapping": "竣工报告", "page_start": 2, "page_end": 2,
                  "pages": "P2", "file_name": "项目甲竣工报告.pdf"},
             ]},
            {"name": "某省某县低压维修", "matched_folder": None,
             "no_folder": True, "start_page": 3, "end_page": 4,
             "page_range": "P3-P4",
             "docs": [{"mapping": "开工报告", "page_start": 3, "page_end": 4,
                       "pages": "P3-P4", "file_name": "开工报告.pdf"}]},
        ],
    }

    # log_undo=False：本套件不验证撤销，且撤销日志是 %APPDATA% 下的全局共享
    # 状态，写入会污染用户数据与其它套件的基线断言（阶段文档 §5.5 第 1 条）
    written, skipped = S.execute_split(analysis, log_undo=False)
    check("全部文档写出成功", len(skipped) == 0,
          "written={} skipped={}".format(len(written), skipped))

    # 1) 前置页面保全
    pre_dir = os.path.join(tmp, "未匹配项目", "_前置页面")
    check("前置页面目录已创建", os.path.isdir(pre_dir), str(pre_dir))
    pre_files = os.listdir(pre_dir) if os.path.isdir(pre_dir) else []
    check("前置页面文件已写出（不再丢失）",
          any("前置页面" in f for f in pre_files), str(pre_files))

    # 2) 项目正常拆分
    check("项目甲两个文档写出",
          os.path.exists(os.path.join(sub, "项目甲开工报告.pdf"))
          and os.path.exists(os.path.join(sub, "项目甲竣工报告.pdf")),
          str(os.listdir(sub)))

    # 3) 页面守恒：所有输出页数之和 == 原 PDF 页数
    total_out = 0
    for item in written:
        r = PdfReader(item["dst"])
        total_out += len(r.pages)
    check("页面守恒（输出总页数 == 源页数 5）", total_out == 5,
          "输出 {} 页".format(total_out))

    # 4) 原文件保留
    check("原整份 PDF 保留不变", os.path.exists(src_pdf))

    # 5) 每个源页都被覆盖恰好一次
    covered = []
    for proj in analysis["projects"]:
        for doc in proj["docs"]:
            covered.extend(range(doc["page_start"], doc["page_end"] + 1))
    check("页面覆盖无遗漏无重复", sorted(covered) == list(range(5)),
          str(sorted(covered)))
finally:
    shutil.rmtree(tmp, ignore_errors=True)

print("== 2. 配置驱动集成（splitter 使用外置配置） ==")
cfg = PC.default_config()
cfg["region_words"] = ["测试省", "测试市"]
cfg["business_suffixes"] = ["光伏并网工程"]
S.set_config(cfg)
name = S._extract_engineering_name([
    [("测试省测试市3225000000000001某某新能源有限公司10kV光伏并网工程",)],
])
check("splitter 使用注入的配置解析名称", name != "" and "光伏并网工程" in name,
      "name={!r}".format(name))

# 恢复默认配置，避免影响其他测试
S.set_config(PC.default_config())
name2 = S._extract_engineering_name([
    [("某省某市某县3225000000000001某县某某建设发展有限公司10kV业扩配套工程",)],
])
check("默认配置仍可解析通用名称", "业扩配套工程" in name2, "name={!r}".format(name2))

print("== 3. 文件夹匹配阈值可配置 ==")
cfg2 = PC.default_config()
cfg2["folder_score_threshold"] = 0.99
S.set_config(cfg2)
check("阈值可配置生效", abs(S._folder_threshold() - 0.99) < 1e-9,
      str(S._folder_threshold()))
S.set_config(PC.default_config())
check("默认阈值回退正确", abs(S._folder_threshold() - 0.55) < 1e-9,
      str(S._folder_threshold()))

print("== 4. 开工报告边界强约束（阶段一 1.4，R-1） ==")
# 背景：_rows_contain_kai 原实现只有「包含开工报告」判断、无排除逻辑。
# 竣工页含「实际开工时间」等字样时会被当作项目边界，导致项目数翻倍。
# 现改为「排除优先于包含」，并对「峻工→竣工」先做 OCR 纠正。
_kai_cases = [
    # (行文本, 期望是否为开工报告页, 说明)
    ("配电网工程开工报告 申请开工", True, "正常开工页"),
    ("配电网工程竣工报告 申请竣工", False, "竣工页不得判为开工页"),
    ("配电网工程峻工报告 申请峻工", False, "竣工页 OCR 错字形态"),
    ("配电网工程竣工报告\n实际开工时间20260718", False,
     "竣工页含开工时间字样（排除优先）"),
    ("配电网工程竣工报告\n计划开工时间20260921\n实际竣工时间20261015", False,
     "竣工页同时含开工/竣工字段（排除优先）"),
    ("新建电缆终端共10套", False, "无关页面"),
]
for _text, _expect, _desc in _kai_cases:
    _rows = [(i * 20, line) for i, line in enumerate(_text.split("\n"))]
    _got = S._rows_contain_kai(_rows)
    check("边界判定：{}".format(_desc), _got == _expect,
          "text={!r} got={} expect={}".format(_text, _got, _expect))

# A1-7：同一项目含「开工+竣工+其他资料」时只切出 1 个项目边界。
# 直接驱动 analyze_pdf 的切分循环需要真实 OCR，这里以「页级分类结果」
# 复现同一逻辑，验证竣工页不会额外产生项目边界。
_page_class = [
    {"page": 0, "mapping": "开工报告", "rows": [(0, "配电网工程开工报告 申请开工")]},
    {"page": 1, "mapping": "竣工报告", "rows": [(0, "配电网工程竣工报告 申请竣工")]},
    {"page": 2, "mapping": "接地电阻测试报告", "rows": [(0, "接地装置接地电阻测试报告")]},
    {"page": 3, "mapping": "工作票", "rows": [(0, "工作票")]},
]
_boundaries = [pc["page"] for pc in _page_class
               if pc["mapping"] == "开工报告" or S._rows_contain_kai(pc["rows"])]
check("开工+竣工+其他资料 → 只切出 1 个项目",
      len(_boundaries) == 1 and _boundaries[0] == 0, str(_boundaries))

# 反过来：两个项目各自「开工+竣工」时，应切出 2 个项目
_page_class2 = [
    {"page": 0, "mapping": "开工报告", "rows": [(0, "配电网工程开工报告 申请开工")]},
    {"page": 1, "mapping": "竣工报告", "rows": [(0, "配电网工程峻工报告 申请峻工")]},
    {"page": 2, "mapping": "开工报告", "rows": [(0, "配电网工程开工报告 申请开工")]},
    {"page": 3, "mapping": "竣工报告", "rows": [(0, "配电网工程竣工报告 申请竣工")]},
]
_boundaries2 = [pc["page"] for pc in _page_class2
                if S._is_project_boundary(pc)]
check("两个项目各含开工+竣工 → 切出 2 个项目（非 4 个）",
      _boundaries2 == [0, 2], str(_boundaries2))

print("== 5. 项目边界排除优先（真实样本暴露的缺陷） ==")
# 背景（真实 30 页样本实测）：_classify_page 按角度顺序尝试旋转、命中即返回。
# 竣工报告扫描页的正确方向是横排，但当该方向顶部裁剪无文字时会 break 换角度，
# 最终落到 angle=0（未旋转侧读），OCR 出竖排乱序文本，其中同时含
# 「实际竣工时间」与「实际开工时间」，模糊匹配抢先命中「开工报告」。
# 若仅凭 mapping == "开工报告" 切项目，竣工页会被当作新项目起点——
# 实测因此把 3 个项目切成 4 个。以下为真实 P1 的 OCR 行。
_P1_ROWS = [
    (10.0, "负责人：建设单位意见：填报人：施工意见建复合材料管，PVC,DN50共227米；"
           "11、新建防火堵料共61千克：12、新建铁构件共0.12吨；线，AC"),
    (28.0, "实际峻工时间：实际开工时间：注："),
    (44.0, "建设单位：工程名称："),
    (73.0, "我单位已按照各项要求施工完毕，"),
    (135.0, "某省某市某县某单位2026年某市某县"
            "低压台区维修项目包36"),
    (185.0, "，50,4芯，备管理2026年9月24日2026年9月21日"),
]
check("竣工页含竣工锚点 → 即使 mapping 误判为开工报告也不作为边界",
      S._is_project_boundary({"mapping": "开工报告", "rows": _P1_ROWS}) is False,
      "jun_anchor={}".format(S._rows_have_jun_anchor(_P1_ROWS)))

_P0_ROWS = [
    (31.0, "配电网工程开工报告"),
    (71.0, "工程名称2026年某市某县低压台区维修项目包36工程编号"),
    (143.0, "某市乙安工程安装有限公司计划开工时间"),
]
check("正常开工页仍作为边界",
      S._is_project_boundary({"mapping": "开工报告", "rows": _P0_ROWS}) is True)

# 真实六页序列（P0/P1/P10/P11/P20/P21）应切出 3 个项目边界
_six = [
    ("P0", "开工报告", _P0_ROWS), ("P1", "开工报告", _P1_ROWS),
    ("P10", "开工报告", _P0_ROWS), ("P11", "竣工报告", _P1_ROWS),
    ("P20", "开工报告", _P0_ROWS), ("P21", "竣工报告", _P1_ROWS),
]
_bounds6 = [lbl for lbl, mp, rw in _six
            if S._is_project_boundary({"mapping": mp, "rows": rw})]
check("真实六页序列 → 3 个项目边界（P0/P10/P20）",
      _bounds6 == ["P0", "P10", "P20"], str(_bounds6))

# 真实 P0 页（完整标题）与 P1 页（残缺标题）——来自 _classify_page 实测输出。
# 关键：P0/P1 是同一项目的开工+竣工两页（工程编号同为 B11000000001），
# P1 标题被 OCR 成「配电网工程工报告」（「竣」字丢失），该文本对
# 「配电网工程开工报告」「配电网工程竣工报告」相似度完全相同（0.9412），
# 属天然歧义，不得据此切分新项目。
_P0_REAL_ROWS = [
    (1.0, "项目负责人2026年9月21日计划开工时间"),
    (22.0, "某市甲安工程安装有限公司施工单位某省某市某县某单位"),
    (43.0, "建设单位"),
    (74.0, "B11000000001工程编号2026年某市某县低压台区维修项目包36"),
    (93.0, "工程名称"),
    (126.0, "配电网工程开工报告"),
]
_P1_REAL_ROWS = [
    (0.0, "吴酸郁项目负责人2026年9月21日实际开工时间："),
    (37.0, "某市甲安工程安装有限公司施工单位：某省某市某县某单位建设单位："),
    (87.0, "B11000000001工程编号：2026年某市某县低压台区维修项目包36工程名称："),
    (128.0, "配电网工程工报告"),
]
check("真实 P0（完整标题『配电网工程开工报告』）→ 是项目边界",
      S._is_project_boundary({"mapping": "开工报告", "rows": _P0_REAL_ROWS}) is True)
check("真实 P1（残缺标题『配电网工程工报告』）→ 不是项目边界",
      S._is_project_boundary({"mapping": "开工报告", "rows": _P1_REAL_ROWS}) is False,
      "degraded={}".format(S._rows_have_degraded_kai_title(_P1_REAL_ROWS)))
check("残缺标题识别：『工报告』为残缺，『开工报告/竣工报告』不是",
      S._rows_have_degraded_kai_title([(0, "配电网工程工报告")]) is True
      and S._rows_have_degraded_kai_title([(0, "配电网工程开工报告")]) is False
      and S._rows_have_degraded_kai_title([(0, "配电网工程竣工报告")]) is False)
# 两页同工程编号 → 同属一个项目，只应有 1 个边界
_pair = [("P0", "开工报告", _P0_REAL_ROWS), ("P1", "开工报告", _P1_REAL_ROWS)]
check("同工程编号的开工+竣工两页 → 只切出 1 个边界",
      len([1 for _l, m, r in _pair
           if S._is_project_boundary({"mapping": m, "rows": r})]) == 1)

print("== 6. 特征词主键与图纸噪声剥离（阶段二 R-3/2.4） ==")
# 真实图纸标题栏文本（取自 2026年低压成本单体第七批7项竣工图纸.pdf P0 实测）
_cfg2 = PC.default_config()
_DRAWING_TITLE = ("某市某某电力勘察设计有限公司2026年某市某县乙镇某堤口2*配变"
                  "低压合区维修项目3226|000000000设计峻工图CAD制图1: 1B11000000001RQ")
_stripped = S.strip_noise_patterns(_DRAWING_TITLE, _cfg2)
check("设计单位被剥离（配置驱动，未硬编码公司名）",
      "某市某某电力勘察设计有限公司" not in _stripped, repr(_stripped[:60]))
check("特征词保留台区/配变描述",
      "某堤口" in S.feature_key(_DRAWING_TITLE, _cfg2),
      repr(S.feature_key(_DRAWING_TITLE, _cfg2)))

# 图纸标题 vs 进度表文件夹名：整串差异极大（多出设计单位/图号噪声）
_FOLDER = "2026年某市某县乙镇某堤口2#配变低压台区维修项目3226000000000007"
_sc = S._folder_score("2026年某市戊县乙镇某堡变低压台区维修项目",
                      "2026年某市某县乙镇某堡变低压台区维修项目", _cfg2)
check("整串近似（含 OCR 错字）得分高于阈值", _sc >= 0.55, "score=%.3f" % _sc)
_sc2 = S._folder_score("2026年某市某县乙镇某堡变低压台区维修项目",
                       "2026年某市某县乙镇某刘庄3#变低压台区维修项目", _cfg2)
check("不同台区不得判定为高相似", _sc2 < 0.95, "score=%.3f" % _sc2)

print("== 7. 用文件夹名反向纠正 OCR 错字（阶段二 2.2） ==")
_FOLDERS = [
    "2026年某市某县乙镇某堤口2#配变低压台区维修项目3226000000000007",
    "2026年某市某县乙镇某堡变低压台区维修项目3225000000010001",
    "2026年某市某县甲镇某楼南台片低压台区维修项目3226000000000012",
]
# 含 OCR 错字（戊县=某县、合区=台区）的名称应被纠正
_ocr_name = "2026年某市戊县乙镇某堡变低压合区维修项目"
_fixed, _info = S.correct_name_by_folders(_ocr_name, _FOLDERS, _cfg2)
check("反纠错把「戊县」修正为「某县」", "某县" in _fixed and "戊县" not in _fixed,
      "fixed={!r}".format(_fixed))
check("反纠错把「合区」修正为「台区」", "台区" in _fixed,
      "fixed={!r}".format(_fixed))
check("反纠错报告命中模板", _info.get("changed") and _info.get("folder"),
      str(_info))

# A2-5：反向纠错不改变已正确的结果
_ok_name = "2026年某市某县乙镇某堡变低压台区维修项目"
_fixed2, _info2 = S.correct_name_by_folders(_ok_name, _FOLDERS, _cfg2)
check("反纠错不改变已正确的名称", _fixed2 == _ok_name and not _info2.get("changed"),
      "fixed={!r} info={}".format(_fixed2, _info2))
# 完全等于某文件夹名时不改动
_exact, _info3 = S.correct_name_by_folders(_FOLDERS[1], _FOLDERS, _cfg2)
check("名称与文件夹完全一致时原样返回", _exact == _FOLDERS[1], repr(_exact))

# 边界：空输入、无候选文件夹
check("反纠错：空名称原样返回",
      S.correct_name_by_folders("", _FOLDERS, _cfg2)[0] == "")
check("反纠错：无候选文件夹原样返回",
      S.correct_name_by_folders(_ocr_name, [], _cfg2)[0] == _ocr_name)
# 差异过大时不强行套模板（避免误改）
_far, _info4 = S.correct_name_by_folders("完全无关的另一段项目文本内容", _FOLDERS, _cfg2)
check("差异过大时不套用模板", not _info4.get("changed"), str(_info4))

print("== 8. 特征词配置驱动（不含硬编码公司名） ==")
# 与 test_name_parser.py 对地域词的口径一致：允许出现在注释/文档字符串中
# 作为示例，但不允许作为**代码字面量**（即引号包裹的值）。
_src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "src", "core", "splitter.py"), encoding="utf-8").read()
_leaked = []
for _w in ("某某电力", "勘察设计有限公司", "电力勘察设计"):
    for _line in _src.splitlines():
        if _line.lstrip().startswith("#"):
            continue
        # 仅当该词被引号包裹成字面量时才视为硬编码
        if re.search(r'["\'][^"\']*' + re.escape(_w) + r'[^"\']*["\']', _line):
            _leaked.append("{}: {}".format(_w, _line.strip()[:70]))
check("splitter.py 未把设计单位公司名硬编码为字面量", not _leaked, str(_leaked[:3]))
check("设计单位模式来自配置", bool(_cfg2.get("design_unit_pattern")))
check("特征词剥离模式来自配置", bool(_cfg2.get("feature_strip_patterns")))
check("特征词/噪声模式均为通用正则（不含具体公司名）",
      "某某电力" not in str(_cfg2.get("design_unit_pattern"))
      and "某某电力" not in str(_cfg2.get("org_noise_pattern")))

print("== 9. 项目编号信号（区分同台区不同批次，阶段二 R-3） ==")
# 实测：第七批7项中有两个「某堡变」记录，特征词完全相同，仅编号不同。
_F_B1 = "2026年某市某县乙镇某堡变低压台区维修项目3225000000010001"
_F_B2 = "2026年某市某县乙镇某堡变低压台区维修项目3226000000090002"
check("同台区两条记录的特征词确实相同（证明必须用编号区分）",
      S.feature_key(_F_B1, _cfg2) == S.feature_key(_F_B2, _cfg2),
      "%s vs %s" % (S.feature_key(_F_B1, _cfg2), S.feature_key(_F_B2, _cfg2)))

# 编号一致（含 OCR 个别数字错字）时正确区分
_N_B1 = "2026年某市戊县乙镇某堡变低压台区维修项目3225000000010011"
_N_B2 = "2026年某市戊县乙镇某堡变低压台区维修项目3226000000090022"
_got1 = S.match_folder(_N_B1, [_F_B1, _F_B2])
_got2 = S.match_folder(_N_B2, [_F_B1, _F_B2])
check("编号（容错）可区分同台区记录 → 第一条",
      _got1[0] == _F_B1, "%s (%.3f)" % (_got1[0], _got1[1]))
check("编号（容错）可区分同台区记录 → 第二条",
      _got2[0] == _F_B2, "%s (%.3f)" % (_got2[0], _got2[1]))

# 编号等价判定（容错）
check("编号等价：完全相等",
      S._codes_equivalent("3225000000010001", "3225000000010001"))
check("编号等价：1~2 位数字错字仍等价",
      S._codes_equivalent("3225000000010011", "3225000000010001"))
check("编号等价：差异过大不等价",
      not S._codes_equivalent("3225000000999999", "3225000000010001"))
check("编号等价：空值不等价", not S._codes_equivalent("", "3225000000010001"))

# 关闭编号信号后应退回原行为（可配置、可回退）
_cfg_nocode = PC.default_config()
_cfg_nocode["code_match_enabled"] = False
_sc_on = S._folder_score(_N_B1, _F_B2, _cfg2)
_sc_off = S._folder_score(_N_B1, _F_B2, _cfg_nocode)
check("编号信号可关闭（关闭后不再降分）", _sc_off > _sc_on,
      "on=%.3f off=%.3f" % (_sc_on, _sc_off))

print("== 10. 编号符号容错（# vs *，阶段二） ==")
# 真实样本：图纸标题「某堤口2#配变」被 OCR 成「某堤口2*配变」，
# '*' 不在名称合法字符集内 → 整个名称曾被判非法、静默返回空串。
_sym_text = ("某市某某电力勘察设计有限公司2026年某市某县乙镇某堤口2*配变"
             "低压台区维修项目3226|000000000设计峻工图CAD制图")
_sym_name, _sym_dbg = NP.extract_name(_sym_text, _cfg2)
check("含 * 的编号不再导致提取失败",
      bool(_sym_name), "how=%s" % _sym_dbg.get("how"))
check("符号容错把 2*配变 归一为 2#配变",
      "2#配变" in (_sym_name or ""), repr(_sym_name))
check("symbol_fixes 来自配置",
      _cfg2.get("symbol_fixes", {}).get("*") == "#")
# 关闭符号容错时应恢复旧行为（该名称因 * 非法而被拒）
_cfg_nosym = PC.default_config()
_cfg_nosym["symbol_fixes"] = {}
_cfg_nosym["name_charset"] = r"\u4e00-\u9fa5A-Za-z0-9#（）()\-\."
_sym_name2, _ = NP.extract_name(_sym_text, _cfg_nosym)
check("关闭符号容错后行为回退（不改变既有判定）",
      _sym_name2 == "", repr(_sym_name2))

print()
if failures:
    print("结果：{} 项失败 -> {}".format(len(failures), failures))
    sys.exit(1)
print("结果：全部通过")
