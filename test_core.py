# -*- coding: utf-8 -*-
"""核心逻辑自测（不启动 GUI，不打包）

覆盖：路径管理 / 规则读写 / 文本匹配 / 文件名工具 / 扫描 / 撤销日志
"""
import os
import sys
import shutil
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from core import paths, rules as R, batch as B

failures = []


def check(name, cond, detail=""):
    if cond:
        print("  [PASS] {}".format(name))
    else:
        failures.append(name)
        print("  [FAIL] {}  {}".format(name, detail))


print("== 1. 路径管理 ==")
app_dir = paths.get_app_dir()
check("应用目录位于 APPDATA", app_dir.lower().startswith(
    os.environ.get("APPDATA", "").lower()), app_dir)
check("规则文件路径拼接", paths.get_rules_file().endswith(
    paths.RULES_FILENAME), paths.get_rules_file())

print("== 2. 文本清洗与匹配 ==")
check("全角转半角+去标点",
      R.normalize_text("测试，ＡＢＣ１２３（报告）") == "测试ABC123报告",
      R.normalize_text("测试，ＡＢＣ１２３（报告）"))

rules = [
    {"识别内容": "配电网工程开工报告", "映射内容": "开工报告"},
    {"识别内容": "竣工图", "映射内容": "竣工图纸"},
]
check("严格包含匹配",
      R.match_rule("2026年度配电网工程开工报告（单位工程）", rules) == "开工报告")
check("模糊匹配（OCR错字容错）",
      R.match_rule("配电网工程开工报靠", rules) == "开工报告",
      R.match_rule("配电网工程开工报靠", rules))
check("未匹配返回 None",
      R.match_rule("完全无关的内容", rules) is None)
check("空标题返回 None", R.match_rule("", rules) is None)
check("空规则列表返回 None", R.match_rule("任意", []) is None)

# 多行 OCR 标题（第一页顶部多行文字拼接）时，按行从上到下匹配，
# 避免下方干扰文字（如表格中其他报告的标题）抢先命中
rules2 = [
    {"识别内容": "配电网工程开工报告", "映射内容": "开工报告"},
    {"识别内容": "配电网工程竣工报告", "映射内容": "竣工报告"},
]
check("多行标题：上方竣工标题优先",
      R.match_rule("配电网工程竣工报告\n2026年度配电网工程开工报告\n施工单位：XX", rules2) == "竣工报告",
      R.match_rule("配电网工程竣工报告\n2026年度配电网工程开工报告\n施工单位：XX", rules2))
check("多行标题：上方开工标题优先",
      R.match_rule("2026年度配电网工程开工报告\n配电网工程竣工报告", rules2) == "开工报告",
      R.match_rule("2026年度配电网工程开工报告\n配电网工程竣工报告", rules2))

# 短关键词（4 字）场景：页面上同时含「实际开工时间」「实际竣工时间」时，
# 不能凭 0.75 相似度把「开工时间」误配成「开工报告」
rules3 = [
    {"识别内容": "开工报告", "映射内容": "开工报告"},
    {"识别内容": "接地装置接地电阻测试报告", "映射内容": "接地电阻测试报告"},
    {"识别内容": "竣工报告", "映射内容": "竣工报告"},
]
check("短关键词：开工/竣工时间不误匹配",
      R.match_rule("新建电缆终端共10套\n实际开工时间20260718\n实际竣工时间20260916\n工程名称2026年XX市XX县", rules3) is None,
      repr(R.match_rule("新建电缆终端共10套\n实际开工时间20260718\n实际竣工时间20260916\n工程名称2026年XX市XX县", rules3)))
check("短关键词：真正包含时正常匹配",
      R.match_rule("配电网工程竣工报告\n实际开工时间20260718", rules3) == "竣工报告",
      R.match_rule("配电网工程竣工报告\n实际开工时间20260718", rules3))

print("== 2b. 竣工报告误判回归（阶段一 R-1） ==")
# 背景（阶段文档 §3.2 证据 1，真实 OCR 文本）：
#   rules.py 原先未接入 project_config.ocr_fixes，标题匹配阶段不做任何
#   OCR 错字纠正，「峻工报告」原样进入匹配；又因模糊匹配平局时「先遍历
#   者胜」，规则表中开工报告在前 → 竣工报告被误判为开工报告，导致拆分
#   模式把竣工页当作新项目起点，项目数翻倍。
#   以下用例使用真实 OCR 文本，两层根因都覆盖。
from core import project_config as PC

_RULES_KAI_JUN = [
    {"识别内容": "配电网工程开工报告", "映射内容": "开工报告"},
    {"识别内容": "配电网工程竣工报告", "映射内容": "竣工报告"},
]

# 真实 OCR 文本（低压包扫描资料.pdf P21 / 包37）
_REAL_JUN = ("某省某市某县某单位2026年某市某县"
             "低压台区维修项目包37配电网工程峻工报告 申请峻工")
_REAL_KAI = ("某省某市某县某单位2026年某市某县"
             "低压台区维修项目包37配电网工程开工报告 申请开工")

check("[R-1] 真实 OCR 峻工报告 → 竣工报告（不得为开工报告）",
      R.match_rule(_REAL_JUN, _RULES_KAI_JUN) == "竣工报告",
      repr(R.match_rule(_REAL_JUN, _RULES_KAI_JUN)))
check("[R-1] 真实 OCR 开工报告 → 开工报告",
      R.match_rule(_REAL_KAI, _RULES_KAI_JUN) == "开工报告",
      repr(R.match_rule(_REAL_KAI, _RULES_KAI_JUN)))
check("[R-1] 标题简洁形态：峻工报告 → 竣工报告",
      R.match_rule("配电网工程峻工报告", _RULES_KAI_JUN) == "竣工报告",
      repr(R.match_rule("配电网工程峻工报告", _RULES_KAI_JUN)))

# A1-6：平局行为确定性 —— 规则表顺序调换后结果必须一致
# （原实现「先遍历者胜」，调换顺序就会得到不同答案）
check("[R-1] 平局消歧与规则表顺序无关",
      R.match_rule(_REAL_JUN, _RULES_KAI_JUN)
      == R.match_rule(_REAL_JUN, list(reversed(_RULES_KAI_JUN))),
      "kai_first={} jun_first={}".format(
          R.match_rule(_REAL_JUN, _RULES_KAI_JUN),
          R.match_rule(_REAL_JUN, list(reversed(_RULES_KAI_JUN)))))
check("[R-1] 平局行为连续 10 次一致",
      len({R.match_rule(_REAL_JUN, _RULES_KAI_JUN) for _ in range(10)}) == 1,
      str({R.match_rule(_REAL_JUN, _RULES_KAI_JUN) for _ in range(10)}))

# A1-4：OCR 纠正接入的向后兼容 —— 纠正表为空/缺失时行为与改动前一致
_saved_fixes = PC.DEFAULT_CONFIG["ocr_fixes"]
try:
    PC.DEFAULT_CONFIG["ocr_fixes"] = {}
    R.clear_ocr_fix_cache()
    check("[R-1] 纠正表为空：不报错且开工文本仍正确",
          R.match_rule("配电网工程开工报告 申请开工", _RULES_KAI_JUN) == "开工报告",
          repr(R.match_rule("配电网工程开工报告 申请开工", _RULES_KAI_JUN)))
    check("[R-1] 纠正表为空：已正确的竣工文本仍正确",
          R.match_rule("配电网工程竣工报告 申请竣工", _RULES_KAI_JUN) == "竣工报告",
          repr(R.match_rule("配电网工程竣工报告 申请竣工", _RULES_KAI_JUN)))
    # 纠正表缺失（None）也不得抛异常，且退化为不纠正
    PC.DEFAULT_CONFIG["ocr_fixes"] = None
    R.clear_ocr_fix_cache()
    check("[R-1] 纠正表缺失(None)：不报错，退化为不纠正",
          R.match_rule("配电网工程峻工报告", _RULES_KAI_JUN) in ("开工报告", "竣工报告"),
          repr(R.match_rule("配电网工程峻工报告", _RULES_KAI_JUN)))
finally:
    PC.DEFAULT_CONFIG["ocr_fixes"] = _saved_fixes
    R.clear_ocr_fix_cache()

# 纠正表内容校验：方向必须是「错误写法 → 正确写法」
check("[R-1] ocr_fixes 含 峻工→竣工",
      PC.DEFAULT_CONFIG["ocr_fixes"].get("峻工") == "竣工")
check("[R-1] ocr_fixes 含 峻工报告→竣工报告（长锚点优先）",
      PC.DEFAULT_CONFIG["ocr_fixes"].get("峻工报告") == "竣工报告")
check("[R-1] ocr_fixes 含 合区→台区",
      PC.DEFAULT_CONFIG["ocr_fixes"].get("合区") == "台区")
# 禁止宽泛替换：不得存在「开工→竣工」这类会破坏正常识别的词对
check("[R-1] ocr_fixes 未含宽泛替换 开工→竣工",
      "开工" not in PC.DEFAULT_CONFIG["ocr_fixes"])

# 长锚点优先：apply_ocr_fixes 必须先替换长键，避免长错词被短键切碎
check("[R-1] 长键优先：峻工报告 整体替换",
      PC.apply_ocr_fixes("配电网工程峻工报告", PC.default_config()) == "配电网工程竣工报告",
      repr(PC.apply_ocr_fixes("配电网工程峻工报告", PC.default_config())))

print("== 3. 文件名工具 ==")
check("非法字符清理", R.sanitize_filename('a<b>:c"d|e?.pdf') == "a_b__c_d_e_.pdf",
      R.sanitize_filename('a<b>:c"d|e?.pdf'))
tmp = tempfile.mkdtemp()
try:
    p1 = os.path.join(tmp, "开工报告.pdf")
    open(p1, "w").close()
    check("冲突加序号", os.path.basename(R.resolve_conflict(p1)) == "开工报告_1.pdf")
    check("无冲突原样返回", os.path.basename(
        R.resolve_conflict(os.path.join(tmp, "不存在.pdf"))) == "不存在.pdf")

    sub = os.path.join(tmp, "模板文件夹")
    os.makedirs(sub)
    pf = os.path.join(sub, "原始扫描件.pdf")
    open(pf, "w").close()
    check("仅映射内容命名", R.build_new_name("开工报告", pf, False) == "开工报告.pdf")
    check("文件夹名+映射内容", R.build_new_name("开工报告", pf, True) == "模板文件夹开工报告.pdf")
finally:
    shutil.rmtree(tmp, ignore_errors=True)

print("== 4. 规则 Excel 读写 ==")
tmp = tempfile.mkdtemp()
try:
    rf = os.path.join(tmp, "规则.xlsx")
    R.create_default_rules(rf)
    loaded = R.read_rules(rf)
    check("默认规则生成与读取", len(loaded) == len(R.DEFAULT_RULES)
          and loaded[0]["识别内容"] == R.DEFAULT_RULES[0]["识别内容"])
    # 保存自定义规则再读回
    custom = [{"识别内容": "关键词A", "映射内容": "甲"}, {"识别内容": "关键词B", "映射内容": ""}]
    R.save_rules(rf, custom)
    loaded2 = R.read_rules(rf)
    check("自定义规则保存读回", loaded2 == custom, str(loaded2))
finally:
    shutil.rmtree(tmp, ignore_errors=True)

print("== 5. 扫描与撤销日志 ==")
# 说明：撤销日志位于 %APPDATA%\PdfOcrRenamer\rename_log.json，是「全局」状态。
# 本用例需要真实验证「执行重命名 → 撤销」的完整链路（不能用 log_undo=False，
# 否则就没有日志可撤销）。因此这里**先备份日志、结束后还原**，
# 既让断言可以写成确定值，又不会污染用户数据与其它套件（阶段文档 §5.5 第 1 条）。
_log_file = paths.get_rename_log_file()
_log_backup = None
if os.path.exists(_log_file):
    _log_backup = _log_file + ".testcore.bak"
    shutil.copy2(_log_file, _log_backup)
if os.path.exists(_log_file):
    os.remove(_log_file)          # 清空，使下面的断言可写成确定值

tmp = tempfile.mkdtemp()
try:
    for d in ("a", "b", "b", "c"):
        pass
    os.makedirs(os.path.join(tmp, "子目录", "孙目录"))
    for name in ("a1.pdf", "a2.PDF", "b.txt",
                 os.path.join("子目录", "c1.pdf"),
                 os.path.join("子目录", "孙目录", "d1.pdf")):
        fp = os.path.join(tmp, name)
        open(fp, "w").close()
    found = B.scan_pdfs(tmp)
    check("递归扫描（含子目录、大小写扩展名）", len(found) == 4, str(found))

    # 日志已清空 → 初始无历史
    check("初始无可撤销记录", B.get_last_rename_count() == 0,
          str(B.get_last_rename_count()))

    # 模拟执行重命名 + 撤销
    target = os.path.join(tmp, "子目录", "c1.pdf")
    fake_results = [{"pdf_path": target, "base_name": "c1.pdf",
                     "mapping": "测试报告", "status": "match"}]
    renamed, skipped = B.execute_renames(fake_results, by_folder=False)
    check("执行重命名成功", len(renamed) == 1 and os.path.exists(
        os.path.join(tmp, "子目录", "测试报告.pdf")))
    check("撤销日志已记录最近一笔", B.get_last_rename_count() == 1)
    restored, failed = B.undo_last()
    check("撤销恢复原名", restored == 1 and os.path.exists(target)
          and not os.path.exists(os.path.join(tmp, "子目录", "测试报告.pdf")))
    check("撤销后日志回到空状态", B.get_last_rename_count() == 0,
          str(B.get_last_rename_count()))
    # 阶段四补齐：log_undo=False 时不得写日志
    target2 = os.path.join(tmp, "a1.pdf")
    renamed2, _ = B.execute_renames(
        [{"pdf_path": target2, "base_name": "a1.pdf",
          "mapping": "不记账报告", "status": "match"}],
        by_folder=False, log_undo=False)
    check("log_undo=False 时不写撤销日志（阶段四）",
          len(renamed2) == 1 and B.get_last_rename_count() == 0,
          str(B.get_last_rename_count()))
finally:
    shutil.rmtree(tmp, ignore_errors=True)
    # 还原撤销日志
    try:
        if _log_backup and os.path.exists(_log_backup):
            shutil.copy2(_log_backup, _log_file)
            os.remove(_log_backup)
        elif os.path.exists(_log_file):
            os.remove(_log_file)
    except OSError:
        pass

print()
if failures:
    print("结果：{} 项失败 -> {}".format(len(failures), failures))
    sys.exit(1)
print("结果：全部通过")
