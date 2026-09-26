# -*- coding: utf-8 -*-
"""拆分命名方式功能自测（不启动 GUI，不打包）

覆盖：
1. _assign_file_names：按命名方式预计算文件名（仅标题 / 文件夹名+标题）
2. execute_split：按预计算文件名写出文档
"""
import os
import sys
import shutil
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from core import splitter as S

failures = []


def check(name, cond, detail=""):
    if cond:
        print("  [PASS] {}".format(name))
    else:
        failures.append(name)
        print("  [FAIL] {}  {}".format(name, detail))


print("== 1. 文件名预计算（_assign_file_names） ==")
# 匹配到文件夹
projects = [
    {"name": "某省某县10kV业扩配套工程", "matched_folder": "模板工程A", "docs": [
        {"mapping": "开工报告"},
        {"mapping": "竣工图纸"},
        {"mapping": ""},           # 未识别标题
        {"mapping": "含:非法/字符"},
    ]},
    # 未匹配到文件夹 → 用工程名称
    {"name": "某市某县业扩工程", "matched_folder": None, "docs": [
        {"mapping": "验收单"},
    ]},
    # 未匹配且工程名称为空 → 仅标题
    {"name": "", "matched_folder": None, "docs": [
        {"mapping": "工作票"},
    ]},
]

S._assign_file_names(projects, by_folder=False)
check("仅标题：匹配项目", projects[0]["docs"][0]["file_name"] == "开工报告.pdf",
      projects[0]["docs"][0].get("file_name"))
check("仅标题：未匹配项目", projects[1]["docs"][0]["file_name"] == "验收单.pdf",
      projects[1]["docs"][0].get("file_name"))
check("仅标题：空映射→未识别标题", projects[0]["docs"][2]["file_name"] == "未识别标题.pdf",
      projects[0]["docs"][2].get("file_name"))

S._assign_file_names(projects, by_folder=True)
check("文件夹名+标题：匹配项目", projects[0]["docs"][0]["file_name"] == "模板工程A开工报告.pdf",
      projects[0]["docs"][0].get("file_name"))
check("文件夹名+标题：未识别标题也带前缀", projects[0]["docs"][2]["file_name"] == "模板工程A未识别标题.pdf",
      projects[0]["docs"][2].get("file_name"))
check("文件夹名+标题：非法字符已清理", projects[0]["docs"][3]["file_name"] == "模板工程A含_非法_字符.pdf",
      projects[0]["docs"][3].get("file_name"))
check("未匹配→工程名称+标题", projects[1]["docs"][0]["file_name"] == "某市某县业扩工程验收单.pdf",
      projects[1]["docs"][0].get("file_name"))
check("未匹配且名称为空→仅标题", projects[2]["docs"][0]["file_name"] == "工作票.pdf",
      projects[2]["docs"][0].get("file_name"))

print("== 2. 执行拆分（execute_split 按 file_name 写出） ==")
tmp = tempfile.mkdtemp()
try:
    # 造一个 3 页的整份 PDF
    from pypdf import PdfWriter
    src_pdf = os.path.join(tmp, "整份扫描.pdf")
    w = PdfWriter()
    for _ in range(3):
        w.add_blank_page(width=595, height=842)
    with open(src_pdf, "wb") as f:
        w.write(f)

    sub = os.path.join(tmp, "模板工程A")
    os.makedirs(sub)

    analysis = {
        "pdf_path": src_pdf,
        "root": tmp,
        "page_count": 3,
        "projects": [
            {"name": "某省某县10kV业扩配套工程", "matched_folder": "模板工程A",
             "start_page": 0, "end_page": 2, "no_folder": False, "docs": [
                {"mapping": "开工报告", "page_start": 0, "page_end": 0,
                 "pages": "P0", "file_name": "模板工程A开工报告.pdf"},
                {"mapping": "竣工图纸", "page_start": 1, "page_end": 2,
                 "pages": "P1-P2", "file_name": "模板工程A竣工图纸.pdf"},
            ]},
        ],
    }
    # log_undo=False：撤销日志是 %APPDATA% 下的全局共享状态，本套件不验证撤销，
    # 写入会污染用户数据与其它套件的基线断言（阶段文档 §5.5 第 1 条）
    written, skipped = S.execute_split(analysis, log_undo=False)
    check("两个文档均写出", len(written) == 2 and len(skipped) == 0,
          "written={} skipped={}".format(written, skipped))
    check("文件名带文件夹前缀", os.path.exists(
        os.path.join(sub, "模板工程A开工报告.pdf")) and os.path.exists(
        os.path.join(sub, "模板工程A竣工图纸.pdf")), str(os.listdir(sub)))
    check("新文件存在且原 PDF 保留", os.path.exists(src_pdf))

    # 未匹配项目 → 未匹配项目/<工程名称>/ 下同样带前缀
    analysis2 = {
        "pdf_path": src_pdf,
        "root": tmp,
        "page_count": 1,
        "projects": [
            {"name": "某市某县业扩工程", "matched_folder": None,
             "start_page": 0, "end_page": 0, "no_folder": True, "docs": [
                {"mapping": "验收单", "page_start": 0, "page_end": 0,
                 "pages": "P0", "file_name": "某市某县业扩工程验收单.pdf"},
            ]},
        ],
    }
    written2, _ = S.execute_split(analysis2, log_undo=False)
    target2 = os.path.join(tmp, "未匹配项目", "某市某县业扩工程", "某市某县业扩工程验收单.pdf")
    check("未匹配项目写出且带前缀", len(written2) == 1 and os.path.exists(target2),
          str(written2))
finally:
    shutil.rmtree(tmp, ignore_errors=True)

print()
if failures:
    print("结果：{} 项失败 -> {}".format(len(failures), failures))
    sys.exit(1)
print("结果：全部通过")
