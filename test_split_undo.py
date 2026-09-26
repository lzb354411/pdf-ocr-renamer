# -*- coding: utf-8 -*-
"""拆分撤销与日志兼容性测试（阶段三 R-5）

覆盖：
1. 拆分日志写入（type="split"）与读取
2. `undo_last_split` 只删除本次新建文件（A3-1/A3-2）
3. 撤销重命名与撤销拆分互不干扰（A3-3）
4. 旧格式日志（无 type 字段）仍可撤销重命名（A3-4）
5. 文件被移动/改名时跳过并报告，不中断

**重要**：撤销日志位于 %APPDATA%\\PdfOcrRenamer\\rename_log.json，是全局共享
状态。本测试在 setUp/tearDown 中**备份并还原**该文件，避免污染用户数据与
其它套件的基线断言（见阶段文档 §5.5 环境陷阱第 1 条）。
"""
import os
import sys
import json
import shutil
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from core import paths, batch as B, splitter as S

failures = []


def check(name, cond, detail=""):
    if cond:
        print("  [PASS] {}".format(name))
    else:
        failures.append(name)
        print("  [FAIL] {}  {}".format(name, detail))


# ============================================================
# 日志文件备份/还原（防止污染全局状态）
# ============================================================
LOG = paths.get_rename_log_file()
_BACKUP = None
if os.path.exists(LOG):
    _BACKUP = LOG + ".p3test.bak"
    shutil.copy2(LOG, _BACKUP)


def restore_log():
    """把撤销日志还原到测试前状态"""
    try:
        if _BACKUP and os.path.exists(_BACKUP):
            shutil.copy2(_BACKUP, LOG)
            os.remove(_BACKUP)
        elif os.path.exists(LOG):
            os.remove(LOG)
    except Exception:
        pass


def clear_log():
    if os.path.exists(LOG):
        os.remove(LOG)


print("== 1. 拆分日志写入与读取 ==")
clear_log()
tmp = tempfile.mkdtemp()
try:
    out1 = os.path.join(tmp, "项目甲", "开工报告.pdf")
    out2 = os.path.join(tmp, "项目甲", "竣工报告.pdf")
    os.makedirs(os.path.dirname(out1), exist_ok=True)
    for p in (out1, out2):
        open(p, "w").close()

    src_pdf = os.path.join(tmp, "整份扫描.pdf")
    open(src_pdf, "w").close()

    n = B.log_split([{"dst": out1}, {"dst": out2}])
    check("log_split 记录条数", n == 2, str(n))
    check("get_last_split_count 读取正确", B.get_last_split_count() == 2,
          str(B.get_last_split_count()))
    check("拆分记录不影响重命名计数（A3-3 基础）",
          B.get_last_rename_count() == 0, str(B.get_last_rename_count()))

    data = json.load(open(LOG, encoding="utf-8"))
    recs = data.get("history", [])
    check("日志记录带 type=split", recs and recs[-1].get("type") == "split",
          str(recs[-1].get("type") if recs else None))
    check("时间戳已写入", bool(recs and recs[-1].get("time")))

    # ---------- A3-1 / A3-2 ----------
    print("== 2. 撤销上次拆分（A3-1/A3-2） ==")
    deleted, skipped = B.undo_last_split()
    check("A3-1 本次新建文件被删除", deleted == 2 and not os.path.exists(out1)
          and not os.path.exists(out2), "deleted={} skipped={}".format(deleted, skipped))
    check("A3-2 原整份 PDF 保留", os.path.exists(src_pdf))
    check("撤销后拆分计数归零", B.get_last_split_count() == 0,
          str(B.get_last_split_count()))
    check("撤销不产生错误报告", not skipped, str(skipped))

    # 目录清理：空目录应被移除，非空目录保留
    check("全空的目录被清理", not os.path.isdir(os.path.join(tmp, "项目甲")),
          str(os.listdir(tmp)))

    # ---------- 用户既有文件不被删除（A3-2 关键） ----------
    print("== 3. 绝不删除清单外文件（A3-2） ==")
    clear_log()
    mine = os.path.join(tmp, "我的重要资料.pdf")
    open(mine, "w").close()
    out3 = os.path.join(tmp, "未匹配项目", "x.pdf")
    os.makedirs(os.path.dirname(out3), exist_ok=True)
    open(out3, "w").close()
    B.log_split([{"dst": out3}])
    deleted2, _ = B.undo_last_split()
    check("清单内文件被删除", deleted2 == 1)
    check("清单外用户文件安然无恙", os.path.exists(mine))

    # ---------- 文件已被移动/改名 → 跳过并报告 ----------
    print("== 4. 文件被移动/改名时跳过不中断 ==")
    clear_log()
    gone = os.path.join(tmp, "gone.pdf")
    kept = os.path.join(tmp, "kept.pdf")
    open(kept, "w").close()
    B.log_split([{"dst": gone}, {"dst": kept}])   # gone 从未存在
    deleted3, skipped3 = B.undo_last_split()
    check("存在的文件被删除、缺失的被跳过",
          deleted3 == 1 and len(skipped3) == 1, "d={} s={}".format(deleted3, skipped3))
    check("跳过项含原因说明", skipped3 and "不存在" in skipped3[0]["info"],
          str(skipped3))

    # ---------- 无记录时不报错 ----------
    print("== 5. 无拆分记录时的行为 ==")
    clear_log()
    deleted4, skipped4 = B.undo_last_split()
    check("无记录时返回 0 且不抛异常", deleted4 == 0 and skipped4,
          "d={} s={}".format(deleted4, skipped4))
    check("无记录提示文案正确", skipped4 and "没有可撤销的拆分记录" in skipped4[0]["info"],
          str(skipped4))

    # ---------- A3-3 撤销重命名与撤销拆分互不干扰 ----------
    print("== 6. 撤销重命名 / 撤销拆分互不干扰（A3-3） ==")
    clear_log()
    # 先做一次重命名
    r_src = os.path.join(tmp, "r1.pdf")
    open(r_src, "w").close()
    fake = [{"pdf_path": r_src, "base_name": "r1.pdf",
             "mapping": "开工报告", "status": "match"}]
    renamed, _ = B.execute_renames(fake, by_folder=False)
    check("重命名已执行", len(renamed) == 1 and B.get_last_rename_count() == 1,
          "renamed={}".format(len(renamed)))
    # 再做一次拆分
    s_out = os.path.join(tmp, "splitout", "a.pdf")
    os.makedirs(os.path.dirname(s_out), exist_ok=True)
    open(s_out, "w").close()
    B.log_split([{"dst": s_out}])
    check("拆分记录未污染重命名计数", B.get_last_rename_count() == 1,
          str(B.get_last_rename_count()))
    check("拆分计数独立可读", B.get_last_split_count() == 1,
          str(B.get_last_split_count()))
    # 撤销重命名：应恢复 r1.pdf，且**不动**拆分文件
    restored, failed = B.undo_last()
    check("撤销重命名成功恢复", restored == 1 and os.path.exists(r_src),
          "restored={} failed={}".format(restored, failed))
    check("撤销重命名不影响拆分文件", os.path.exists(s_out))
    check("撤销重命名后拆分计数仍在（仍可撤销拆分）",
          B.get_last_split_count() == 1, str(B.get_last_split_count()))
    # 再撤销拆分
    deleted5, _ = B.undo_last_split()
    check("随后仍可正常撤销拆分", deleted5 == 1 and not os.path.exists(s_out),
          str(deleted5))

    # ---------- A3-4 旧格式日志 ----------
    print("== 7. 旧格式日志兼容（A3-4：无 type 字段） ==")
    old_dir = os.path.join(tmp, "old")
    os.makedirs(old_dir, exist_ok=True)
    old_new = os.path.join(old_dir, "旧报告.pdf")
    open(old_new, "w").close()
    old_orig = os.path.join(old_dir, "Scan9.pdf")
    # 手工写一份「阶段三之前」的日志：无 type 字段
    with open(LOG, "w", encoding="utf-8") as f:
        json.dump({"history": [
            {"time": "2026-01-01 00:00:00",
             "items": [{"old": old_orig, "new": old_new}]}
        ]}, f, ensure_ascii=False, indent=2)
    check("旧格式记录被识别为重命名（无 type）",
          B._record_type({"items": []}) == B.REC_RENAME)
    check("旧格式下 get_last_rename_count 可读",
          B.get_last_rename_count() == 1, str(B.get_last_rename_count()))
    check("旧格式下 get_last_split_count 为 0（不误读）",
          B.get_last_split_count() == 0, str(B.get_last_split_count()))
    r2, f2 = B.undo_last()
    check("旧格式日志仍可正常撤销重命名",
          r2 == 1 and os.path.exists(old_orig), "r={} f={}".format(r2, f2))

    # ---------- 混合日志（新旧并存） ----------
    print("== 8. 新旧格式混合日志 ==")
    clear_log()
    newf = os.path.join(tmp, "mix_new.pdf")
    open(newf, "w").close()
    with open(LOG, "w", encoding="utf-8") as f:
        json.dump({"history": [
            {"time": "t1", "items": [{"old": os.path.join(tmp, "mix_old.pdf"),
                                      "new": newf}]},          # 旧格式 rename
            {"time": "t2", "type": "split",
             "items": [{"dst": os.path.join(tmp, "nope.pdf")}]},  # 新格式 split
        ]}, f, ensure_ascii=False, indent=2)
    check("混合日志：重命名计数=1", B.get_last_rename_count() == 1,
          str(B.get_last_rename_count()))
    check("混合日志：拆分计数=1", B.get_last_split_count() == 1,
          str(B.get_last_split_count()))
    r3, _ = B.undo_last()
    check("混合日志：撤销重命名取到的是 rename 记录",
          r3 == 1 and os.path.exists(os.path.join(tmp, "mix_old.pdf")),
          str(r3))
    check("混合日志：拆分记录未被误删",
          B.get_last_split_count() == 1, str(B.get_last_split_count()))

    # ---------- 历史长度上限 ----------
    print("== 9. 日志长度上限（防无限增长） ==")
    clear_log()
    for i in range(25):
        p = os.path.join(tmp, "cap%d.pdf" % i)
        open(p, "w").close()
        B.log_split([{"dst": p}])
    data2 = json.load(open(LOG, encoding="utf-8"))
    check("历史记录被裁剪到上限 20 笔", len(data2["history"]) == 20,
          str(len(data2["history"])))

    print("== 10. execute_split 的 log_undo 开关（防污染全局日志） ==")
    # 背景：阶段三给 execute_split 加了撤销日志写入。既有套件
    # （test_split_naming / test_split_flow）会调用它，若不关掉就会把记录
    # 写进 %APPDATA%，污染用户数据与 test_core.py 的基线断言。
    clear_log()
    from pypdf import PdfWriter
    sp = os.path.join(tmp, "guard.pdf")
    w2 = PdfWriter()
    w2.add_blank_page(width=595, height=842)
    with open(sp, "wb") as f2:
        w2.write(f2)
    gsub = os.path.join(tmp, "守卫工程")
    os.makedirs(gsub, exist_ok=True)
    g_analysis = {
        "pdf_path": sp, "root": tmp, "page_count": 1,
        "projects": [{"name": "守卫工程", "matched_folder": "守卫工程",
                      "no_folder": False, "start_page": 0, "end_page": 0,
                      "docs": [{"mapping": "开工报告", "page_start": 0,
                                "page_end": 0, "pages": "P0",
                                "file_name": "守卫工程开工报告.pdf"}]}],
    }
    S.execute_split(g_analysis, log_undo=False)
    check("log_undo=False 时不写撤销日志",
          B.get_last_split_count() == 0, str(B.get_last_split_count()))
    # 默认行为仍应记账（GUI 依赖）
    S.execute_split(g_analysis, log_undo=True)
    check("log_undo=True（默认）时正常记账",
          B.get_last_split_count() == 1, str(B.get_last_split_count()))

finally:
    shutil.rmtree(tmp, ignore_errors=True)
    restore_log()

print()
# 还原校验：确保没有污染全局撤销日志
_after = B.get_last_rename_count()
print("（已还原撤销日志，当前可撤销重命名数量 = {}）".format(_after))
if failures:
    print("结果：{} 项失败 -> {}".format(len(failures), failures))
    sys.exit(1)
print("结果：全部通过")
