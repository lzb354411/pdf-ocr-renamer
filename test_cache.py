# -*- coding: utf-8 -*-
"""页级 OCR 缓存测试（阶段四 R-7 / A4-2、A4-3、A4-4）

覆盖：
1. 缓存键：覆盖文件身份 + 页码 + 渲染参数 + 算法参数；相关输入变化必换键
2. 存取与还原：rows 结构、mapping、angle 完整往返
3. A4-3 失效：文件大小/mtime 变化 → 键变化 → 旧条目不再命中
4. A4-4 容量：条目上限淘汰（LRU）、清空、损坏条目自愈
5. 开关与降级：禁用缓存时行为等价于无缓存；异常不抛出
6. A4-2（结构级）：_classify_page_cached 命中时**不再调用 OCR**

**不依赖真实 PDF/OCR 引擎**：用假引擎与合成图像验证逻辑，
保证在无 GPU/无模型环境下也能跑通（真实样本回归见 §5.2）。
"""
import os
import sys
import json
import shutil
import tempfile
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from core import cache as C
from core import paths

failures = []


def check(name, cond, detail=""):
    if cond:
        print("  [PASS] {}".format(name))
    else:
        failures.append(name)
        print("  [FAIL] {}  {}".format(name, detail))


print("== 1. 缓存键：覆盖全部影响结果的输入 ==")
tmp = tempfile.mkdtemp()
try:
    pdf = os.path.join(tmp, "样本.pdf")
    with open(pdf, "wb") as f:
        f.write(b"%PDF-1.4 fake" + b"x" * 200)
    st = os.stat(pdf)

    base = C.make_page_key(pdf, 0, 1.0, extra={"a": 1}, file_stat=st)
    check("同输入生成同键（确定性）",
          base == C.make_page_key(pdf, 0, 1.0, extra={"a": 1}, file_stat=st))
    check("页码变化 → 换键",
          base != C.make_page_key(pdf, 1, 1.0, extra={"a": 1}, file_stat=st))
    check("渲染参数变化 → 换键",
          base != C.make_page_key(pdf, 0, 0.5, extra={"a": 1}, file_stat=st))
    check("算法参数(extra)变化 → 换键",
          base != C.make_page_key(pdf, 0, 1.0, extra={"a": 2}, file_stat=st))
    check("文件路径变化 → 换键",
          base != C.make_page_key(os.path.join(tmp, "别的.pdf"), 0, 1.0,
                                  extra={"a": 1}, file_stat=st))
    check("scale 1 与 1.0 视为同一参数（不产生重复条目）",
          C.make_page_key(pdf, 0, 1) == C.make_page_key(pdf, 0, 1.0))
    check("extra 字典键序无关（稳定序列化）",
          C.make_page_key(pdf, 0, 1.0, extra={"a": 1, "b": 2}) ==
          C.make_page_key(pdf, 0, 1.0, extra={"b": 2, "a": 1}))
    check("空路径返回 None（调用方据此跳过缓存）",
          C.make_page_key("", 0, 1.0) is None)

    # ---------- A4-3：文件修改 → 键变化 ----------
    print("== 2. A4-3 文件修改后缓存自动失效 ==")
    k_before = C.make_page_key(pdf, 0, 1.0, extra={"a": 1})
    C.enable_cache(True)
    C.clear_cache()
    check("修改前已写入缓存", C.save_page(k_before, "开工报告",
                                          [(1.0, "甲")], 0))
    check("修改前可命中", C.load_page(k_before) is not None)
    # 修改文件内容（大小 + mtime 均变）
    time.sleep(0.01)
    with open(pdf, "ab") as f:
        f.write(b"CHANGED")
    k_after = C.make_page_key(pdf, 0, 1.0, extra={"a": 1})
    check("文件修改后缓存键变化", k_before != k_after,
          "before={} after={}".format(k_before, k_after))
    check("旧键条目不再被新键命中", C.load_page(k_after) is None)

    # 仅 mtime 变化（内容不变）也应失效
    k_m1 = C.make_page_key(pdf, 0, 1.0, extra={"a": 1})
    time.sleep(0.01)
    os.utime(pdf, None)
    k_m2 = C.make_page_key(pdf, 0, 1.0, extra={"a": 1})
    check("仅 mtime 变化也应换键（保守失效）", k_m1 != k_m2)

    # ---------- 3. 存取往返 ----------
    print("== 3. 存取往返：mapping / rows / angle 完整还原 ==")
    C.clear_cache()
    k = C.make_page_key(pdf, 0, 1.0, extra={"a": 1})
    rows_in = [(12.0, "配电网工程开工报告"), (48.5, "工程名称某工程"), (99, "B1109")]
    check("保存成功", C.save_page(k, "开工报告", rows_in, 90))
    e = C.load_page(k)
    check("读取成功", e is not None)
    check("mapping 还原", e.get("mapping") == "开工报告", str(e.get("mapping")))
    check("angle 还原", e.get("angle") == 90, str(e.get("angle")))
    rows_out = C.rows_from_entry(e)
    check("rows 数量与顺序还原", len(rows_out) == 3, str(rows_out))
    check("rows 内容还原（y 为 float，text 完整）",
          rows_out == [(12.0, "配电网工程开工报告"), (48.5, "工程名称某工程"),
                       (99.0, "B1109")], str(rows_out))
    # mapping 为 None 也应能缓存（"未命中"是合法结果）
    k_none = C.make_page_key(pdf, 2, 1.0, extra={"a": 1})
    C.save_page(k_none, None, [], 0)
    e2 = C.load_page(k_none)
    check("mapping=None 可缓存并还原（未命中也是结果）",
          e2 is not None and e2.get("mapping") is None
          and C.rows_from_entry(e2) == [], str(e2))

    # ---------- 4. 损坏条目自愈 ----------
    print("== 4. 损坏缓存自愈（不得抛出） ==")
    k_bad = C.make_page_key(pdf, 3, 1.0, extra={"a": 1})
    p_bad = os.path.join(C.get_cache_dir(), k_bad + ".json")
    with open(p_bad, "w", encoding="utf-8") as f:
        f.write("{ 这不是合法 JSON ")
    got = C.load_page(k_bad)
    check("损坏条目返回 None 而非抛异常", got is None)
    check("损坏条目被清理（避免反复失败）", not os.path.exists(p_bad))
    # 版本不匹配的旧缓存应失效
    k_old = C.make_page_key(pdf, 4, 1.0, extra={"a": 1})
    p_old = os.path.join(C.get_cache_dir(), k_old + ".json")
    with open(p_old, "w", encoding="utf-8") as f:
        json.dump({"v": 0, "mapping": "过时"}, f, ensure_ascii=False)
    check("旧版本缓存条目失效（v 不匹配）", C.load_page(k_old) is None)

    # ---------- 5. 开关与降级 ----------
    print("== 5. 开关与降级 ==")
    C.clear_cache()
    k5 = C.make_page_key(pdf, 5, 1.0, extra={"a": 1})
    C.enable_cache(False)
    check("禁用缓存时不读取（恒 miss）", C.load_page(k5) is None)
    check("禁用缓存时不写入", C.save_page(k5, "甲", [], 0) is False)
    C.enable_cache(True)
    check("重新启用后可写入", C.save_page(k5, "甲", [], 0) is True)
    check("重新启用后可命中", C.load_page(k5) is not None)
    C.reset_stats()
    C.load_page("不存在的键")
    check("统计可记录 miss", C.get_stats()["miss"] >= 1, str(C.get_stats()))

    # ---------- 6. A4-4 容量与清理 ----------
    print("== 6. A4-4 容量上限与清理 ==")
    C.clear_cache()
    for i in range(12):
        kk = C.make_page_key(pdf, 100 + i, 1.0, extra={"a": 1})
        C.save_page(kk, "报告%d" % i, [(1.0, "x")], 0)
        time.sleep(0.005)   # 保证 mtime 可区分，便于 LRU 判定
    n_before, _ = C.cache_size()
    check("已写入 12 条", n_before == 12, str(n_before))
    removed = C.enforce_limit(max_entries=5)
    n_after, _ = C.cache_size()
    check("超出上限时淘汰到上限", n_after == 5, "removed={} now={}".format(
        removed, n_after))
    check("淘汰返回被删数量", removed == 7, str(removed))
    check("未超上限时不淘汰", C.enforce_limit(max_entries=5) == 0)
    # 清空
    n_cleared, freed = C.clear_cache()
    n_final, size_final = C.cache_size()
    check("清空缓存后条目数为 0", n_final == 0, str(n_final))
    check("清空返回删除数与释放字节", n_cleared == 5 and freed > 0,
          "n={} freed={}".format(n_cleared, freed))
    check("清空后目录仍可继续使用",
          C.save_page(C.make_page_key(pdf, 200, 1.0), "甲", [], 0) is True)

    # ---------- 7. A4-2 结构级：命中缓存不得调用 OCR ----------
    print("== 7. A4-2 结构级：缓存命中时不再调用 OCR ==")
    from core import splitter as S

    class _FakeEngine:
        pass

    calls = {"n": 0}

    def _fake_ocr(engine, pil_image):
        calls["n"] += 1
        return [(((0, 0), (10, 10), (20, 20), (0, 20)), "配电网工程开工报告")]

    real_ocr = S._ocr_image
    real_build = S._build_rows
    real_orient = S._detect_orientation
    try:
        S._ocr_image = _fake_ocr
        S._build_rows = lambda result, limit=60: [(1.0, "配电网工程开工报告")]
        S._detect_orientation = lambda img: 0
        C.clear_cache()
        C.enable_cache(True)
        S.set_rules_signature([{"识别内容": "配电网工程开工报告",
                                "映射内容": "开工报告"}])

        class _FakeImg:
            size = (100, 100)

            def crop(self, box):
                return self

            def rotate(self, angle, expand=True):
                return self

        img = _FakeImg()
        rules = [{"识别内容": "配电网工程开工报告", "映射内容": "开工报告"}]

        calls["n"] = 0
        k = C.make_page_key(pdf, 300, 1.0, extra=S._cache_extra_signature())
        out1 = S._classify_page_cached(img, rules, _FakeEngine(), cache_key=k)
        n_first = calls["n"]
        check("首次未命中确实调用了 OCR", n_first > 0, str(n_first))
        check("首次返回 mapping=开工报告", out1[0] == "开工报告", str(out1[0]))

        calls["n"] = 0
        out2 = S._classify_page_cached(img, rules, _FakeEngine(), cache_key=k)
        check("二次命中时 OCR 调用次数为 0", calls["n"] == 0, str(calls["n"]))
        check("命中结果与首次完全一致", out1 == out2,
              "{} vs {}".format(out1, out2))

        # 无 cache_key（cache_key=None）→ 恒走真实 OCR
        calls["n"] = 0
        S._classify_page_cached(img, rules, _FakeEngine(), cache_key=None)
        check("cache_key=None 时不走缓存（仍调 OCR）", calls["n"] > 0,
              str(calls["n"]))
    finally:
        S._ocr_image = real_ocr
        S._build_rows = real_build
        S._detect_orientation = real_orient
        C.clear_cache()

    # ---------- 8. 预筛（阶段四 4.3）不改变有文字页结果 ----------
    print("== 8. 4.3 预筛保守性：有文字页结果不受影响 ==")
    try:
        S._ocr_image = _fake_ocr
        S._build_rows = lambda result, limit=60: [(1.0, "配电网工程开工报告")]
        S._detect_orientation = lambda img: 0
        C.enable_cache(False)     # 关缓存，纯看预筛影响
        rules = [{"识别内容": "配电网工程开工报告", "映射内容": "开工报告"}]

        class _Img2:
            size = (100, 100)

            def crop(self, box):
                return self

            def rotate(self, angle, expand=True):
                return self

        r_nopre = S._classify_page_cached(_Img2(), rules, _FakeEngine(),
                                          cache_key=None, prescan=False)
        r_pre = S._classify_page_cached(_Img2(), rules, _FakeEngine(),
                                        cache_key=None, prescan=True)
        check("预筛开启/关闭结果一致（有文字页）", r_nopre == r_pre,
              "{} vs {}".format(r_nopre, r_pre))

        # 无文字的页：预筛应短路为 (None, [], 0)
        S._build_rows = lambda result, limit=60: []
        r_empty = S._classify_page_cached(_Img2(), rules, _FakeEngine(),
                                          cache_key=None, prescan=True)
        check("预筛对无文字页短路为 (None, [], 0)",
              r_empty == (None, [], 0), str(r_empty))
    finally:
        S._ocr_image = real_ocr
        S._build_rows = real_build
        S._detect_orientation = real_orient
        C.enable_cache(True)
        C.clear_cache()

    # ---------- 9. 路径与目录 ----------
    print("== 9. 缓存位置与目录管理 ==")
    d = C.get_cache_dir()
    check("缓存位于 %APPDATA% 下 PdfOcrRenamer\\cache 中",
          d.lower().startswith(paths.get_app_dir().lower())
          and "cache" in d.lower(), d)
    check("目录自动创建", os.path.isdir(d))
    C.clear_all()
    check("clear_all 可彻底删除 cache 目录",
          not os.path.isdir(os.path.join(paths.get_app_dir(), "cache")))
    check("删除后再次访问可自动重建", os.path.isdir(C.get_cache_dir()))

finally:
    C.clear_cache()
    C.enable_cache(True)
    shutil.rmtree(tmp, ignore_errors=True)

print()
if failures:
    print("结果：{} 项失败 -> {}".format(len(failures), failures))
    sys.exit(1)
print("结果：全部通过")
