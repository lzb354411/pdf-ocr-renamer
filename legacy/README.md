# legacy/ —— 历史遗留脚本（勿使用）

本目录存放**迁移前的热拔插脚本**，仅作历史留存，**不属于当前应用的一部分**。

## pdf_ocr_renamer.py（1031 行）

- **来源**：v2.0 时代的单文件热拔插脚本；其模块化版本已迁入 `src/core/`
  （`rules.py` / `batch.py` / `ocr.py` / `splitter.py` / `paths.py`）。
- **为何移到此处**：
  1. `_get_rules_dir()` 引用了**不存在的** `src.platform_api`（第 143 行），
     在独立运行时必然抛 `ImportError`；
  2. 缺少**整份 PDF 拆分**功能（无 `analyze_pdf` / `execute_split`），
     而拆分是当前主功能之一；
  3. 与 `src/core/` 的逻辑已经分叉，继续留在根目录会造成"两个入口"的困惑。
- **状态**：**已废弃，请勿调用**。当前应用入口是 `src/main.py`。
- **可恢复性**：git 有完整历史（初始提交 `0295120`），`git log -- legacy/` 可查。

## 注意

- 本目录**不参与**打包（见 `build/` 与 PyInstaller 配置），
  也不会被 `src/` 下任何模块导入。
- 若你确实需要该脚本的某个旧行为，请参照上表迁移到 `src/core/` 对应模块，
  而**不要**重新 import 本文件。
