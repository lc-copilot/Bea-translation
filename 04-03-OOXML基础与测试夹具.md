# 任务 04-03：OOXML 基础与测试夹具

## 状态

已完成。`ooxml.py`、`fixture_docx.py`、`test_ooxml.py` 全绿；提交见下。

## 实际结果

- `fixture_docx.py`：`make_fixture(path, tracked_changes=False)` 生成最小合法 DOCX（9 个必需成员），覆盖正文标题/段落/编号列表项/表格单元格/文本框（`w:txbxContent`）/页眉/页脚/图片/原脚注（ID 2），`tracked_changes=True` 时把“第一段。”包进 `<w:ins w:id="1" w:author="test">`。
- `ooxml.py`：实现全部绑定接口。关键决策：
  - `paragraph_text` 只收集段落直接子级 run 及 `w:ins/del/moveFrom/moveTo/hyperlink/smartTag` 容器内 run 的文本，`_run_text` 只取 run 直接子级 `w:t/w:tab/w:br`，避免文本框/绘图嵌套文本重复计数（首跑失败 2 次后修复）。
  - `iter_paragraphs` 以文档序给所有 `w:p`（含空段）编号，产出非空且不在分隔符脚注（id=-1/0）内的段落。
  - `write_parts` 按 document → 排序 headers → 排序 footers → footnotes → endnotes → 其余 的顺序写出。
  - `validate_docx` 校验扩展名、必需成员、加密、修订痕迹、chart/diagram part。
- 测试：`test_fixture_exposes_all_supported_text_and_keeps_media_hash`、`test_validation_rejects_wrong_extension_and_tracked_changes`、`test_validation_accepts_clean_fixture` 3/3 PASS。
- 测试文件顶部把 `scripts/` 加入 `sys.path`（计划运行命令不带 PYTHONPATH）。

## 提交

`git add '04-03-OOXML基础与测试夹具.md' skills/translate-zh-es-docx/scripts/ooxml.py skills/translate-zh-es-docx/tests/fixture_docx.py skills/translate-zh-es-docx/tests/test_ooxml.py && git commit -m "feat: add deterministic docx ooxml primitives"`（见 git log）。

## 任务目标

用纯 Python 标准库实现 `scripts/ooxml.py`（校验、ZIP/XML 读写、段落枚举、稳定 ID、格式化、媒体哈希）与 `tests/fixture_docx.py`（确定性夹具构建器），并让 `tests/test_ooxml.py` 通过。

## 绑定接口（Task 3–8 共用）

- `DocxError(ValueError)`；`qn`、`source_sha256`、`read_parts`、`write_parts`、
  `translatable_parts`、`has_tracked_changes`、`validate_docx`、`iter_paragraphs`、
  `paragraph_text`、`paragraph_style`、`block_kind`、`stable_block_id`、`media_hashes`。
- 常量：`W_NS`、`R_NS`、`TRACKED_TAGS`、`PRIMARY_PARTS`、`OPTIONAL_PREFIXES`、`OPTIONAL_PARTS`。
- 实现规则：保留未编辑 ZIP 成员字节；part 顺序为 document → 排序 headers → 排序 footers → footnotes → endnotes；
  忽略空段落与脚注分隔符 ID -1/0；收集所有后代 `w:p`（含表格与 `w:txbxContent`）；
  稳定 ID = `b-` + SHA-256(part + NUL + 十进制索引 + NUL + 源文本) 前 16 个 hex；
  对缺失成员、坏 ZIP、加密、修订痕迹、图表/图形文字抛 `DocxError`。

## 验收标准

- `test_ooxml.py` 全绿：夹具暴露全部 8 个文本块、媒体哈希一致；错误扩展名与修订痕迹被拒。
- `make_fixture` 生成最小合法 DOCX：正文标题/段落/编号列表/表格单元格/文本框/页眉/页脚/图片/原脚注（ID 2），可选 tracked changes。

## 实际结果

（见上方状态节）

## 提交

（见上方状态节）
