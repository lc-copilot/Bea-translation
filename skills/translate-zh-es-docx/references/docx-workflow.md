# DOCX 工作流（docx-workflow）

本文件规定每次文件操作必须遵循的精确 CLI 流程。所有文件操作前先读本文件。命令中的路径按实际环境替换。

## 前置条件

- 输入必须是单个可读、未加密的 `.docx`；其他格式（`.doc`、`.pdf`、`.txt`）一律拒绝。
- 含修订痕迹（tracked changes）的文档必须拒绝，不得在其上翻译。
- 目标语言固定为中文 → 中性国际西班牙语，不得改为其他语言对。

## 提取命令（Extract）

```bash
python3 skills/translate-zh-es-docx/scripts/extract_docx.py \
  SOURCE.docx --manifest manifest.json
```

- 成功时只打印 manifest 路径；失败时打印 `error:` 消息到 stderr 并退出码 2。
- manifest 是唯一允许编辑的文件；绝不直接修改源 DOCX 的 XML。

## Manifest 契约

manifest 是 UTF-8 JSON（`ensure_ascii=false`、2 空格缩进），顶层结构：

```json
{
  "schema_version": 1,
  "source": {"filename": "article.docx", "sha256": "64-hex"},
  "blocks": [
    {
      "id": "b-16-hex",
      "part": "word/document.xml",
      "paragraph_index": 0,
      "kind": "heading",
      "style_id": "Title",
      "source_text": "长白山下的足球梦",
      "translation": "",
      "notes": []
    }
  ]
}
```

- 每个可编辑文字块（标题、正文、列表、表格单元格、文本框、页眉、页脚、脚注、尾注）对应一个 block。
- block 顺序与数量不得改动；`id`、`part`、`paragraph_index`、`source_text` 都是构建与审计的绑定依据。

## 翻译编辑

- 只编辑每个 block 的 `translation` 字段（写入西语译文）与 `notes` 列表（新增解释脚注）。
- 不得编辑 `source_text` 或任何标识字段；不得增删 block。
- 每个 block 的 `translation` 必须非空；无法翻译的空段落除外（按脚本行为处理）。

## 新增注释标记（Added note markers）

- 需要新增解释脚注时，在 `translation` 中标点之后写入 `{{fn:key}}`，并在该 block 的 `notes` 中加：

```json
{"key": "li", "text": "Unidad tradicional china de longitud, equivalente hoy a unos quinientos metros."}
```

- `key` 只在当前 block 内唯一；标记必须紧跟 `.` `?` `!` `:` `;` 之后。
- 构建时 `{{fn:key}}` 会被替换为真正的 Word 脚注引用（superscript）。

## 构建命令（Build）

```bash
python3 skills/translate-zh-es-docx/scripts/build_bilingual_docx.py \
  SOURCE.docx manifest.json OUTPUT.docx
```

- 构建只写出到独立的 OUTPUT 路径，绝不覆盖源文件。
- 插入的西语段格式固定：Times New Roman、12 pt（`w:sz` 24）、1.5 倍行距（`w:spacing w:line="360"`）。
- 所有 `word/media/*` 字节必须逐字节保持不变。

## 审计命令（Audit）

```bash
python3 skills/translate-zh-es-docx/scripts/audit_bilingual_docx.py \
  SOURCE.docx manifest.json OUTPUT.docx
```

- 干净输出：`OK: N blocks, M media file`，退出码 0；有缺陷：`ERROR: <首条缺陷> (and N more)`，退出码 1。
- 加 `--json` 可输出机器可读报告（内部使用）。

## 不支持的 OOXML

- 图片像素与图表/图形文字不得 OCR 或翻译；`word/media/*` 一律原样保留。
- 检测到无法安全处理的可见可编辑图表文字时，必须停止并说明原因，不得猜测处理。

## 渲染与检查（Render and inspect）

构建后必须渲染检查：

```bash
RENDER_PY=/Users/lingchao/.codex/plugins/cache/openai-primary-runtime/documents/*/skills/documents/render_docx.py
python3 $RENDER_PY OUTPUT.docx --output_dir rendered/
```

- 必须渲染每一页（render every page），逐页检查：文字裁切/重叠、表格破损、图片移位、孤立标题、双语相邻性、脚注位置、字体替换、分页异常。
- 发现缺陷先加失败回归测试再修复，修复后重跑测试、审计、渲染并复查受影响页面。

## 最终交付

- 只交付最终 `.docx` 文件；manifest、日志、渲染产物均为内部文件，不交付。
- 交付前必须完成三遍语言质检、结构审计（退出码 0）与逐页渲染检查。
