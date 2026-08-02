---
name: translate-zh-es-docx
description: 把可编辑的中文 DOCX 出版物翻译为出版级中西双语西班牙语 Word 文档，保留中文段落、图片、表格、页眉页脚、注释与版式结构。当用户请求把中文 .docx 翻译成中性西班牙语、生成逐段配对的双语 Word 输出、保留 DOCX 格式，或执行出版级中译西翻译与质检时使用。
---

# 中西双语 DOCX 翻译器

## Outcome（成果）

交付一个可编辑、出版级的中西双语 `.docx`：每个中文块原位保留，紧接其下插入一段西语；标题、正文、列表、表格单元格、文本框、页眉、页脚、脚注、尾注全部覆盖；图片字节逐字节不变；西语统一 Times New Roman、12 pt、1.5 倍行距；超过五个西语词的文化解释使用真正的 Word 脚注。只交付最终 `.docx`。

## Required inputs and stop conditions（必需输入与停止条件）

- 输入：单个可读、未加密、无修订痕迹（tracked changes）的 `.docx`。
- 遇到 `.doc/.pdf/.txt` 等其他格式、加密文档、修订痕迹文档：立即停止并向用户说明，不做任何处理。
- 目标语言固定为中文 → 中性国际西班牙语（neutral international Spanish），不得改为其他语言对。
- 检测到无法安全处理的可见可编辑图表/图形文字：停止并说明，禁止猜测处理。

## 1. Prepare a clean working copy（准备干净工作副本）

- 把源文件复制到独立工作目录（如 `/private/tmp/zh-es-docx-work/`），此后只操作副本，绝不修改用户原件。
- 记录源副本、manifest、输出三条路径，后续命令统一使用。

## 2. Extract the manifest（提取 manifest）

```bash
python3 skills/translate-zh-es-docx/scripts/extract_docx.py \
  SOURCE.docx --manifest manifest.json
```

- 成功只打印 manifest 路径；失败打印 `error:` 并退出 2，按提示处理。
- manifest 是唯一允许编辑的文件；绝不直接改 DOCX 内部的 XML。

## 3. Classify genre and load references（分类语体并加载规范）

- 每次运行**必读** `references/translation-standard.md`。
- 按文档语体读取 `references/genre-style-guide.md` 的对应章节（叙事 / 机构 / 学术与技术 / 古典诗词与文学散文）。
- 选择出版措辞时读取 `references/sample-derived-patterns.md`。
- 每次文件操作前读取 `references/docx-workflow.md`（含精确命令与 manifest 契约）。

## 4. Research names and terminology（核实专名与术语）

- 对不确定的人名、机构名、地名，先做核实研究再下笔；优先官方西语名。
- 无通行译名的中文人名：无声调汉语拼音，姓在名前（surname before given name），不重复汉字。
- 术语全文一致，同一专名不换译法。

## 5. Translate every block（翻译每个块）

- 只编辑每个 block 的 `translation`（西语）与 `notes`（新增脚注）；不得改动 `id`/`part`/`paragraph_index`/`source_text`，不得增删块。
- 每个 block 的 `translation` 必须非空；译成中性国际西班牙语，匹配语体档案。
- 文化词解释阈值：不超过五个西语词（five Spanish words）用括注且节奏自然；否则用 `{{fn:key}}` 标记（紧跟 `.` `?` `!` `:` `;` 之后）并在该块 `notes` 中定义真正的 Word 脚注（true Word footnote），不写 `Nota del traductor` 导语。
- 原文已有脚注/尾注：保留原文内容并在其后追加西语译文。

## 6. Run three-pass linguistic QA（三遍语言质检）

1. 忠实度：逐块对照中文与西语，无遗漏、无篡改。
2. 母语润色：以西班牙语母语者视角通读，修正翻译腔。
3. 独立校对：专查性数一致、动词变位、拼写与标点（含脚注）。

三遍完成前不得进入构建。

## 7. Build the bilingual DOCX（构建双语 DOCX）

```bash
python3 skills/translate-zh-es-docx/scripts/build_bilingual_docx.py \
  SOURCE.docx manifest.json OUTPUT.docx
```

- 构建只写独立输出路径，绝不覆盖源文件；`{{fn:key}}` 会变成真正的 `w:footnoteReference`。
- 西语段落自动强制 Times New Roman、12 pt、1.5 倍行距；图片字节保持不变。

## 8. Audit structure and formatting（审计结构与格式）

```bash
python3 skills/translate-zh-es-docx/scripts/audit_bilingual_docx.py \
  SOURCE.docx manifest.json OUTPUT.docx
```

- 必须输出 `OK: N blocks, M media file` 且退出 0；有缺陷先修复再继续，禁止带缺陷交付。

## 9. Render every page and repair defects（渲染每一页并修复缺陷）

```bash
RUNTIME_PY=/Users/lingchao/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3
RENDER_PY=/Users/lingchao/.codex/plugins/cache/openai-primary-runtime/documents/*/skills/documents/render_docx.py
$RUNTIME_PY $RENDER_PY OUTPUT.docx --output_dir rendered/
```

- 必须渲染每一页（render every page），逐页检查：文字裁切/重叠、表格破损、图片移位、孤立标题、双语相邻性、脚注位置、字体替换、分页异常。
- 发现缺陷：先诊断根因（脚本行为或规则指引），再修复；修复后重跑审计、渲染并复查受影响页面。

## 10. Deliver only the DOCX（只交付 DOCX）

- 返回最终 `.docx` 的路径/链接；manifest、日志、渲染产物均为内部文件，不交付。
- 交付前必须完成：三遍语言质检、结构审计退出 0、逐页渲染检查。

## Commands（命令速查）

```bash
# 提取
python3 skills/translate-zh-es-docx/scripts/extract_docx.py SOURCE.docx --manifest manifest.json
# 构建
python3 skills/translate-zh-es-docx/scripts/build_bilingual_docx.py SOURCE.docx manifest.json OUTPUT.docx
# 审计（--json 输出机器可读报告）
python3 skills/translate-zh-es-docx/scripts/audit_bilingual_docx.py SOURCE.docx manifest.json OUTPUT.docx [--json]
# 渲染（用 documents runtime 的 python，含 soffice 与 pdf2image）
/Users/lingchao/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/lingchao/.codex/plugins/cache/openai-primary-runtime/documents/*/skills/documents/render_docx.py \
  OUTPUT.docx --output_dir rendered/
```

## Non-negotiable checks（不可妥协的检查）

- 禁止 OCR 或翻译图片像素；`word/media/*` 字节必须逐字节原样保留。
- 禁止覆盖源文件；输出必须指向不同路径。
- 禁止使用 `Nota del traductor` 之类导语；长解释必须用真正的 Word 脚注。
- 西语必须 Times New Roman、12 pt、1.5 倍行距（由构建脚本强制，构建后必须审计确认）。
- 禁止空 `translation`；每个块都必须有西语译文。
- 禁止改动 manifest 的标识字段或块数量；禁止直接编辑 DOCX XML。
