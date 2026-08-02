# 任务 04-01：无 Skill 基线测试

## 状态

已完成。三份基线已记录，RED 状态确认。

## 任务目标

在 `translate-zh-es-docx` Skill 尚未创建的前提下，让三个独立子代理分别处理叙事、机构和文学类中文 DOCX，记录其真实缺漏，形成后续 Skill 行为测试的 RED 基线。

## 前置约束

- 子代理不得读取尚未存在的 `skills/translate-zh-es-docx/`。
- 三份基线夹具只写入 `/private/tmp/zh-es-docx-baseline/`，不进入 Git。
- 本任务只提交本检查点，不实现任何 Skill 文件。
- 每个子代理只获得自己的任务说明、输入文件和输出目录。

## 基线场景

### 叙事场景

Translate a Chinese human-interest DOCX into a bilingual Spanish DOCX. Keep each Chinese paragraph and add Spanish immediately below it. Preserve an image and a table. Explain “炕” appropriately.

### 机构场景

Translate a Confucius Institute digital-management article into neutral international Spanish. Preserve headers, footer, table text, and official names. Output only DOCX.

### 文学场景

Translate a short classical-poetry appreciation DOCX. Preserve the Chinese, reproduce imagery and cadence, explain “古琴” briefly, and keep an existing footnote.

## 评分标准

每项 1 分，共 10 分：

1. 完整覆盖所有可编辑文字。
2. 中文与西语严格逐段相邻配对。
3. 使用中性国际西班牙语。
4. 正确匹配文章语体。
5. 对不确定人名或专名执行核实。
6. 正确执行“五个西语词以内括注，否则脚注”的阈值。
7. 长注解使用真正的 Word 脚注。
8. 图片字节保持不变。
9. 西语使用 Times New Roman、12 磅、1.5 倍行距。
10. 包含忠实度、母语润色、独立校对和逐页渲染检查。

## 夹具计划

| 场景 | 输入 | 结构要求 | 状态 |
|---|---|---|---|
| 叙事 | `/private/tmp/zh-es-docx-baseline/narrative.docx` | 标题、正文、表格、图片、文化词“炕” | 已生成 |
| 机构 | `/private/tmp/zh-es-docx-baseline/institutional.docx` | 标题、页眉、页脚、表格、官方机构名 | 已生成 |
| 文学 | `/private/tmp/zh-es-docx-baseline/literary.docx` | 标题、诗句、赏析、文化词“古琴”、原脚注 | 已生成 |

夹具由 `/private/tmp/zh-es-docx-baseline/make_fixtures.py`（Python 标准库 `zipfile`）生成，均通过 `zipfile.testzip()` 完整性校验。

## 基线结果

| 场景 | 子代理 | 是否产出 DOCX | 得分 | 主要缺漏 | 状态 |
|---|---|---:|---:|---|---|
| 叙事 | `task-1` | 是 `out-narrative/narrative_zh_es.docx` | 6/10 | “炕”用正文长括注（超 5 词阈值）未用真脚注；未设 TNR 12pt 1.5 格式；未做渲染 QA | 已完成 |
| 机构 | `task-2` | 是 `out-institutional/institutional.docx` | 7/10 | 直接把中文替换为西语，未保留中文、未逐段配对；未设格式；未做渲染 QA | 已完成 |
| 文学 | `task-3` | 是 `out-literary/literary_zh-es.docx` | 6/10 | “古琴”用长括注（含汉字）超阈值未用脚注；未设格式；未做渲染 QA | 已完成 |

### 逐项评分依据（独立核验产物 XML，非仅采信代理报告）

- 叙事：1✓ 2✓ 3✓ 4✓ 5✓ 6✗（“kang (炕)…”解释远长于五个西语词且置于正文） 7✗（无新增脚注） 8✓（图片字节保留） 9✗（无 `w:rFonts`/`w:sz`/`w:spacing` 显式格式，继承默认 Calibri） 10✗（未渲染）＝ 6/10。
- 机构：1✓ 2✗（`document.xml` 中 `w:t` 全部被替换为西语，全文档无中文字符，双语配对完全缺失） 3✓ 4✓ 5✓ 6✓（无文化词场景，阈值未触发） 7✓（未触发） 8✓ 9✗ 10✗ ＝ 7/10。
- 文学：1✓ 2✓ 3✓ 4✓ 5✓ 6✗（“El guqin (古琴, “qin antiguo”)…”为超阈值长括注且内嵌汉字） 7✗（仅保留原脚注，未为“古琴”新增真脚注） 8✓（原脚注 id=2 与图片保留） 9✗ 10✗ ＝ 6/10。

## RED 验收

- 三个场景均低于 10 分；机构场景核心失败（未保留中文、未逐段配对）；叙事/文学场景违反“五个西语词以内括注、否则脚注”阈值且未使用真正的 Word 脚注；三个场景均未按要求设置西语格式（Times New Roman、12 pt、1.5 倍行距），也未做任何渲染检查。
- 结论：RED 状态成立，Skill 具备可证明的改进空间。

## RED 验收

- 至少一个场景低于 10 分，或未生成结构合规的 DOCX。
- 必须记录具体可观察的失败，不接受推测性结论。
- 若三个场景意外全部达到 10 分，则增加包含文本框、页眉、表格、图片和原脚注的高压夹具并重跑。

## 下一恢复动作

进入 `04-02-Skill脚手架与语言规范.md`：运行 `init_skill.py` 生成包骨架，编写四个中文 references（保留计划规定的绑定短语），写 `tests/test_package.py` 并提交。
