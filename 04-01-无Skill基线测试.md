# 任务 04-01：无 Skill 基线测试

## 状态

进行中。

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
| 叙事 | `/private/tmp/zh-es-docx-baseline/narrative.docx` | 标题、正文、表格、图片、文化词“炕” | 未生成 |
| 机构 | `/private/tmp/zh-es-docx-baseline/institutional.docx` | 标题、页眉、页脚、表格、官方机构名 | 未生成 |
| 文学 | `/private/tmp/zh-es-docx-baseline/literary.docx` | 标题、诗句、赏析、文化词“古琴”、原脚注 | 未生成 |

## 基线结果

| 场景 | 子代理 | 是否产出 DOCX | 得分 | 主要缺漏 | 状态 |
|---|---|---:|---:|---|---|
| 叙事 | 未分配 | 未运行 | — | — | 未运行 |
| 机构 | 未分配 | 未运行 | — | — | 未运行 |
| 文学 | 未分配 | 未运行 | — | — | 未运行 |

## RED 验收

- 至少一个场景低于 10 分，或未生成结构合规的 DOCX。
- 必须记录具体可观察的失败，不接受推测性结论。
- 若三个场景意外全部达到 10 分，则增加包含文本框、页眉、表格、图片和原脚注的高压夹具并重跑。

## 下一恢复动作

生成三份临时夹具，随后并行派发三个无 Skill 子代理；完成后把代理标识、产物、逐项评分和失败理由回写本文件。
