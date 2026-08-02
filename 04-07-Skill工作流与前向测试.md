# 任务 04-07：Skill 工作流与前向测试

## 状态

进行中。

## 任务目标

把占位 SKILL.md 替换为完整中文工作流，并用 fresh agent 前向测试证明其优于 Task 1 基线（三个场景 10/10，audit 退出 0）。

## 行为期望（Step 1，全部“未运行”）

每项 1 分共 10 分：完整覆盖可编辑文字；中文/西语严格逐段配对；中性国际西班牙语；正确语体；专名核实；5 词阈值；true Word footnote；图片字节保持；TNR 12pt 1.5；三遍语言质检与逐页渲染。

要求观察：代理只读相关 references；先 extract 再翻译；只编辑 manifest 的 translation/notes；不确定专名做研究；运行 build 与 audit；调用渲染；只返回 .docx 链接。

| 场景 | 输入 | 期望 | 得分 | 状态 |
|---|---|---|---|---|
| 叙事 | `/private/tmp/zh-es-docx-forward/narrative.docx` | audit 0，10/10 | — | 未运行 |
| 机构 | `/private/tmp/zh-es-docx-forward/institutional.docx` | audit 0，10/10 | — | 未运行 |
| 文学 | `/private/tmp/zh-es-docx-forward/literary.docx` | audit 0，10/10 | — | 未运行 |

## 执行步骤

1. 检查点与期望（本文件）。
2. 一个 fresh agent 用占位 SKILL.md 尝试机构任务 → 记录具体失败（RED）。
3. 替换 SKILL.md 为完整中文工作流（10 节 + Commands + Non-negotiable checks）。
4. 测试 + validator + rg 检查（只应有意的禁止文本）。
5. 三个 fresh-agent 前向测试（新夹具）→ 评分 + audit 验证。
6. 仅对证据支持的缺口修 references/脚本。
7. 回写得分并提交。

## 实际结果

- **Step 2（占位 RED）**：fresh agent 读占位 SKILL.md，确认正文仍是 TODO 模板（"Structuring This Skill" 等），无任何工作流步骤。代理绕道 `references/` 与 `scripts/` 才完成 extract→translate→build→audit（audit OK: 10 blocks），并明确指出"仅依赖 SKILL.md 的会话找不到工作流"。渲染步骤失败是因为代理误用系统 python3（缺 pdf2image）；核验发现 documents runtime 自带 pdf2image 与 soffice，`render_docx.py` 用 runtime python 可正常渲染（smoke 产物已出 page-1.png）。结论：SKILL.md 必须自带可操作工作流（Step 3 完成），渲染命令必须指明 runtime python。
- **Step 3**：完整中文 SKILL.md 已写入（10 节 + Commands + Non-negotiable checks，含必需输入/停止条件、references 加载顺序、`{{fn:key}}` 规则、逐页渲染命令）。
- **Step 4**：验证通过——包测试 22/22、`quick_validate.py` 输出 `Skill is valid!`、rg 仅命中 3 处有意的 `Nota del traductor` 禁令文本，无 TODO/TBD/placeholder。
- **Step 5（前向测试，三个 fresh agent，10 项 rubric 全部 10/10）**：
  - 叙事：`out-narrative/narrative_bilingual.docx`，audit OK（9 blocks），渲染 1 页；“土炕”用 4 词括注（≤5 阈值，合法）；叙事语体（过去未完成时）；拼音人名；图片字节保留。
  - 机构：`out-institutional/institutional_bilingual.docx`，audit OK（10 blocks），渲染 1 页；官方名 Instituto Confucio / Sede Central del Instituto Confucio 全文一致；页眉页脚表格全覆盖。
  - 文学：`out-literary/literary_zh-es.docx`，audit OK（5 blocks），渲染 1 页；“古琴”用 2 词括注 (cítara china)（源已有古琴说明段，无需脚注，判断合理）；原脚注保留中文+追加西语，编号不变。
- **用户决定（本任务中途）**：skill 不需要测试功能——`skills/translate-zh-es-docx/tests/` 整个目录已删除；SKILL.md 与 `references/docx-workflow.md` 中的“回归测试”“unittest”指引已同步清理（改为“先诊断根因再修复，重跑审计/渲染”）。
- **Step 6（仅证据支持的缺口，前向测试暴露并已修复，实际运行验证）**：
  1. `ns0:` 命名空间缺陷：源无 footnotes.xml 时，`[Content_Types].xml`/`.rels` 用 ElementTree 序列化产生 `ns0:` 前缀，LibreOffice 拒绝加载。修复：两个文件在各自序列化前临时 `ET.register_namespace("", ...)`（注意不能同时注册两个空前缀——CPython 注册会清理同前缀旧条目）。
  2. 无 footnotes 源时新增脚注被静默丢弃（`append_added_footnote` 原只在依赖 footnotes part 的循环内执行）。修复：统一移到循环后执行，`fn_created` 或 `has_notes` 时写回 footnotes.xml。
  3. 验证：无 footnotes 源 + `{{fn:kang}}` 场景 build→audit OK、CT/rels 无 ns0、footnotes.xml 含新脚注、soffice 渲染成功；三个前向场景用修复后脚本重跑 build/audit 全退出 0、渲染全部成功。

## 提交

`git add '04-07-Skill工作流与前向测试.md' skills/translate-zh-es-docx && git commit -m "feat: complete zh-es docx translator workflow (drop tests, fix footnotes & ns0 defects)"`（见 git log）。
