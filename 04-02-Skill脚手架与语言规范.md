# 任务 04-02：Skill 脚手架与语言规范

## 状态

已完成。包骨架、四个中文 references、中文 UI 元数据、包测试均通过；提交见下。

## 任务目标

生成可直接安装的 `translate-zh-es-docx` Skill 包骨架，编写四个聚焦的中文 references（保留计划规定的绑定短语），写 `tests/test_package.py` 固定包级不变量，并让包测试与官方校验器通过。

## 语言决定（用户要求：skill 用中文编写）

- `SKILL.md`、四个 `references/*.md` 的指导性正文全部用简体中文。
- 计划中 `tests/test_package.py` 校验的绑定短语保留原文，嵌入中文句中：
  `neutral international Spanish`、`Times New Roman`、`12 pt`、`1.5`、`five Spanish words`、
  `true Word footnote`、`surname before given name`、`motoroes`、`render every page`。
- `agents/openai.yaml` 的界面元数据用中文，`default_prompt` 保留 `$translate-zh-es-docx` 触发引用。
- 代码、命令、manifest 键名、OOXML 标签等技术术语保持原文。

## 执行事项

1. 运行 `init_skill.py` 生成骨架（记录命令与输出）。
2. 写 `tests/test_package.py`，先跑出 RED。
3. 按计划规定标题写四个中文 references 与 openai.yaml。
4. 跑包测试与 `quick_validate.py` 至 GREEN。

## 验收标准

- `python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_package.py' -v` 通过。
- `quick_validate.py skills/translate-zh-es-docx` 输出成功。
- 四个 references 全部存在且包含全部绑定短语。

## 实际结果

- `init_skill.py` 命令执行成功，生成 `skills/translate-zh-es-docx/SKILL.md`（占位，Task 7 完成）与 `agents/openai.yaml`。
- `tests/test_package.py` 首跑 RED（references 缺失），补全后 PASS。
- 四个 references 与 openai.yaml 均以简体中文编写，保留全部绑定短语；`agents/openai.yaml` 界面元数据改为中文，`default_prompt` 含 `$translate-zh-es-docx`。
- 占位 SKILL.md frontmatter 的 `description` 原为 YAML 列表导致 validator 报错，已改为中文描述字符串。
- `quick_validate.py` 依赖 `yaml`，系统与 codex 运行时的 python 均缺失；已 `pip install pyyaml` 到 codex-primary-runtime（仅工具链，未改动项目依赖），validator 输出 `Skill is valid!`。

## 提交

`git add '04-02-Skill脚手架与语言规范.md' skills/translate-zh-es-docx && git commit -m "feat: scaffold zh-es docx translation skill"`（见 git log）。
