# Translate ZH–ES Editorial Style Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Improve direct Chinese-to-Spanish first drafts using the approved editorial boundaries while making the skill internally consistent.

**Architecture:** Keep the entrypoint short and use the existing four focused references. Update scope and output exceptions in `SKILL.md`, genre-specific choices in the style guide, meaning boundaries in translation principles, and verification in quality assurance. Do not add scripts, samples, or a DOCX workflow.

**Tech Stack:** Markdown skill files, YAML frontmatter, Git, application-scenario checks with fresh agents.

## Global Constraints

- The workflow translates complete Chinese articles directly into aligned Chinese/Spanish units. It does not review existing Spanish drafts or process DOCX files.
- Cultural travel and narrative reporting may gain modest expressive energy without changing facts, stance, speaker tone, or factual certainty.
- Standalone classical poem text remains in Chinese. Translate its title, author, introduction, modern prose explanation, and source attribution as applicable.
- Poem introductions may be moderately literary; modern prose explanations prioritize clarity and accuracy.
- Do not turn the professor's one-off factual expansion into a reusable rule.

---

### Task 1: Establish baseline behavior

**Files:** Read `skills/translate-zh-es/SKILL.md` and its four references. Create no persistent test files.

**Interfaces:** Produces observed failure notes for Tasks 2 and 3.

- [ ] **Step 1: Run a fresh-agent cultural-reportage scenario against the current skill.** Use this input, without referring to the supplied manuscripts:

```text
按现有 translate-zh-es skill 翻译这段拟刊登的文旅报道，输出逐段中西对照：
绿茵场边，孩子们迎来了属于自己的狂欢。县里的足球比赛让安静的小城热闹起来。
母亲曾严肃地说：“成绩一降，就停训。”
```

Record whether the heading image becomes a literal color term, whether the prose is natural and moderately vivid, and whether the mother's warning is overstated.

- [ ] **Step 2: Run a fresh-agent standalone-poem scenario against the current skill.** Use this input:

```text
按现有 translate-zh-es skill 处理独立古诗材料，输出逐段中西对照：
《送友人》
甲诗人
青山横北郭，白水绕东城。
题解：这首诗可能写于诗人送别友人之时，情感含蓄。
译文：青山横在城北，白水绕过城东。
（译文据某资料改写）
```

Record whether the poem body is left Chinese-only, the uncertainty in `可能` remains, and the source-attribution line is translated. Expected baseline risk: the entrypoint's every-unit rule conflicts with the poetry exception.

### Task 2: Align scope and poem handling

**Files:** Modify `skills/translate-zh-es/SKILL.md`, `skills/translate-zh-es/references/genre-style-guide.md`, and `skills/translate-zh-es/references/quality-assurance.md`.

**Interfaces:** Consumes Task 1 notes. Produces a consistent unit rule and poetry style rule for Task 3.

- [ ] **Step 1: Update the entrypoint.** Remove existing-translation review from the frontmatter description. Clarify that an independent classical poem body is retained Chinese-only, whereas surrounding content is translated. Distinguish original source attribution from an unrequested research bibliography.
- [ ] **Step 2: Update the poetry section of the style guide.** State that introductions may use measured literary language, modern prose explanations stay clear and accurate, and uncertainty words remain uncertain.
- [ ] **Step 3: Update quality assurance.** Replace the unconditional every-poetry-unit check with the standalone-poem exception. Check that source attribution is retained and translated, and factual certainty is not changed.
- [ ] **Step 4: Run the standalone-poem scenario with a fresh agent and compare its behavior with Task 1.** Expected: original poem Chinese-only; all other requested units aligned; no invented background; `可能` retained.
- [ ] **Step 5: Commit the focused poetry/scope changes.**

### Task 3: Calibrate editorial style

**Files:** Modify `skills/translate-zh-es/references/genre-style-guide.md`, `skills/translate-zh-es/references/translation-principles.md`, and `skills/translate-zh-es/references/quality-assurance.md`.

**Interfaces:** Consumes Task 1 notes and Task 2's consistent unit rule. Produces genre-sensitive first-draft guidance.

- [ ] **Step 1: Add a short cultural-travel/reportage rule.** Favor concrete scenes, idiomatic images, domain-accurate terms, and varied rhythm. Allow modest expressive emphasis while preserving factual and quoted content.
- [ ] **Step 2: Clarify the translation boundary.** Render conventional images by their function in Spanish rather than word-for-word; permit stylistic energy without adding new claims or changing certainty.
- [ ] **Step 3: Extend the three-pass check.** Check natural domain terminology and whether vivid language fits the passage's function; retain the existing grammar and punctuation checks.
- [ ] **Step 4: Run the cultural-reportage scenario with a fresh agent.** Expected: no literal color metaphor, readable energy, no exaggerated warning or added facts.
- [ ] **Step 5: Check file consistency and syntax.** Run `git diff --check` and a search across the five skill files for contradictory rules about poems, review mode, source attribution, and evaluation intensity. Commit the changes.

## Self-review

- Spec coverage: scope, reportage style, poem function, uncertainty, original attribution, and QA all map to Tasks 2–3.
- No new processing mode, sample archive, or DOCX procedure is planned.
