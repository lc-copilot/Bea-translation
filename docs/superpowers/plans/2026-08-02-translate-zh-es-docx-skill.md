# Chinese-to-Spanish DOCX Translator Skill Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a directly installable Codex skill that converts an editable Chinese `.docx` into a complete, publication-quality Chinese–Spanish bilingual `.docx` while preserving source content, images, editable structure, and paragraph pairing.

**Architecture:** Keep translation judgment in `SKILL.md` and four focused references. Use dependency-free Python OOXML utilities to extract a stable JSON manifest, rebuild a bilingual copy with true Word footnotes, and audit structure, Spanish formatting, and unchanged image bytes. Validate language behavior with no-skill baselines and fresh-agent forward tests, then validate document layout by rendering every page.

**Tech Stack:** Python 3.10+ standard library (`argparse`, `copy`, `hashlib`, `json`, `pathlib`, `re`, `unittest`, `xml.etree.ElementTree`, `zipfile`), WordprocessingML/OOXML, Codex skill metadata YAML, bundled `documents` render tooling.

## Global Constraints

- Accept one readable, unencrypted `.docx`; reject other formats and tracked-change sources before mutation.
- Translate Chinese into neutral international Spanish only; match narrative, institutional, academic/technical, or literary register automatically.
- Preserve every Chinese block and insert exactly one Spanish block immediately after it.
- Cover titles, bylines, headings, body paragraphs, lists, table cells, captions, headers, footers, footnotes, endnotes, and Word text boxes; reject visible editable chart/diagram text that cannot be handled safely.
- Never OCR or translate image pixels; preserve every `word/media/*` byte exactly.
- Format every inserted Spanish run as Times New Roman, 12 pt; format every inserted Spanish paragraph with 1.5 line spacing.
- Prefer verified official names; otherwise use toneless Hanyu Pinyin, surname before given name, without repeating Chinese characters.
- Use parenthetical cultural explanations only when no longer than five Spanish words and rhythm remains natural; otherwise use a true Word footnote without `Nota del traductor`, with its reference after punctuation.
- Never overwrite the source file; build only to a distinct output path.
- Deliver only the final `.docx`; validation logs and render artifacts remain internal.
- Before each implementation task, create its named root-level Markdown checkpoint; after the task, record artifacts, commands, results, and commit hash there.
- Do not modify the user-owned `task.md`.

---

## File map and stable interfaces

### Root recovery records

- `04-01-无Skill基线测试.md`: baseline prompts, outputs, scored failures, and commit.
- `04-02-Skill脚手架与语言规范.md`: scaffold and reference validation record.
- `04-03-OOXML基础与测试夹具.md`: package/fixture test record.
- `04-04-DOCX文本提取.md`: extraction test record.
- `04-05-双语DOCX构建.md`: paragraph, formatting, and true-footnote test record.
- `04-06-DOCX结构审计.md`: audit test record.
- `04-07-Skill工作流与前向测试.md`: skill behavior test record.
- `05-验收与文档渲染.md`: complete automated and visual validation record.
- `99-最终整合记录.md`: deduplicated final merge of all root checkpoint files.

### Skill package

- `skills/translate-zh-es-docx/SKILL.md`: trigger, routing, mandatory workflow, linguistic QA, document QA, and deliverable contract.
- `skills/translate-zh-es-docx/agents/openai.yaml`: UI metadata whose default prompt invokes `$translate-zh-es-docx`.
- `skills/translate-zh-es-docx/references/translation-standard.md`: binding language, names, punctuation, notes, numbers, quotations, and QA rules.
- `skills/translate-zh-es-docx/references/genre-style-guide.md`: genre classifier and four Spanish register profiles.
- `skills/translate-zh-es-docx/references/sample-derived-patterns.md`: positive sample tendencies and prohibited V3 errors.
- `skills/translate-zh-es-docx/references/docx-workflow.md`: exact CLI flow, manifest editing rules, unsupported structures, and render QA.
- `skills/translate-zh-es-docx/scripts/ooxml.py`: reusable validation, ZIP/XML, paragraph enumeration, stable IDs, formatting, and media hashing.
- `skills/translate-zh-es-docx/scripts/extract_docx.py`: `extract(source: Path) -> dict[str, object]` and CLI.
- `skills/translate-zh-es-docx/scripts/build_bilingual_docx.py`: `build(source: Path, manifest_path: Path, output: Path) -> None` and CLI.
- `skills/translate-zh-es-docx/scripts/audit_bilingual_docx.py`: `audit(source: Path, manifest_path: Path, output: Path) -> dict[str, object]` and CLI.
- `skills/translate-zh-es-docx/tests/fixture_docx.py`: standard-library fixture builder used only in tests.
- `skills/translate-zh-es-docx/tests/test_package.py`: package and metadata tests.
- `skills/translate-zh-es-docx/tests/test_ooxml.py`: validation, traversal, and fixture tests.
- `skills/translate-zh-es-docx/tests/test_extract.py`: manifest tests.
- `skills/translate-zh-es-docx/tests/test_build.py`: bilingual construction and footnote tests.
- `skills/translate-zh-es-docx/tests/test_audit.py`: positive and negative audit tests.

### Manifest contract

The extractor emits UTF-8 JSON with `ensure_ascii=false` and two-space indentation:

```json
{
  "schema_version": 1,
  "source": {
    "filename": "article.docx",
    "sha256": "64-lowercase-hex-characters"
  },
  "blocks": [
    {
      "id": "b-16-lowercase-hex-characters",
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

An added explanatory footnote uses one marker exactly once in `translation` and one matching note object:

```json
{
  "translation": "Recorrió quinientos li.{{fn:li}}",
  "notes": [
    {
      "key": "li",
      "text": "Unidad tradicional china de longitud, equivalente hoy a unos quinientos metros."
    }
  ]
}
```

The builder replaces `{{fn:li}}` with a real `w:footnoteReference`. Because the marker follows the period, the generated superscript also follows punctuation. Word assigns display numbers by anchor order; internal `w:id` values remain implementation identifiers.

---

### Task 1: Record failing no-skill baselines

**Files:**
- Create: `04-01-无Skill基线测试.md`

**Interfaces:**
- Consumes: requirements and sample conclusions from root checkpoints 01 and 02.
- Produces: three baseline transcripts scored against a fixed rubric; later Task 7 reruns equivalent prompts with the skill.

- [ ] **Step 1: Create the checkpoint before dispatching any baseline**

Write `04-01-无Skill基线测试.md` with status `进行中`, the three prompts below, the 10-point rubric, and an empty results table whose status cells say `未运行`.

```text
Narrative prompt: Translate a Chinese human-interest DOCX into a bilingual Spanish DOCX. Keep each Chinese paragraph and add Spanish immediately below it. Preserve an image and a table. Explain “炕” appropriately.

Institutional prompt: Translate a Confucius Institute digital-management article into neutral international Spanish. Preserve headers, footer, table text, and official names. Output only DOCX.

Literary prompt: Translate a short classical-poetry appreciation DOCX. Preserve the Chinese, reproduce imagery and cadence, explain “古琴” briefly, and keep an existing footnote.
```

Rubric, one point each: complete editable-text coverage; strict paragraph pairing; neutral Spanish; correct genre; name-verification step; parenthesis/footnote threshold; true Word footnotes; image-byte preservation; TNR 12 pt plus 1.5 spacing; three-pass linguistic and render QA.

- [ ] **Step 2: Dispatch three fresh agents without exposing the new skill**

Use one fresh agent per prompt. Give each agent only the prompt, a generated fixture path, and the output directory. Explicitly forbid reading `skills/translate-zh-es-docx/` because it does not yet exist.

- [ ] **Step 3: Score the baselines and confirm the RED state**

Record every observed omission and each rubric score in `04-01-无Skill基线测试.md`.

Expected: at least one scenario scores below 10 or fails to produce a structurally compliant document. If all score 10, replace one scenario with a harder fixture containing a header, table, text box, original footnote, and image, then rerun until a genuine gap is identified.

- [ ] **Step 4: Commit the baseline evidence**

```bash
git add '04-01-无Skill基线测试.md'
git commit -m "test: record zh-es docx baseline failures"
```

Expected: one commit containing only the baseline checkpoint.

---

### Task 2: Scaffold the installable package and codify language rules

**Files:**
- Create: `04-02-Skill脚手架与语言规范.md`
- Create: `skills/translate-zh-es-docx/agents/openai.yaml`
- Create: `skills/translate-zh-es-docx/references/translation-standard.md`
- Create: `skills/translate-zh-es-docx/references/genre-style-guide.md`
- Create: `skills/translate-zh-es-docx/references/sample-derived-patterns.md`
- Create: `skills/translate-zh-es-docx/references/docx-workflow.md`
- Create: `skills/translate-zh-es-docx/tests/test_package.py`
- Generated now but completed in Task 7: `skills/translate-zh-es-docx/SKILL.md`

**Interfaces:**
- Consumes: the confirmed translation contract and baseline gaps.
- Produces: package directories, valid UI metadata, four focused references, and a package test used by all later tasks.

- [ ] **Step 1: Create the checkpoint and scaffold with the official generator**

Record the command and status in `04-02-Skill脚手架与语言规范.md`, then run:

```bash
/Users/lingchao/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/lingchao/.codex/skills/.system/skill-creator/scripts/init_skill.py \
  translate-zh-es-docx --path skills --resources scripts,references \
  --interface 'display_name=Chinese–Spanish DOCX Translator' \
  --interface 'short_description=Create polished bilingual Spanish Word documents' \
  --interface 'default_prompt=Use $translate-zh-es-docx to translate this Chinese DOCX into a bilingual Spanish Word document.'
```

Expected: `skills/translate-zh-es-docx/` exists with `SKILL.md`, `agents/openai.yaml`, `scripts/`, and `references/`.

- [ ] **Step 2: Write the failing package test**

Create `tests/test_package.py` with these exact invariants:

```python
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]


class PackageTests(unittest.TestCase):
    def test_required_files_and_metadata(self):
        required = [
            "SKILL.md",
            "agents/openai.yaml",
            "references/translation-standard.md",
            "references/genre-style-guide.md",
            "references/sample-derived-patterns.md",
            "references/docx-workflow.md",
        ]
        for relative in required:
            self.assertTrue((ROOT / relative).is_file(), relative)
        metadata = (ROOT / "agents/openai.yaml").read_text(encoding="utf-8")
        self.assertIn('$translate-zh-es-docx', metadata)

    def test_references_encode_binding_requirements(self):
        merged = "\n".join(
            (ROOT / "references" / name).read_text(encoding="utf-8")
            for name in (
                "translation-standard.md",
                "genre-style-guide.md",
                "sample-derived-patterns.md",
                "docx-workflow.md",
            )
        )
        for phrase in (
            "neutral international Spanish",
            "Times New Roman",
            "12 pt",
            "1.5",
            "five Spanish words",
            "true Word footnote",
            "surname before given name",
            "motoroes",
            "render every page",
        ):
            self.assertIn(phrase, merged)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 3: Run the test to verify RED**

```bash
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_package.py' -v
```

Expected: FAIL because the four reference files are absent or do not yet contain the binding phrases.

- [ ] **Step 4: Write the four complete references and final metadata**

Write the requirements from the design spec under these exact headings:

```text
translation-standard.md: Quality hierarchy; Neutral Spanish; Names and proper nouns; Numbers and units; Quotations and punctuation; Cultural explanations; Original notes; Three-pass language QA
genre-style-guide.md: Classification; Narrative; Institutional; Academic and technical; Classical poetry and literary prose; Conflict resolution
sample-derived-patterns.md: Evidence status; Preferred idioms; Domain terminology; Cultural precision; Verified names; Prohibited errors; Anti-inflation rule
docx-workflow.md: Preconditions; Extract command; Manifest contract; Translation editing; Added note markers; Build command; Audit command; Unsupported OOXML; Render and inspect; Final delivery
```

Quote all scalar values in `agents/openai.yaml` and keep its default prompt exactly:

```yaml
interface:
  display_name: "Chinese–Spanish DOCX Translator"
  short_description: "Create polished bilingual Spanish Word documents"
  default_prompt: "Use $translate-zh-es-docx to translate this Chinese DOCX into a bilingual Spanish Word document."
```

- [ ] **Step 5: Run package tests and skill validator**

```bash
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_package.py' -v
python3 /Users/lingchao/.codex/skills/.system/skill-creator/scripts/quick_validate.py skills/translate-zh-es-docx
```

Expected: package test PASS; validator prints a success result. If system Python lacks `yaml`, rerun the validator with the bundled runtime used for `init_skill.py`; do not install project dependencies.

- [ ] **Step 6: Record results and commit**

```bash
git add '04-02-Skill脚手架与语言规范.md' skills/translate-zh-es-docx
git commit -m "feat: scaffold zh-es docx translation skill"
```

Expected: scaffold, metadata, references, and passing package test committed together.

---

### Task 3: Build OOXML primitives and deterministic fixtures

**Files:**
- Create: `04-03-OOXML基础与测试夹具.md`
- Create: `skills/translate-zh-es-docx/scripts/ooxml.py`
- Create: `skills/translate-zh-es-docx/tests/fixture_docx.py`
- Create: `skills/translate-zh-es-docx/tests/test_ooxml.py`

**Interfaces:**
- Consumes: Python 3.10+ standard library only.
- Produces: `DocxError`, `validate_docx`, `read_parts`, `write_parts`, `translatable_parts`, `iter_paragraphs`, `paragraph_text`, `paragraph_style`, `block_kind`, `stable_block_id`, `source_sha256`, `media_hashes`, `has_tracked_changes`, and namespace helper `qn`.

- [ ] **Step 1: Create the checkpoint and the failing OOXML tests**

The fixture builder must create a minimal valid DOCX ZIP containing body title and paragraphs, a numbered list item, table cell, text box, header, footer, image relationship, PNG bytes, original footnote, and an optional tracked-change tag. Test these behaviors:

```python
class OoxmlTests(unittest.TestCase):
    def test_fixture_exposes_all_supported_text_and_keeps_media_hash(self):
        source = self.tmp / "fixture.docx"
        make_fixture(source)
        parts = read_parts(source)
        found = [paragraph_text(p) for part in translatable_parts(parts)
                 for _, p in iter_paragraphs(parts[part])]
        self.assertEqual(
            found,
            ["测试标题", "第一段。", "列表项", "表格文字", "文本框文字",
             "页眉文字", "页脚文字", "原脚注文字"],
        )
        self.assertEqual(media_hashes(source)["word/media/image1.png"],
                         hashlib.sha256(PNG_BYTES).hexdigest())

    def test_validation_rejects_wrong_extension_and_tracked_changes(self):
        wrong = self.tmp / "fixture.zip"
        wrong.write_bytes(b"not a docx")
        with self.assertRaisesRegex(DocxError, r"\.docx"):
            validate_docx(wrong)
        tracked = self.tmp / "tracked.docx"
        make_fixture(tracked, tracked_changes=True)
        with self.assertRaisesRegex(DocxError, "tracked changes"):
            validate_docx(tracked)
```

- [ ] **Step 2: Run the tests to verify RED**

```bash
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_ooxml.py' -v
```

Expected: FAIL with import errors because `fixture_docx.py` and `ooxml.py` do not exist.

- [ ] **Step 3: Implement the fixture writer**

Use `zipfile.ZipFile(path, "w", ZIP_DEFLATED)` and literal minimal OOXML strings. Required fixture members are:

```python
REQUIRED_MEMBERS = {
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
    "word/_rels/document.xml.rels",
    "word/styles.xml",
    "word/header1.xml",
    "word/footer1.xml",
    "word/footnotes.xml",
    "word/media/image1.png",
}
```

Put the text-box paragraph inside `w:txbxContent`; assign original footnote ID `2`; use a 1×1 PNG constant; when `tracked_changes=True`, wrap `第一段。` in `<w:ins w:id="1" w:author="test"><w:r><w:t>…</w:t></w:r></w:ins>`.

- [ ] **Step 4: Implement shared OOXML functions**

Use these exact signatures and constants; the listed return types and exceptions are binding across later tasks:

```python
from collections.abc import Iterator
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
TRACKED_TAGS = {qn("w:ins"), qn("w:del"), qn("w:moveFrom"), qn("w:moveTo")}
PRIMARY_PARTS = ("word/document.xml",)
OPTIONAL_PREFIXES = ("word/header", "word/footer")
OPTIONAL_PARTS = ("word/footnotes.xml", "word/endnotes.xml")

class DocxError(ValueError):
    pass

qn(name: str) -> str
source_sha256(path: Path) -> str
read_parts(path: Path) -> dict[str, bytes]
write_parts(parts: dict[str, bytes], output: Path) -> None
translatable_parts(parts: dict[str, bytes]) -> tuple[str, ...]
has_tracked_changes(parts: dict[str, bytes]) -> bool
validate_docx(path: Path) -> dict[str, bytes]
iter_paragraphs(xml_bytes: bytes) -> Iterator[tuple[int, ET.Element]]
paragraph_text(paragraph: ET.Element) -> str
paragraph_style(paragraph: ET.Element) -> str
block_kind(part: str, paragraph: ET.Element) -> str
stable_block_id(part: str, paragraph_index: int, text: str) -> str
media_hashes(path: Path) -> dict[str, str]
```

Implementation rules: preserve ZIP member bytes unless explicitly edited; order parts as document, sorted headers, sorted footers, footnotes, endnotes; ignore empty paragraphs and footnote separator IDs `-1` and `0`; collect all descendant `w:p`, including table and `w:txbxContent`; hash IDs as `b-` plus the first 16 hex characters of SHA-256 over `part + NUL + decimal index + NUL + source text`; raise `DocxError` for missing package members, bad ZIP, encrypted/unreadable input, tracked changes, chart text, or diagram text.

- [ ] **Step 5: Run OOXML tests to verify GREEN**

```bash
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_ooxml.py' -v
```

Expected: all OOXML tests PASS.

- [ ] **Step 6: Record exact test output and commit**

```bash
git add '04-03-OOXML基础与测试夹具.md' \
  skills/translate-zh-es-docx/scripts/ooxml.py \
  skills/translate-zh-es-docx/tests/fixture_docx.py \
  skills/translate-zh-es-docx/tests/test_ooxml.py
git commit -m "feat: add deterministic docx ooxml primitives"
```

Expected: one independently passing OOXML foundation commit.

---

### Task 4: Extract every supported editable block into a stable manifest

**Files:**
- Create: `04-04-DOCX文本提取.md`
- Create: `skills/translate-zh-es-docx/scripts/extract_docx.py`
- Create: `skills/translate-zh-es-docx/tests/test_extract.py`

**Interfaces:**
- Consumes: all Task 3 functions from `ooxml.py`.
- Produces: `extract(source: Path) -> dict[str, object]`, `write_manifest(manifest: dict[str, object], output: Path) -> None`, and CLI `extract_docx.py SOURCE --manifest MANIFEST`.

- [ ] **Step 1: Create the checkpoint and failing extraction tests**

```python
class ExtractTests(unittest.TestCase):
    def test_manifest_is_complete_stable_and_ready_for_translation(self):
        source = self.tmp / "article.docx"
        make_fixture(source)
        first = extract(source)
        second = extract(source)
        self.assertEqual(first, second)
        self.assertEqual(first["schema_version"], 1)
        self.assertEqual(first["source"]["filename"], "article.docx")
        self.assertRegex(first["source"]["sha256"], r"^[0-9a-f]{64}$")
        self.assertEqual(
            [b["source_text"] for b in first["blocks"]],
            ["测试标题", "第一段。", "列表项", "表格文字", "文本框文字",
             "页眉文字", "页脚文字", "原脚注文字"],
        )
        self.assertTrue(all(b["translation"] == "" for b in first["blocks"]))
        self.assertTrue(all(b["notes"] == [] for b in first["blocks"]))
        self.assertEqual(len({b["id"] for b in first["blocks"]}), 8)

    def test_cli_refuses_to_overwrite_source(self):
        source = self.tmp / "article.docx"
        make_fixture(source)
        with self.assertRaisesRegex(DocxError, "distinct"):
            write_manifest(extract(source), source)
```

- [ ] **Step 2: Run extraction tests to verify RED**

```bash
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_extract.py' -v
```

Expected: FAIL because `extract_docx.py` cannot be imported.

- [ ] **Step 3: Implement extraction and CLI**

Build each block with exactly these keys and no XML serialization in the manifest:

```python
block = {
    "id": stable_block_id(part, paragraph_index, text),
    "part": part,
    "paragraph_index": paragraph_index,
    "kind": block_kind(part, paragraph),
    "style_id": paragraph_style(paragraph),
    "source_text": text,
    "translation": "",
    "notes": [],
}
```

Use deterministic part/paragraph order from Task 3. Write JSON with `json.dumps(manifest, ensure_ascii=False, indent=2) + "\n"`. Refuse a manifest path that resolves to the source path. Print only the created manifest path on success; print a concise `error:` message to stderr and exit 2 on `DocxError`.

- [ ] **Step 4: Run unit and CLI smoke tests**

```bash
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_extract.py' -v
PYTHONPATH='skills/translate-zh-es-docx/tests:skills/translate-zh-es-docx/scripts' \
  python3 -c 'from pathlib import Path; from fixture_docx import make_fixture; make_fixture(Path("/private/tmp/zh-es-fixture.docx"))'
python3 skills/translate-zh-es-docx/scripts/extract_docx.py \
  /private/tmp/zh-es-fixture.docx --manifest /private/tmp/zh-es-manifest.json
```

Expected: unit tests PASS; smoke command exits 0 after the test fixture is copied to the stated temporary path and prints `/private/tmp/zh-es-manifest.json`.

- [ ] **Step 5: Record results and commit**

```bash
git add '04-04-DOCX文本提取.md' \
  skills/translate-zh-es-docx/scripts/extract_docx.py \
  skills/translate-zh-es-docx/tests/test_extract.py
git commit -m "feat: extract stable docx translation manifests"
```

Expected: deterministic extractor and its tests committed.

---

### Task 5: Build bilingual DOCX paragraphs and true explanatory footnotes

**Files:**
- Create: `04-05-双语DOCX构建.md`
- Create: `skills/translate-zh-es-docx/scripts/build_bilingual_docx.py`
- Create: `skills/translate-zh-es-docx/tests/test_build.py`

**Interfaces:**
- Consumes: Task 3 OOXML functions and Task 4 manifest schema.
- Produces: `load_and_validate_manifest`, `make_translation_paragraph`, `ensure_footnotes_part`, `append_added_footnote`, `build`, and CLI `build_bilingual_docx.py SOURCE MANIFEST OUTPUT`.

- [ ] **Step 1: Create the checkpoint and failing builder tests**

Fill the fixture manifest translations deterministically, add `{{fn:li}}` after a period in one body translation, and assert:

```python
class BuildTests(unittest.TestCase):
    def test_build_pairs_blocks_formats_spanish_and_preserves_images(self):
        source, manifest_path, output = self.prepared_case(with_note=True)
        before = media_hashes(source)
        build(source, manifest_path, output)
        self.assertEqual(media_hashes(output), before)
        report = inspect_bilingual_xml(output)
        self.assertEqual(report["source_texts"], self.source_texts)
        self.assertEqual(report["spanish_texts"], self.translations_without_markers)
        self.assertTrue(all(report["paired_immediately"]))
        self.assertTrue(all(font == "Times New Roman" for font in report["fonts"]))
        self.assertTrue(all(size == "24" for size in report["sizes"]))
        self.assertTrue(all(line == "360" for line in report["line_spacing"]))
        self.assertEqual(report["added_footnote_texts"], [self.li_note])
        self.assertTrue(report["note_reference_follows_period"])

    def test_build_rejects_blank_mismatch_and_source_overwrite(self):
        source, manifest_path, output = self.prepared_case(with_note=False)
        blank_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        blank_manifest["blocks"][0]["translation"] = ""
        manifest_path.write_text(json.dumps(blank_manifest), encoding="utf-8")
        with self.assertRaisesRegex(DocxError, "blank translation"):
            build(source, manifest_path, output)
        with self.assertRaisesRegex(DocxError, "distinct"):
            build(source, manifest_path, source)
```

Add separate cases for source SHA mismatch, changed source text, missing/duplicate note marker, note marker before punctuation, and unchanged original footnote ID `2`.

- [ ] **Step 2: Run builder tests to verify RED**

```bash
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_build.py' -v
```

Expected: FAIL because the builder module does not exist.

- [ ] **Step 3: Implement manifest validation**

Validate schema version `1`, exact source SHA-256, exact block count/order/IDs/part/index/source text, nonblank Spanish, `notes` as a list, unique note keys within each block, one occurrence of each `{{fn:key}}`, no undeclared markers, nonblank note text, and marker placement immediately after `.`, `?`, `!`, `:`, or `;`. Resolve paths before comparing source and output.

- [ ] **Step 4: Implement translation paragraph construction**

Use a deep copy of the source `w:pPr`, remove copied `w:numPr` so the translation does not create a duplicate list bullet, and force spacing/run formatting:

```xml
<w:spacing w:line="360" w:lineRule="auto"/>
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"
          w:eastAsia="Times New Roman" w:cs="Times New Roman"/>
<w:sz w:val="24"/>
<w:szCs w:val="24"/>
<w:lang w:val="es-ES"/>
```

Split translation text at note markers. Emit ordinary segments as `w:r/w:t` with `xml:space="preserve"` when needed; emit a note segment as a formatted run containing `w:footnoteReference`. Add `w:bookmarkStart`/`w:bookmarkEnd` named `zes_<block-id-without-hyphen>` around Spanish content so the auditor can locate it deterministically.

- [ ] **Step 5: Implement true Word footnotes and ZIP rebuild**

When `word/footnotes.xml` is absent, create separator footnotes IDs `-1` and `0`, add its content type override, and add one document relationship of type `http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes`. Assign new positive IDs starting above every existing footnote ID. Append each note as a `w:footnote` containing one Spanish paragraph whose first run carries `w:footnoteRef`, followed by the formatted note text. If the part already exists, preserve all original footnote elements byte-equivalently at the semantic XML level and append only new ones.

For each XML part, build a parent map, identify original paragraphs by their pre-insertion ordinal, and insert translations in reverse ordinal order. Repackage every untouched ZIP member unchanged; only edited XML, content types, and the document relationship part may differ. Write to a temporary sibling and atomically replace the requested output, never the source.

- [ ] **Step 6: Run builder and regression tests**

```bash
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_build.py' -v
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_*.py' -v
```

Expected: builder tests PASS and all earlier tests remain PASS.

- [ ] **Step 7: Record results and commit**

```bash
git add '04-05-双语DOCX构建.md' \
  skills/translate-zh-es-docx/scripts/build_bilingual_docx.py \
  skills/translate-zh-es-docx/tests/test_build.py
git commit -m "feat: build formatted bilingual docx output"
```

Expected: bilingual construction, true footnotes, and regression tests committed.

---

### Task 6: Audit completeness, pairing, formatting, notes, and media

**Files:**
- Create: `04-06-DOCX结构审计.md`
- Create: `skills/translate-zh-es-docx/scripts/audit_bilingual_docx.py`
- Create: `skills/translate-zh-es-docx/tests/test_audit.py`

**Interfaces:**
- Consumes: source DOCX, completed Task 4 manifest, Task 5 output, and bookmark contract `zes_<block-id-without-hyphen>`.
- Produces: `audit(source: Path, manifest_path: Path, output: Path) -> dict[str, object]`, CLI exit `0` for clean output and `1` for defects, with `--json` for machine-readable internal use.

- [ ] **Step 1: Create the checkpoint and failing audit tests**

```python
class AuditTests(unittest.TestCase):
    def test_clean_bilingual_document_passes(self):
        source, manifest, output = self.built_case()
        report = audit(source, manifest, output)
        self.assertTrue(report["ok"])
        self.assertEqual(report["errors"], [])
        self.assertEqual(report["checked_blocks"], 8)
        self.assertEqual(report["checked_media"], 1)

    def test_detects_translation_format_and_media_damage(self):
        source, manifest, output = self.built_case()
        mutate_spanish_font(output, "Arial")
        mutate_media_byte(output)
        report = audit(source, manifest, output)
        self.assertFalse(report["ok"])
        self.assertIn("Spanish font is not Times New Roman", "\n".join(report["errors"]))
        self.assertIn("media hash mismatch", "\n".join(report["errors"]))
```

Add corruption cases for a missing bookmark, a Spanish paragraph moved away from its source, changed Chinese source text, wrong 12-point size, wrong 1.5 spacing, missing explanatory footnote, and source hash mismatch.

- [ ] **Step 2: Run audit tests to verify RED**

```bash
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_audit.py' -v
```

Expected: FAIL because `audit_bilingual_docx.py` does not exist.

- [ ] **Step 3: Implement a non-mutating auditor**

Return this stable report shape:

```python
report = {
    "ok": not errors,
    "source": str(source),
    "output": str(output),
    "checked_blocks": len(blocks),
    "checked_media": len(expected_media),
    "errors": errors,
}
```

For each manifest block, locate its source by original part/index and exact text, then require the next element sibling to be the bookmarked Spanish `w:p`. Compare its visible text after removing footnote references with `translation` after removing `{{fn:*}}`. Check every Spanish text run for all four Times New Roman font attributes, `w:sz=24`, `w:szCs=24`, and `w:lang=es-ES`; check paragraph `w:spacing` for `line=360` and `lineRule=auto`. Resolve each added reference ID to a footnote whose text equals its manifest note text. Compare source and output media maps exactly by path and hash, and verify every pre-existing image relationship ID/target remains unchanged.

- [ ] **Step 4: Implement CLI output and exit codes**

```text
Success, default: OK: 8 blocks, 1 media file
Failure, default: ERROR: <first defect> (and N more)
--json: UTF-8 JSON report with ensure_ascii=false and indent=2
Argument or DOCX error: exit 2 with concise stderr message
Audit defects: exit 1
Clean audit: exit 0
```

- [ ] **Step 5: Run audit and complete regression suite**

```bash
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_audit.py' -v
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_*.py' -v
```

Expected: all audit cases and the full suite PASS.

- [ ] **Step 6: Record results and commit**

```bash
git add '04-06-DOCX结构审计.md' \
  skills/translate-zh-es-docx/scripts/audit_bilingual_docx.py \
  skills/translate-zh-es-docx/tests/test_audit.py
git commit -m "feat: audit bilingual docx integrity"
```

Expected: non-mutating structural auditor committed.

---

### Task 7: Complete SKILL.md and prove behavior with fresh agents

**Files:**
- Create: `04-07-Skill工作流与前向测试.md`
- Modify: `skills/translate-zh-es-docx/SKILL.md`
- Modify only if a forward test exposes a real rule gap: `skills/translate-zh-es-docx/references/*.md`

**Interfaces:**
- Consumes: baseline evidence, all reference files, and the three working scripts.
- Produces: a concise trigger description, mandatory end-to-end procedure, and forward-test evidence showing improvement over Task 1.

- [ ] **Step 1: Create the checkpoint and write failing behavior expectations**

In `04-07-Skill工作流与前向测试.md`, copy the Task 1 rubric, name three new fixtures with the same difficulty classes, and set each expectation to `未运行`. Add these required observations: the agent reads only relevant references, runs extractor before translating, edits only manifest translation/note fields, performs proper-name research when uncertain, runs build and audit, invokes document rendering, and returns only a `.docx` link.

- [ ] **Step 2: Run one fresh agent against the scaffold to confirm the remaining RED state**

Give a fresh agent the institutional fixture and access to the scaffolded skill. Expected: it cannot reliably complete the workflow because `SKILL.md` is still generated placeholder content. Record the concrete failure; do not accept a purely hypothetical failure.

- [ ] **Step 3: Replace generated SKILL.md with the complete workflow**

Use this frontmatter shape and keep the body concise enough to load operationally:

```yaml
---
name: translate-zh-es-docx
description: Translate editable Chinese DOCX publication articles into polished Chinese–Spanish bilingual Word documents while preserving Chinese paragraphs, images, tables, headers, footers, notes, and layout structure. Use when a user asks to translate a Chinese .docx into neutral international Spanish, produce paragraph-paired bilingual Word output, preserve DOCX formatting, or apply publication-quality Chinese-to-Spanish translation and QA.
---
```

The body must contain these sections and imperative gates:

```text
Outcome
Required inputs and stop conditions
1. Prepare a clean working copy
2. Extract the manifest
3. Classify genre and load references
4. Research names and terminology
5. Translate every block
6. Run three-pass linguistic QA
7. Build the bilingual DOCX
8. Audit structure and formatting
9. Render every page and repair defects
10. Deliver only the DOCX
Commands
Non-negotiable checks
```

Require `references/translation-standard.md` for every run; load exactly one or more relevant genre sections, `sample-derived-patterns.md` when selecting publication wording, and `docx-workflow.md` for every file operation. State that added notes use `{{fn:key}}`, note keys are local to a block, and markers must follow punctuation. State explicitly that image OCR is prohibited.

- [ ] **Step 4: Validate package structure and wording**

```bash
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_*.py' -v
python3 /Users/lingchao/.codex/skills/.system/skill-creator/scripts/quick_validate.py skills/translate-zh-es-docx
rg -n 'T[D]O|T[B]D|placeholder|Nota del traductor|OCR image' skills/translate-zh-es-docx
```

Expected: tests and validator PASS; `rg` finds only intentional prohibition text for `Nota del traductor` and image OCR, with no generated placeholder language.

- [ ] **Step 5: Run three fresh-agent forward tests with the completed skill**

Dispatch narrative, institutional, and literary tasks using fixtures not seen in Task 1. Require each agent to use the skill. Score the same 10-point rubric and inspect produced DOCX files with `audit_bilingual_docx.py`. Expected: each audit exits 0, every scenario scores 10/10, and the literary scenario contains a short parenthesis for `古琴` while the longer unit explanation is a true Word footnote.

- [ ] **Step 6: Close only evidence-backed rule gaps**

If a forward test fails, first record the exact failure and diagnose whether the cause is instruction ambiguity, missing reference detail, or script behavior. Add the smallest explicit rule or regression test, rerun the failed scenario, and retain both before/after evidence in the checkpoint. Do not add speculative sections unrelated to observed failures.

- [ ] **Step 7: Record final scores and commit**

```bash
git add '04-07-Skill工作流与前向测试.md' skills/translate-zh-es-docx
git commit -m "feat: complete zh-es docx translator workflow"
```

Expected: complete installable skill and its passing behavior evidence committed.

---

### Task 8: Run real-sample acceptance, render every page, and integrate records

**Files:**
- Create: `05-验收与文档渲染.md`
- Create: `99-最终整合记录.md`
- Modify: `00-工作恢复索引.md`
- Modify only for evidenced defects: files under `skills/translate-zh-es-docx/`
- Temporary only, never commit: `/private/tmp/zh-es-docx-acceptance/`

**Interfaces:**
- Consumes: completed installable package and one source sample copied from the user-provided corpus.
- Produces: final automated/visual evidence, a deduplicated recovery record, and the reviewed skill package. It does not keep or deliver a sample translation as part of the package.

- [ ] **Step 1: Create the acceptance checkpoint before touching a real sample**

Write `05-验收与文档渲染.md` with source sample path `/Users/lingchao/Desktop/C26 长白山下的足球梦(1).docx`, temporary workspace, commands to run, expected artifacts, page-inspection table, and status `进行中`. Never edit the original Desktop or WeChat file.

```bash
mkdir -p /private/tmp/zh-es-docx-acceptance
cp '/Users/lingchao/Desktop/C26 长白山下的足球梦(1).docx' \
  /private/tmp/zh-es-docx-acceptance/source.docx
```

- [ ] **Step 2: Run full automated validation from repository root**

```bash
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_*.py' -v
python3 /Users/lingchao/.codex/skills/.system/skill-creator/scripts/quick_validate.py skills/translate-zh-es-docx
git diff --check
```

Expected: all tests PASS, skill validator succeeds, and `git diff --check` produces no output.

- [ ] **Step 3: Execute extract, translate, build, and audit on the copied sample**

```bash
python3 skills/translate-zh-es-docx/scripts/extract_docx.py \
  /private/tmp/zh-es-docx-acceptance/source.docx \
  --manifest /private/tmp/zh-es-docx-acceptance/manifest.json
python3 skills/translate-zh-es-docx/scripts/build_bilingual_docx.py \
  /private/tmp/zh-es-docx-acceptance/source.docx \
  /private/tmp/zh-es-docx-acceptance/manifest.json \
  /private/tmp/zh-es-docx-acceptance/bilingual.docx
python3 skills/translate-zh-es-docx/scripts/audit_bilingual_docx.py \
  /private/tmp/zh-es-docx-acceptance/source.docx \
  /private/tmp/zh-es-docx-acceptance/manifest.json \
  /private/tmp/zh-es-docx-acceptance/bilingual.docx --json
```

Before the build command, translate every manifest block according to the skill and complete all three language passes. Expected: no blank translations, build exits 0, audit JSON has `"ok": true`.

- [ ] **Step 4: Render every page using the documents runtime**

```bash
/Users/lingchao/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/lingchao/.codex/plugins/cache/openai-primary-runtime/documents/26.801.11242/skills/documents/render_docx.py \
  /private/tmp/zh-es-docx-acceptance/bilingual.docx \
  --output_dir /private/tmp/zh-es-docx-acceptance/rendered
```

Expected: one PNG per page and no renderer error.

- [ ] **Step 5: Inspect 100% of pages and repair any defect test-first**

Record every page number as `通过` or describe the defect. Check clipped/overlapping text, broken tables, displaced images, orphan headings, incorrect bilingual adjacency, footnote placement, font substitution, and unreadable page breaks. For a mechanical defect, add a failing regression test before changing a script; for an instruction defect, add a failing fresh-agent scenario before editing prose. Rerun tests, audit, render, and affected page inspection after each fix.

- [ ] **Step 6: Merge all root checkpoints into the final record**

Create `99-最终整合记录.md` with these deduplicated sections:

```text
Final outcome
User requirements
Sample-derived translation standard
Approved architecture
Implementation task results and commit hashes
Automated test evidence
Fresh-agent behavior evidence
Real-sample structural and visual evidence
Installable package contents
Known scope boundaries
Recovery and maintenance instructions
```

Do not concatenate repeated headings mechanically. Preserve exact commands and final results, but collapse repeated requirements to one authoritative statement. Update every row in `00-工作恢复索引.md` to `已完成` and set current recovery point to the final commit.

- [ ] **Step 7: Run final verification and commit only durable artifacts**

```bash
python3 -m unittest discover -s skills/translate-zh-es-docx/tests -p 'test_*.py' -v
python3 /Users/lingchao/.codex/skills/.system/skill-creator/scripts/quick_validate.py skills/translate-zh-es-docx
git diff --check
git status --short
git add '00-工作恢复索引.md' '05-验收与文档渲染.md' '99-最终整合记录.md' skills/translate-zh-es-docx
git commit -m "test: verify installable zh-es docx translator skill"
```

Expected: tests and validator PASS, no whitespace errors, `task.md` remains untracked and unmodified, temporary sample/render files are absent from Git, and the final commit contains only durable package or checkpoint changes.

---

## Plan self-review checklist

- [x] Spec coverage: every design section—scope, language, names, notes, samples, package structure, flow, error handling, testing, and installation—maps to at least one task above.
- [x] Placeholder scan: `rg -n 'T[B]D|T[D]O|implement l[a]ter|fill in d[e]tails|Similar to T[a]sk|Add appropriate error h[a]ndling|Write tests for the a[b]ove' docs/superpowers/plans/2026-08-02-translate-zh-es-docx-skill.md` returned no matches.
- [x] Interface consistency: all three CLIs, manifest keys, function signatures, bookmark names, note-marker syntax, format values, and exit codes agree across Tasks 3–8.
- [x] Recovery consistency: every implementation task names a root-level checkpoint that is written before action and updated after verification.
- [x] Scope safety: no task modifies `task.md`, original samples, image bytes, or a source DOCX.
