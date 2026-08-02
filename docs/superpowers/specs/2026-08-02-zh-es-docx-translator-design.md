# Chinese-to-Spanish DOCX Translator Skill Design

## Goal

Build a directly installable Codex skill that translates Chinese `.docx` publication articles into polished, neutral international Spanish while preserving the source document, images, editable structure, and paragraph-level bilingual alignment.

## Chosen approach

Use a hybrid skill package:

- Keep linguistic judgment in concise skill instructions and focused references.
- Use deterministic Python/OOXML scripts for extracting editable text, producing a translation manifest, inserting Spanish paragraphs, preserving document relationships, and auditing the result.
- Require the existing `documents` skill for render-based visual verification.

This is preferred over a prompt-only skill because DOCX traversal, insertion, footnotes, headers, tables, text boxes, and validation are fragile mechanical operations. It is preferred over a standalone translation service because the language model must retain genre-sensitive judgment, research proper names, and adapt style to each article.

## Alternatives considered

1. **Prompt-only skill:** smallest package, but unreliable for full-document coverage and format preservation.
2. **Hybrid instructions plus deterministic DOCX tooling:** selected; separates linguistic judgment from repeatable document manipulation.
3. **External translation pipeline or service:** potentially scalable but unnecessary, harder to install, and less adaptable to literary and publication prose.

## Scope

### Inputs

- Accept one editable `.docx` file containing primarily Chinese publication content.
- Reject non-DOCX inputs with a concise explanation.
- Treat image pixels as opaque. Do not OCR or translate text embedded inside images.
- Translate all visible editable text: titles, bylines, headings, body paragraphs, list items, table cells, captions, footnotes, endnotes, headers, footers, and editable drawing/text-box text when supported by OOXML.

### Output

- Produce one bilingual `.docx` file.
- Keep every Chinese source paragraph in place.
- Insert its Spanish translation as the immediately following paragraph or equivalent block.
- Keep source images unchanged and preserve their relationships and relative ordering.
- Apply Times New Roman, Chinese small-four size (12 pt), and 1.5 line spacing to inserted Spanish text.
- Permit natural repagination, paragraph reflow, and image movement. Translation quality takes priority over matching the original page count.
- Deliver only the bilingual Word document; do not create a separate QA report.

## Translation contract

### Language and register

- Translate only Chinese into Spanish.
- Use neutral international Spanish without conspicuous regionalisms.
- Infer the genre and match its register:
  - human-interest and narrative articles: vivid, readable, and natural;
  - institutional reports: formal, concise, and publication-ready;
  - academic or technical prose: rigorous and terminologically precise;
  - classical poetry and literary passages: accurate while recreating imagery, cadence, rhetoric, and rhyme where feasible.
- Interpret unclear source sentences from context and complete the translation without pausing for confirmation.

### Quality hierarchy

Apply this order when goals conflict:

1. Complete and accurate meaning with no omissions or unsupported additions.
2. Correct Spanish grammar, spelling, punctuation, agreement, and syntax.
3. Native, idiomatic collocations and domain terminology.
4. Genre-appropriate publication style and coherent discourse.
5. Elevated vocabulary and rhetoric only when the source register supports them.

“Advanced” Spanish means precise vocabulary, idiomatic combinations, varied syntax, and smooth cohesion. It does not mean rare vocabulary, inflated formality, or embellishment absent from the source.

### Names and terminology

- Prefer verified official Spanish names for institutions, places, events, products, titles, and specialized terms.
- Prefer a person's verified official Latin-script name.
- If no official form exists, use Hanyu Pinyin without tone marks and preserve surname-before-given-name order.
- Do not repeat Chinese characters after a romanized personal name.
- Research uncertain proper names using authoritative official sources first, then reputable Spanish-language sources and corpora.
- If no definitive form exists, choose the most reliable translation and continue without interrupting the task.

### Explanations and notes

- Preserve and translate every original note.
- Add a short parenthetical explanation when it is no more than five Spanish words and does not disturb the reading rhythm, for example `guqin (cítara china)`.
- Use a true Word footnote when an explanation exceeds five Spanish words or interrupts the sentence.
- Do not label added notes `Nota del traductor`.
- Number all footnotes consecutively by order of appearance. When an explanatory note is inserted between original notes, renumber later notes automatically while preserving their anchors and content.
- Place superscript note markers consistently after punctuation.

## Sample-derived editorial guidance

Treat the reviewed `V3` documents as preference evidence, not infallible ground truth.

Positive patterns to generalize:

- Replace literal calques with idiomatic expressions such as `fiesta sobre el césped` rather than `fiesta verde`.
- Prefer domain-native terms such as `cantera femenina`, `pago de tasas`, `confirmación de matrícula`, and `herramientas de evaluación en línea`.
- Preserve semantic nuance: `humilde gratitud` is preferable to a negatively connoted `débil gratitud` in a filial-poetry context.
- Use culturally precise forms such as `guqin (cítara china)` and explain titles or appellations when necessary for comprehension.
- Preserve verified personal names such as `Andreas Janousch` rather than replacing them mechanically with pinyin.

Negative patterns to reject:

- Typos and agreement errors introduced during editing, including forms like `motoroes`, `del información`, and `optimización refinida`.
- Missing contractions such as `a el` where `al` is required.
- Punctuation artifacts such as duplicated colon-period sequences.
- Artificial elevation that weakens dialogue, overstates a modest source expression, or replaces a natural collocation with bureaucratic wording.

## Package structure

Create the package as `translate-zh-es-docx/` with:

- `SKILL.md`: trigger description, required workflow, routing to references, output contract, and stop conditions.
- `agents/openai.yaml`: generated UI metadata.
- `references/translation-standard.md`: language hierarchy, names, numbers, punctuation, cultural terms, quotations, and footnotes.
- `references/genre-style-guide.md`: narrative, institutional, academic, and literary profiles.
- `references/sample-derived-patterns.md`: approved tendencies and anti-patterns extracted from the six reviewed samples.
- `references/docx-workflow.md`: extraction, manifest editing, insertion, OOXML boundaries, and render QA.
- `scripts/extract_docx.py`: inspect the DOCX and emit a deterministic JSON manifest of translatable editable blocks.
- `scripts/build_bilingual_docx.py`: consume the source DOCX and completed manifest, then insert formatted Spanish blocks without modifying image pixels.
- `scripts/audit_bilingual_docx.py`: verify source preservation, one-to-one pairing, untranslated manifest entries, images, footnote references, and Spanish formatting.
- `tests/`: automated unit and integration tests with generated DOCX fixtures.

No README, installation guide, changelog, or unrelated assets are included.

## Data flow

1. Validate that the input is a readable `.docx`.
2. Inspect tracked changes; require a clean source view or accept changes into a temporary working copy before extraction.
3. Extract translatable editable blocks into JSON with stable IDs, locations, source text, block kind, style metadata, and translation fields.
4. Classify the article genre and read the relevant style reference.
5. Research proper names and domain terms when needed.
6. Translate each manifest block, keeping paragraph-level one-to-one alignment and recording Spanish only in the translation field.
7. Run linguistic self-review in three passes: fidelity, native-language editing, independent proofreading.
8. Build the bilingual DOCX from the source and completed manifest.
9. Audit structural completeness and image preservation.
10. Render every page to PNG using the `documents` skill, inspect at 100%, fix defects, and repeat until clean.
11. Return only the final bilingual `.docx`.

## Error handling

- Fail before mutation when the file is not DOCX, is encrypted, is unreadable, or contains unsupported corruption.
- Fail the build when any required manifest translation is blank or when manifest IDs do not match the source.
- Report unsupported editable OOXML constructs explicitly instead of silently omitting their text.
- Keep all temporary extraction, accepted-revision, and render files outside the final deliverable.
- Never overwrite the source document; require a distinct output path.

## Testing strategy

### Script tests

- Generate fixtures containing headings, body text, lists, tables, captions, headers, footers, images, and footnotes.
- Verify extraction order and stable block IDs.
- Verify source paragraphs remain unchanged and Spanish paragraphs follow them.
- Verify Spanish font size and line spacing.
- Verify image binary hashes and relationship targets are unchanged.
- Verify incomplete or mismatched manifests fail clearly.
- Verify source and output paths cannot be identical.

### Skill behavior tests

- Baseline-test an agent without the skill on a representative bilingual DOCX task and record omissions, formatting mistakes, and translation-process gaps.
- Forward-test the completed skill on narrative, institutional, and literary scenarios with fresh agents.
- Check that agents load the correct references, research proper names, preserve images, avoid translating image text, and perform all three QA passes.

### Final validation

- Run the skill-creator validator.
- Run all automated tests from a clean environment.
- Run extraction/build/audit on a representative sample copy.
- Render and visually inspect every resulting page.

## Installation target

Create the version-controlled package at `skills/translate-zh-es-docx/` in this repository. This keeps the package reviewable and directly copyable into a Codex skills directory. Installing a copy into a user-level runtime directory is a separate optional step because that location is outside the project workspace.
