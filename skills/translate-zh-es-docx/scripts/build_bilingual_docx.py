"""把填好译文的 manifest 构建为中西双语 DOCX。

用法：build_bilingual_docx.py SOURCE.docx manifest.json OUTPUT.docx
"""
import argparse
import copy
import json
import os
import re
import sys
from pathlib import Path
from xml.etree import ElementTree as ET

from ooxml import (
    DocxError,
    R_NS,
    W_NS,
    qn,
    read_parts,
    source_sha256,
    translatable_parts,
    validate_docx,
    write_parts,
)

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
FOOTNOTES_REL_TYPE = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                      "relationships/footnotes")
FOOTNOTES_CT = ("application/vnd.openxmlformats-officedocument."
                "wordprocessingml.footnotes+xml")
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
MARKER_RE = re.compile(r"\{\{fn:([A-Za-z0-9_-]+)\}\}")
PUNCT = ".?!:;"

# 注册常用 OOXML 命名空间，保证序列化保留 w:/r:/wp:/a:/pic: 等前缀。
ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("wp",
                      "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("v", "urn:schemas-microsoft-com:vml")
ET.register_namespace("o", "urn:schemas-microsoft-com:office:office")
ET.register_namespace("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")
ET.register_namespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math")
ET.register_namespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml")

_SEPARATOR_FOOTNOTES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:footnotes xmlns:w="%s">'
    '<w:footnote w:type="separator" w:id="-1">'
    "<w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
    '<w:footnote w:type="continuationSeparator" w:id="0">'
    "<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
    "</w:footnotes>" % W_NS
)


def _serialize(root) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _spanish_run_props(rpr: ET.Element) -> ET.Element:
    rfonts = ET.SubElement(rpr, qn("w:rFonts"))
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn("w:" + attr), "Times New Roman")
    sz = ET.SubElement(rpr, qn("w:sz"))
    sz.set(qn("w:val"), "24")
    szcs = ET.SubElement(rpr, qn("w:szCs"))
    szcs.set(qn("w:val"), "24")
    lang = ET.SubElement(rpr, qn("w:lang"))
    lang.set(qn("w:val"), "es-ES")
    return rpr


def split_translation(translation: str) -> list[tuple[str, str]]:
    """把 translation 按 {{fn:key}} 切成 [("text", ...) | ("fn", key), ...]。"""
    segments: list[tuple[str, str]] = []
    pos = 0
    for m in MARKER_RE.finditer(translation):
        if m.start() > pos:
            segments.append(("text", translation[pos:m.start()]))
        segments.append(("fn", m.group(1)))
        pos = m.end()
    if pos < len(translation):
        segments.append(("text", translation[pos:]))
    return segments


def make_translation_paragraph(source_paragraph: ET.Element, segments,
                               bookmark_id: str, note_ids: dict,
                               block_index: int, bookmark_seq: list[int]) -> ET.Element:
    """构造西语段落：deep copy 源 pPr（移除 numPr），强制格式，插入 bookmark 与 run。"""
    para = ET.Element(qn("w:p"))
    ppr = source_paragraph.find(qn("w:pPr"))
    if ppr is not None:
        ppr = copy.deepcopy(ppr)
        numpr = ppr.find(qn("w:numPr"))
        if numpr is not None:
            ppr.remove(numpr)
        old_spacing = ppr.find(qn("w:spacing"))
        if old_spacing is not None:
            ppr.remove(old_spacing)
    else:
        ppr = ET.Element(qn("w:pPr"))
    spacing = ET.SubElement(ppr, qn("w:spacing"))
    spacing.set(qn("w:line"), "360")
    spacing.set(qn("w:lineRule"), "auto")
    para.append(ppr)

    bm_id = str(bookmark_seq[0])
    bookmark_seq[0] += 1
    bm_start = ET.SubElement(para, qn("w:bookmarkStart"))
    bm_start.set(qn("w:id"), bm_id)
    bm_start.set(qn("w:name"), bookmark_id)

    for kind, payload in segments:
        run = ET.SubElement(para, qn("w:r"))
        rpr = ET.SubElement(run, qn("w:rPr"))
        _spanish_run_props(rpr)
        if kind == "fn":
            ref = ET.SubElement(run, qn("w:footnoteReference"))
            ref.set(qn("w:id"), str(note_ids[(block_index, payload)]))
        else:
            t = ET.SubElement(run, qn("w:t"))
            t.set(XML_SPACE, "preserve")
            t.text = payload

    bm_end = ET.SubElement(para, qn("w:bookmarkEnd"))
    bm_end.set(qn("w:id"), bm_id)
    return para


def ensure_footnotes_part(parts: dict[str, bytes]):
    """确保 word/footnotes.xml 存在。返回 (footnotes_root, next_id, created)。"""
    if "word/footnotes.xml" in parts:
        root = ET.fromstring(parts["word/footnotes.xml"])
        positive = [
            int(f.get(qn("w:id"))) for f in root.iter(qn("w:footnote"))
            if f.get(qn("w:id")) is not None and f.get(qn("w:id")).lstrip("-").isdigit()
            and int(f.get(qn("w:id"))) > 0
        ]
        return root, (max(positive) + 1 if positive else 1), False
    root = ET.fromstring(_SEPARATOR_FOOTNOTES)
    return root, 1, True


def append_added_footnote(footnotes_root: ET.Element, note_text: str,
                          footnote_id: int) -> None:
    """追加一个真正的 Word 脚注：footnoteRef + 格式化西语文本。"""
    fn = ET.SubElement(footnotes_root, qn("w:footnote"))
    fn.set(qn("w:id"), str(footnote_id))
    p = ET.SubElement(fn, qn("w:p"))
    r1 = ET.SubElement(p, qn("w:r"))
    rpr1 = ET.SubElement(r1, qn("w:rPr"))
    vert = ET.SubElement(rpr1, qn("w:vertAlign"))
    vert.set(qn("w:val"), "superscript")
    ET.SubElement(r1, qn("w:footnoteRef"))
    r2 = ET.SubElement(p, qn("w:r"))
    rpr2 = ET.SubElement(r2, qn("w:rPr"))
    _spanish_run_props(rpr2)
    t = ET.SubElement(r2, qn("w:t"))
    t.set(XML_SPACE, "preserve")
    t.text = note_text


def load_and_validate_manifest(source: Path, manifest_path: Path) -> dict[str, object]:
    from extract_docx import extract  # 本地导入避免循环依赖

    source = Path(source)
    manifest_path = Path(manifest_path)
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise DocxError(f"cannot read manifest: {exc}") from exc
    if manifest.get("schema_version") != 1:
        raise DocxError("unsupported manifest schema_version")
    if manifest["source"]["sha256"] != source_sha256(source):
        raise DocxError("manifest sha256 does not match the source docx")

    fresh = extract(source)
    fresh_blocks = fresh["blocks"]
    blocks = manifest["blocks"]
    if len(blocks) != len(fresh_blocks):
        raise DocxError("manifest block count does not match the source docx")
    for manifest_block, fresh_block in zip(blocks, fresh_blocks):
        for key in ("id", "part", "paragraph_index", "source_text"):
            if manifest_block[key] != fresh_block[key]:
                raise DocxError(
                    f"manifest {key.replace('_', ' ')} mismatch for block {fresh_block['id']}"
                )

    for block in blocks:
        translation = block["translation"]
        if not isinstance(translation, str) or not translation.strip():
            raise DocxError(f"blank translation for block {block['id']}")
        notes = block["notes"]
        if not isinstance(notes, list):
            raise DocxError(f"notes must be a list for block {block['id']}")
        keys = []
        for note in notes:
            if not isinstance(note, dict) or "key" not in note or "text" not in note:
                raise DocxError(f"malformed note for block {block['id']}")
            if not isinstance(note["text"], str) or not note["text"].strip():
                raise DocxError(f"blank note text for block {block['id']}")
            keys.append(note["key"])
        if len(keys) != len(set(keys)):
            raise DocxError(f"duplicate note key in block {block['id']}")
        markers = [m.group(1) for m in MARKER_RE.finditer(translation)]
        if len(markers) != len(set(markers)):
            raise DocxError(f"duplicate note marker in block {block['id']}")
        declared = set(keys)
        undeclared = set(markers) - declared
        if undeclared:
            raise DocxError(f"undeclared note marker in block {block['id']}")
        missing = declared - set(markers)
        if missing:
            raise DocxError(f"missing note marker in block {block['id']}")
        for m in MARKER_RE.finditer(translation):
            if m.start() == 0 or translation[m.start() - 1] not in PUNCT:
                raise DocxError(
                    f"note marker must follow punctuation in block {block['id']}"
                )
    return manifest


def build(source: Path, manifest_path: Path, output: Path) -> None:
    source = Path(source).resolve()
    output = Path(output).resolve()
    if source == output:
        raise DocxError(
            "output path resolves to the source docx; use a distinct path"
        )
    manifest = load_and_validate_manifest(source, Path(manifest_path))
    blocks = manifest["blocks"]
    parts = validate_docx(source)

    fn_root, next_id, fn_created = ensure_footnotes_part(parts)
    block_index = {id(block): i for i, block in enumerate(blocks)}
    note_ids: dict[tuple[int, str], int] = {}
    for i, block in enumerate(blocks):
        for note in block["notes"]:
            note_ids[(i, note["key"])] = next_id
            next_id += 1

    bookmark_seq = [1000]
    for part in translatable_parts(parts):
        part_blocks = [b for b in blocks if b["part"] == part]
        if not part_blocks and part != "word/footnotes.xml":
            continue
        root = fn_root if part == "word/footnotes.xml" else ET.fromstring(parts[part])
        parent_map = {id(child): parent for parent in root.iter() for child in parent}
        paras = [p for p in root.iter(qn("w:p"))]
        for block in reversed(part_blocks):
            ordinal = block["paragraph_index"]
            source_para = paras[ordinal]
            segments = split_translation(block["translation"])
            new_para = make_translation_paragraph(
                source_para,
                segments,
                "zes_" + block["id"].replace("-", ""),
                note_ids,
                block_index[id(block)],
                bookmark_seq,
            )
            parent = parent_map[id(source_para)]
            index = list(parent).index(source_para)
            parent.insert(index + 1, new_para)
        if part == "word/footnotes.xml":
            for i, block in enumerate(blocks):
                for note in block["notes"]:
                    append_added_footnote(fn_root, note["text"],
                                          note_ids[(i, note["key"])])
        parts[part] = _serialize(root)

    if fn_created:
        parts["word/footnotes.xml"] = _serialize(fn_root)
        parts["[Content_Types].xml"] = _add_footnotes_content_type(
            parts["[Content_Types].xml"]
        )
        parts["word/_rels/document.xml.rels"] = _add_footnotes_relationship(
            parts["word/_rels/document.xml.rels"]
        )

    tmp = output.with_name(output.name + ".tmp-build")
    write_parts(parts, tmp)
    os.replace(tmp, output)


def _add_footnotes_content_type(content_types_xml: bytes) -> bytes:
    root = ET.fromstring(content_types_xml)
    if not any(
        node.get("PartName") == "/word/footnotes.xml"
        for node in root.iter(f"{{{CT_NS}}}Override")
    ):
        override = ET.SubElement(root, f"{{{CT_NS}}}Override")
        override.set("PartName", "/word/footnotes.xml")
        override.set("ContentType", FOOTNOTES_CT)
    return _serialize(root)


def _add_footnotes_relationship(document_rels_xml: bytes) -> bytes:
    root = ET.fromstring(document_rels_xml)
    if not any(
        node.get("Type") == FOOTNOTES_REL_TYPE
        for node in root.iter(f"{{{REL_NS}}}Relationship")
    ):
        used = {
            node.get("Id") for node in root.iter(f"{{{REL_NS}}}Relationship")
            if node.get("Id")
        }
        rid = "rIdFootnotes"
        while rid in used:
            rid += "x"
        rel = ET.SubElement(root, f"{{{REL_NS}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", FOOTNOTES_REL_TYPE)
        rel.set("Target", "footnotes.xml")
    return _serialize(root)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="build_bilingual_docx.py",
        description="Build a bilingual Chinese-Spanish DOCX from a manifest.",
    )
    parser.add_argument("source", help="source .docx path")
    parser.add_argument("manifest", help="translated manifest JSON path")
    parser.add_argument("output", help="output .docx path")
    args = parser.parse_args(argv)
    try:
        build(args.source, args.manifest, args.output)
    except DocxError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    print(args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
