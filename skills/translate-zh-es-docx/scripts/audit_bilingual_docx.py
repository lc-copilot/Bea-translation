"""非破坏性审计双语 DOCX：配对、西语文本、格式、脚注、媒体字节与图片关系。

用法：audit_bilingual_docx.py SOURCE.docx manifest.json OUTPUT.docx [--json]
退出码：0 干净；1 存在缺陷；2 参数或 DOCX 错误。
"""
import argparse
import json
import sys
from pathlib import Path
from xml.etree import ElementTree as ET

from ooxml import (
    DocxError,
    R_NS,
    W_NS,
    qn,
    media_hashes,
    paragraph_text,
    read_parts,
    source_sha256,
)

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MARKER_RE = __import__("re").compile(r"\{\{fn:[A-Za-z0-9_-]+\}\}")
BOOKMARK_PREFIX = "zes_"


def _serialize(root) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _bookmark_name(block_id: str) -> str:
    return BOOKMARK_PREFIX + block_id.replace("-", "")


def _image_relationships(parts: dict[str, bytes]) -> dict[str, str]:
    rels_xml = parts.get("word/_rels/document.xml.rels")
    if rels_xml is None:
        return {}
    root = ET.fromstring(rels_xml)
    return {
        rel.get("Id"): rel.get("Target")
        for rel in root.findall(f".//{{{REL_NS}}}Relationship")
        if "/image" in (rel.get("Type") or "")
    }


def _spanish_runs_ok(para: ET.Element, errors: list[str], block_id: str) -> None:
    fonts = [rf for rf in para.findall(".//" + qn("w:rFonts"))]
    if fonts and any(
        rf.get(qn("w:ascii")) != "Times New Roman"
        or rf.get(qn("w:hAnsi")) != "Times New Roman"
        or rf.get(qn("w:eastAsia")) != "Times New Roman"
        or rf.get(qn("w:cs")) != "Times New Roman"
        for rf in fonts
    ):
        errors.append(f"Spanish font is not Times New Roman at block {block_id}")
    sizes = [sz.get(qn("w:val")) for sz in para.findall(".//" + qn("w:sz"))]
    if sizes and any(v != "24" for v in sizes):
        errors.append(f"Spanish size is not 12 pt at block {block_id}")
    sizes_cs = [sz.get(qn("w:val")) for sz in para.findall(".//" + qn("w:szCs"))]
    if sizes_cs and any(v != "24" for v in sizes_cs):
        errors.append(f"Spanish size is not 12 pt at block {block_id}")
    langs = [lg.get(qn("w:val")) for lg in para.findall(".//" + qn("w:lang"))]
    if langs and any(v != "es-ES" for v in langs):
        errors.append(f"Spanish language is not es-ES at block {block_id}")
    ppr = para.find(qn("w:pPr"))
    if ppr is not None:
        spacing = ppr.find(qn("w:spacing"))
        if spacing is not None:
            if spacing.get(qn("w:line")) != "360" \
                    or spacing.get(qn("w:lineRule")) != "auto":
                errors.append(f"Spanish line spacing is not 1.5 at block {block_id}")


def audit(source: Path, manifest_path: Path, output: Path) -> dict[str, object]:
    source = Path(source).resolve()
    output = Path(output).resolve()
    manifest_path = Path(manifest_path)
    errors: list[str] = []

    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise DocxError(f"cannot read manifest: {exc}") from exc
    if manifest.get("schema_version") != 1:
        raise DocxError("unsupported manifest schema_version")

    if manifest["source"]["sha256"] != source_sha256(source):
        errors.append("source sha256 mismatch")

    try:
        out_parts = read_parts(output)
    except DocxError as exc:
        raise DocxError(f"cannot read output docx: {exc}") from exc

    blocks = manifest["blocks"]
    fn_texts: dict[str, str] = {}
    if "word/footnotes.xml" in out_parts:
        fn_root = ET.fromstring(out_parts["word/footnotes.xml"])
        for fn in fn_root.findall(".//" + qn("w:footnote")):
            fid = fn.get(qn("w:id"))
            fn_texts[fid] = "".join(
                n.text or "" for n in fn.findall(".//" + qn("w:t"))
            )

    part_cache: dict[str, tuple[ET.Element, list[ET.Element], dict[int, ET.Element]]] = {}

    for block in blocks:
        part = block["part"]
        block_id = block["id"]
        if part not in out_parts:
            errors.append(f"missing part {part} for block {block_id}")
            continue
        if part not in part_cache:
            root = ET.fromstring(out_parts[part])
            paras = root.findall(".//" + qn("w:p"))
            parent_map = {id(c): parent for parent in root.iter() for c in parent}
            part_cache[part] = (root, paras, parent_map)
        root, paras, parent_map = part_cache[part]

        # 反向定位：先找 bookmark 西语段，其前一兄弟必须是源段。
        # 不能用 paras[paragraph_index]：构建插入西语段后段落索引已漂移。
        spanish_para = None
        for candidate in paras:
            bm = candidate.find(qn("w:bookmarkStart"))
            if bm is not None and bm.get(qn("w:name")) == _bookmark_name(block_id):
                spanish_para = candidate
                break
        if spanish_para is None:
            errors.append(f"missing Spanish bookmark at block {block_id}")
            continue

        parent = parent_map.get(id(spanish_para))
        if parent is None:
            errors.append(
                f"Spanish paragraph not immediately after source at block {block_id}"
            )
            continue
        index = list(parent).index(spanish_para)
        siblings = list(parent)
        if index == 0 or siblings[index - 1].tag != qn("w:p"):
            errors.append(
                f"Spanish paragraph not immediately after source at block {block_id}"
            )
            continue
        source_para = siblings[index - 1]
        if paragraph_text(source_para).strip() != block["source_text"]:
            errors.append(f"Chinese source text changed at block {block_id}")
            continue

        expected_text = MARKER_RE.sub("", block["translation"])
        if paragraph_text(spanish_para).strip() != expected_text:
            errors.append(f"Spanish text mismatch at block {block_id}")

        _spanish_runs_ok(spanish_para, errors, block_id)

        for note in block["notes"]:
            refs = [
                r.get(qn("w:id"))
                for r in spanish_para.findall(".//" + qn("w:footnoteReference"))
            ]
            if not any(fn_texts.get(rid) == note["text"] for rid in refs):
                errors.append(f"missing explanatory footnote at block {block_id}")

    # 媒体字节与图片关系
    src_media = media_hashes(source)
    out_media = media_hashes(output)
    if src_media != out_media:
        errors.append("media hash mismatch")
    src_parts = read_parts(source)
    if _image_relationships(src_parts) != _image_relationships(out_parts):
        errors.append("image relationship changed")

    return {
        "ok": not errors,
        "source": str(source),
        "output": str(output),
        "checked_blocks": len(blocks),
        "checked_media": len(src_media),
        "errors": errors,
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="audit_bilingual_docx.py",
        description="Audit a bilingual Chinese-Spanish DOCX against its manifest.",
    )
    parser.add_argument("source", help="source .docx path")
    parser.add_argument("manifest", help="translated manifest JSON path")
    parser.add_argument("output", help="built bilingual .docx path")
    parser.add_argument("--json", action="store_true",
                        help="print the full report as JSON")
    args = parser.parse_args(argv)
    try:
        report = audit(args.source, args.manifest, args.output)
    except DocxError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0 if report["ok"] else 1
    if report["ok"]:
        print(f"OK: {report['checked_blocks']} blocks, "
              f"{report['checked_media']} media file")
        return 0
    first = report["errors"][0]
    rest = len(report["errors"]) - 1
    suffix = f" (and {rest} more)" if rest else ""
    print(f"ERROR: {first}{suffix}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
