"""把源 DOCX 的全部可编辑文本块提取为确定性 JSON manifest。

用法：extract_docx.py SOURCE.docx --manifest manifest.json
"""
import argparse
import json
import sys
from pathlib import Path

from ooxml import (
    DocxError,
    block_kind,
    iter_paragraphs,
    paragraph_style,
    paragraph_text,
    source_sha256,
    stable_block_id,
    translatable_parts,
    validate_docx,
)


def extract(source: Path) -> dict[str, object]:
    source = Path(source)
    parts = validate_docx(source)
    manifest: dict[str, object] = {
        "schema_version": 1,
        "source": {"filename": source.name, "sha256": source_sha256(source)},
        "blocks": [],
    }
    blocks = []
    for part in translatable_parts(parts):
        for index, paragraph in iter_paragraphs(parts[part]):
            text = paragraph_text(paragraph)
            blocks.append({
                "id": stable_block_id(part, index, text),
                "part": part,
                "paragraph_index": index,
                "kind": block_kind(part, paragraph),
                "style_id": paragraph_style(paragraph),
                "source_text": text,
                "translation": "",
                "notes": [],
            })
    manifest["blocks"] = blocks
    return manifest


def write_manifest(manifest: dict[str, object], output: Path) -> None:
    """写 manifest；manifest 路径解析到源 DOCX 时拒绝。"""
    output = Path(output)
    src_sha = manifest["source"]["sha256"]
    if output.is_file() and source_sha256(output) == src_sha:
        raise DocxError(
            "manifest path resolves to the source docx; use a distinct path"
        )
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="extract_docx.py",
        description="Extract a stable translation manifest from a Chinese DOCX.",
    )
    parser.add_argument("source", help="source .docx path")
    parser.add_argument("--manifest", required=True, help="output manifest JSON path")
    args = parser.parse_args(argv)
    try:
        source = Path(args.source)
        manifest_out = Path(args.manifest)
        if source.resolve() == manifest_out.resolve():
            raise DocxError("manifest path resolves to the source docx; use a distinct path")
        manifest = extract(source)
        write_manifest(manifest, manifest_out)
    except DocxError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    print(args.manifest)
    return 0


if __name__ == "__main__":
    sys.exit(main())
