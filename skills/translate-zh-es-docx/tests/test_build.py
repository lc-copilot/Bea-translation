import json
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from build_bilingual_docx import build
from extract_docx import extract, write_manifest
from fixture_docx import make_fixture
from ooxml import DocxError, qn, media_hashes, paragraph_text, read_parts, translatable_parts


def _parent_map(root):
    return {id(child): parent for parent in root.iter() for child in parent}


def _first_font(para):
    for rf in para.findall(".//" + qn("w:rFonts")):
        return rf.get(qn("w:ascii"))
    return None


def _first_size(para):
    for sz in para.findall(".//" + qn("w:sz")):
        return sz.get(qn("w:val"))
    return None


def _first_line(para):
    ppr = para.find(qn("w:pPr"))
    if ppr is None:
        return None
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        return None
    return spacing.get(qn("w:line"))


def inspect_bilingual_xml(output):
    """从构建产物提取双语结构与格式报告（测试辅助）。"""
    parts = read_parts(output)
    report = {
        "source_texts": [],
        "spanish_texts": [],
        "paired_immediately": [],
        "fonts": [],
        "sizes": [],
        "line_spacing": [],
        "added_footnote_texts": [],
        "note_reference_follows_period": False,
    }
    saw_fn_ref = False
    fn_ref_after_period = False
    for part in translatable_parts(parts):
        root = ET.fromstring(parts[part])
        pmap = _parent_map(root)
        # 用 findall(".//tag") 而非 iter(tag)：CPython 3.9 的 iter 生成器存在共享状态 bug，
        # 连续/嵌套 iter 会漏节点；findall 返回独立 list，安全。
        footnotes = root.findall(".//" + qn("w:footnote"))
        fn_info = [
            {
                "para_ids": {id(n) for n in fn.findall(".//" + qn("w:p"))},
                "has_bm": bool(fn.findall(".//" + qn("w:bookmarkStart"))),
                "has_es": any(
                    rpr.find(qn("w:rFonts")) is not None
                    for rpr in fn.findall(".//" + qn("w:rPr"))
                ),
            }
            for fn in footnotes
        ]

        def in_added_footnote(para):
            """脚注 part 中、新增解释脚注（无 bookmark、带西语格式）内的段落。"""
            if part != "word/footnotes.xml":
                return False
            for info in fn_info:
                if id(para) in info["para_ids"]:
                    return info["has_es"] and not info["has_bm"]
            return False

        for para in root.findall(".//" + qn("w:p")):
            if in_added_footnote(para):
                continue
            bm = para.find(qn("w:bookmarkStart"))
            if bm is not None and (bm.get(qn("w:name")) or "").startswith("zes_"):
                report["spanish_texts"].append(paragraph_text(para).strip())
                report["fonts"].append(_first_font(para))
                report["sizes"].append(_first_size(para))
                report["line_spacing"].append(_first_line(para))
                parent = pmap.get(id(para))
                idx = list(parent).index(para) if parent is not None else -1
                prev = parent[idx - 1] if parent is not None and idx > 0 else None
                report["paired_immediately"].append(
                    prev is not None and prev.tag == qn("w:p")
                    and prev.find(qn("w:bookmarkStart")) is None
                )
                for run in para.findall(qn("w:r")):
                    if run.find(qn("w:footnoteReference")) is not None:
                        saw_fn_ref = True
                        p2 = pmap.get(id(run))
                        i2 = list(p2).index(run) if p2 is not None else -1
                        prev_run = p2[i2 - 1] if p2 is not None and i2 > 0 else None
                        if prev_run is not None and prev_run.tag == qn("w:r"):
                            t = prev_run.find(qn("w:t"))
                            if t is not None and (t.text or "").endswith("."):
                                fn_ref_after_period = True
            elif paragraph_text(para).strip():
                report["source_texts"].append(paragraph_text(para).strip())
    if "word/footnotes.xml" in parts:
        fn_root = ET.fromstring(parts["word/footnotes.xml"])
        footnotes = fn_root.findall(".//" + qn("w:footnote"))
        fn_records = [
            {
                "fid": fn.get(qn("w:id")),
                "has_bookmark": bool(fn.findall(".//" + qn("w:bookmarkStart"))),
                "has_spanish_format": any(
                    rpr.find(qn("w:rFonts")) is not None
                    for rpr in fn.findall(".//" + qn("w:rPr"))
                ),
                "text": "".join(n.text or "" for n in fn.findall(".//" + qn("w:t"))),
            }
            for fn in footnotes
        ]
        for rec in fn_records:
            if rec["fid"] in ("-1", "0"):
                continue
            if rec["has_spanish_format"] and not rec["has_bookmark"]:
                report["added_footnote_texts"].append(rec["text"])
    report["note_reference_follows_period"] = saw_fn_ref and fn_ref_after_period
    return report


class BuildTests(unittest.TestCase):
    def setUp(self):
        self._tmp = TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.source_texts = [
            "测试标题", "第一段。", "列表项", "表格文字", "文本框文字",
            "页眉文字", "页脚文字", "原脚注文字",
        ]
        self.translations = {
            "测试标题": "El sueño del fútbol al pie de la montaña Changbai",
            "第一段。": "En una noche nevada, los niños corren tras un balón.",
            "列表项": "Elemento de lista",
            "表格文字": "Texto de tabla",
            "文本框文字": "Texto de cuadro de texto",
            "页眉文字": "Texto de encabezado",
            "页脚文字": "Texto de pie de página",
            "原脚注文字": "Texto de nota original",
        }
        self.li_note = ("Unidad tradicional china de longitud, "
                        "equivalente hoy a unos quinientos metros.")

    def tearDown(self):
        self._tmp.cleanup()

    def prepared_case(self, with_note=False):
        source = self.tmp / "article.docx"
        make_fixture(source)
        manifest_path = self.tmp / "manifest.json"
        manifest = extract(source)
        for text, tr in self.translations.items():
            block = next(b for b in manifest["blocks"] if b["source_text"] == text)
            block["translation"] = tr
            if with_note and text == "第一段。":
                block["translation"] = tr + "{{fn:li}}"
                block["notes"] = [{"key": "li", "text": self.li_note}]
        write_manifest(manifest, manifest_path)
        output = self.tmp / "bilingual.docx"
        return source, manifest_path, output

    def test_build_pairs_blocks_formats_spanish_and_preserves_images(self):
        source, manifest_path, output = self.prepared_case(with_note=True)
        before = media_hashes(source)
        build(source, manifest_path, output)
        self.assertEqual(media_hashes(output), before)
        report = inspect_bilingual_xml(output)
        self.assertEqual(report["source_texts"], self.source_texts)
        self.assertEqual(report["spanish_texts"], list(self.translations.values()))
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

    def test_rejects_source_sha_mismatch(self):
        source, manifest_path, output = self.prepared_case()
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        manifest["source"]["sha256"] = "0" * 64
        manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
        with self.assertRaisesRegex(DocxError, "sha256"):
            build(source, manifest_path, output)

    def test_rejects_changed_source_text(self):
        source, manifest_path, output = self.prepared_case()
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        manifest["blocks"][1]["source_text"] = "被改动的文本。"
        manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
        with self.assertRaisesRegex(DocxError, "source text"):
            build(source, manifest_path, output)

    def test_rejects_missing_or_duplicate_note_marker(self):
        source, manifest_path, output = self.prepared_case()
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        manifest["blocks"][1]["notes"] = [{"key": "li", "text": "Nota de prueba."}]
        manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
        with self.assertRaisesRegex(DocxError, "marker"):
            build(source, manifest_path, output)

        source2, manifest_path2, output2 = self.prepared_case()
        manifest2 = json.loads(manifest_path2.read_text(encoding="utf-8"))
        manifest2["blocks"][1]["translation"] = "Texto.{{fn:li}} y {{fn:li}}"
        manifest2["blocks"][1]["notes"] = [{"key": "li", "text": "Nota de prueba."}]
        manifest_path2.write_text(json.dumps(manifest2), encoding="utf-8")
        with self.assertRaisesRegex(DocxError, "marker"):
            build(source2, manifest_path2, output2)

    def test_rejects_marker_before_punctuation(self):
        source, manifest_path, output = self.prepared_case()
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        manifest["blocks"][1]["translation"] = "Texto{{fn:li}} sin punto"
        manifest["blocks"][1]["notes"] = [{"key": "li", "text": "Nota de prueba."}]
        manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
        with self.assertRaisesRegex(DocxError, "punct"):
            build(source, manifest_path, output)

    def test_preserves_original_footnote_id(self):
        source, manifest_path, output = self.prepared_case(with_note=True)
        build(source, manifest_path, output)
        parts = read_parts(output)
        fn_root = ET.fromstring(parts["word/footnotes.xml"])
        texts = {
            fn.get(qn("w:id")): "".join(n.text or "" for n in fn.iter(qn("w:t")))
            for fn in fn_root.iter(qn("w:footnote"))
        }
        self.assertIn("2", texts)
        self.assertIn("原脚注文字", texts["2"])
        self.assertIn("Texto de nota original", texts["2"])


if __name__ == "__main__":
    unittest.main()
