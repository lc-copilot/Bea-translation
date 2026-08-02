import json
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from audit_bilingual_docx import audit
from build_bilingual_docx import build
from extract_docx import extract, write_manifest
from fixture_docx import make_fixture
from ooxml import qn, read_parts, write_parts


def _serialize(root):
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _rewrite(path, mutate):
    parts = read_parts(path)
    mutate(parts)
    write_parts(parts, path)


def mutate_spanish_font(path, font):
    def mutate(parts):
        root = ET.fromstring(parts["word/document.xml"])
        for para in root.findall(".//" + qn("w:p")):
            bm = para.find(qn("w:bookmarkStart"))
            if bm is not None and (bm.get(qn("w:name")) or "").startswith("zes_"):
                for rf in para.findall(".//" + qn("w:rFonts")):
                    rf.set(qn("w:ascii"), font)
        parts["word/document.xml"] = _serialize(root)
    _rewrite(path, mutate)


def mutate_media_byte(path):
    def mutate(parts):
        data = bytearray(parts["word/media/image1.png"])
        data[0] ^= 0xFF
        parts["word/media/image1.png"] = bytes(data)
    _rewrite(path, mutate)


def mutate_document(path, fn):
    def mutate(parts):
        root = ET.fromstring(parts["word/document.xml"])
        fn(root)
        parts["word/document.xml"] = _serialize(root)
    _rewrite(path, mutate)


def mutate_footnotes(path, fn):
    def mutate(parts):
        root = ET.fromstring(parts["word/footnotes.xml"])
        fn(root)
        parts["word/footnotes.xml"] = _serialize(root)
    _rewrite(path, mutate)


def _spanish_para(root):
    for para in root.findall(".//" + qn("w:p")):
        bm = para.find(qn("w:bookmarkStart"))
        if bm is not None and (bm.get(qn("w:name")) or "").startswith("zes_"):
            return para
    return None


class AuditTests(unittest.TestCase):
    def setUp(self):
        self._tmp = TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
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

    def built_case(self, with_note=True):
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
        build(source, manifest_path, output)
        return source, manifest_path, output

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
        joined = "\n".join(report["errors"])
        self.assertIn("Spanish font is not Times New Roman", joined)
        self.assertIn("media hash mismatch", joined)

    def test_detects_missing_bookmark(self):
        source, manifest, output = self.built_case()
        mutate_document(output, lambda root: _spanish_para(root).remove(
            _spanish_para(root).find(qn("w:bookmarkStart"))))
        report = audit(source, manifest, output)
        self.assertFalse(report["ok"])
        self.assertIn("missing Spanish bookmark", "\n".join(report["errors"]))

    def test_detects_spanish_paragraph_moved_away(self):
        source, manifest, output = self.built_case()
        mutate_document(output, _move_first_spanish_to_front)
        report = audit(source, manifest, output)
        self.assertFalse(report["ok"])
        self.assertIn("not immediately after", "\n".join(report["errors"]))

    def test_detects_changed_chinese_source_text(self):
        source, manifest, output = self.built_case()
        mutate_document(output, lambda root: _set_first_source_text(root, "被改动的中文。"))
        report = audit(source, manifest, output)
        self.assertFalse(report["ok"])
        self.assertIn("Chinese source text", "\n".join(report["errors"]))

    def test_detects_wrong_size_and_spacing(self):
        source, manifest, output = self.built_case()
        mutate_document(output, lambda root: _set_size(root, "20"))
        report = audit(source, manifest, output)
        self.assertFalse(report["ok"])
        self.assertIn("12 pt", "\n".join(report["errors"]))
        mutate_document(output, lambda root: _set_line(root, "240"))
        report2 = audit(source, manifest, output)
        self.assertFalse(report2["ok"])
        self.assertIn("1.5", "\n".join(report2["errors"]))

    def test_detects_missing_explanatory_footnote(self):
        source, manifest, output = self.built_case()
        mutate_footnotes(output, lambda root: _remove_new_footnotes(root))
        report = audit(source, manifest, output)
        self.assertFalse(report["ok"])
        self.assertIn("missing explanatory footnote", "\n".join(report["errors"]))

    def test_detects_source_sha_mismatch(self):
        source, manifest_path, output = self.built_case()
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        manifest["source"]["sha256"] = "0" * 64
        manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
        report = audit(source, manifest_path, output)
        self.assertFalse(report["ok"])
        self.assertIn("source sha256", "\n".join(report["errors"]))


def _set_first_source_text(root, new_text):
    target = None
    for para in root.findall(".//" + qn("w:p")):
        if para.find(qn("w:bookmarkStart")) is None:
            t = para.find(".//" + qn("w:t"))
            if t is not None and "第一段" in (t.text or ""):
                target = t
                break
    if target is not None:
        target.text = new_text


def _set_size(root, val):
    for para in root.findall(".//" + qn("w:p")):
        bm = para.find(qn("w:bookmarkStart"))
        if bm is not None and (bm.get(qn("w:name")) or "").startswith("zes_"):
            for sz in para.findall(".//" + qn("w:sz")):
                sz.set(qn("w:val"), val)


def _set_line(root, val):
    for para in root.findall(".//" + qn("w:p")):
        bm = para.find(qn("w:bookmarkStart"))
        if bm is not None and (bm.get(qn("w:name")) or "").startswith("zes_"):
            spacing = para.find(".//" + qn("w:spacing"))
            if spacing is not None:
                spacing.set(qn("w:line"), val)


def _remove_new_footnotes(root):
    for fn in root.findall(".//" + qn("w:footnote")):
        fid = fn.get(qn("w:id"))
        if fid not in ("-1", "0", "2"):
            root.remove(fn)


def _move_first_spanish_to_front(root):
    body = root.find(".//" + qn("w:body"))
    para = _spanish_para(root)
    if body is not None and para is not None:
        body.remove(para)
        body.insert(0, para)


if __name__ == "__main__":
    unittest.main()
