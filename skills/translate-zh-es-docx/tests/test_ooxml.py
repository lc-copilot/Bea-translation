import base64
import hashlib
import sys
import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from fixture_docx import PNG_BYTES, make_fixture
from ooxml import (
    DocxError,
    iter_paragraphs,
    media_hashes,
    paragraph_text,
    read_parts,
    translatable_parts,
    validate_docx,
)


class OoxmlTests(unittest.TestCase):
    def setUp(self):
        self._tmp = TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def test_fixture_exposes_all_supported_text_and_keeps_media_hash(self):
        source = self.tmp / "fixture.docx"
        make_fixture(source)
        parts = read_parts(source)
        found = [
            paragraph_text(p)
            for part in translatable_parts(parts)
            for _, p in iter_paragraphs(parts[part])
        ]
        self.assertEqual(
            found,
            ["测试标题", "第一段。", "列表项", "表格文字", "文本框文字",
             "页眉文字", "页脚文字", "原脚注文字"],
        )
        self.assertEqual(
            media_hashes(source)["word/media/image1.png"],
            hashlib.sha256(PNG_BYTES).hexdigest(),
        )

    def test_validation_rejects_wrong_extension_and_tracked_changes(self):
        wrong = self.tmp / "fixture.zip"
        wrong.write_bytes(b"not a docx")
        with self.assertRaisesRegex(DocxError, r"\.docx"):
            validate_docx(wrong)
        tracked = self.tmp / "tracked.docx"
        make_fixture(tracked, tracked_changes=True)
        with self.assertRaisesRegex(DocxError, "tracked changes"):
            validate_docx(tracked)

    def test_validation_accepts_clean_fixture(self):
        clean = self.tmp / "clean.docx"
        make_fixture(clean)
        parts = validate_docx(clean)
        self.assertIn("word/document.xml", parts)
        with zipfile.ZipFile(clean) as z:
            self.assertIsNone(z.testzip())


if __name__ == "__main__":
    unittest.main()
