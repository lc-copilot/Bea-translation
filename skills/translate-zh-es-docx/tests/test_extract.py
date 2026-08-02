import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from extract_docx import extract, write_manifest
from fixture_docx import make_fixture
from ooxml import DocxError


class ExtractTests(unittest.TestCase):
    def setUp(self):
        self._tmp = TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

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


if __name__ == "__main__":
    unittest.main()
