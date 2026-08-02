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
