import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import compare_docs


class DocDiffSmokeTests(unittest.TestCase):
    def test_word_diff_marks_changed_words(self):
        left, right = compare_docs.word_diff(
            "The generation plan is ready.",
            "The resource plan is ready.",
        )
        self.assertIn("w-del", left)
        self.assertIn("generation", left)
        self.assertIn("w-ins", right)
        self.assertIn("resource", right)

    def test_text_extraction_and_report_rendering(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            first = Path(tmpdir) / "first.txt"
            second = Path(tmpdir) / "second.txt"
            first.write_text("Alpha paragraph\n\nOriginal wording", encoding="utf-8")
            second.write_text("Alpha paragraph\n\nRevised wording", encoding="utf-8")

            blocks_a = compare_docs.extract_blocks(str(first))
            blocks_b = compare_docs.extract_blocks(str(second))
            rows = compare_docs.align_blocks(blocks_a, blocks_b)
            summary = compare_docs.summarise(rows)
            report = compare_docs.build_html(rows, summary, first.name, second.name)

            self.assertGreaterEqual(summary["changed"], 1)
            self.assertIn("first.txt", report)
            self.assertIn("second.txt", report)


if __name__ == "__main__":
    unittest.main()
