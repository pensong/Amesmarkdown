import unittest

from mdconvert_app.markdown_utils import markdown_table, slugify


class MarkdownUtilsTests(unittest.TestCase):
    def test_slugify(self) -> None:
        self.assertEqual(slugify("Quarterly Review 2026"), "quarterly-review-2026")

    def test_markdown_table(self) -> None:
        value = markdown_table([["Name", "Score"], ["Alice", "10"]])
        self.assertIn("| Name | Score |", value)
        self.assertIn("| Alice | 10 |", value)


if __name__ == "__main__":
    unittest.main()
