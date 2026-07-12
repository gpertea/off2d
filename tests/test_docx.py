"""docx -> markdown tests (stdlib only). Run: python -m unittest -v

Uses the real sample document plus a tiny generated fixture (with an image)
so image extraction is covered without committing a binary."""

import base64
import os
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from off2d import docx2md  # noqa: E402

HERE = os.path.dirname(__file__)
SAMPLE = os.path.join(
    HERE, "..", "test", "Modern-Genotype-Storage-Formats-for-Large-Cohorts.docx"
)

_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mP8"
    "z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def _make_image_docx(path):
    doc = (
        '<?xml version="1.0"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main" xmlns:r="http://schemas.'
        'openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        "<w:body>"
        '<w:p><w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr>'
        "<w:t>Title Here</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Before.</w:t></w:r></w:p>"
        "<w:p><w:r><w:drawing><a:blip r:embed=\"rId5\"/></w:drawing>"
        "</w:r></w:p>"
        "<w:p><w:r><w:t>After.</w:t></w:r></w:p>"
        "</w:body></w:document>"
    )
    rels = (
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.'
        'openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/image" Target="media/image1.png"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", rels)
        z.writestr("word/media/image1.png", _PNG)


class SampleDocTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.md, _ = docx2md.convert(SAMPLE)

    def test_title_is_h1(self):
        self.assertTrue(self.md.startswith("# Modern Genotype Storage Formats"))

    def test_has_h2_sections(self):
        self.assertIn("\n## Executive Summary\n", self.md)

    def test_italic_preserved(self):
        self.assertIn("*and*", self.md)

    def test_hyperlinks_present(self):
        # 28 external hyperlinks in the source; expect inline [text](http...)
        self.assertGreaterEqual(self.md.count("](http"), 20)

    def test_tables_rendered(self):
        self.assertIn("| --- |", self.md)
        self.assertIn("| Format |", self.md)

    def test_bullet_list(self):
        self.assertIn("\n- ", self.md)

    def test_ordered_list(self):
        self.assertIn("\n1. ", self.md)

    def test_code_fence(self):
        self.assertIn("```", self.md)

    def test_code_lines_not_merged(self):
        self.assertNotIn("txttiledbvcf", self.md)

    def test_inline_citations_linked(self):
        # plain [1] marker links to reference 1's URL, keeping visible "[1]"
        self.assertIn(
            "[\\[1\\]](https://www.tiledb.com/blog/"
            "population-genomics-data-with-tiledb)",
            self.md,
        )

    def test_caret_citations_linked(self):
        # footnote-style [^N] markers share the reference numbering
        self.assertRegex(self.md, r"\[\\\[\^\d+\\\]\]\(https?://")

    def test_references_section_not_rewritten(self):
        # numeric citation markers must not be linkified inside the list
        # (escaped brackets around non-numeric text like "[PDF]" are fine)
        import re
        tail = self.md[self.md.index("## References"):]
        self.assertIsNone(re.search(r"\[\\\[\^?\d+\\\]\]\(", tail))

    def test_no_cite_links_option(self):
        plain, _ = docx2md.convert(SAMPLE, cite_links=False)
        self.assertIn("\\[1\\]", plain)
        self.assertNotIn(
            "[\\[1\\]](https://www.tiledb.com", plain
        )


class ImageTests(unittest.TestCase):
    def test_extract_with_images(self):
        tmp = tempfile.mkdtemp()
        docx = os.path.join(tmp, "img.docx")
        _make_image_docx(docx)
        img_dir = os.path.join(tmp, "media_out")
        md, images = docx2md.convert(
            docx, extract_images=True, image_dir=img_dir
        )
        self.assertIn("![[image1.png]]", md)
        self.assertEqual(len(images), 1)
        self.assertTrue(os.path.exists(os.path.join(img_dir, "image1.png")))

    def test_no_extract_without_flag(self):
        tmp = tempfile.mkdtemp()
        docx = os.path.join(tmp, "img.docx")
        _make_image_docx(docx)
        md, images = docx2md.convert(docx, extract_images=False)
        self.assertNotIn("image1.png", md)
        self.assertEqual(images, [])


if __name__ == "__main__":
    unittest.main()
