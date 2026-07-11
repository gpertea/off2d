"""Stdlib unittest suite (no third-party deps). Run: python -m unittest -v"""

import io
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from off2d import detect, xlsx, sheet2txt  # noqa: E402

HERE = os.path.dirname(__file__)
FIXture = os.path.join(HERE, "..", "test", "ST7-coloc_results.xlsx")


class DetectTests(unittest.TestCase):
    def test_extension(self):
        self.assertEqual(detect.detect_type("a.xlsx"), detect.XLSX)
        self.assertEqual(detect.detect_type("a.docx"), detect.DOCX)

    def test_fixture_sniff(self):
        self.assertEqual(detect.detect_type(FIXture), detect.XLSX)


class SheetTests(unittest.TestCase):
    def test_sheet_names_and_order(self):
        with xlsx.Xlsx(FIXture) as xl:
            names = [n for n, _ in xl.sheets()]
        self.assertEqual(names[:2], ["coloc_pass", "coloc_summary"])

    def test_first_sheet_default(self):
        name, rows = xlsx.read_sheet_rows(FIXture)
        self.assertEqual(name, "coloc_pass")
        self.assertGreater(len(rows), 1)

    def test_select_by_index(self):
        name, _ = xlsx.read_sheet_rows(FIXture, "2")
        self.assertEqual(name, "coloc_summary")

    def test_select_by_name(self):
        name, _ = xlsx.read_sheet_rows(FIXture, "coloc_summary")
        self.assertEqual(name, "coloc_summary")

    def test_index_out_of_range(self):
        with self.assertRaises(ValueError):
            xlsx.read_sheet_rows(FIXture, "99")

    def test_rows_rectangular(self):
        _, rows = xlsx.read_sheet_rows(FIXture)
        widths = {len(r) for r in rows}
        self.assertEqual(len(widths), 1, "rows should be padded to one width")


class FormatGuessTests(unittest.TestCase):
    def test_guess(self):
        self.assertEqual(sheet2txt._guess_format("x.csv"), (",", False))
        self.assertEqual(sheet2txt._guess_format("x.tab"), ("\t", False))
        self.assertEqual(sheet2txt._guess_format("x.tsv"), ("\t", False))
        self.assertEqual(sheet2txt._guess_format("x.tab.gz"), ("\t", True))
        self.assertEqual(sheet2txt._guess_format("x.csv.gz"), (",", True))
        self.assertEqual(sheet2txt._guess_format("x.unknown"), ("\t", False))


class ConvertTests(unittest.TestCase):
    def test_tab_output_to_file(self):
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".tab", delete=False) as tf:
            out = tf.name
        try:
            sheet2txt.convert(FIXture, None, out)
            with open(out) as fh:
                first = fh.readline().rstrip("\n")
            self.assertIn("\t", first)
        finally:
            os.unlink(out)

    def test_gzip_output(self):
        import gzip
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".csv.gz", delete=False) as tf:
            out = tf.name
        try:
            sheet2txt.convert(FIXture, None, out)
            with gzip.open(out, "rt") as fh:
                first = fh.readline()
            self.assertIn(",", first)
        finally:
            os.unlink(out)


if __name__ == "__main__":
    unittest.main()
