import unittest
from exex import extract

TEST_FILE_PATH = "tests/files/sample.xlsx"


class ExtractTest(unittest.TestCase):
    def test_all(self):
        ext = extract.Extractor(TEST_FILE_PATH)

        self.assertEqual(
            ext.all(), [["name", "age"], ["alpha", 15], ["beta", 28], ["gamma", 32]]
        )

    def test_range_single_cell(self):
        ext = extract.Extractor(TEST_FILE_PATH)
        self.assertEqual(ext.range("A1"), [["name"]])

    def test_range_horizontal(self):
        ext = extract.Extractor(TEST_FILE_PATH)
        self.assertEqual(ext.range("A1:B1"), [["name", "age"]])

    def test_range_vertical(self):
        ext = extract.Extractor(TEST_FILE_PATH)
        self.assertEqual(ext.range("A1:A4"), [["name"], ["alpha"], ["beta"], ["gamma"]])

    def test_range_square(self):
        ext = extract.Extractor(TEST_FILE_PATH)
        self.assertEqual(ext.range("A3:B4"), [["beta", 28], ["gamma", 32]])

    def test_range_multiple_rows(self):
        ext = extract.Extractor(TEST_FILE_PATH)
        self.assertEqual(ext.range("1:2"), [["name", "age"], ["alpha", 15]])

    @unittest.skip
    def test_range_single_row(self):
        ext = extract.Extractor(TEST_FILE_PATH)
        self.assertEqual(ext.range("1"), [["name", "age"]])
        self.assertEqual(ext.range("2"), [["alpha", 15]])

    # TODO
    @unittest.skip
    def test_range_single_column(self):
        ext = extract.Extractor(TEST_FILE_PATH)
        self.assertEqual(ext.range("A"), [["name"], ["alpha"], ["beta"], ["gamma"]])

    # TODO
    @unittest.skip
    def test_range_multiple_columns(self):
        ext = extract.Extractor(TEST_FILE_PATH)
        self.assertEqual(
            ext.range("A:B"),
            [["name", "age"], ["alpha", 15], ["beta", 28], ["gamma", 32]],
        )
