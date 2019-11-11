import unittest
from exex import extract

TEST_FILE_PATH = "tests/files/sample.xlsx"


class ExtractTest(unittest.TestCase):
    def test_all(self):
        ext = extract.Extractor(TEST_FILE_PATH)

        self.assertEqual(
            ext.all(), [["name", "age"], ["alpha", 15], ["beta", 28], ["gamma", 32]]
        )

    def test_range_cell(self):
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

    # TODO

    def ztest_rows_single(self):
        ext = extract.Extractor(TEST_FILE_PATH)
        self.assertEqual(ext.rows(1), [["name", "age"]])
        self.assertEqual(ext.rows([1]), [["name", "age"]])

    def ztest_rows_multiple(self):
        ext = extract.Extractor(TEST_FILE_PATH)
        self.assertEqual(ext.rows(1, 3), [["name", "age"], ["beta", 28]])
        self.assertEqual(ext.rows([1, 3]), [["name", "age"], ["beta", 28]])

    def ztest_columns_single(self):
        ext = extract.Extractor(TEST_FILE_PATH)
        self.assertEqual(ext.columns([1]), [["name", "age"]])
        self.assertEqual(ext.columns(["name"]), [["name", "age"]])

    def ztest_columns_multiple(self):
        ext = extract.Extractor(TEST_FILE_PATH)
        self.assertEqual(ext.columns(1, 2), [["name", "age"], ["beta", 28]])
        self.assertEqual(ext.columns([1, 2]), [["name", "age"], ["beta", 28]])
        self.assertEqual(ext.columns("name", "age"), [["name", "age"], ["beta", 28]])
        self.assertEqual(ext.columns(["name", "age"]), [["name", "age"], ["beta", 28]])
