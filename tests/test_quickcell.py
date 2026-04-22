import importlib
import os
import sys
import unittest
from datetime import date, datetime, time

os.environ["GDK_BACKEND"] = "broadway"

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))


class TestVersion(unittest.TestCase):
    def test_version_is_set(self):
        module = importlib.import_module("quickcell")
        self.assertEqual(module.VERSION, "1.0.0")


class TestFormatCellValue(unittest.TestCase):
    def setUp(self):
        self.fmt = importlib.import_module("quickcell").format_cell_value

    def test_none(self):
        self.assertEqual(self.fmt(None), "")

    def test_bool(self):
        self.assertEqual(self.fmt(True), "TRUE")
        self.assertEqual(self.fmt(False), "FALSE")

    def test_int(self):
        self.assertEqual(self.fmt(42), "42")

    def test_float_integer(self):
        self.assertEqual(self.fmt(3.0), "3")

    def test_float_fraction(self):
        self.assertEqual(self.fmt(3.14), "3.14")

    def test_string(self):
        self.assertEqual(self.fmt("hello"), "hello")

    def test_date(self):
        self.assertEqual(self.fmt(date(2024, 1, 2)), "2024-01-02")

    def test_datetime_midnight(self):
        self.assertEqual(self.fmt(datetime(2024, 1, 2, 0, 0, 0)), "2024-01-02")

    def test_datetime_with_time(self):
        self.assertEqual(
            self.fmt(datetime(2024, 1, 2, 3, 4, 5)), "2024-01-02 03:04:05"
        )

    def test_time(self):
        self.assertEqual(self.fmt(time(9, 30, 0)), "09:30:00")


class TestAppClass(unittest.TestCase):
    def test_has_expected_methods(self):
        module = importlib.import_module("quickcell")
        self.assertTrue(hasattr(module, "QuickCellApp"))
        self.assertTrue(hasattr(module, "SheetView"))
        for m in [
            "load_file",
            "open_dialog",
            "copy_selection",
            "zoom_delta",
            "zoom_set",
            "show_help",
        ]:
            self.assertTrue(
                hasattr(module.QuickCellApp, m), f"QuickCellApp missing {m}"
            )
        self.assertTrue(callable(getattr(module, "main", None)))


if __name__ == "__main__":
    unittest.main()
