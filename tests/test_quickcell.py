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
        self.assertEqual(module.VERSION, "1.1.0")


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


class _FakeCtx:
    """Minimal ctx matching the interface the parser expects."""

    def __init__(self, cells, max_row=1000, max_col=100):
        self.cells = cells
        self.max_row = max_row
        self.max_col = max_col

    def get_value(self, sheet, row, col):
        return self.cells.get((row, col))

    def get_sheet_dims(self, sheet):
        return (self.max_row, self.max_col)


class TestFormulaEvaluator(unittest.TestCase):
    def setUp(self):
        self.qc = importlib.import_module("quickcell")
        self.cells = {}
        self.ctx = _FakeCtx(self.cells)

    def eval(self, src):
        return self.qc._FormulaParser(src, self.ctx, None).parse()

    def test_literal_number(self):
        self.assertEqual(self.eval("=42"), 42)
        self.assertEqual(self.eval("=3.14"), 3.14)

    def test_arithmetic(self):
        self.assertEqual(self.eval("=1+2*3"), 7)
        self.assertEqual(self.eval("=(1+2)*3"), 9)
        self.assertEqual(self.eval("=10/4"), 2.5)
        self.assertEqual(self.eval("=2^10"), 1024)
        self.assertEqual(self.eval("=-5 + 3"), -2)

    def test_divide_by_zero(self):
        with self.assertRaises(self.qc.FormulaError):
            self.eval("=1/0")

    def test_strings_and_concat(self):
        self.assertEqual(self.eval('="foo" & "bar"'), "foobar")
        self.assertEqual(self.eval('="he said ""hi"""'), 'he said "hi"')

    def test_booleans_and_compare(self):
        self.assertTrue(self.eval("=1<2"))
        self.assertFalse(self.eval("=1>=2"))
        self.assertTrue(self.eval("=1=1"))
        self.assertTrue(self.eval("=1<>2"))

    def test_cell_ref(self):
        self.cells[(1, 1)] = 10
        self.cells[(2, 1)] = 5
        self.assertEqual(self.eval("=A1+A2"), 15)
        self.assertEqual(self.eval("=$A$1*2"), 20)

    def test_range_and_sum(self):
        for r in range(1, 6):
            self.cells[(r, 1)] = r
        self.assertEqual(self.eval("=SUM(A1:A5)"), 15)
        self.assertEqual(self.eval("=AVERAGE(A1:A5)"), 3)
        self.assertEqual(self.eval("=MIN(A1:A5)"), 1)
        self.assertEqual(self.eval("=MAX(A1:A5)"), 5)
        self.assertEqual(self.eval("=COUNT(A1:A5)"), 5)

    def test_if(self):
        self.assertEqual(self.eval('=IF(1>0, "yes", "no")'), "yes")
        self.assertEqual(self.eval('=IF(1<0, "yes", "no")'), "no")

    def test_string_funcs(self):
        self.assertEqual(self.eval('=LEN("hello")'), 5)
        self.assertEqual(self.eval('=UPPER("abc")'), "ABC")
        self.assertEqual(self.eval('=LEFT("hello", 3)'), "hel")
        self.assertEqual(self.eval('=RIGHT("hello", 2)'), "lo")
        self.assertEqual(self.eval('=MID("hello", 2, 3)'), "ell")
        self.assertEqual(self.eval('=TRIM("  a  b  ")'), "a b")

    def test_unknown_function_errors(self):
        with self.assertRaises(self.qc.FormulaError):
            self.eval("=VLOOKUP(1,A1:B2,2,FALSE)")

    def test_if_short_circuits_true_branch(self):
        # A1=0 so 1/A1 would raise — but IF should skip the else branch.
        self.cells[(1, 1)] = 0
        self.assertEqual(self.eval('=IF(A1=0, "zero", 1/A1)'), "zero")

    def test_if_short_circuits_false_branch(self):
        self.cells[(1, 1)] = 5
        self.assertEqual(self.eval('=IF(A1=0, 1/A1, "ok")'), "ok")

    def test_iferror_catches(self):
        self.cells[(1, 1)] = 0
        self.assertEqual(self.eval('=IFERROR(1/A1, "oops")'), "oops")
        self.assertEqual(self.eval('=IFERROR(10/2, "oops")'), 5)

    def test_whole_column_sum(self):
        self.cells[(1, 1)] = 10
        self.cells[(2, 1)] = 20
        self.cells[(3, 1)] = 30
        self.ctx.max_row = 3
        self.assertEqual(self.eval("=SUM(A:A)"), 60)

    def test_match_and_index(self):
        for i, v in enumerate(["apple", "banana", "cherry"], 1):
            self.cells[(i, 1)] = v
        for i, v in enumerate([1.0, 2.0, 3.0], 1):
            self.cells[(i, 2)] = v
        self.ctx.max_row = 3
        self.assertEqual(self.eval('=MATCH("banana",A:A,0)'), 2)
        self.assertEqual(self.eval('=INDEX(B:B, MATCH("cherry",A:A,0))'), 3.0)

    def test_value(self):
        self.assertEqual(self.eval('=VALUE("42")'), 42)

    def test_countif_exact(self):
        for r, v in enumerate(["a", "b", "a", "c", "a"], 1):
            self.cells[(r, 1)] = v
        self.ctx.max_row = 5
        self.assertEqual(self.eval('=COUNTIF(A:A,"a")'), 3)

    def test_countif_comparison(self):
        for r, v in enumerate([1, 5, 10, 3, 8], 1):
            self.cells[(r, 1)] = v
        self.ctx.max_row = 5
        self.assertEqual(self.eval('=COUNTIF(A:A,">4")'), 3)
        self.assertEqual(self.eval('=COUNTIF(A:A,"<=3")'), 2)
        self.assertEqual(self.eval('=COUNTIF(A:A,"<>5")'), 4)

    def test_countif_wildcard(self):
        for r, v in enumerate(["apple", "apricot", "banana", "avocado"], 1):
            self.cells[(r, 1)] = v
        self.ctx.max_row = 4
        self.assertEqual(self.eval('=COUNTIF(A:A,"a*")'), 3)
        self.assertEqual(self.eval('=COUNTIF(A:A,"?a*")'), 1)  # only banana

    def test_sumif(self):
        for r, (k, v) in enumerate([("x", 1), ("y", 2), ("x", 3), ("y", 4)], 1):
            self.cells[(r, 1)] = k
            self.cells[(r, 2)] = v
        self.ctx.max_row = 4
        self.assertEqual(self.eval('=SUMIF(A:A,"x",B:B)'), 4)

    def test_double_equals_prefix(self):
        # Some files store formulas with a leading "==" — tolerate it.
        self.cells[(1, 1)] = 10
        self.assertEqual(self.eval("==A1+5"), 15)

    def test_nested_iferror_user_style(self):
        self.cells[(1, 1)] = 0
        self.cells[(2, 1)] = ""
        # mirrors the user's pattern: OR(A=0, B="") → ""
        self.assertEqual(
            self.eval('=IF(OR(A1=0,A2=""),"",ROUND(A1/A2,2))'), ""
        )


class _MultiSheetCtx:
    def __init__(self, sheets):
        self.sheets = sheets  # {name: {(r,c): value}}

    def get_value(self, sheet, row, col):
        s = self.sheets.get(sheet)
        if s is None:
            raise importlib.import_module("quickcell").FormulaError(
                f"unknown sheet: {sheet!r}"
            )
        return s.get((row, col))

    def get_sheet_dims(self, sheet):
        s = self.sheets.get(sheet)
        if s is None:
            return None
        if not s:
            return (1, 1)
        rows = [r for r, _ in s.keys()]
        cols = [c for _, c in s.keys()]
        return (max(rows), max(cols))


class TestCrossSheet(unittest.TestCase):
    def setUp(self):
        self.qc = importlib.import_module("quickcell")
        self.ctx = _MultiSheetCtx(
            {
                "Main": {(1, 1): 10},
                "tbl_afmetingen": {
                    (1, 1): "X",
                    (2, 1): "Y",
                    (3, 1): "Z",
                    (1, 3): 100,
                    (2, 3): 200,
                    (3, 3): 300,
                },
                "Sheet With Space": {(5, 5): "hello"},
            }
        )

    def eval(self, src, sheet="Main"):
        return self.qc._FormulaParser(src, self.ctx, sheet).parse()

    def test_cross_sheet_ref(self):
        self.assertEqual(self.eval("=tbl_afmetingen!C2"), 200)

    def test_cross_sheet_quoted(self):
        self.assertEqual(self.eval("='Sheet With Space'!E5"), "hello")

    def test_cross_sheet_range(self):
        self.assertEqual(
            self.eval('=INDEX(tbl_afmetingen!C:C,MATCH("Y",tbl_afmetingen!A:A,0))'),
            200,
        )

    def test_nested_iferror_cross_sheet(self):
        # user's formula: match on string key, fall back to VALUE-converted
        # key, fall back to "" — exercises IFERROR recovery after a
        # deeply-nested raise across cross-sheet refs.
        self.ctx.sheets["Main"][(1, 1)] = "missing"
        self.assertEqual(
            self.eval(
                '=IFERROR(INDEX(tbl_afmetingen!C:C,MATCH(A1,tbl_afmetingen!A:A,0)),'
                'IFERROR(INDEX(tbl_afmetingen!C:C,MATCH(VALUE(A1),tbl_afmetingen!A:A,0)),""))'
            ),
            "",
        )


if __name__ == "__main__":
    unittest.main()
