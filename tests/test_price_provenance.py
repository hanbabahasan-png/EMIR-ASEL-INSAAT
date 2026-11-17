"""Regression tests: imported Excel prices must never acquire invented provenance."""
import ast
import math
import unittest
import numpy as np
from datetime import datetime
from pathlib import Path

SOURCE = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
TREE = ast.parse(SOURCE)
NODE = next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == "normalize_import_price_provenance")


def missing(value):
    return value is None or (isinstance(value, float) and math.isnan(value))


def num(value, default=0.0):
    if missing(value):
        return default
    try:
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return default


def now_tr():
    return datetime(2026, 11, 16)

NS = {"math": math, "np": np, "missing": missing, "num": num, "now_tr": now_tr}
exec(compile(ast.Module(body=[NODE], type_ignores=[]), "<price-provenance>", "exec"), NS)
normalize = NS["normalize_import_price_provenance"]


class ImportedPriceProvenanceTests(unittest.TestCase):
    def test_missing_year_and_source_never_become_current(self):
        price, year, source, note = normalize(30500, None, None)
        self.assertTrue(math.isnan(price))
        self.assertTrue(math.isnan(year))
        self.assertEqual(source, "")
        self.assertIn("fiyat yılı", note)
        self.assertIn("kaynak", note)
        self.assertIn("30500", note)

    def test_missing_source_blocks_price(self):
        price, year, source, note = normalize(3700, 2026, "")
        self.assertTrue(math.isnan(price))
        self.assertTrue(math.isnan(year))
        self.assertEqual(source, "")
        self.assertIn("kaynak", note)

    def test_explicit_recent_source_and_year_are_preserved(self):
        self.assertEqual(normalize(4200, 2026, "Supplier quotation 2026-10-20"),
                         (4200.0, 2026, "Supplier quotation 2026-10-20", ""))

    def test_old_price_is_held_for_requotation(self):
        price, year, source, note = normalize(30500, 2024, "Old workbook quote")
        self.assertTrue(math.isnan(price))
        self.assertTrue(math.isnan(year))
        self.assertEqual(source, "")
        self.assertIn("eski kabul edildi", note)

    def test_future_year_is_not_accepted(self):
        price, year, source, note = normalize(30500, 2027, "Future dated source")
        self.assertTrue(math.isnan(price))
        self.assertTrue(math.isnan(year))
        self.assertIn("fiyat yılı", note)

    def test_excel_import_calls_provenance_guard(self):
        node = next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == "build_excel_takeoff_rows")
        calls = {n.func.id for n in ast.walk(node) if isinstance(n, ast.Call) and isinstance(n.func, ast.Name)}
        self.assertIn("normalize_import_price_provenance", calls)
        source_segment = ast.get_source_segment(SOURCE, node)
        self.assertNotIn('f"Excel {Path(source_name).name}', source_segment)


if __name__ == "__main__":
    unittest.main()
