from pathlib import Path
import importlib.util
import tempfile
import unittest

from tools.demo_openpyxl import run_demo


@unittest.skipUnless(importlib.util.find_spec("openpyxl"), "optional demo dependency not installed")
class OpenpyxlDemoTests(unittest.TestCase):
    def test_same_cells_and_locators_can_hide_loss_of_formula_evidence(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            first = run_demo(root / "first")
            second = run_demo(root / "second")
            self.assertEqual((root / "first" / "comparison.json").read_bytes(),
                             (root / "second" / "comparison.json").read_bytes())
            self.assertEqual({"cells": 5, "formula_cells": 1}, first["expected"])
            self.assertEqual(1, first["modes"]["formulas"]["formula_cells"])
            self.assertEqual(0, first["modes"]["cached_values"]["formula_cells"])
            self.assertTrue(first["expected_tradeoff_observed"])
            self.assertEqual(0, first["network_attempts"])
            self.assertEqual(0, first["private_sources_opened"])
            self.assertEqual(first["report_sha256"], second["report_sha256"])
            with self.assertRaises(FileExistsError):
                run_demo(root / "first")
