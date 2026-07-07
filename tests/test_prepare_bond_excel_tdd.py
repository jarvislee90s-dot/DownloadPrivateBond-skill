"""Tests for prepare_bond_excel module."""

import json
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "skills" / "download-private-bond" / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from prepare_bond_excel import (
    load_ocr_records,
    extract_company_from_bond_full_name,
)


class LoadOcrRecordsTests(unittest.TestCase):
    """Test loading OCR records with various formats."""

    def test_load_records_with_company_and_short_name(self):
        """Should accept records with both company_name and bond_short_name."""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
            json.dump([
                {"company_name": "杭州临平开发投资集团有限公司", "bond_short_name": "26杭开F1"},
            ], f, ensure_ascii=False)
            f.flush()

            records = load_ocr_records(f.name)

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["company_name"], "杭州临平开发投资集团有限公司")
        self.assertEqual(records[0]["bond_short_name"], "26杭开F1")

    def test_load_records_with_bond_short_name_only(self):
        """Should accept records with only bond_short_name (company_name via WIND)."""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
            json.dump([
                {"bond_short_name": "26杭开F1"},
            ], f, ensure_ascii=False)
            f.flush()

            records = load_ocr_records(f.name)

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["company_name"], "")
        self.assertEqual(records[0]["bond_short_name"], "26杭开F1")

    def test_load_records_mixed_format(self):
        """Should accept mixed format records."""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
            json.dump([
                {"company_name": "杭州临平", "bond_short_name": "26杭开F1"},
                {"bond_short_name": "26嘉州01"},
            ], f, ensure_ascii=False)
            f.flush()

            records = load_ocr_records(f.name)

        self.assertEqual(len(records), 2)
        # First record with company name
        self.assertEqual(records[0]["company_name"], "杭州临平")
        self.assertEqual(records[0]["bond_short_name"], "26杭开F1")
        # Second record without company name
        self.assertEqual(records[1]["company_name"], "")
        self.assertEqual(records[1]["bond_short_name"], "26嘉州01")

    def test_rejects_records_without_bond_short_name(self):
        """Should reject records without bond_short_name."""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
            json.dump([
                {"company_name": "杭州临平"},
            ], f, ensure_ascii=False)
            f.flush()

            with self.assertRaises(ValueError) as ctx:
                load_ocr_records(f.name)

            self.assertIn("债券简称", str(ctx.exception))


class ExtractCompanyFromBondFullNameTests(unittest.TestCase):
    """Test extracting company name from bond full name."""

    def test_extract_from_standard_format(self):
        """Should extract company name from standard bond full name."""
        bond_full_name = "杭州临平开发投资集团有限公司2026年面向专业投资者非公开发行公司债券(第一期)"
        result = extract_company_from_bond_full_name(bond_full_name)
        self.assertEqual(result, "杭州临平开发投资集团有限公司")

    def test_extract_with_varieties(self):
        """Should extract company name even with variety suffix."""
        bond_full_name = "杭州高新技术产业开发区资产经营有限公司2026年面向专业投资者非公开发行公司债券(第一期)(品种一)"
        result = extract_company_from_bond_full_name(bond_full_name)
        self.assertEqual(result, "杭州高新技术产业开发区资产经营有限公司")

    def test_empty_input(self):
        """Should return empty string for empty input."""
        result = extract_company_from_bond_full_name("")
        self.assertEqual(result, "")

    def test_none_input(self):
        """Should return empty string for None input."""
        result = extract_company_from_bond_full_name(None)
        self.assertEqual(result, "")


if __name__ == "__main__":
    unittest.main()
