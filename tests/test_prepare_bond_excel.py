import json
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "skills" / "download-private-bond" / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from prepare_bond_excel import (
    load_ocr_records,
    normalize_name,
    remove_public_and_sort_rows,
)


class PrepareBondExcelTests(unittest.TestCase):
    def test_load_ocr_records_requires_company_and_short_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bond_list.json"
            path.write_text(
                json.dumps(
                    [
                        {
                            "company_name": "乐山国有资产投资运营(集团)有限公司",
                            "bond_short_name": "26嘉州01",
                        }
                    ],
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            records = load_ocr_records(path)

        self.assertEqual(records[0]["company_name"], "乐山国有资产投资运营(集团)有限公司")
        self.assertEqual(records[0]["bond_short_name"], "26嘉州01")

    def test_normalize_name_unifies_parentheses_and_spaces(self):
        self.assertEqual(
            normalize_name(" 乐山国有资产投资运营（集团）有限公司 "),
            "乐山国有资产投资运营(集团)有限公司",
        )

    def test_remove_public_and_sort_rows(self):
        rows = [
            {
                "company_name": "乙公司",
                "bond_short_name": "26乙债01",
                "bond_code": "2",
                "bond_full_name": "乙公司2026年非公开发行债券",
                "issue_method": "私募",
            },
            {
                "company_name": "甲公司",
                "bond_short_name": "26甲债01",
                "bond_code": "1",
                "bond_full_name": "甲公司2026年公开发行债券",
                "issue_method": "公募",
            },
            {
                "company_name": "甲公司",
                "bond_short_name": "26甲私01",
                "bond_code": "3",
                "bond_full_name": "甲公司2026年非公开发行债券",
                "issue_method": "私募",
            },
        ]

        result = remove_public_and_sort_rows(rows)

        self.assertEqual([row["company_name"] for row in result], ["甲公司", "乙公司"])
        self.assertEqual([row["bond_short_name"] for row in result], ["26甲私01", "26乙债01"])


if __name__ == "__main__":
    unittest.main()
