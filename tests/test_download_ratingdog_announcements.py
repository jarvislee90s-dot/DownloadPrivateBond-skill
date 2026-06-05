import unittest
import os
import sys
import tempfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "skills" / "download-private-bond" / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from download_ratingdog_announcements import (
    extract_year_from_bond_full_name,
    find_latest_prepared_excel,
    group_consecutive_issuers,
    is_target_prospectus_title,
    load_skill_env,
    normalize_for_match,
    resolve_excel_path,
    title_matches_bond,
)


class DownloadRatingdogAnnouncementsTests(unittest.TestCase):
    def test_extract_year_from_bond_full_name(self):
        self.assertEqual(
            extract_year_from_bond_full_name("某公司2025年非公开发行公司债券"),
            2025,
        )

    def test_normalize_for_match_unifies_symbols(self):
        self.assertEqual(
            normalize_for_match("A（第一期） “测试”"),
            "A(第一期)测试",
        )

    def test_title_matches_bond_allows_prefix_text(self):
        bond_full_name = "乐山国有资产投资运营(集团)有限公司2026年面向专业投资者非公开发行低碳转型挂钩公司债券(第一期)募集说明书"
        title = "中诚信低碳转型挂钩债券独立评估报告-乐山国有资产投资运营（集团）有限公司2026年面向专业投资者非公开发行低碳转型挂钩公司债券（第一期）募集说明书"
        self.assertTrue(title_matches_bond(title, bond_full_name))

    def test_target_title_requires_prospectus(self):
        bond_full_name = "安徽省铁路发展基金股份有限公司2026年面向专业投资者非公开发行公司债券(第一期)"
        prospectus_title = "安徽省铁路发展基金股份有限公司2026年面向专业投资者非公开发行公司债券（第一期）募集说明书"
        result_title = "安徽省铁路发展基金股份有限公司2026年面向专业投资者非公开发行公司债券（第一期）发行结果公告"

        self.assertTrue(is_target_prospectus_title(prospectus_title, bond_full_name))
        self.assertFalse(is_target_prospectus_title(result_title, bond_full_name))

    def test_group_consecutive_issuers_keeps_adjacent_same_company(self):
        rows = [
            {"company_name": "甲公司", "bond_full_name": "甲公司2026年债券A"},
            {"company_name": "甲公司", "bond_full_name": "甲公司2026年债券B"},
            {"company_name": "乙公司", "bond_full_name": "乙公司2026年债券A"},
        ]

        groups = group_consecutive_issuers(rows)

        self.assertEqual([group[0]["company_name"] for group in groups], ["甲公司", "乙公司"])
        self.assertEqual(len(groups[0]), 2)
        self.assertEqual(len(groups[1]), 1)

    def test_find_latest_prepared_excel_uses_newest_matching_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_dir = Path(tmp)
            old_file = output_dir / "信评需求私募债_20260101_000000.xlsx"
            new_file = output_dir / "信评需求私募债_20260102_000000.xlsx"
            ignored_file = output_dir / "ratingdog_smoke_20260103_000000.xlsx"
            old_file.write_text("old", encoding="utf-8")
            new_file.write_text("new", encoding="utf-8")
            ignored_file.write_text("ignored", encoding="utf-8")
            os.utime(old_file, (100, 100))
            os.utime(new_file, (200, 200))
            os.utime(ignored_file, (300, 300))

            self.assertEqual(find_latest_prepared_excel(output_dir), new_file)

    def test_resolve_excel_path_prefers_explicit_path(self):
        explicit_path = Path("custom.xlsx")

        self.assertEqual(resolve_excel_path(explicit_path, "missing-output-dir"), explicit_path)

    def test_load_skill_env_accepts_colon_prefixed_key(self):
        with tempfile.TemporaryDirectory() as tmp:
            env_path = Path(tmp) / ".env"
            env_path.write_text(
                "RATINGDOG_USERNAME=user\n:RATINGDOG_PASSWORD=pass\n",
                encoding="utf-8",
            )
            old_username = os.environ.pop("RATINGDOG_USERNAME", None)
            old_password = os.environ.pop("RATINGDOG_PASSWORD", None)
            try:
                load_skill_env(env_path)

                self.assertEqual(os.environ["RATINGDOG_USERNAME"], "user")
                self.assertEqual(os.environ["RATINGDOG_PASSWORD"], "pass")
            finally:
                os.environ.pop("RATINGDOG_USERNAME", None)
                os.environ.pop("RATINGDOG_PASSWORD", None)
                if old_username is not None:
                    os.environ["RATINGDOG_USERNAME"] = old_username
                if old_password is not None:
                    os.environ["RATINGDOG_PASSWORD"] = old_password


if __name__ == "__main__":
    unittest.main()
