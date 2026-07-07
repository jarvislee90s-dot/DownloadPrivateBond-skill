import unittest
import os
import sys
import tempfile
import inspect
import json
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "skills" / "download-private-bond" / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from download_ratingdog_announcements import (
    build_prospectus_search_keyword,
    download_button_xpaths,
    extract_year_from_bond_full_name,
    find_latest_prepared_json,
    group_consecutive_issuers,
    is_target_prospectus_title,
    first_visible_locator,
    visible_table_rows,
    read_pagination_total_count,
    read_bond_rows,
    load_skill_env,
    placeholder_xpath,
    search_button_xpaths,
    normalize_for_match,
    resolve_json_path,
    title_matches_bond,
    search_tenant_announcements,
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

    def test_find_latest_prepared_json_uses_newest_matching_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_dir = Path(tmp)
            old_file = output_dir / "信评需求私募债_20260101_000000.json"
            new_file = output_dir / "信评需求私募债_20260102_000000.json"
            ignored_file = output_dir / "ratingdog_smoke_20260103_000000.json"
            old_file.write_text("old", encoding="utf-8")
            new_file.write_text("new", encoding="utf-8")
            ignored_file.write_text("ignored", encoding="utf-8")
            os.utime(old_file, (100, 100))
            os.utime(new_file, (200, 200))
            os.utime(ignored_file, (300, 300))

            self.assertEqual(find_latest_prepared_json(output_dir), new_file)

    def test_resolve_json_path_prefers_explicit_json_path(self):
        explicit_path = Path("custom.json")

        self.assertEqual(resolve_json_path(explicit_path, None, "missing-output-dir"), explicit_path)

    def test_resolve_json_path_accepts_legacy_excel_path(self):
        excel_path = Path("custom.xlsx")

        self.assertEqual(resolve_json_path(None, excel_path, "missing-output-dir"), Path("custom.json"))

    def test_read_bond_rows_uses_json_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            json_path = Path(tmp) / "bonds.json"
            json_path.write_text(
                json.dumps(
                    [
                        {
                            "bond_short_name": "26嘉州01",
                            "company_name": "乐山国资",
                            "bond_code": "CODE1",
                            "bond_full_name": "乐山国资2026年非公开发行债券",
                            "issue_method": "私募",
                        },
                        {
                            "bond_short_name": "26公募01",
                            "company_name": "公募公司",
                            "bond_code": "CODE2",
                            "bond_full_name": "公募公司2026年公开发行债券",
                            "issue_method": "公募",
                        },
                        {
                            "bond_short_name": "缺全称",
                            "company_name": "坏数据",
                        },
                    ],
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            rows = read_bond_rows(json_path)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["bond_short_name"], "26嘉州01")
        self.assertEqual(rows[0]["company_name"], "乐山国资")
        self.assertEqual(rows[0]["bond_code"], "CODE1")
        self.assertEqual(rows[0]["bond_full_name"], "乐山国资2026年非公开发行债券")

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

    def test_placeholder_xpath_can_include_type(self):
        self.assertEqual(
            placeholder_xpath("请输入密码", input_type="password"),
            "//input[@placeholder='请输入密码' and @type='password']",
        )

    def test_build_prospectus_search_keyword(self):
        self.assertEqual(
            build_prospectus_search_keyword("某公司2026年非公开发行公司债券"),
            "某公司2026年非公开发行公司债券募集说明书",
        )

    def test_search_button_xpath_prefers_input_group_append_button(self):
        self.assertIn(
            "yyep-input-group__append",
            search_button_xpaths()[0],
        )

    def test_download_button_xpath_prefers_file_desc(self):
        self.assertIn("file-desc", download_button_xpaths()[0])

    def test_tenant_search_waits_for_result_count_only_after_search_click(self):
        source = inspect.getsource(search_tenant_announcements)

        self.assertEqual(source.count("wait_for_search_results"), 1)
        self.assertLess(
            source.index("click_search_button_for_input"),
            source.index("wait_for_search_results"),
        )

    def test_first_visible_locator_returns_first_visible_candidate(self):
        class Element:
            def __init__(self, visible):
                self.visible = visible

            def is_displayed(self):
                return self.visible

        class Driver:
            def find_elements(self, by, locator):
                return {
                    "missing": [],
                    "hidden": [Element(False)],
                    "visible": [Element(True)],
                }[locator]

        self.assertIsNot(
            first_visible_locator(
                Driver(),
                [("css", "missing"), ("css", "hidden"), ("css", "visible")],
            ),
            False,
        )

    def test_visible_table_rows_counts_only_displayed_rows(self):
        class Row:
            def __init__(self, visible):
                self.visible = visible

            def is_displayed(self):
                return self.visible

        class Driver:
            def find_elements(self, by, locator):
                return [Row(True), Row(False), Row(True)]

        self.assertEqual(len(visible_table_rows(Driver())), 2)

    def test_read_pagination_total_count_parses_total_text(self):
        class Total:
            text = "共 12284 条"

            def is_displayed(self):
                return True

        class Driver:
            def find_elements(self, by, locator):
                return [Total()]

        self.assertEqual(read_pagination_total_count(Driver()), 12284)


if __name__ == "__main__":
    unittest.main()
