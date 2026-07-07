import sys
import tempfile
import unittest
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "skills" / "download-private-bond" / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from cleanup_residual_files import collect_residual_files


class CleanupResidualFilesTests(unittest.TestCase):
    def test_collect_residual_files_includes_output_json(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "data").mkdir()
            (root / "reference").mkdir()
            (root / "output").mkdir()
            (root / "data" / "bond_list.json").write_text("[]", encoding="utf-8")
            (root / "output" / "信评需求私募债_20260606_133633.xlsx").write_text("xlsx", encoding="utf-8")
            (root / "output" / "信评需求私募债_20260606_133633.json").write_text("[]", encoding="utf-8")
            (root / "output" / "download_log_20260606_133633.txt").write_text("log", encoding="utf-8")

            files = {path.name for path in collect_residual_files(root)}

        self.assertIn("bond_list.json", files)
        self.assertIn("信评需求私募债_20260606_133633.xlsx", files)
        self.assertIn("信评需求私募债_20260606_133633.json", files)
        self.assertNotIn("download_log_20260606_133633.txt", files)


if __name__ == "__main__":
    unittest.main()
