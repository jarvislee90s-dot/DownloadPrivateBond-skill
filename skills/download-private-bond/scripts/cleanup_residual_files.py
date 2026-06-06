"""清理项目临时文件

本脚本在完整流程结束后运行，清理临时文件（保留Download文件夹和已下载的PDF）
"""

import argparse
import glob
from datetime import datetime
from pathlib import Path


def find_project_root():
    """查找项目根目录"""
    current = Path.cwd()
    for path in [current, *current.parents]:
        if (path / "data").exists() and (path / "reference").exists():
            return path
    return current


def clean_residual_files(dry_run=False):
    """清理临时文件（保留Download文件夹和日志）"""
    root = find_project_root()

    patterns_to_clean = [
        "data/*.json",
        "data/*.tmp",
        "output/*.xlsx",
        # 注意：不清理日志文件 (*.txt)，保留 download_log_*.txt 用于排查问题
    ]

    files_to_delete = []
    for pattern in patterns_to_clean:
        for file_path in glob.glob(str(root / pattern)):
            # 保留 .gitkeep 文件
            if Path(file_path).name == ".gitkeep":
                continue
            files_to_delete.append(file_path)

    if not files_to_delete:
        print("[INFO] 无临时文件需要清理")
        return

    print(f"[INFO] 发现 {len(files_to_delete)} 个临时文件")

    if dry_run:
        print("[INFO] 模拟运行模式，以下文件将被删除：")
        for f in files_to_delete:
            print(f"  - {Path(f).name}")
        return

    # 执行清理
    cleaned_count = 0
    for file_path in files_to_delete:
        try:
            Path(file_path).unlink()
            cleaned_count += 1
            print(f"[DEBUG] 已删除: {Path(file_path).name}")
        except Exception as e:
            print(f"[WARN] 无法删除 {file_path}: {e}")

    print(f"[INFO] 清理完成，共删除 {cleaned_count} 个文件")


def main():
    parser = argparse.ArgumentParser(
        description="清理项目临时文件（保留Download文件夹和日志）"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="模拟运行，只显示将要删除的文件而不实际删除",
    )
    args = parser.parse_args()

    print("=" * 50)
    print("开始清理临时文件")
    print(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)

    clean_residual_files(dry_run=args.dry_run)

    print("=" * 50)
    print("清理完成")
    print("=" * 50)


if __name__ == "__main__":
    main()
