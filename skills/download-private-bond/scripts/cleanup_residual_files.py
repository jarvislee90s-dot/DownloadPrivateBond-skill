"""清理项目临时文件

本脚本在完整流程结束后运行，清理临时文件（保留Download文件夹和已下载的PDF）
"""

import argparse
import glob
from datetime import datetime
from pathlib import Path


def find_project_root():
    """查找项目根目录
    优先使用当前工作目录（agent 调用时的项目目录）
    """
    current = Path.cwd()
    if (current / "data").exists() or (current / "output").exists() or (current / "Download").exists():
        return current
    # 回退到 skill 目录
    for path in [current, *current.parents]:
        if (path / "data").exists() and (path / "reference").exists():
            return path
    return current


def collect_residual_files(root):
    root = Path(root)
    patterns_to_clean = [
        "data/*.json",
        "data/*.tmp",
        "output/*.xlsx",
        "output/*.json",
        # 注意：不清理日志文件 (*.txt)，保留 download_log_*.txt 用于排查问题
    ]

    files_to_delete = []
    for pattern in patterns_to_clean:
        for file_path in glob.glob(str(root / pattern)):
            path = Path(file_path)
            if path.name == ".gitkeep":
                continue
            files_to_delete.append(path)
    return files_to_delete


def clean_residual_files(dry_run=False):
    """清理临时文件（保留Download文件夹和日志）"""
    root = find_project_root()

    files_to_delete = collect_residual_files(root)

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
            file_path.unlink()
            cleaned_count += 1
            print(f"[DEBUG] 已删除: {file_path.name}")
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
