#!/usr/bin/env python3
"""
Excel翻译全自动流水线
======================
直接运行即可完成全流程，无需选择：
    1. 从 0_待翻译文件 提取所有Excel文本 → 生成 待翻译项.csv 及分片
  2. 将 3_翻译后文件 里已填好的译文增量写入 翻译对照.json
  3. 根据 翻译对照.json 对 0_待翻译文件 里的Excel输出翻译版到 4_输出文件excel

重复运行：每次运行都会重新提取并更新待翻译项，同时把 3_翻译后文件 里新增的译文
增量并入 json，再重新输出翻译后的Excel，实现持续更新。
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "0_待翻译文件"
MERGED_DIR = BASE_DIR / "2_提取文件合并文件"
PREPARE_SOURCE_FILE = MERGED_DIR / "待翻译项.csv"
TRANSLATE_DIR = BASE_DIR / "3_翻译后文件"
JSON_PATH = BASE_DIR / "翻译对照.json"
OUTPUT_DIR = BASE_DIR / "4_输出文件excel"
OUTPUT_EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}
OUTPUT_PREFIX = "翻译"


def run_step(label: str, args: list[str]) -> None:
    print(f"\n{'='*60}")
    print(f"[步骤] {label}")
    print(f"{'='*60}")
    subprocess.run(args, check=True, cwd=BASE_DIR)


def cleanup_md_files(target_dir: Path) -> int:
    if not target_dir.exists():
        return 0

    deleted = 0
    for md_file in target_dir.rglob("*.md"):
        if md_file.is_file():
            md_file.unlink()
            deleted += 1
    return deleted


def add_prefix_to_output_excels(target_dir: Path, prefix: str) -> int:
    if not target_dir.exists():
        return 0

    renamed = 0
    for excel_file in sorted(target_dir.rglob("*")):
        if not excel_file.is_file():
            continue
        if excel_file.suffix.lower() not in OUTPUT_EXCEL_SUFFIXES:
            continue
        if excel_file.name.startswith("~$") or excel_file.name.startswith(prefix):
            continue

        new_path = excel_file.with_name(f"{prefix}{excel_file.name}")
        if new_path.exists():
            continue
        excel_file.rename(new_path)
        renamed += 1
    return renamed


def main() -> None:
    print(f"待翻译Excel目录: {INPUT_DIR}")
    print(f"翻译对照JSON:    {JSON_PATH}")
    print(f"输出目录:        {OUTPUT_DIR}")

    steps: list[tuple[str, list[str]]] = [
        (
            "提取Excel文本 → 生成待翻译项.csv",
            [
                sys.executable,
                str(BASE_DIR / "1_excel_to_txt_all_cells.py"),
                str(INPUT_DIR),
            ],
        ),
        (
            "将翻译后文件增量写入翻译对照.json",
            [
                sys.executable,
                str(BASE_DIR / "2_compare_translation.py"),
                "--source-file", str(PREPARE_SOURCE_FILE),
                "--translate-dir", str(TRANSLATE_DIR),
                "--json-path", str(JSON_PATH),
                "--output-dir", str(OUTPUT_DIR),
                "--update-existing",
            ],
        ),
        (
            "根据翻译对照.json 输出翻译后的Excel",
            [
                sys.executable,
                str(BASE_DIR / "3_excel_apply_translation_from_json.py"),
                str(INPUT_DIR),
                "--json-path", str(JSON_PATH),
                "--output-dir", str(OUTPUT_DIR),
            ],
        ),
    ]

    for label, args in steps:
        run_step(label, args)

    renamed_excel = add_prefix_to_output_excels(OUTPUT_DIR, OUTPUT_PREFIX)
    deleted_md = cleanup_md_files(MERGED_DIR)

    print(f"\n{'='*60}")
    print("全流程完成！")
    print(f"  待翻译项.csv → {PREPARE_SOURCE_FILE}")
    print(f"  翻译JSON    → {JSON_PATH}")
    print(f"  输出Excel   → {OUTPUT_DIR}")
    print(f"  已加前缀    → {OUTPUT_PREFIX} (处理 {renamed_excel} 个Excel)")
    print(f"  已清理MD    → {MERGED_DIR} (删除 {deleted_md} 个)")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()