#!/usr/bin/env python3
"""
Excel翻译全自动流水线
======================
直接运行即可完成全流程，无需选择：
  1. 从 0_待翻译文件 提取所有Excel文本 → 生成 待翻译项.md 及分片
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
PREPARE_SOURCE_FILE = BASE_DIR / "2_提取文件合并文件" / "待翻译项.md"
TRANSLATE_DIR = BASE_DIR / "3_翻译后文件"
JSON_PATH = BASE_DIR / "翻译对照.json"
OUTPUT_DIR = BASE_DIR / "4_输出文件excel"
LINES_PER_SPLIT = 200


def run_step(label: str, args: list[str]) -> None:
    print(f"\n{'='*60}")
    print(f"[步骤] {label}")
    print(f"{'='*60}")
    subprocess.run(args, check=True, cwd=BASE_DIR)


def main() -> None:
    print(f"待翻译Excel目录: {INPUT_DIR}")
    print(f"翻译对照JSON:    {JSON_PATH}")
    print(f"输出目录:        {OUTPUT_DIR}")

    # 步骤1：提取Excel文本 → 生成 待翻译项.md 及分片
    run_step(
        "提取Excel文本 → 生成待翻译项.md 及分片",
        [
            sys.executable,
            str(BASE_DIR / "1_excel_to_txt_all_cells.py"),
            str(INPUT_DIR),
        ],
    )
    run_step(
        "将 待翻译项.md 按行数拆分（供翻译使用）",
        [
            sys.executable,
            str(BASE_DIR / "split_md.py"),
            "--input-file", str(PREPARE_SOURCE_FILE),
            "--lines-per-file", str(LINES_PER_SPLIT),
        ],
    )

    # 步骤2：将 3_翻译后文件 里已填好的译文增量写入 翻译对照.json
    run_step(
        "将翻译后md增量写入翻译对照.json",
        [
            sys.executable,
            str(BASE_DIR / "2_compare_translation.py"),
            "--source-file", str(PREPARE_SOURCE_FILE),
            "--translate-dir", str(TRANSLATE_DIR),
            "--json-path", str(JSON_PATH),
            "--output-dir", str(OUTPUT_DIR),
        ],
    )

    # 步骤3：根据 json 对 0_待翻译文件 里的Excel输出翻译版
    run_step(
        "根据翻译对照.json 输出翻译后的Excel",
        [
            sys.executable,
            str(BASE_DIR / "3_excel_apply_translation_from_json.py"),
            str(INPUT_DIR),
            "--json-path", str(JSON_PATH),
            "--output-dir", str(OUTPUT_DIR),
        ],
    )

    print(f"\n{'='*60}")
    print("全流程完成！")
    print(f"  待翻译项.md → {PREPARE_SOURCE_FILE}")
    print(f"  分片文件    → {BASE_DIR / '2_提取文件合并文件' / '拆分文件'}")
    print(f"  翻译JSON    → {JSON_PATH}")
    print(f"  输出Excel   → {OUTPUT_DIR}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()