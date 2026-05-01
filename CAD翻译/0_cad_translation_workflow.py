#!/usr/bin/env python3
"""
CAD翻译全自动流水线（DXF）
直接运行即可完成：提取 -> 生成待翻译项 -> 导入翻译 -> 输出翻译CAD。
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "00_dxf文件"
PENDING_FILE = BASE_DIR / "2_提取文件合并文件" / "待翻译项.md"
SPLIT_DIR = BASE_DIR / "2_提取文件合并文件" / "拆分文件"
TRANSLATE_DIR = BASE_DIR / "3_翻译后文件"
JSON_PATH = BASE_DIR / "翻译对照.json"
OUTPUT_DIR = BASE_DIR / "4_输出文件cad"
LINES_PER_SPLIT = 200


def run_step(label: str, args: list[str]) -> None:
    print(f"\n{'=' * 60}")
    print(f"[步骤] {label}")
    print(f"{'=' * 60}")
    subprocess.run(args, check=True, cwd=BASE_DIR)


def split_pending_file(
    input_path: Path = PENDING_FILE,
    output_dir: Path = SPLIT_DIR,
    translate_dir: Path = TRANSLATE_DIR,
    lines_per_file: int = LINES_PER_SPLIT,
) -> None:
    if not input_path.exists():
        print(f"文件不存在: {input_path}")
        return

    lines = input_path.read_text(encoding="utf-8").splitlines()
    if not lines:
        print("待翻译项为空，跳过拆分")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    translate_dir.mkdir(parents=True, exist_ok=True)

    part = 0
    total = len(lines)
    for start in range(0, total, lines_per_file):
        part += 1
        chunk = lines[start : start + lines_per_file]
        out_path = output_dir / f"待翻译项_第{part}部分.md"
        out_path.write_text("\n".join(chunk) + "\n", encoding="utf-8")

        trans_path = translate_dir / f"翻译_第{part}部分.md"
        if not trans_path.exists():
            trans_path.write_text("", encoding="utf-8")

        print(f"  [{part}] 第 {start + 1}-{start + len(chunk)} 行 -> {out_path.name} | {trans_path.name}")

    print(f"拆分完成: 共 {total} 行 -> {part} 个文件")


def main() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"待翻译CAD目录: {INPUT_DIR}")
    print(f"翻译JSON:      {JSON_PATH}")
    print(f"输出目录:      {OUTPUT_DIR}")

    run_step(
        "提取CAD泰文文本并生成待翻译项",
        [
            sys.executable,
            str(BASE_DIR / "1_cad_extract_thai_text.py"),
            str(INPUT_DIR),
        ],
    )

    print(f"\n{'=' * 60}")
    print("[步骤] 拆分待翻译项（供翻译填充）")
    print(f"{'=' * 60}")
    split_pending_file()

    run_step(
        "将翻译后md增量写入翻译JSON",
        [
            sys.executable,
            str(BASE_DIR / "2_compare_translation.py"),
            "--source-file", str(PENDING_FILE),
            "--translate-dir", str(TRANSLATE_DIR),
            "--json-path", str(JSON_PATH),
            "--output-dir", str(OUTPUT_DIR),
        ],
    )

    run_step(
        "根据翻译JSON输出CAD",
        [
            sys.executable,
            str(BASE_DIR / "3_cad_apply_translation_from_json.py"),
            str(INPUT_DIR),
            "--json-path", str(JSON_PATH),
            "--output-dir", str(OUTPUT_DIR),
        ],
    )

    print(f"\n{'=' * 60}")
    print("CAD翻译全流程完成")
    print(f"待翻译项: {PENDING_FILE}")
    print(f"翻译分片: {SPLIT_DIR}")
    print(f"翻译JSON: {JSON_PATH}")
    print(f"输出CAD:  {OUTPUT_DIR}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
