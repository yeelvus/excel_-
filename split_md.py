#!/usr/bin/env python3
"""将待翻译项.md按固定行数拆分为多个md文件。"""
import argparse
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
LINES_PER_FILE = 200
INPUT_PATH = BASE_DIR / "2_提取文件合并文件" / "待翻译项.md"
OUTPUT_DIR = BASE_DIR / "2_提取文件合并文件" / "拆分文件"
TRANSLATE_DIR = BASE_DIR / "3_翻译后文件"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="将待翻译项拆分成多个md，并在翻译目录创建对应空文件"
    )
    parser.add_argument(
        "--input-file",
        default=str(INPUT_PATH),
        help="待拆分文件，默认使用 2_提取文件合并文件/待翻译项.md",
    )
    parser.add_argument(
        "--output-dir",
        default=str(OUTPUT_DIR),
        help="拆分后的原文md输出目录",
    )
    parser.add_argument(
        "--translate-dir",
        default=str(TRANSLATE_DIR),
        help="翻译结果md目录",
    )
    parser.add_argument(
        "--lines-per-file",
        type=int,
        default=LINES_PER_FILE,
        help="每个拆分文件的行数，默认 200",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input_file).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    translate_dir = Path(args.translate_dir).expanduser().resolve()

    if args.lines_per_file <= 0:
        raise ValueError("--lines-per-file 必须大于 0")

    if not input_path.exists():
        print(f"文件不存在: {input_path}")
        return

    lines = input_path.read_text(encoding="utf-8").splitlines()
    total = len(lines)
    if total == 0:
        print("文件为空")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    translate_dir.mkdir(parents=True, exist_ok=True)

    part = 0
    for start in range(0, total, args.lines_per_file):
        part += 1
        chunk = lines[start : start + args.lines_per_file]
        out_path = output_dir / f"待翻译项_第{part}部分.md"
        out_path.write_text("\n".join(chunk) + "\n", encoding="utf-8")
        # 在翻译目录创建对应的空文件
        trans_path = translate_dir / f"翻译_第{part}部分.md"
        if not trans_path.exists():
            trans_path.write_text("", encoding="utf-8")
        print(f"[{part}] 第 {start + 1}-{start + len(chunk)} 行 → {out_path.name} | {trans_path.name}")

    print(f"\n完成！共 {total} 行，拆分为 {part} 个文件 → {output_dir}")


if __name__ == "__main__":
    main()
