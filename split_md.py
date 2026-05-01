#!/usr/bin/env python3
"""将过滤后文本.md按每300行拆分为多个md文件。"""
from pathlib import Path

LINES_PER_FILE = 200
INPUT_PATH = Path(__file__).resolve().parent / "2_提取文件合并文件" / "过滤后文本.md"
OUTPUT_DIR = Path(__file__).resolve().parent / "2_提取文件合并文件" / "拆分文件"
TRANSLATE_DIR = Path(__file__).resolve().parent / "3_翻译后文件"


def main() -> None:
    if not INPUT_PATH.exists():
        print(f"文件不存在: {INPUT_PATH}")
        return

    lines = INPUT_PATH.read_text(encoding="utf-8").splitlines()
    total = len(lines)
    if total == 0:
        print("文件为空")
        return

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    TRANSLATE_DIR.mkdir(parents=True, exist_ok=True)

    part = 0
    for start in range(0, total, LINES_PER_FILE):
        part += 1
        chunk = lines[start : start + LINES_PER_FILE]
        out_path = OUTPUT_DIR / f"过滤后文本_第{part}部分.md"
        out_path.write_text("\n".join(chunk) + "\n", encoding="utf-8")
        # 在翻译目录创建对应的空文件
        trans_path = TRANSLATE_DIR / f"翻译_第{part}部分.md"
        if not trans_path.exists():
            trans_path.write_text("", encoding="utf-8")
        print(f"[{part}] 第 {start + 1}-{start + len(chunk)} 行 → {out_path.name} | {trans_path.name}")

    print(f"\n完成！共 {total} 行，拆分为 {part} 个文件 → {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
