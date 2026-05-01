#!/usr/bin/env python3
"""
对比待翻译项与翻译文件，增量更新翻译对照 JSON。
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
SOURCE_FILE = BASE_DIR / "2_提取文件合并文件" / "待翻译项.md"
TRANSLATE_DIR = BASE_DIR / "3_翻译后文件"
OUTPUT_DIR = BASE_DIR / "4_输出文件cad"
JSON_PATH = BASE_DIR / "翻译对照.json"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="将翻译md增量写入翻译对照json")
    parser.add_argument("--source-file", default=str(SOURCE_FILE), help="待匹配原文文件")
    parser.add_argument("--translate-dir", default=str(TRANSLATE_DIR), help="翻译md目录")
    parser.add_argument("--json-path", default=str(JSON_PATH), help="翻译对照json路径")
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR), help="输出目录")
    return parser.parse_args()


def normalize_for_match(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())


def parse_translation_file(path: Path) -> dict[str, str]:
    text = path.read_text(encoding="utf-8")
    lines = text.splitlines()
    pairs: dict[str, str] = {}

    i = 0
    while i < len(lines):
        if not lines[i].strip():
            i += 1
            continue
        original = lines[i].strip()
        i += 1
        while i < len(lines) and not lines[i].strip():
            i += 1
        if i >= len(lines):
            break
        translation = lines[i].strip()
        i += 1

        if original and translation and original != translation:
            pairs[original] = translation

    return pairs


def load_existing_json(json_path: Path) -> list[dict]:
    if not json_path.exists():
        return []
    try:
        data = json.loads(json_path.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return [
                {"original": k, "translation": v}
                for k, v in data.items()
                if isinstance(k, str) and isinstance(v, str)
            ]
        if isinstance(data, list):
            return data
        return []
    except Exception:
        return []


def save_json_incremental(new_items: list[dict], json_path: Path) -> None:
    existing = load_existing_json(json_path)
    merged = {item["original"]: item for item in existing if isinstance(item, dict) and "original" in item}

    added = 0
    for item in new_items:
        original = item["original"]
        if original not in merged:
            merged[original] = item
            added += 1

    final_list = list(merged.values())
    json_path.write_text(
        json.dumps(final_list, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"JSON更新完成: 新增 {added} 条，总计 {len(final_list)} 条")
    print(f"翻译JSON: {json_path}")


def main() -> None:
    args = parse_args()
    source_file = Path(args.source_file).expanduser().resolve()
    translate_dir = Path(args.translate_dir).expanduser().resolve()
    json_path = Path(args.json_path).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()

    output_dir.mkdir(parents=True, exist_ok=True)

    if not source_file.exists():
        print(f"待匹配原文不存在: {source_file}")
        return
    if not translate_dir.exists():
        print(f"翻译目录不存在: {translate_dir}")
        return

    source_lines = [
        line.strip() for line in source_file.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]

    all_pairs: dict[str, str] = {}
    for md_file in sorted(translate_dir.glob("*.md")):
        pairs = parse_translation_file(md_file)
        print(f"{md_file.name}: {len(pairs)} 对")
        all_pairs.update(pairs)

    norm_map = {normalize_for_match(orig): orig for orig in all_pairs}

    matched: list[dict[str, str]] = []
    unmatched: list[str] = []

    for line in source_lines:
        norm = normalize_for_match(line)
        if line in all_pairs:
            matched.append({"original": line, "translation": all_pairs[line]})
        elif norm in norm_map:
            key = norm_map[norm]
            matched.append({"original": line, "translation": all_pairs[key]})
        else:
            unmatched.append(line)

    save_json_incremental(matched, json_path)

    unmatched_path = output_dir / "未翻译行.md"
    unmatched_path.write_text("\n".join(unmatched) + ("\n" if unmatched else ""), encoding="utf-8")

    print(f"匹配成功: {len(matched)}")
    print(f"未翻译: {len(unmatched)}")
    print(f"未翻译输出: {unmatched_path}")


if __name__ == "__main__":
    main()
