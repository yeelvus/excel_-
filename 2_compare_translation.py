#!/usr/bin/env python3
"""
对比过滤后文本与翻译文件，输出：
1. 已匹配的翻译对 → JSON文件（增量更新）
2. 未翻译的行 → 单独md文件
"""
import json
import re
import argparse
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
FILTER_FILE = BASE_DIR / "2_提取文件合并文件" / "待翻译项.md"
TRANSLATE_DIR = BASE_DIR / "3_翻译后文件"
OUTPUT_DIR = BASE_DIR / "4_输出文件excel"

JSON_PATH = BASE_DIR / "翻译对照.json"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="将翻译后的md增量写入json，并输出仍未翻译的行"
    )
    parser.add_argument(
        "--source-file",
        default=str(FILTER_FILE),
        help="待匹配原文文件，默认使用 2_提取文件合并文件/待翻译项.md",
    )
    parser.add_argument(
        "--translate-dir",
        default=str(TRANSLATE_DIR),
        help="翻译结果md文件夹",
    )
    parser.add_argument(
        "--json-path",
        default=str(JSON_PATH),
        help="翻译对照json路径",
    )
    parser.add_argument(
        "--output-dir",
        default=str(OUTPUT_DIR),
        help="输出目录，用于保存未翻译行和json副本",
    )
    return parser.parse_args()


def parse_translation_file(path: Path) -> dict[str, str]:
    """解析翻译文件，返回 {原文: 译文} 映射"""
    text = path.read_text(encoding="utf-8")
    lines = text.splitlines()
    pairs: dict[str, str] = {}
    i = 0
    total = len(lines)

    # 跳过开头的AI废话
    while i < total:
        line = lines[i].strip()
        if line == "" or (len(line) > 10 and re.search(r"[\u4e00-\u9fff].*[\u4e00-\u9fff].*[\u4e00-\u9fff]", line) and not re.search(r"[\u0e00-\u0e7f]", line)):
            i += 1
            continue
        break

    while i < total:
        if lines[i].strip() == "":
            i += 1
            continue
        original = lines[i].strip()
        i += 1

        while i < total and lines[i].strip() == "":
            i += 1
        if i >= total:
            break
        translation = lines[i].strip()
        i += 1

        while i < total and lines[i].strip() == "":
            i += 1

        if original and translation and original != translation:
            pairs[original] = translation

    return pairs


def normalize_for_match(text: str) -> str:
    t = text.strip()
    t = re.sub(r"\s+", " ", t)
    return t


def load_existing_json(json_path: Path) -> list[dict]:
    """加载已存在的JSON，返回已有的匹配列表"""
    if json_path.exists():
        try:
            data = json.loads(json_path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                data = [
                    {"original": original, "translation": translation}
                    for original, translation in data.items()
                    if isinstance(original, str) and isinstance(translation, str)
                ]
            print(f"已加载现有翻译对照: {len(data)} 条")
            return data
        except Exception:
            print("⚠️ 现有JSON文件损坏，将重新生成")
    return []


def save_json_incremental(new_matched: list[dict], json_path: Path, output_dir: Path):
    """增量保存：合并旧数据 + 新数据（去重）"""
    existing = load_existing_json(json_path)

    # 用 (original) 作为唯一键，构建已有集合
    existing_dict = {item["original"]: item for item in existing}

    added_count = 0
    for item in new_matched:
        orig = item["original"]
        if orig not in existing_dict:
            existing_dict[orig] = item
            added_count += 1

    final_list = list(existing_dict.values())

    output_dir.mkdir(parents=True, exist_ok=True)
    json_text = json.dumps(final_list, ensure_ascii=False, indent=2) + "\n"
    json_path.write_text(
        json_text,
        encoding="utf-8"
    )

    output_json_path = output_dir / "翻译对照.json"
    output_json_path.write_text(
        json_text,
        encoding="utf-8"
    )

    print(f"✅ JSON增量更新完成！本次新增 {added_count} 条，总计 {len(final_list)} 条")
    print(f"主JSON: {json_path}")
    print(f"副本JSON: {output_json_path}")


def main() -> None:
    args = parse_args()

    source_file = Path(args.source_file).expanduser().resolve()
    translate_dir = Path(args.translate_dir).expanduser().resolve()
    json_path = Path(args.json_path).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()

    if not source_file.exists():
        print(f"❌ 待匹配原文不存在: {source_file}")
        return

    if not translate_dir.exists():
        print(f"❌ 翻译文件夹不存在: {translate_dir}")
        return

    source_lines = [
        ln.strip() for ln in source_file.read_text(encoding="utf-8").splitlines()
        if ln.strip()
    ]
    print(f"原始待匹配文本共 {len(source_lines)} 行\n")

    # 解析所有翻译文件
    all_pairs: dict[str, str] = {}
    trans_files = sorted(translate_dir.glob("*.md"))
    for tf in trans_files:
        pairs = parse_translation_file(tf)
        print(f" {tf.name}: {len(pairs)} 对翻译")
        all_pairs.update(pairs)

    print(f"\n所有翻译文件总计: {len(all_pairs)} 对\n")

    # 构建规范化索引
    norm_map: dict[str, str] = {}
    for orig in all_pairs:
        norm_map[normalize_for_match(orig)] = orig

    # 匹配
    matched: list[dict[str, str]] = []
    unmatched: list[str] = []

    for line in source_lines:
        norm_line = normalize_for_match(line)
        if line in all_pairs:
            matched.append({"original": line, "translation": all_pairs[line]})
            continue
        if norm_line in norm_map:
            orig_key = norm_map[norm_line]
            matched.append({"original": line, "translation": all_pairs[orig_key]})
            continue
        unmatched.append(line)

    # 增量保存JSON
    save_json_incremental(matched, json_path, output_dir)

    # 保存未翻译行
    unmatched_path = output_dir / "未翻译行.md"
    unmatched_path.write_text("\n".join(unmatched) + ("\n" if unmatched else ""), encoding="utf-8")

    print(f"\n处理完成！")
    print(f"匹配成功: {len(matched)} 行")
    print(f"未翻译: {len(unmatched)} 行")
    print(f"未翻译行 → {unmatched_path}")


if __name__ == "__main__":
    main()