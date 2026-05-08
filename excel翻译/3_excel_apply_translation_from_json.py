#!/usr/bin/env python3
"""
Excel翻译替换工具
输入目录: INPUT_DIR 中所有 .xlsx 文件
翻译对照: TRANS_JSON (original/translation 键值对列表)
输出目录: OUTPUT_DIR

规则：
- 跳过公式单元格（值以 '=' 开头）
- 对纯文本单元格做精确匹配替换（去首尾空格后比较）
- 合并单元格的占位格（MergedCell）跳过，不修改
"""

import json
import re
import unicodedata
import argparse
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

BASE_DIR = Path(__file__).resolve().parent
EXCEL_SUFFIXES = {'.xlsx', '.xlsm', '.xltx', '.xltm'}
TRANS_JSON = BASE_DIR / '翻译对照.json'
OUTPUT_DIR = BASE_DIR / '4_输出文件excel'
REPLACEMENT_REPORT_JSON = '翻译替换报告.json'


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description='根据翻译对照json批量输出已翻译Excel'
    )
    parser.add_argument(
        'input_path',
        nargs='?',
        help='待输出翻译的Excel文件或文件夹路径',
    )
    parser.add_argument(
        '--json-path',
        default=str(TRANS_JSON),
        help='翻译对照json路径',
    )
    parser.add_argument(
        '--output-dir',
        default=str(OUTPUT_DIR),
        help='输出目录',
    )
    return parser.parse_args()


def _norm(text: str) -> str:
    """NFKC 归一化 + 折叠空白，兼容泰语表现形式字符（如 ）差异。"""
    t = unicodedata.normalize("NFKC", text.strip())
    return re.sub(r"\s+", " ", t)


def normalize_path_text(path_text: str) -> str:
    return path_text.strip().strip('"').strip("'")


def collect_excel_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() in EXCEL_SUFFIXES and not input_path.name.startswith('~$'):
            return [input_path]
        print(f'跳过非Excel文件: {input_path.name}')
        return []
    if input_path.is_dir():
        return sorted(
            file
            for file in input_path.rglob('*')
            if file.suffix.lower() in EXCEL_SUFFIXES and not file.name.startswith('~$')
        )
    raise FileNotFoundError(f'路径不存在: {input_path}')


def build_lookup(path: Path) -> tuple[dict[str, str], dict[str, str], list[str]]:
    """返回 (精确lookup, 归一化lookup, 按长度降序的归一化key列表)。"""
    data = json.loads(path.read_text(encoding='utf-8'))
    exact: dict[str, str] = {}

    pairs: list[tuple[str, str]] = []
    if isinstance(data, dict):
        pairs = [(k, v) for k, v in data.items() if isinstance(k, str) and isinstance(v, str)]
    else:
        pairs = [
            (item.get('original', ''), item.get('translation', ''))
            for item in data if isinstance(item, dict)
        ]

    for orig, trans in pairs:
        orig, trans = orig.strip(), trans.strip()
        if orig and trans:
            exact[orig] = trans

    # 归一化 lookup（NFKC + 折叠空白）
    norm: dict[str, str] = {}
    for orig, trans in exact.items():
        nk = _norm(orig)
        if nk not in norm:
            norm[nk] = trans

        # 兼容回退：收录 NFC 形态，避免历史数据中混用不同规范化形式
        nfc_key = re.sub(r"\s+", " ", unicodedata.normalize("NFC", orig.strip()))
        if nfc_key and nfc_key not in norm:
            norm[nfc_key] = trans

    # 按 key 长度降序，用于子串替换时优先匹配较长条目
    sorted_norm_keys = sorted(norm.keys(), key=len, reverse=True)

    return exact, norm, sorted_norm_keys


def replace_multiline_text(
    text: str,
    exact: dict[str, str],
    norm: dict[str, str],
    sorted_norm_keys: list[str],
) -> tuple[str, int, int, int, int]:
    normalized = text.replace('\r\n', '\n').replace('\r', '\n')
    parts = normalized.split('\n')
    replaced = 0
    exact_hits = 0
    norm_hits = 0
    substring_hits = 0
    new_parts: list[str] = []
    unmatched_count = 0

    for part in parts:
        stripped = part.strip()
        if not stripped:
            new_parts.append(part)
            continue

        leading = len(part) - len(part.lstrip())
        trailing = len(part) - len(part.rstrip())
        prefix = part[:leading]
        suffix = part[len(part) - trailing:] if trailing else ''

        # 1. 精确匹配
        if stripped in exact:
            new_parts.append(f"{prefix}{exact[stripped]}{suffix}")
            replaced += 1
            exact_hits += 1
            continue

        # 2. 归一化精确匹配（NFC + 折叠空白）
        nk = _norm(stripped)
        if nk in norm:
            new_parts.append(f"{prefix}{norm[nk]}{suffix}")
            replaced += 1
            norm_hits += 1
            continue

        # 3. 子串替换：把归一化后的行文本里包含的所有 key 依次替换
        #    适用于翻译者把一行拆成多条、或录入时空格不统一的情况
        result = nk
        sub_count = 0
        for key in sorted_norm_keys:
            if len(key) < 6:  # 过短的 key 不做子串替换，避免误伤
                continue
            if key in result:
                result = result.replace(key, norm[key])
                sub_count += 1
        if sub_count:
            new_parts.append(f"{prefix}{result}{suffix}")
            replaced += sub_count
            substring_hits += sub_count
        else:
            new_parts.append(part)
            if any(ch.isalpha() for ch in stripped) or re.search(r"[\u0e00-\u0e7f]", stripped):
                unmatched_count += 1

    if replaced == 0:
        return text, 0, 0, 0, unmatched_count
    return '\n'.join(new_parts), exact_hits, norm_hits, substring_hits, unmatched_count


def translate_workbook(
    src: Path,
    dst: Path,
    exact: dict[str, str],
    norm: dict[str, str],
    sorted_norm_keys: list[str],
) -> dict[str, int]:
    wb = load_workbook(src, data_only=False)
    stats = {
        'exact_hits': 0,
        'norm_hits': 0,
        'substring_hits': 0,
        'total_replacements': 0,
        'unmatched_cells': 0,
    }
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                val = cell.value
                if val is None:
                    continue
                if isinstance(val, str) and val.lstrip().startswith('='):
                    continue
                if isinstance(val, str):
                    replaced_text, exact_hits, norm_hits, substring_hits, unmatched_count = replace_multiline_text(
                        val, exact, norm, sorted_norm_keys
                    )
                    replaced_count = exact_hits + norm_hits + substring_hits
                    if replaced_count > 0:
                        cell.value = replaced_text
                        stats['exact_hits'] += exact_hits
                        stats['norm_hits'] += norm_hits
                        stats['substring_hits'] += substring_hits
                        stats['total_replacements'] += replaced_count
                    elif unmatched_count:
                        stats['unmatched_cells'] += 1
    wb.save(dst)
    return stats


def write_reports(
    output_dir: Path,
    file_reports: list[dict[str, object]],
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    json_path = output_dir / REPLACEMENT_REPORT_JSON

    total = {
        'files': len(file_reports),
        'exact_hits': 0,
        'norm_hits': 0,
        'substring_hits': 0,
        'total_replacements': 0,
        'unmatched_cells': 0,
    }

    def _to_int(value: Any) -> int:
        if isinstance(value, bool):
            return int(value)
        if isinstance(value, int):
            return value
        if isinstance(value, float):
            return int(value)
        if isinstance(value, str):
            s = value.strip()
            if not s:
                return 0
            try:
                return int(float(s))
            except ValueError:
                return 0
        return 0

    for r in file_reports:
        total['exact_hits'] += _to_int(r.get('exact_hits', 0))
        total['norm_hits'] += _to_int(r.get('norm_hits', 0))
        total['substring_hits'] += _to_int(r.get('substring_hits', 0))
        total['total_replacements'] += _to_int(r.get('total_replacements', 0))
        total['unmatched_cells'] += _to_int(r.get('unmatched_cells', 0))

    payload = {
        'summary': total,
        'files': file_reports,
        'unmatched_cells_preview_count': total['unmatched_cells'],
    }
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    return json_path


def main() -> None:
    args = parse_args()

    raw_input = args.input_path
    if raw_input is None:
        raw_input = input('请输入Excel文件或文件夹路径: ').strip()

    raw_input = normalize_path_text(raw_input)
    if not raw_input:
        raise ValueError('路径不能为空')

    input_path = Path(raw_input).expanduser().resolve()
    json_path = Path(args.json_path).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()

    excel_files = collect_excel_files(input_path)
    if not excel_files:
        print('未找到任何Excel文件')
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    exact, norm, sorted_norm_keys = build_lookup(json_path)
    print(f"已加载翻译词条: {len(exact)} 条（精确），{len(norm)} 条（归一化）")

    source_root = input_path if input_path.is_dir() else input_path.parent
    print(f"发现 Excel 文件: {len(excel_files)} 个\n")

    total = 0
    total_exact_hits = 0
    total_norm_hits = 0
    total_substring_hits = 0
    total_unmatched_cells = 0
    file_reports: list[dict[str, object]] = []
    for src in excel_files:
        relative_path = src.relative_to(source_root) if input_path.is_dir() else Path(src.name)
        dst = output_dir / relative_path
        dst.parent.mkdir(parents=True, exist_ok=True)
        stats = translate_workbook(src, dst, exact, norm, sorted_norm_keys)
        n = int(stats['total_replacements'])
        total += n
        total_exact_hits += int(stats['exact_hits'])
        total_norm_hits += int(stats['norm_hits'])
        total_substring_hits += int(stats['substring_hits'])
        total_unmatched_cells += int(stats['unmatched_cells'])

        file_reports.append(
            {
                'file': str(relative_path),
                **stats,
            }
        )

        print(
            f"  [{n:>4} 处替换] {relative_path} "
            f"(精确 {stats['exact_hits']}, 归一化 {stats['norm_hits']}, 子串 {stats['substring_hits']}, 未命中单元格 {stats['unmatched_cells']})"
        )

    report_json = write_reports(output_dir, file_reports)

    # 兼容旧版本：若目录中仍有历史 md 报告则删除
    stale_md = output_dir / '未命中单元格.md'
    if stale_md.exists() and stale_md.is_file():
        stale_md.unlink()

    print(
        f"\n完成！共替换 {total} 处 "
        f"(精确 {total_exact_hits}, 归一化 {total_norm_hits}, 子串 {total_substring_hits})"
    )
    print(f"未命中单元格: {total_unmatched_cells} 个")
    print(f"替换报告(JSON): {report_json}")
    print(f"输出目录: {output_dir}")


if __name__ == '__main__':
    main()

