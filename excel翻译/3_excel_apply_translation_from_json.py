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

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

BASE_DIR = Path(__file__).resolve().parent
EXCEL_SUFFIXES = {'.xlsx', '.xlsm', '.xltx', '.xltm'}
TRANS_JSON = BASE_DIR / '翻译对照.json'
OUTPUT_DIR = BASE_DIR / '4_输出文件excel'


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
    """NFC 归一化 + 折叠空白，用于模糊匹配。"""
    t = unicodedata.normalize("NFC", text.strip())
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

    # 归一化 lookup（NFC + 折叠空白）
    norm: dict[str, str] = {}
    for orig, trans in exact.items():
        nk = _norm(orig)
        if nk not in norm:
            norm[nk] = trans

    # 按 key 长度降序，用于子串替换时优先匹配较长条目
    sorted_norm_keys = sorted(norm.keys(), key=len, reverse=True)

    return exact, norm, sorted_norm_keys


def replace_multiline_text(
    text: str,
    exact: dict[str, str],
    norm: dict[str, str],
    sorted_norm_keys: list[str],
) -> tuple[str, int]:
    normalized = text.replace('\r\n', '\n').replace('\r', '\n')
    parts = normalized.split('\n')
    replaced = 0
    new_parts: list[str] = []

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
            continue

        # 2. 归一化精确匹配（NFC + 折叠空白）
        nk = _norm(stripped)
        if nk in norm:
            new_parts.append(f"{prefix}{norm[nk]}{suffix}")
            replaced += 1
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
        else:
            new_parts.append(part)

    if replaced == 0:
        return text, 0
    return '\n'.join(new_parts), replaced


def translate_workbook(
    src: Path,
    dst: Path,
    exact: dict[str, str],
    norm: dict[str, str],
    sorted_norm_keys: list[str],
) -> int:
    wb = load_workbook(src, data_only=False)
    count = 0
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
                    replaced_text, replaced_count = replace_multiline_text(
                        val, exact, norm, sorted_norm_keys
                    )
                    if replaced_count > 0:
                        cell.value = replaced_text
                        count += replaced_count
    wb.save(dst)
    return count


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
    for src in excel_files:
        relative_path = src.relative_to(source_root) if input_path.is_dir() else Path(src.name)
        dst = output_dir / relative_path
        dst.parent.mkdir(parents=True, exist_ok=True)
        n = translate_workbook(src, dst, exact, norm, sorted_norm_keys)
        total += n
        print(f"  [{n:>4} 处替换]  {relative_path}")

    print(f"\n完成！共替换 {total} 处，输出目录: {output_dir}")


if __name__ == '__main__':
    main()

