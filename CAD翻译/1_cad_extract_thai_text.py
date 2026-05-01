#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Iterable

try:
    import ezdxf
except ImportError as exc:
    raise SystemExit(
        "未安装依赖 ezdxf，请先执行: pip install ezdxf"
    ) from exc

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "00_dxf文件"
EXTRACT_DIR = BASE_DIR / "1_提取文本"
MERGE_DIR = BASE_DIR / "2_提取文件合并文件"
JSON_CANDIDATES = [
    BASE_DIR / "翻译对照.json",
    BASE_DIR / "4_输出文件cad" / "翻译对照.json",
]
DXF_SUFFIXES = {".dxf"}

THAI_RE = re.compile(r"[\u0e00-\u0e7f]")
PURE_NUMBER_RE = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")
INDEX_NUMBER_RE = re.compile(r"^\d+(?:\.\d+)+$")
SHORT_ALNUM_CODE_RE = re.compile(r"^[A-Za-z]{1,6}\d+(?:\.\d+)*$|^\d+[A-Za-z]{1,4}$")


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())


def is_noise_line(text: str) -> bool:
    compact = text.strip()
    if not compact:
        return True
    normalized = compact.replace(",", "")
    if PURE_NUMBER_RE.fullmatch(normalized):
        return True
    if INDEX_NUMBER_RE.fullmatch(compact):
        return True
    if SHORT_ALNUM_CODE_RE.fullmatch(compact):
        return True
    return False


def split_entity_text(text: str) -> list[str]:
    normalized = text.replace("\\P", "\n").replace("\r\n", "\n").replace("\r", "\n")
    return [line.strip() for line in normalized.split("\n") if line.strip()]


def iter_entity_texts(doc) -> Iterable[str]:
    for layout in doc.layouts:
        for entity in layout:
            etype = entity.dxftype()
            if etype == "TEXT":
                yield str(entity.dxf.text)
            elif etype == "MTEXT":
                yield str(entity.text)
            elif etype in {"ATTRIB", "ATTDEF"}:
                yield str(entity.dxf.text)


def collect_dxf_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path] if input_path.suffix.lower() in DXF_SUFFIXES else []
    if input_path.is_dir():
        return sorted(
            f for f in input_path.rglob("*")
            if f.suffix.lower() in DXF_SUFFIXES and not f.name.startswith("~$")
        )
    raise FileNotFoundError(f"路径不存在: {input_path}")


def extract_file(dxf_path: Path, out_path: Path, dedupe: bool = True) -> int:
    doc = ezdxf.readfile(dxf_path)
    seen: set[str] = set()
    lines: list[str] = []

    for raw_text in iter_entity_texts(doc):
        for line in split_entity_text(raw_text):
            if not THAI_RE.search(line):
                continue
            if is_noise_line(line):
                continue
            normalized = normalize_text(line)
            if dedupe and normalized in seen:
                continue
            seen.add(normalized)
            lines.append(line)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")
    return len(lines)


def find_json_path() -> Path:
    for path in JSON_CANDIDATES:
        if path.exists():
            return path
    return JSON_CANDIDATES[0]


def load_originals(json_path: Path) -> tuple[set[str], set[str]]:
    if not json_path.exists():
        return set(), set()

    data = json.loads(json_path.read_text(encoding="utf-8"))
    originals: list[str] = []

    if isinstance(data, list):
        for item in data:
            if isinstance(item, dict):
                orig = item.get("original")
                if isinstance(orig, str) and orig.strip():
                    originals.append(orig.strip())
    elif isinstance(data, dict):
        for key in data:
            if isinstance(key, str) and key.strip():
                originals.append(key.strip())

    exact = set(originals)
    normalized = {normalize_text(item) for item in originals}
    return exact, normalized


def merge_and_filter_pending() -> None:
    md_files = sorted(EXTRACT_DIR.glob("*.md"))
    MERGE_DIR.mkdir(parents=True, exist_ok=True)

    merged_path = MERGE_DIR / "合并文本.md"
    filtered_path = MERGE_DIR / "过滤后文本.md"
    pending_path = MERGE_DIR / "待翻译项.md"

    seen: set[str] = set()
    unique_lines: list[str] = []

    for md_file in md_files:
        for line in md_file.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if not stripped:
                continue
            norm = normalize_text(stripped)
            if norm in seen:
                continue
            seen.add(norm)
            unique_lines.append(stripped)

    merged_path.write_text("\n".join(unique_lines) + ("\n" if unique_lines else ""), encoding="utf-8")
    filtered_path.write_text("\n".join(unique_lines) + ("\n" if unique_lines else ""), encoding="utf-8")

    json_path = find_json_path()
    exact, normalized = load_originals(json_path)
    pending = [
        line for line in unique_lines
        if line not in exact and normalize_text(line) not in normalized
    ]

    pending_path.write_text("\n".join(pending) + ("\n" if pending else ""), encoding="utf-8")
    print(f"合并后 {len(unique_lines)} 行，待翻译 {len(pending)} 行")
    print(f"待翻译项输出: {pending_path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="提取CAD(DXF)中的泰文文本，生成待翻译项")
    parser.add_argument(
        "input_path",
        nargs="?",
        default=str(INPUT_DIR),
        help="待处理DXF文件或目录，默认 00_dxf文件",
    )
    parser.add_argument(
        "--no-dedupe",
        action="store_true",
        help="保留重复行",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input_path).expanduser().resolve()
    files = collect_dxf_files(input_path)

    if not files:
        print("未找到可处理的DXF文件")
        return

    EXTRACT_DIR.mkdir(parents=True, exist_ok=True)

    print(f"找到 {len(files)} 个DXF文件")
    total_lines = 0
    for i, dxf_path in enumerate(files, start=1):
        out_name = dxf_path.stem + ".md"
        out_path = EXTRACT_DIR / out_name
        if out_path.exists():
            out_name = f"{dxf_path.parent.name}_{dxf_path.stem}.md"
            out_path = EXTRACT_DIR / out_name
        count = extract_file(dxf_path, out_path, dedupe=not args.no_dedupe)
        total_lines += count
        print(f"[{i}/{len(files)}] {dxf_path.name} -> {out_name} ({count} 行)")

    print(f"提取完成，共 {total_lines} 行")
    merge_and_filter_pending()


if __name__ == "__main__":
    main()
