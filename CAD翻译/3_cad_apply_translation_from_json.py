#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

try:
    import ezdxf
except ImportError as exc:
    raise SystemExit(
        "未安装依赖 ezdxf，请先执行: pip install ezdxf"
    ) from exc

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "00_dxf文件"
JSON_PATH = BASE_DIR / "翻译对照.json"
OUTPUT_DIR = BASE_DIR / "4_输出文件cad"
DXF_SUFFIXES = {".dxf"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="根据翻译json输出CAD(DXF)翻译文件")
    parser.add_argument("input_path", nargs="?", default=str(INPUT_DIR), help="待处理DXF文件或目录")
    parser.add_argument("--json-path", default=str(JSON_PATH), help="翻译对照json路径")
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR), help="输出目录")
    return parser.parse_args()


def collect_dxf_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path] if input_path.suffix.lower() in DXF_SUFFIXES else []
    if input_path.is_dir():
        return sorted(
            f for f in input_path.rglob("*")
            if f.suffix.lower() in DXF_SUFFIXES and not f.name.startswith("~$")
        )
    raise FileNotFoundError(f"路径不存在: {input_path}")


def build_lookup(path: Path) -> dict[str, str]:
    data = json.loads(path.read_text(encoding="utf-8"))
    lookup: dict[str, str] = {}

    if isinstance(data, dict):
        for orig, trans in data.items():
            if isinstance(orig, str) and isinstance(trans, str):
                if orig.strip() and trans.strip():
                    lookup[orig.strip()] = trans.strip()
        return lookup

    if isinstance(data, list):
        for item in data:
            if not isinstance(item, dict):
                continue
            orig = str(item.get("original", "")).strip()
            trans = str(item.get("translation", "")).strip()
            if orig and trans:
                lookup[orig] = trans
    return lookup


def replace_multiline_text(text: str, lookup: dict[str, str]) -> tuple[str, int]:
    if "\\P" in text:
        sep = "\\P"
    elif "\r\n" in text:
        sep = "\r\n"
    elif "\n" in text:
        sep = "\n"
    else:
        sep = None

    if sep is None:
        stripped = text.strip()
        if stripped in lookup:
            leading = len(text) - len(text.lstrip())
            trailing = len(text) - len(text.rstrip())
            prefix = text[:leading]
            suffix = text[len(text) - trailing:] if trailing else ""
            return f"{prefix}{lookup[stripped]}{suffix}", 1
        return text, 0

    parts = text.split(sep)
    replaced = 0
    new_parts: list[str] = []

    for part in parts:
        stripped = part.strip()
        if stripped and stripped in lookup:
            leading = len(part) - len(part.lstrip())
            trailing = len(part) - len(part.rstrip())
            prefix = part[:leading]
            suffix = part[len(part) - trailing:] if trailing else ""
            new_parts.append(f"{prefix}{lookup[stripped]}{suffix}")
            replaced += 1
        else:
            new_parts.append(part)

    if replaced == 0:
        return text, 0
    return sep.join(new_parts), replaced


def translate_dxf(src: Path, dst: Path, lookup: dict[str, str]) -> int:
    doc = ezdxf.readfile(src)
    replaced_count = 0

    for layout in doc.layouts:
        for entity in layout:
            etype = entity.dxftype()
            if etype == "TEXT":
                raw = str(entity.dxf.text)
                new_text, count = replace_multiline_text(raw, lookup)
                if count > 0:
                    entity.dxf.text = new_text
                    replaced_count += count
            elif etype == "MTEXT":
                raw = str(entity.text)
                new_text, count = replace_multiline_text(raw, lookup)
                if count > 0:
                    entity.text = new_text
                    replaced_count += count
            elif etype in {"ATTRIB", "ATTDEF"}:
                raw = str(entity.dxf.text)
                new_text, count = replace_multiline_text(raw, lookup)
                if count > 0:
                    entity.dxf.text = new_text
                    replaced_count += count

    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.saveas(dst)
    return replaced_count


def main() -> None:
    args = parse_args()
    input_path = Path(args.input_path).expanduser().resolve()
    json_path = Path(args.json_path).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()

    files = collect_dxf_files(input_path)
    if not files:
        print("未找到可处理的DXF文件")
        return

    lookup = build_lookup(json_path)
    print(f"已加载翻译词条: {len(lookup)} 条")

    source_root = input_path if input_path.is_dir() else input_path.parent
    total = 0
    for src in files:
        rel = src.relative_to(source_root) if input_path.is_dir() else Path(src.name)
        dst = output_dir / rel
        count = translate_dxf(src, dst, lookup)
        total += count
        print(f"[{count:>4} 处替换] {rel}")

    print(f"完成，总替换 {total} 处，输出目录: {output_dir}")


if __name__ == "__main__":
    main()
