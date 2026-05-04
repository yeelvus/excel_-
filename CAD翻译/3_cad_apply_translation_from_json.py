#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

try:
    import ezdxf
except ImportError as exc:
    raise SystemExit(
        "未安装依赖 ezdxf，请先执行: pip install ezdxf"
    ) from exc

try:
    from opencc import OpenCC
except ImportError:
    OpenCC = None

CHINESE_RE = re.compile(r"[\u4e00-\u9fff\u3400-\u4dbf]")
TRAD_HINT_CHARS = set(
    "萬與專業東絲兩嚴喪個豐臨為麗舉麼義烏樂喬習鄉書買亂爭於虧雲亞產畝親億僅從倉儀們價眾優會傘偉傳傷倫偽體餘佈來係俠倀倆傾僅僉僑僞僥僱儲儷兒兌兗內冊冪凍凜幾鳳凱別刪則剋剎剛剥剮創劃劇劉劊劍劑勁動務勛勝勞勢勵勸區醫華協單賣盧鹵臥衛卻卷厭厲壓參雙發變"
)


def build_simplifier():
    if OpenCC is None:
        print("未安装 opencc，繁体转简体功能将跳过。可安装: pip install opencc-python-reimplemented")
        return None
    return OpenCC("t2s")


def is_traditional(text: str) -> bool:
    if not CHINESE_RE.search(text):
        return False
    return any(ch in TRAD_HINT_CHARS for ch in text)


def to_simplified(text: str, simplifier) -> str:
    if simplifier is None or not CHINESE_RE.search(text):
        return text
    return simplifier.convert(text)

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
    if not path.exists():
        print(f"翻译JSON不存在，按空词条处理: {path}")
        return {}

    raw = path.read_text(encoding="utf-8")
    if not raw.strip():
        print(f"翻译JSON为空，按空词条处理: {path}")
        return {}

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        print(f"翻译JSON格式异常，按空词条处理: {path}")
        return {}

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


def replace_multiline_text(text: str, lookup: dict[str, str], simplifier=None) -> tuple[str, int]:
    if "\\P" in text:
        sep = "\\P"
    elif "\r\n" in text:
        sep = "\r\n"
    elif "\n" in text:
        sep = "\n"
    else:
        sep = None

    def _lookup_or_simplify(segment: str) -> tuple[str, int]:
        """先尝试查翻译，再尝试简体化，否则原样返回。"""
        stripped = segment.strip()
        if not stripped:
            return segment, 0
        # 优先直接匹配
        if stripped in lookup:
            leading = len(segment) - len(segment.lstrip())
            trailing = len(segment) - len(segment.rstrip())
            prefix = segment[:leading]
            suffix = segment[len(segment) - trailing:] if trailing else ""
            return f"{prefix}{lookup[stripped]}{suffix}", 1
        # 尝试将繁体转简体后再查
        simp = to_simplified(stripped, simplifier)
        if simp != stripped and simp in lookup:
            leading = len(segment) - len(segment.lstrip())
            trailing = len(segment) - len(segment.rstrip())
            prefix = segment[:leading]
            suffix = segment[len(segment) - trailing:] if trailing else ""
            return f"{prefix}{lookup[simp]}{suffix}", 1
        # 无翻译条目时，若含汉字则直接转简体（无需判断是否繁体，OpenCC 对简体无影响）
        if CHINESE_RE.search(stripped):
            converted = to_simplified(segment, simplifier)
            if converted != segment:
                return converted, 1
        return segment, 0

    if sep is None:
        new_text, count = _lookup_or_simplify(text)
        return new_text, count

    parts = text.split(sep)
    replaced = 0
    new_parts: list[str] = []

    for part in parts:
        new_part, count = _lookup_or_simplify(part)
        new_parts.append(new_part)
        replaced += count

    if replaced == 0:
        return text, 0
    return sep.join(new_parts), replaced


def process_text(text: str, lookup: dict[str, str], simplifier) -> tuple[str, int]:
    """查翻译 + 繁→简双重处理。

    先做基于 lookup 的段落级替换，再对整体文本做一次全文简化，
    确保 MTEXT 格式代码内嵌的繁体汉字（如 {\\fFont|...;繁體}）也能被转换。
    """
    new_text, count = replace_multiline_text(text, lookup, simplifier)
    # 兜底：对整体文本再做一次简化，覆盖格式代码内嵌字符
    if simplifier and CHINESE_RE.search(new_text):
        simplified = to_simplified(new_text, simplifier)
        if simplified != new_text:
            new_text = simplified
            count += 1
    return new_text, count


def iter_spaces(doc):
    for layout in doc.layouts:
        yield layout
    for block in doc.blocks:
        name = getattr(block, "name", "")
        if isinstance(name, str) and name.startswith("*"):
            continue
        yield block


def translate_dxf(src: Path, dst: Path, lookup: dict[str, str], simplifier=None) -> int:
    doc = ezdxf.readfile(src)
    replaced_count = 0

    for space in iter_spaces(doc):
        for entity in space:
            etype = entity.dxftype()
            if etype == "TEXT":
                raw = str(entity.dxf.text)
                new_text, count = process_text(raw, lookup, simplifier)
                if count > 0:
                    entity.dxf.text = new_text
                    replaced_count += count
            elif etype == "MTEXT":
                raw = str(getattr(entity, "text", ""))
                new_text, count = process_text(raw, lookup, simplifier)
                if count > 0 and hasattr(entity, "text"):
                    setattr(entity, "text", new_text)
                    replaced_count += count
            elif etype in {"ATTRIB", "ATTDEF"}:
                raw = str(entity.dxf.text)
                new_text, count = process_text(raw, lookup, simplifier)
                if count > 0:
                    entity.dxf.text = new_text
                    replaced_count += count
            elif etype == "INSERT":
                for attrib in getattr(entity, "attribs", []):
                    raw = str(attrib.dxf.text)
                    new_text, count = process_text(raw, lookup, simplifier)
                    if count > 0:
                        attrib.dxf.text = new_text
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

    simplifier = build_simplifier()
    if simplifier:
        print("繁体转简体功能已启用")

    source_root = input_path if input_path.is_dir() else input_path.parent
    total = 0
    for src in files:
        rel = src.relative_to(source_root) if input_path.is_dir() else Path(src.name)
        dst = output_dir / rel
        count = translate_dxf(src, dst, lookup, simplifier)
        total += count
        print(f"[{count:>4} 处替换] {rel}")

    print(f"完成，总替换 {total} 处，输出目录: {output_dir}")


if __name__ == "__main__":
    main()
