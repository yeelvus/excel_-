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

try:
    from opencc import OpenCC
except ImportError:
    OpenCC = None

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
ENGLISH_RE = re.compile(r"[A-Za-z]")
ENGLISH_WORD_RE = re.compile(r"[A-Za-z]{3,}")
CHINESE_RE = re.compile(r"[\u4e00-\u9fff]")
TRAD_HINT_CHARS = set(
    "萬與專業東絲兩嚴喪個豐臨為麗舉麼義烏樂喬習鄉書買亂爭於虧雲亞產畝親億僅從倉儀們價眾優會傘偉傳傷倫偽體餘佈來係俠倀倆傾僅僉僑僞僥僱儲儷兒兌兗內冊冪凍凜幾鳳凱別刪則剋剎剛剝剮創劃劇劉劊劍劑勁動務勛勝勞勢勵勸區醫華協單賣盧鹵臥衛卻卷厭厲壓參雙發變"
)
PURE_NUMBER_RE = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")
INDEX_NUMBER_RE = re.compile(r"^\d+(?:\.\d+)+$")
SHORT_ALNUM_CODE_RE = re.compile(r"^[A-Za-z]{1,6}\d+(?:\.\d+)*$|^\d+[A-Za-z]{1,4}$")
CAD_META_RE = re.compile(r"^(?:AcDb|AcCm|ByLayer|ByBlock|Model|Paper|STANDARD|Standard|Layer|LAYER|BLOCK|INSERT|LINE|CIRCLE|ARC|DIM|STYLE|UCS|VIEW|TABLE|XREF|HANDLE|OWNER)")


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())


def is_noise_line(text: str) -> bool:
    compact = text.strip()
    if not compact:
        return True
    if len(compact) > 300:
        return True
    normalized = compact.replace(",", "")
    if PURE_NUMBER_RE.fullmatch(normalized):
        return True
    if INDEX_NUMBER_RE.fullmatch(compact):
        return True
    if SHORT_ALNUM_CODE_RE.fullmatch(compact):
        return True
    return False


def is_readable_text(text: str) -> bool:
    compact = text.strip()
    if not compact:
        return False
    if CAD_META_RE.match(compact):
        return False

    # 默认保留：泰语、中文（含繁简）
    if THAI_RE.search(compact):
        return True
    if CHINESE_RE.search(compact):
        return True
    return False


def is_readable_text_with_english(text: str) -> bool:
    compact = text.strip()
    if not compact:
        return False
    if CAD_META_RE.match(compact):
        return False

    if THAI_RE.search(compact):
        return True
    if CHINESE_RE.search(compact):
        return True
    if ENGLISH_WORD_RE.search(compact):
        return True
    return False


def split_entity_text(text: str) -> list[str]:
    normalized = text.replace("\\P", "\n").replace("\r\n", "\n").replace("\r", "\n")
    return [line.strip() for line in normalized.split("\n") if line.strip()]


def iter_string_values(value):
    if value is None:
        return
    if isinstance(value, str):
        yield value
        return
    if isinstance(value, (list, tuple)):
        for item in value:
            if isinstance(item, str):
                yield item


def is_traditional_line(text: str) -> bool:
    if not CHINESE_RE.search(text):
        return False
    return any(ch in TRAD_HINT_CHARS for ch in text)


def build_simplifier(enable: bool):
    if not enable:
        return None
    if OpenCC is None:
        print("未安装 opencc，繁体转简体功能将跳过。可安装: pip install opencc-python-reimplemented")
        return None
    return OpenCC("t2s")


def to_simplified(text: str, simplifier) -> str:
    if simplifier is None:
        return text
    if not CHINESE_RE.search(text):
        return text
    return simplifier.convert(text)


def iter_spaces(doc):
    # layouts 覆盖模型空间/图纸空间；blocks 覆盖用户自定义块定义。
    for layout in doc.layouts:
        yield layout
    for block in doc.blocks:
        name = getattr(block, "name", "")
        if isinstance(name, str) and name.startswith("*"):
            continue
        yield block


def iter_entity_texts(doc) -> Iterable[str]:
    for space in iter_spaces(doc):
        for entity in space:
            etype = entity.dxftype()
            if etype == "TEXT":
                yield str(entity.dxf.text)
            elif etype == "MTEXT":
                yield str(entity.text)
            elif etype in {"ATTRIB", "ATTDEF"}:
                yield str(entity.dxf.text)
            elif etype == "INSERT":
                for attrib in getattr(entity, "attribs", []):
                    yield str(attrib.dxf.text)

            # 兜底扫描实体所有 dxf 字段中的字符串值，覆盖更多对象类型与自定义字段。
            for value in entity.dxfattribs().values():
                for text in iter_string_values(value):
                    yield text


def collect_dxf_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path] if input_path.suffix.lower() in DXF_SUFFIXES else []
    if input_path.is_dir():
        return sorted(
            f for f in input_path.rglob("*")
            if f.suffix.lower() in DXF_SUFFIXES and not f.name.startswith("~$")
        )
    raise FileNotFoundError(f"路径不存在: {input_path}")


def extract_file(
    dxf_path: Path,
    out_path: Path,
    dedupe: bool = True,
    include_english: bool = False,
) -> int:
    doc = ezdxf.readfile(dxf_path)
    seen: set[str] = set()
    lines: list[str] = []

    for raw_text in iter_entity_texts(doc):
        for line in split_entity_text(raw_text):
            if is_noise_line(line):
                continue
            if include_english:
                keep = is_readable_text_with_english(line)
            else:
                keep = is_readable_text(line)
            if not keep:
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

    raw = json_path.read_text(encoding="utf-8")
    if not raw.strip():
        print(f"翻译JSON为空，按无历史词条处理: {json_path}")
        return set(), set()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        print(f"翻译JSON格式异常，按无历史词条处理: {json_path}")
        return set(), set()

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


def merge_and_filter_pending(include_english: bool = False, convert_trad_to_simp: bool = True) -> None:
    md_files = sorted(EXTRACT_DIR.glob("*.md"))
    MERGE_DIR.mkdir(parents=True, exist_ok=True)

    merged_path = MERGE_DIR / "合并文本.md"
    filtered_path = MERGE_DIR / "过滤后文本.md"
    pending_path = MERGE_DIR / "待翻译项.md"
    category_thai_path = MERGE_DIR / "分类_泰语.md"
    category_trad_path = MERGE_DIR / "分类_繁体.md"
    category_english_path = MERGE_DIR / "分类_英文.md"
    category_simp_path = MERGE_DIR / "分类_繁体转简体.md"

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

    simplifier = build_simplifier(convert_trad_to_simp)

    converted_pairs: list[tuple[str, str]] = []
    converted_unique_lines: list[str] = []
    seen_converted: set[str] = set()
    for line in unique_lines:
        converted = to_simplified(line, simplifier)
        converted_pairs.append((line, converted))
        converted_norm = normalize_text(converted)
        if converted_norm in seen_converted:
            continue
        seen_converted.add(converted_norm)
        converted_unique_lines.append(converted)

    merged_path.write_text("\n".join(converted_unique_lines) + ("\n" if converted_unique_lines else ""), encoding="utf-8")
    filtered_path.write_text("\n".join(converted_unique_lines) + ("\n" if converted_unique_lines else ""), encoding="utf-8")

    thai_lines: list[str] = []
    trad_lines: list[str] = []
    english_lines: list[str] = []
    trad_simplified_lines: list[str] = []
    for line in unique_lines:
        if THAI_RE.search(line):
            thai_lines.append(line)
        elif is_traditional_line(line):
            trad_lines.append(line)
            trad_simplified_lines.append(to_simplified(line, simplifier))
        elif include_english and ENGLISH_WORD_RE.search(line):
            english_lines.append(line)

    category_thai_path.write_text("\n".join(thai_lines) + ("\n" if thai_lines else ""), encoding="utf-8")
    category_trad_path.write_text("\n".join(trad_lines) + ("\n" if trad_lines else ""), encoding="utf-8")
    category_english_path.write_text("\n".join(english_lines) + ("\n" if english_lines else ""), encoding="utf-8")
    category_simp_path.write_text(
        "\n".join(dict.fromkeys(trad_simplified_lines)) + ("\n" if trad_simplified_lines else ""),
        encoding="utf-8",
    )

    json_path = find_json_path()
    exact, normalized = load_originals(json_path)
    pending: list[str] = []
    for original, converted in converted_pairs:
        original_norm = normalize_text(original)
        converted_norm = normalize_text(converted)
        if original in exact or original_norm in normalized:
            continue
        if converted in exact or converted_norm in normalized:
            continue
        pending.append(converted)

    pending = list(dict.fromkeys(pending))

    pending_path.write_text("\n".join(pending) + ("\n" if pending else ""), encoding="utf-8")
    print(f"合并后 {len(converted_unique_lines)} 行，待翻译 {len(pending)} 行")
    print(f"待翻译项输出: {pending_path}")
    print(
        f"分类输出: 泰语 {len(thai_lines)} 行, 繁体 {len(trad_lines)} 行, "
        f"繁转简 {len(trad_simplified_lines)} 行, 英文 {len(english_lines)} 行"
    )


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
    parser.add_argument(
        "--include-english",
        action="store_true",
        help="包含英文字段（默认不包含，避免 AACD 等代码样式）",
    )
    parser.add_argument(
        "--no-trad-to-simp",
        action="store_true",
        help="关闭繁体转简体（默认开启）",
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
        count = extract_file(
            dxf_path,
            out_path,
            dedupe=not args.no_dedupe,
            include_english=args.include_english,
        )
        total_lines += count
        print(f"[{i}/{len(files)}] {dxf_path.name} -> {out_name} ({count} 行)")

    print(f"提取完成，共 {total_lines} 行")
    merge_and_filter_pending(
        include_english=args.include_english,
        convert_trad_to_simp=not args.no_trad_to_simp,
    )


if __name__ == "__main__":
    main()
