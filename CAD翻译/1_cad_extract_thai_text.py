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

try:
    from openpyxl import Workbook, load_workbook
    _HAS_OPENPYXL = True
except ImportError:
    _HAS_OPENPYXL = False

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "00_dxf文件"
EXTRACT_DIR = BASE_DIR / "1_提取文本"
MERGE_DIR = BASE_DIR / "2_提取文件合并文件"
JSON_CANDIDATES = [
    BASE_DIR / "翻译对照.json",
    BASE_DIR / "4_输出文件cad" / "翻译对照.json",
]
DXF_SUFFIXES = {".dxf"}
DXF_INDEX_MD = BASE_DIR / "00_dxf文件" / "dxf文件路径.md"
EXCEL_OUTPUT = BASE_DIR / "1_提取文本" / "文字提取翻译表.xlsx"
CAD_OUTPUT_DIR = BASE_DIR / "4_输出文件cad"
_MD_PATH_STRIP_RE = re.compile(r"^[\-\*\d\.)\s]*")

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
HEX_HANDLE_RE = re.compile(r"^[A-F0-9]{3,10}$")
STAR_D_RE = re.compile(r"^\*D\d+$")
AUTOCAD_ANON_RE = re.compile(r"^A\$C[0-9A-F]{6,}$")
UPPER_CODE_RE = re.compile(r"^[A-Z][A-Z0-9_\-]{1,15}$")
HAS_LETTER_RE = re.compile(r"[A-Za-z]")
GROUP_SPLIT_RE = re.compile(r"[;；|｜、，]+|\s{2,}|\s/\s")


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
    if STAR_D_RE.fullmatch(compact):
        return True
    if AUTOCAD_ANON_RE.fullmatch(compact):
        return True
    if HEX_HANDLE_RE.fullmatch(compact):
        return True

    # 过滤单个大写代码词（如 PEGL/PDGL/PVGRID），保留含空格的自然短语。
    if " " not in compact and UPPER_CODE_RE.fullmatch(compact):
        return True

    # 过滤纯符号与格式控制碎片。
    if not HAS_LETTER_RE.search(compact) and not THAI_RE.search(compact) and not CHINESE_RE.search(compact):
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


def is_readable_text_any(text: str) -> bool:
    compact = text.strip()
    if not compact:
        return False
    if CAD_META_RE.match(compact):
        return False
    return True


def split_entity_text(text: str) -> list[str]:
    normalized = text.replace("\\P", "\n").replace("\r\n", "\n").replace("\r", "\n")
    return [line.strip() for line in normalized.split("\n") if line.strip()]


def split_grouped_line(line: str) -> list[str]:
    """把一行内的编组内容拆开：分号/分隔符/多空格/编号分段。"""
    chunks = [line]

    # 先按常见分组分隔符拆分。
    next_chunks: list[str] = []
    for chunk in chunks:
        for part in GROUP_SPLIT_RE.split(chunk):
            part = part.strip()
            if part:
                next_chunks.append(part)
    chunks = next_chunks

    # 再按行内编号拆分，例如: "1. ... 2. ..."
    numbered: list[str] = []
    for chunk in chunks:
        # 仅在“编号. 空格”场景拆分，避免把 0.20 这类小数误拆。
        parts = re.split(r"(?<!^)\s*(?=\d+\.\s+)", chunk)
        for part in parts:
            part = part.strip()
            if part:
                numbered.append(part)
    return numbered


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
        direct_files = sorted(
            f for f in input_path.rglob("*")
            if f.suffix.lower() in DXF_SUFFIXES and not f.name.startswith("~$")
        )
        if direct_files:
            return direct_files

        # 目录下若没有实际 DXF，尝试从 dxf文件路径.md 指向的外部目录收集。
        index_md = input_path / "dxf文件路径.md"
        if index_md.exists():
            return collect_dxf_from_index_md(index_md)

        return []
    raise FileNotFoundError(f"路径不存在: {input_path}")


def extract_file(
    dxf_path: Path,
    out_path: Path,
    dedupe: bool = True,
    include_english: bool = False,
    all_fields: bool = True,
) -> int:
    doc = ezdxf.readfile(dxf_path)
    seen: set[str] = set()
    lines: list[str] = []

    for raw_text in iter_entity_texts(doc):
        for line in split_entity_text(raw_text):
            for piece in split_grouped_line(line):
                if is_noise_line(piece):
                    continue
                if all_fields:
                    keep = is_readable_text_any(piece)
                elif include_english:
                    keep = is_readable_text_with_english(piece)
                else:
                    keep = is_readable_text(piece)
                if not keep:
                    continue
                normalized = normalize_text(piece)
                if dedupe and normalized in seen:
                    continue
                seen.add(normalized)
                lines.append(piece)

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


# ── Excel 工作流：从 dxf文件路径.md 提取所有文字 ────────────────────────────────

def _clean_md_line_to_path(line: str) -> str:
    text = line.strip()
    if not text:
        return ""
    text = _MD_PATH_STRIP_RE.sub("", text, count=1).strip().strip("` ")
    if (text.startswith('"') and text.endswith('"')) or (
        text.startswith("'") and text.endswith("'")
    ):
        text = text[1:-1].strip()
    return text


def collect_dxf_from_index_md(md_path: Path) -> list[Path]:
    """读取 dxf文件路径.md，收集其中文件夹下所有 DXF 文件。"""
    if not md_path.exists():
        return []
    folders: list[Path] = []
    seen_f: set[Path] = set()
    for raw in md_path.read_text(encoding="utf-8").splitlines():
        path_str = _clean_md_line_to_path(raw)
        if not path_str:
            continue
        p = Path(path_str).expanduser()
        if p not in seen_f:
            seen_f.add(p)
            folders.append(p)
    files: list[Path] = []
    for folder in folders:
        if not folder.exists() or not folder.is_dir():
            print(f"⚠️  目录不存在，跳过: {folder}")
            continue
        for fp in sorted(folder.rglob("*")):
            if fp.is_file() and fp.suffix.lower() in DXF_SUFFIXES:
                if not any(part.startswith(".") for part in fp.relative_to(folder).parts):
                    files.append(fp)
    return files


def _load_excel_translations(excel_path: Path) -> dict[tuple[str, str], str]:
    data: dict[tuple[str, str], str] = {}
    if not excel_path.exists() or not _HAS_OPENPYXL:
        return data
    wb = load_workbook(excel_path)
    for ws in wb.worksheets:
        headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
        hmap = {h: i + 1 for i, h in enumerate(headers) if h}
        if not {"absolute_path", "original_text", "translation"}.issubset(hmap):
            continue
        pc, tc, rc = hmap["absolute_path"], hmap["original_text"], hmap["translation"]
        for row in range(2, ws.max_row + 1):
            p = ws.cell(row=row, column=pc).value
            t = ws.cell(row=row, column=tc).value
            r = ws.cell(row=row, column=rc).value
            if p and t and r:
                data[(str(p).strip(), str(t).strip())] = str(r).strip()
    wb.close()
    return data


def extract_all_texts_to_excel(dxf_files: list[Path], excel_path: Path) -> tuple[int, int]:
    """提取所有文字字段到 Excel（优先完整性），保留已有译文。返回 (总条数, 已有译文条数)。"""
    if not _HAS_OPENPYXL:
        raise SystemExit("未安装 openpyxl，请执行: pip install openpyxl")
    existing = _load_excel_translations(excel_path)
    max_excel_rows = 1_048_576
    headers = ["absolute_path", "file_name", "original_text", "translation"]

    wb = Workbook(write_only=True)
    sheet_index = 1
    ws = wb.create_sheet(title="dxf_texts")
    ws.append(headers)
    rows_in_sheet = 1

    def append_row(values: list[str]) -> tuple[int, int, object]:
        nonlocal rows_in_sheet, ws, sheet_index
        if rows_in_sheet >= max_excel_rows:
            sheet_index += 1
            ws = wb.create_sheet(title=f"dxf_texts_{sheet_index}")
            ws.append(headers)
            rows_in_sheet = 1
        ws.append(values)
        rows_in_sheet += 1
        return rows_in_sheet, sheet_index, ws

    total = 0
    translated = 0
    for dxf_path in dxf_files:
        try:
            doc = ezdxf.readfile(dxf_path)
        except Exception as exc:
            print(f"  ⚠️  读取失败: {dxf_path.name}: {exc}")
            continue
        seen_in_file: set[str] = set()
        for raw_text in iter_entity_texts(doc):
            for line in split_entity_text(raw_text):
                for piece in split_grouped_line(line):
                    norm = normalize_text(piece)
                    if not norm:
                        continue
                    # Excel主表优先完整性：仅过滤明显CAD元字段，避免漏提可翻译文本。
                    if CAD_META_RE.match(norm):
                        continue
                    if norm in seen_in_file:
                        continue
                    seen_in_file.add(norm)
                    key = (str(dxf_path), norm)
                    trans = existing.get(key, "")
                    append_row([str(dxf_path), dxf_path.name, norm, trans])
                    total += 1
                    if trans:
                        translated += 1
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(excel_path)
    wb.close()
    return total, translated


def _translate_text_with_map(text: str, trans_map: dict[str, str]) -> str:
    parts = text.replace("\\P", "\n").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    new_parts: list[str] = []
    for part in parts:
        norm = normalize_text(part.strip())
        if norm in trans_map:
            leading = len(part) - len(part.lstrip())
            trailing = len(part) - len(part.rstrip())
            prefix = part[:leading]
            suffix = part[len(part) - trailing:] if trailing else ""
            new_parts.append(f"{prefix}{trans_map[norm]}{suffix}")
        else:
            new_parts.append(part)
    return "\n".join(new_parts)


def _apply_trans_to_doc(doc, trans_map: dict[str, str]) -> int:
    replaced = 0
    for space in iter_spaces(doc):
        for entity in space:
            etype = entity.dxftype()
            if etype == "TEXT":
                old = str(entity.dxf.text)
                new = _translate_text_with_map(old, trans_map)
                if new != old:
                    entity.dxf.text = new
                    replaced += 1
            elif etype == "MTEXT":
                old = str(entity.text)
                new = _translate_text_with_map(old, trans_map)
                if new != old:
                    entity.text = new
                    replaced += 1
            elif etype in {"ATTRIB", "ATTDEF"}:
                old = str(entity.dxf.text)
                new = _translate_text_with_map(old, trans_map)
                if new != old:
                    entity.dxf.text = new
                    replaced += 1
    return replaced


def apply_translations_from_excel(excel_path: Path, output_dir: Path) -> None:
    """读取 Excel translation 列，替换对应 DXF 文件，输出到 output_dir。"""
    if not _HAS_OPENPYXL:
        raise SystemExit("未安装 openpyxl，请执行: pip install openpyxl")
    if not excel_path.exists():
        print(f"❌ Excel 不存在: {excel_path}")
        return
    wb = load_workbook(excel_path)
    file_trans: dict[str, dict[str, str]] = {}
    valid_sheet_found = False
    for ws in wb.worksheets:
        headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
        hmap = {h: i + 1 for i, h in enumerate(headers) if h}
        if not {"absolute_path", "original_text", "translation"}.issubset(hmap):
            continue
        valid_sheet_found = True
        pc, tc, rc = hmap["absolute_path"], hmap["original_text"], hmap["translation"]
        for row in range(2, ws.max_row + 1):
            p = ws.cell(row=row, column=pc).value
            t = ws.cell(row=row, column=tc).value
            r = ws.cell(row=row, column=rc).value
            if p and t and r:
                p, t, r = str(p).strip(), str(t).strip(), str(r).strip()
                if r:
                    file_trans.setdefault(p, {})[t] = r
    if not valid_sheet_found:
        print("❌ Excel 缺少必要列: absolute_path / original_text / translation")
        wb.close()
        return
    wb.close()
    if not file_trans:
        print("未发现任何翻译条目，translation 列全为空。")
        return
    total_entries = sum(len(v) for v in file_trans.values())
    print(f"找到 {len(file_trans)} 个文件、{total_entries} 条翻译")
    output_dir.mkdir(parents=True, exist_ok=True)
    replaced_total = 0
    for dxf_path_str, trans_map in sorted(file_trans.items()):
        dxf_path = Path(dxf_path_str)
        if not dxf_path.exists():
            print(f"  ⚠️  文件不存在，跳过: {dxf_path_str}")
            continue
        dst = output_dir / dxf_path.name
        try:
            doc = ezdxf.readfile(dxf_path)
        except Exception as exc:
            print(f"  ⚠️  读取失败: {dxf_path.name}: {exc}")
            continue
        replaced = _apply_trans_to_doc(doc, trans_map)
        doc.saveas(dst)
        replaced_total += replaced
        print(f"  ✅ {dxf_path.name} → {dst.name}  ({replaced} 处替换)")
    print(f"\n翻译替换完成，共 {replaced_total} 处，输出目录: {output_dir}")


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
        help="在仅中文/泰文模式下，额外包含英文字段",
    )
    parser.add_argument(
        "--thai-cn-only",
        action="store_true",
        help="仅保留中文/泰文（可配合 --include-english），默认提取所有文本字段",
    )
    parser.add_argument(
        "--no-trad-to-simp",
        action="store_true",
        help="关闭繁体转简体（默认开启）",
    )
    parser.add_argument(
        "--excel",
        action="store_true",
        help="兼容旧参数：强制走Excel流程（当前默认即为Excel流程）",
    )
    parser.add_argument(
        "--excel-output",
        default=str(EXCEL_OUTPUT),
        help=f"Excel 输出路径，默认: {EXCEL_OUTPUT}",
    )
    parser.add_argument(
        "--no-excel-output",
        action="store_true",
        help="关闭Excel输出（一般不建议）",
    )
    parser.add_argument(
        "--generate-md",
        action="store_true",
        help="额外生成md提取与合并结果（默认关闭，仅保留Excel）",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    # ── 默认流程：直接输出Excel（不再强制生成md） ───────────────────────────────
    input_path = Path(args.input_path).expanduser().resolve()
    files = collect_dxf_files(input_path)

    if not files:
        print("未找到可处理的DXF文件")
        if input_path.is_dir():
            print("提示: 请确认目录内存在 .dxf，或在 dxf文件路径.md 中填写可访问目录。")
        return

    print(f"找到 {len(files)} 个DXF文件")

    if args.generate_md:
        EXTRACT_DIR.mkdir(parents=True, exist_ok=True)
        total_lines = 0
        all_fields = not args.thai_cn_only
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
                all_fields=all_fields,
            )
            total_lines += count
            print(f"[{i}/{len(files)}] {dxf_path.name} -> {out_name} ({count} 行)")

        print(f"MD提取完成，共 {total_lines} 行")
        merge_and_filter_pending(
            include_english=(not args.thai_cn_only) or args.include_english,
            convert_trad_to_simp=not args.no_trad_to_simp,
        )

    if not args.no_excel_output:
        excel_path = Path(args.excel_output).expanduser().resolve()
        if _HAS_OPENPYXL:
            total, translated = extract_all_texts_to_excel(files, excel_path)
            print(f"Excel已输出: {excel_path}")
            print(f"Excel提取条数: {total}，已有译文: {translated}")
            if translated == 0:
                print("translation 列全为空，请填写译文后再次运行。")
        else:
            print("未安装 openpyxl，已跳过 Excel 输出。可执行: pip install openpyxl")


if __name__ == "__main__":
    main()
