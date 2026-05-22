#!/usr/bin/env python3
from __future__ import annotations
import argparse
import json
import re
import tempfile
from pathlib import Path
from typing import Iterable

try:
    import ezdxf
except ImportError as exc:
    raise SystemExit("未安装依赖 ezdxf，请先执行: pip install ezdxf") from exc

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
EXCEL_OUTPUT = BASE_DIR / "1_提取文本" / "文字提取翻译表.xlsx"

# ==================== 正则 ====================
_MD_PATH_STRIP_RE = re.compile(r"^[\-\*\d\.)\s]*")
THAI_RE = re.compile(r"[\u0e00-\u0e7f]")
ENGLISH_WORD_RE = re.compile(r"[A-Za-z]{3,}")
CHINESE_RE = re.compile(r"[\u4e00-\u9fff]")
PURE_NUMBER_RE = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")
INDEX_NUMBER_RE = re.compile(r"^\d+(?:\.\d+)+$")
SHORT_ALNUM_CODE_RE = re.compile(r"^[A-Za-z]{1,6}\d+(?:\.\d+)*$|^\d+[A-Za-z]{1,4}$")
CAD_META_RE = re.compile(r"^(?:AcDb|AcCm|ByLayer|ByBlock|Model|Paper|STANDARD|Standard|Layer|LAYER|BLOCK|INSERT|LINE|CIRCLE|ARC|DIM|STYLE|UCS|VIEW|TABLE|XREF|HANDLE|OWNER)")
HEX_HANDLE_RE = re.compile(r"^[A-F0-9]{3,10}$")
STAR_D_RE = re.compile(r"^\*D\d+$")
AUTOCAD_ANON_RE = re.compile(r"^A\$C[0-9A-F]{6,}$")
UPPER_CODE_RE = re.compile(r"^[A-Z][A-Z0-9_\-]{1,15}$")
HAS_LETTER_RE = re.compile(r"[A-Za-z]")

# ==================== 最终强化过滤（已加入高程、C-ROAD-TEXT 等） ====================
SHORT_CODE_RE = re.compile(r"^(?:0-|fc-|gv-|pc:|pt:|ep:|pi:|l=|\d+\+|sta\.|km\.|phase|p\d+|road|line_)[a-z0-9\-]+$", re.IGNORECASE)
GARBAGE_RE = re.compile(r"[=[\]{}<>^`~\\|]+|wx |ctg=|v\"|yf=|9e\[|mk'|s,kpg|ang|sym_|a\$c[0-9a-f]{6,}", re.IGNORECASE)
PLOT_STAMP_RE = re.compile(r"^(?:plot stamp|pl ot|rev\.|app\.|date|no\.|sheet|title|logo|ieat)", re.IGNORECASE)
STATION_RE = re.compile(r"(?:BP\.|STA\.|PC:|PT:|EP:|PI:)\s*\d+\+\d", re.IGNORECASE)
RAI_RE = re.compile(r"\d+\.?\d*\s*RAI", re.IGNORECASE)
LENGTH_RE = re.compile(r"L=\s*\d+\.?\d*\s*M\.", re.IGNORECASE)
ELEVATION_RE = re.compile(r"^(?:ELVE\.|ELE |EL\.|ELEV)\s*\d+\.?\d*", re.IGNORECASE)
C_ROAD_TEXT_RE = re.compile(r"C-ROAD-TEXT-", re.IGNORECASE)
DIM_RE = re.compile(r"^Dim-\d", re.IGNORECASE)
TYPICAL_RE = re.compile(r"^typical-rd", re.IGNORECASE)
FAC_CON_RE = re.compile(r"^(?:0-|Fac Con|Factory Con|Fire Hydrant|Blowe Off|Drop Off|D-off)", re.IGNORECASE)
LAYER_LIKE_RE = re.compile(r"^0-[A-Za-z0-9_\-$&\.]+$")
SHORT_ASCII_TOKEN_RE = re.compile(r"^[A-Za-z]{1,2}$|^[A-Za-z0-9]{1,2}$")
DIM_MARK_RE = re.compile(r"^Dim\s*\d+$", re.IGNORECASE)
SECTION_MARK_RE = re.compile(r"^[A-Za-z]\s*-\s*[A-Za-z]$")
TEXT_STYLE_RE = re.compile(r"^Text_t\d+(?:\s*CN)?$", re.IGNORECASE)
ARROW_MARK_RE = re.compile(r"^Arrow_\d+$", re.IGNORECASE)
GROUP_SPLIT_RE = re.compile(r"[;；|｜、，]+|\s{2,}|\s/\s")
STAR_X_RE = re.compile(r"^\*X\d+$")
EXCEL_ILLEGAL_CHAR_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())

def sanitize_excel_text(text: str) -> str:
    cleaned = EXCEL_ILLEGAL_CHAR_RE.sub("", text)
    if len(cleaned) > 32767:
        cleaned = cleaned[:32767]
    return cleaned

def is_noise_line(text: str) -> bool:
    compact = text.strip()
    if not compact or len(compact) < 2:
        return True
    if len(compact) > 300:
        return True
    norm = compact.replace(" ", "").replace(",", "")
    if PURE_NUMBER_RE.fullmatch(norm) or INDEX_NUMBER_RE.fullmatch(compact):
        return True
    if SHORT_ALNUM_CODE_RE.fullmatch(compact) or SHORT_CODE_RE.fullmatch(compact):
        return True
    if GARBAGE_RE.search(compact):
        return True
    if PLOT_STAMP_RE.search(compact):
        return True
    if STATION_RE.search(compact):
        return True
    if RAI_RE.search(compact):
        return True
    if LENGTH_RE.search(compact):
        return True
    if ELEVATION_RE.search(compact):
        return True
    if C_ROAD_TEXT_RE.search(compact):
        return True
    if DIM_RE.search(compact):
        return True
    if TYPICAL_RE.search(compact):
        return True
    if FAC_CON_RE.search(compact):
        return True
    if AUTOCAD_ANON_RE.fullmatch(compact) or STAR_D_RE.fullmatch(compact):
        return True
    if HEX_HANDLE_RE.fullmatch(compact):
        return True
    if UPPER_CODE_RE.fullmatch(compact) and " " not in compact and len(compact) < 20:
        return True
    if not HAS_LETTER_RE.search(compact) and not THAI_RE.search(compact) and not CHINESE_RE.search(compact):
        return True
    return False

def is_low_value_translation_text(text: str) -> bool:
    t = text.strip()
    if not t:
        return True
    if is_noise_line(t):
        return True
    if STAR_X_RE.fullmatch(t) or LAYER_LIKE_RE.fullmatch(t):
        return True
    if t.lower().endswith((".shx", ".dwg", ".dxf")):
        return True
    if "{\\" in t or "%%" in t or "\\P" in t:
        return True
    if SHORT_ASCII_TOKEN_RE.fullmatch(t):
        return True
    if DIM_MARK_RE.fullmatch(t) or SECTION_MARK_RE.fullmatch(t):
        return True
    if TEXT_STYLE_RE.fullmatch(t) or ARROW_MARK_RE.fullmatch(t):
        return True
    if len(t) < 5 and " " not in t and not THAI_RE.search(t):
        return True
    return False

def is_readable_text_any(text: str) -> bool:
    compact = text.strip()
    if not compact or len(compact) < 3:
        return False
    if CAD_META_RE.match(compact):
        return False
    if is_noise_line(compact) or is_low_value_translation_text(compact):
        return False
    has_thai = THAI_RE.search(compact)
    has_chinese = CHINESE_RE.search(compact)
    has_english_phrase = ENGLISH_WORD_RE.search(compact) and " " in compact and len(compact) >= 8
    return bool(has_thai or has_chinese or has_english_phrase)

def split_entity_text(text: str) -> list[str]:
    normalized = text.replace("\\P", "\n").replace("\r\n", "\n").replace("\r", "\n")
    return [line.strip() for line in normalized.split("\n") if line.strip()]

def split_grouped_line(line: str) -> list[str]:
    chunks = [line]
    next_chunks: list[str] = []
    for chunk in chunks:
        for part in GROUP_SPLIT_RE.split(chunk):
            part = part.strip()
            if part:
                next_chunks.append(part)
    chunks = next_chunks
    numbered: list[str] = []
    for chunk in chunks:
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
    if not enable or OpenCC is None:
        return None
    return OpenCC("t2s")

def to_simplified(text: str, simplifier) -> str:
    if simplifier is None or not CHINESE_RE.search(text):
        return text
    return simplifier.convert(text)

def iter_spaces(doc):
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
        index_md = input_path / "dxf文件路径.md"
        if index_md.exists():
            return collect_dxf_from_index_md(index_md)
        return []
    raise FileNotFoundError(f"路径不存在: {input_path}")

def collect_dxf_from_index_md(md_path: Path) -> list[Path]:
    if not md_path.exists():
        return []
    folders: list[Path] = []
    seen_f: set[Path] = set()
    for raw in md_path.read_text(encoding="utf-8").splitlines():
        path_str = _MD_PATH_STRIP_RE.sub("", raw.strip(), count=1).strip().strip("` '\"")
        if not path_str:
            continue
        p = Path(path_str).expanduser()
        if p not in seen_f:
            seen_f.add(p)
            folders.append(p)
    files: list[Path] = []
    for folder in folders:
        if not folder.exists() or not folder.is_dir():
            continue
        for fp in sorted(folder.rglob("*")):
            if fp.is_file() and fp.suffix.lower() in DXF_SUFFIXES and not any(part.startswith(".") for part in fp.relative_to(folder).parts):
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

    def append_row(values: list[str]):
        nonlocal rows_in_sheet, ws, sheet_index
        if rows_in_sheet >= max_excel_rows:
            sheet_index += 1
            ws = wb.create_sheet(title=f"dxf_texts_{sheet_index}")
            ws.append(headers)
            rows_in_sheet = 1
        ws.append(values)
        rows_in_sheet += 1

    total = 0
    translated = 0
    total_files = len(dxf_files)
    for idx, dxf_path in enumerate(dxf_files, start=1):
        try:
            doc = ezdxf.readfile(dxf_path)
        except Exception as exc:
            print(f" ⚠️ 读取失败: {dxf_path.name}: {exc}")
            continue
        before_total = total
        seen_in_file: set[str] = set()
        for raw_text in iter_entity_texts(doc):
            for line in split_entity_text(raw_text):
                for piece in split_grouped_line(line):
                    piece = sanitize_excel_text(piece)
                    norm = normalize_text(piece)
                    if not norm:
                        continue
                    if CAD_META_RE.match(norm):
                        continue
                    if is_noise_line(norm):
                        continue
                    if is_low_value_translation_text(norm):
                        continue
                    if not is_readable_text_any(norm):
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
        added = total - before_total
        if idx == 1 or idx % 5 == 0 or idx == total_files:
            print(f"[{idx}/{total_files}] {dxf_path.name} 新增 {added} 条，累计 {total} 条")

    excel_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_file = None
    try:
        with tempfile.NamedTemporaryFile(prefix=f"{excel_path.stem}.", suffix=".tmp.xlsx", dir=str(excel_path.parent), delete=False) as tf:
            tmp_file = Path(tf.name)
        wb.save(tmp_file)
        tmp_file.replace(excel_path)
    finally:
        if tmp_file and tmp_file.exists() and tmp_file != excel_path:
            tmp_file.unlink(missing_ok=True)
    wb.close()
    return total, translated

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="提取CAD(DXF)中的泰文文本，生成待翻译项")
    parser.add_argument("input_path", nargs="?", default=str(INPUT_DIR), help="待处理DXF文件或目录")
    parser.add_argument("--no-dedupe", action="store_true", help="保留重复行")
    parser.add_argument("--include-english", action="store_true", help="额外包含英文字段")
    parser.add_argument("--thai-cn-only", action="store_true", help="仅保留中文/泰文")
    parser.add_argument("--no-trad-to-simp", action="store_true", help="关闭繁体转简体")
    parser.add_argument("--no-excel-output", action="store_true", help="关闭Excel输出")
    return parser.parse_args()

def main() -> None:
    args = parse_args()
    input_path = Path(args.input_path).expanduser().resolve()
    files = collect_dxf_files(input_path)
    if not files:
        print("未找到可处理的DXF文件")
        return
    print(f"找到 {len(files)} 个DXF文件")
    if not args.no_excel_output:
        excel_path = EXCEL_OUTPUT
        if _HAS_OPENPYXL:
            print("开始提取并写入Excel，请稍候（会周期输出进度）...")
            total, translated = extract_all_texts_to_excel(files, excel_path)
            print(f"✅ Excel已输出: {excel_path}")
            print(f"提取条数: {total}，已有译文: {translated}")
        else:
            print("未安装 openpyxl，请执行: pip install openpyxl")

if __name__ == "__main__":
    main()