#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
from pathlib import Path

from openpyxl import Workbook, load_workbook

DEFAULT_MD_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT = DEFAULT_MD_DIR / "文件名翻译表.xlsx"

PATH_LINE_RE = re.compile(r"^[\-\*\d\.)\s]*")


def clean_line_to_path(line: str) -> str:
    """Normalize one markdown line to a possible folder path string."""
    text = line.strip()
    if not text:
        return ""

    # Remove markdown list prefixes like '-', '*', '1.'
    text = PATH_LINE_RE.sub("", text, count=1).strip()

    # Remove wrapping markdown/code punctuation
    text = text.strip("` ")
    if (text.startswith("\"") and text.endswith("\"")) or (
        text.startswith("'") and text.endswith("'")
    ):
        text = text[1:-1].strip()

    return text


def read_folders_from_md(md_dir: Path) -> list[Path]:
    folders: list[Path] = []
    seen: set[Path] = set()

    for md_file in sorted(md_dir.glob("*.md")):
        for raw_line in md_file.read_text(encoding="utf-8").splitlines():
            path_str = clean_line_to_path(raw_line)
            if not path_str:
                continue

            folder = Path(path_str).expanduser()
            if folder in seen:
                continue

            seen.add(folder)
            folders.append(folder)

    return folders


def scan_entries(folders: list[Path]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []

    for folder in folders:
        if not folder.exists() or not folder.is_dir():
            continue

        for fp in sorted(folder.rglob("*")):
            # Skip hidden files/folders like .DS_Store, .git, .svn.
            if any(part.startswith(".") for part in fp.relative_to(folder).parts):
                continue

            if fp.is_dir():
                rows.append(
                    {
                        "item_type": "folder",
                        "source_folder": str(folder),
                        "absolute_path": str(fp),
                        "relative_path": str(fp.relative_to(folder)),
                        "original_name": fp.name,
                        "stem": fp.name,
                        "suffix": "",
                        "translation": "",
                        "name_with_translation": "",
                    }
                )
                continue

            if not fp.is_file():
                continue

            rows.append(
                {
                    "item_type": "file",
                    "source_folder": str(folder),
                    "absolute_path": str(fp),
                    "relative_path": str(fp.relative_to(folder)),
                    "original_name": fp.name,
                    "stem": fp.stem,
                    "suffix": fp.suffix,
                    "translation": "",
                    "name_with_translation": "",
                }
            )

    return rows


def read_existing_excel(excel_path: Path) -> dict[tuple[str, str], dict[str, str]]:
    """Return mapping: (item_type, absolute_path) -> existing row data."""
    if not excel_path.exists():
        return {}

    wb = load_workbook(excel_path)
    ws = wb.active

    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    header_map = {h: i + 1 for i, h in enumerate(headers) if h}

    required = {"absolute_path", "translation"}
    if not required.issubset(header_map):
        wb.close()
        return {}

    type_col = header_map.get("item_type")
    path_col = header_map["absolute_path"]
    trans_col = header_map["translation"]
    source_folder_col = header_map.get("source_folder")
    relative_path_col = header_map.get("relative_path")
    original_name_col = header_map.get("original_name")
    stem_col = header_map.get("stem")
    suffix_col = header_map.get("suffix")
    name_with_translation_col = header_map.get("name_with_translation")

    data: dict[tuple[str, str], dict[str, str]] = {}
    for row_idx in range(2, ws.max_row + 1):
        abs_path = ws.cell(row=row_idx, column=path_col).value
        translation = ws.cell(row=row_idx, column=trans_col).value

        if abs_path is None:
            continue
        key = str(abs_path).strip()
        if not key:
            continue

        trans_text = "" if translation is None else str(translation).strip()
        item_type = "file"
        if type_col is not None:
            cell_type = ws.cell(row=row_idx, column=type_col).value
            if cell_type is not None and str(cell_type).strip().lower() in {"file", "folder"}:
                item_type = str(cell_type).strip().lower()

        row_data = {
            "item_type": item_type,
            "source_folder": "" if source_folder_col is None else str(ws.cell(row=row_idx, column=source_folder_col).value or "").strip(),
            "absolute_path": key,
            "relative_path": "" if relative_path_col is None else str(ws.cell(row=row_idx, column=relative_path_col).value or "").strip(),
            "original_name": "" if original_name_col is None else str(ws.cell(row=row_idx, column=original_name_col).value or "").strip(),
            "stem": "" if stem_col is None else str(ws.cell(row=row_idx, column=stem_col).value or "").strip(),
            "suffix": "" if suffix_col is None else str(ws.cell(row=row_idx, column=suffix_col).value or "").strip(),
            "translation": trans_text,
            "name_with_translation": "" if name_with_translation_col is None else str(ws.cell(row=row_idx, column=name_with_translation_col).value or "").strip(),
        }
        data[(item_type, key)] = row_data

    wb.close()
    return data


def build_name_with_translation(stem: str, suffix: str, translation: str) -> str:
    if not translation:
        return ""
    if stem.endswith(f"({translation})"):
        return f"{stem}{suffix}"
    return f"{stem}({translation}){suffix}"


def ask_yes_no(prompt: str) -> bool:
    answer = input(prompt).strip().lower()
    return answer in {"y", "yes"}


def update_paths_after_folder_rename(
    rows: list[dict[str, str]], old_folder: Path, new_folder: Path
) -> None:
    old_prefix = str(old_folder)
    new_prefix = str(new_folder)

    for row in rows:
        p = row.get("absolute_path", "")
        if p == old_prefix or p.startswith(old_prefix + "/"):
            new_abs = new_prefix + p[len(old_prefix) :]
            row["absolute_path"] = new_abs
            source_folder = row.get("source_folder", "")
            if source_folder:
                sf = Path(source_folder)
                try:
                    row["relative_path"] = str(Path(new_abs).relative_to(sf))
                except ValueError:
                    pass


def rename_entries(rows: list[dict[str, str]]) -> tuple[int, int, int]:
    """Rename files/folders using name_with_translation. Returns (renamed, skipped, conflicts)."""
    renamed = 0
    skipped = 0
    conflicts = 0

    candidates = [r for r in rows if r.get("name_with_translation", "").strip()]
    candidates.sort(key=lambda r: len(Path(r.get("absolute_path", "")).parts), reverse=True)

    for row in candidates:
        new_name = row.get("name_with_translation", "").strip()
        old_path_str = row.get("absolute_path", "").strip()
        item_type = row.get("item_type", "file").strip().lower()

        if not new_name or not old_path_str:
            continue

        old_path = Path(old_path_str)
        if not old_path.exists():
            skipped += 1
            continue

        if item_type == "file" and not old_path.is_file():
            skipped += 1
            continue

        if item_type == "folder" and not old_path.is_dir():
            skipped += 1
            continue

        new_path = old_path.with_name(new_name)
        if new_path == old_path:
            continue

        if new_path.exists():
            conflicts += 1
            continue

        old_path.rename(new_path)

        if item_type == "folder":
            update_paths_after_folder_rename(rows, old_path, new_path)

        row["absolute_path"] = str(new_path)
        try:
            row["relative_path"] = str(new_path.relative_to(Path(row["source_folder"])))
        except ValueError:
            pass
        row["original_name"] = new_path.name
        if item_type == "file":
            row["stem"] = new_path.stem
            row["suffix"] = new_path.suffix
        else:
            row["stem"] = new_path.name
            row["suffix"] = ""
        renamed += 1

    return renamed, skipped, conflicts


def write_excel(excel_path: Path, rows: list[dict[str, str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "file_names"

    headers = [
        "item_type",
        "source_folder",
        "absolute_path",
        "relative_path",
        "original_name",
        "stem",
        "suffix",
        "translation",
        "name_with_translation",
    ]

    ws.append(headers)

    for row in rows:
        ws.append([row.get(h, "") for h in headers])

    ws.freeze_panes = "A2"
    wb.save(excel_path)
    wb.close()


def run(md_dir: Path, output_excel: Path) -> None:
    folders = read_folders_from_md(md_dir)
    existing_rows = read_existing_excel(output_excel)
    rows = scan_entries(folders)

    scanned_keys = {(r["item_type"], r["absolute_path"]) for r in rows}

    for row in rows:
        key = (row["item_type"], row["absolute_path"])
        tr = existing_rows.get(key, {}).get("translation", "")
        row["translation"] = tr
        row["name_with_translation"] = build_name_with_translation(
            row["stem"], row["suffix"], tr
        )

    # Incremental mode: keep historical rows that are no longer currently scanned.
    for key, old_row in existing_rows.items():
        if key in scanned_keys:
            continue
        rows.append(old_row)

    write_excel(output_excel, rows)

    print(f"读取到目录数量: {len(folders)}")
    print(f"扫描到文件数量: {len(rows)}")
    print(f"输出Excel: {output_excel}")
    print("说明: 支持 file/folder 两类条目，translation可人工填写，name_with_translation会自动生成")

    rename_candidates = [r for r in rows if r.get("name_with_translation", "").strip()]
    if not rename_candidates:
        print("未发现可重命名项（translation列为空）。")
        return

    print(f"可重命名文件数量: {len(rename_candidates)}")
    should_rename = ask_yes_no("是否执行文件重命名？输入 y 确认，其它键取消: ")
    if not should_rename:
        print("已取消重命名。")
        return

    renamed, skipped, conflicts = rename_entries(rows)
    write_excel(output_excel, rows)
    print(f"重命名完成: 成功 {renamed}，跳过 {skipped}，冲突 {conflicts}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="读取md中的目录路径，导出文件名到Excel，并复用译名生成括号后缀名称"
    )
    parser.add_argument(
        "--md-dir",
        default=str(DEFAULT_MD_DIR),
        help="存放md路径文件的目录，默认脚本所在目录",
    )
    parser.add_argument(
        "--output-excel",
        default=str(DEFAULT_OUTPUT),
        help="输出Excel路径，默认: 文件名翻译表.xlsx",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    md_dir = Path(args.md_dir).expanduser().resolve()
    output_excel = Path(args.output_excel).expanduser().resolve()

    if not md_dir.exists() or not md_dir.is_dir():
        raise FileNotFoundError(f"md目录不存在: {md_dir}")

    output_excel.parent.mkdir(parents=True, exist_ok=True)
    run(md_dir, output_excel)


if __name__ == "__main__":
    main()
