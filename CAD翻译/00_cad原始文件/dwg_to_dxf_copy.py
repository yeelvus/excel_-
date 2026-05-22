#!/usr/bin/env python3
"""
dwg_to_dxf_copy.py
==================
读取同目录 cad文件路径.md 中的源文件夹路径，将整个目录复制到 00_dxf文件/ 下：
  - .dwg 文件 → 转换为同名 .dxf
  - 其余文件  → 直接复制原样

转换优先级（自动检测可用工具）：
  1. ODA File Converter  (/Applications/ODAFileConverter*.app  或 PATH)
  2. LibreDWG dwg2dxf   (brew install libredwg)
  3. 均不可用时：报告哪些文件无法转换，其余文件仍会复制
"""
from __future__ import annotations

import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# ── 路径配置 ──────────────────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
MD_FILE = SCRIPT_DIR / "cad文件路径.md"
CAD_BASE = SCRIPT_DIR.parent             # CAD翻译/
DXF_INDEX_DIR = CAD_BASE / "00_dxf文件"  # 用于存放输出路径 md
DXF_INDEX_MD = DXF_INDEX_DIR / "dxf文件路径.md"

PATH_LINE_RE = re.compile(r"^[\-\*\d\.)\s]*")


# ── 读取 md 里的路径 ──────────────────────────────────────────────────────────
def clean_line_to_path(line: str) -> str:
    text = line.strip()
    if not text:
        return ""
    text = PATH_LINE_RE.sub("", text, count=1).strip()
    text = text.strip("` ")
    if (text.startswith('"') and text.endswith('"')) or \
       (text.startswith("'") and text.endswith("'")):
        text = text[1:-1].strip()
    return text


def read_source_folders(md_path: Path) -> list[Path]:
    folders: list[Path] = []
    seen: set[Path] = set()
    for raw in md_path.read_text(encoding="utf-8").splitlines():
        path_str = clean_line_to_path(raw)
        if not path_str:
            continue
        p = Path(path_str).expanduser()
        if p not in seen:
            seen.add(p)
            folders.append(p)
    return folders


# ── 检测可用的 DWG→DXF 转换工具 ─────────────────────────────────────────────
def find_oda_converter() -> Path | None:
    # 检查 PATH
    oda = shutil.which("ODAFileConverter")
    if oda:
        return Path(oda)
    # 检查 /Applications 下的 .app
    for app in sorted(Path("/Applications").glob("ODAFileConverter*.app"), reverse=True):
        exe = app / "Contents" / "MacOS" / "ODAFileConverter"
        if exe.exists():
            return exe
    return None


def find_dwg2dxf() -> Path | None:
    cmd = shutil.which("dwg2dxf")
    return Path(cmd) if cmd else None


def convert_dwg_oda(oda_exe: Path, dwg_path: Path, out_dir: Path) -> Path | None:
    """ODA File Converter: 输入目录 → 输出目录，批量转换。返回预期 dxf 路径。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [
            str(oda_exe),
            str(dwg_path.parent),  # 输入目录
            str(out_dir),          # 输出目录
            "ACAD2018",            # 版本（转换目标格式标识，DXF）
            "DXF",
            "0",                   # 递归：0=只当前目录
            "1",                   # 审计模式
            f"{dwg_path.name}",    # 只转这一个文件
        ],
        capture_output=True,
        text=True,
    )
    dxf_path = out_dir / (dwg_path.stem + ".dxf")
    if dxf_path.exists():
        return dxf_path
    print(f"  ⚠️  ODA 转换失败: {dwg_path.name}\n    {result.stderr.strip()[:200]}")
    return None


def convert_dwg_libredwg(dwg2dxf_exe: Path, dwg_path: Path, dxf_path: Path) -> bool:
    dxf_path.parent.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [str(dwg2dxf_exe), str(dwg_path), "-o", str(dxf_path)],
        capture_output=True,
        text=True,
    )
    if result.returncode == 0 and dxf_path.exists():
        return True
    print(f"  ⚠️  LibreDWG 转换失败: {dwg_path.name}\n    {result.stderr.strip()[:200]}")
    return False


# ── 主复制 + 转换逻辑 ─────────────────────────────────────────────────────────
def make_dxf_output_path(src_folder: Path) -> Path:
    """输出目录与源目录同级，文件夹名追加 （dxf）"""
    return src_folder.parent / f"{src_folder.name}（dxf）"


def write_dxf_index_md(index_dir: Path, index_md: Path, output_folders: list[Path]) -> None:
    """把所有输出目录路径写入 dxf文件路径.md（每次覆盖）"""
    index_dir.mkdir(parents=True, exist_ok=True)
    lines = [f"'{p}'" for p in output_folders]
    index_md.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"\n已写入路径索引: {index_md}")


def process_folder(
    src_folder: Path,
    oda_exe: Path | None,
    dwg2dxf_exe: Path | None,
) -> tuple[int, int, int, list[str], Path]:
    """
    把 src_folder 下所有内容复制到同级的 src_folder（dxf）/ 下。
    返回 (copied_files, converted_dwg, failed_dwg, failed_names, dst_root)
    """
    dst_root = make_dxf_output_path(src_folder)
    copied = 0
    converted = 0
    failed = 0
    failed_names: list[str] = []

    all_items = sorted(src_folder.rglob("*"))

    # ODA 需要按目录批量调用，先收集每个 dwg 所在的（唯一）父目录
    dwg_files: list[Path] = []

    for src in all_items:
        # 跳过隐藏文件/目录
        if any(part.startswith(".") for part in src.relative_to(src_folder).parts):
            continue
        if not src.is_file():
            continue

        rel = src.relative_to(src_folder)
        dst = dst_root / rel

        if src.suffix.lower() != ".dwg":
            # 普通文件直接复制
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            copied += 1
            continue

        # DWG 文件：转换为 DXF
        dxf_dst = dst.with_suffix(".dxf")
        dxf_dst.parent.mkdir(parents=True, exist_ok=True)

        ok = False
        if oda_exe:
            with tempfile.TemporaryDirectory() as tmp_dir:
                result_dxf = convert_dwg_oda(oda_exe, src, Path(tmp_dir))
                if result_dxf and result_dxf.exists():
                    shutil.move(str(result_dxf), str(dxf_dst))
                    ok = True
        elif dwg2dxf_exe:
            ok = convert_dwg_libredwg(dwg2dxf_exe, src, dxf_dst)

        if ok:
            converted += 1
            print(f"  ✅ 转换: {rel}  →  {dxf_dst.name}")
        else:
            failed += 1
            failed_names.append(str(rel))
            if not oda_exe and not dwg2dxf_exe:
                # 没有转换工具，复制原 dwg
                shutil.copy2(src, dst)
                print(f"  ⚠️  无转换工具，复制原文件: {rel}")

    return copied, converted, failed, failed_names, dst_root


# ── 入口 ──────────────────────────────────────────────────────────────────────
def main() -> None:
    if not MD_FILE.exists():
        print(f"❌ md文件不存在: {MD_FILE}")
        sys.exit(1)

    oda_exe = find_oda_converter()
    dwg2dxf_exe = find_dwg2dxf()

    if oda_exe:
        print(f"✅ 使用 ODA File Converter: {oda_exe}")
    elif dwg2dxf_exe:
        print(f"✅ 使用 LibreDWG dwg2dxf: {dwg2dxf_exe}")
    else:
        print("⚠️  未找到 DWG→DXF 转换工具")
        print("   安装方式（二选一）：")
        print("   1. ODA File Converter: https://www.opendesign.com/guestfiles/oda_file_converter")
        print("   2. LibreDWG: brew install libredwg")
        print("   将直接复制 .dwg 原文件，不做转换。\n")

    source_folders = read_source_folders(MD_FILE)
    if not source_folders:
        print("❌ md文件中未找到有效路径")
        sys.exit(1)

    total_copied = total_converted = total_failed = 0
    all_failed: list[str] = []
    output_folders: list[Path] = []

    for folder in source_folders:
        if not folder.exists() or not folder.is_dir():
            print(f"⚠️  目录不存在，跳过: {folder}")
            continue

        dst = make_dxf_output_path(folder)
        print(f"\n{'='*60}")
        print(f"处理目录: {folder}")
        print(f"输出到:   {dst}")
        print(f"{'='*60}")

        copied, converted, failed, failed_names, dst_root = process_folder(
            folder, oda_exe, dwg2dxf_exe
        )
        total_copied += copied
        total_converted += converted
        total_failed += failed
        all_failed.extend(failed_names)
        output_folders.append(dst_root)

    write_dxf_index_md(DXF_INDEX_DIR, DXF_INDEX_MD, output_folders)

    print(f"\n{'='*60}")
    print("完成！")
    print(f"  普通文件复制:  {total_copied}")
    print(f"  DWG→DXF 成功: {total_converted}")
    print(f"  DWG 转换失败: {total_failed}")
    if all_failed:
        print("  失败文件列表:")
        for name in all_failed:
            print(f"    - {name}")


if __name__ == "__main__":
    main()
