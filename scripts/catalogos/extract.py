#!/usr/bin/env python3
"""Extract SAT CFDI catalog XLS files into per-catalog CSV files.

Reads output/catalog.csv to find the latest XLS catalog files scraped by
scrape.py (those stored under hf/xls/), then extracts every sheet whose name
starts with "c_" into hf/csv/ mirroring the same directory structure.

Example outputs:
    hf/xls/anexo20/catCFDI40.xls        → hf/csv/anexo20/c_uso_cfdi.csv …
    hf/xls/complementos/carta-porte/x.xls → hf/csv/complementos/carta-porte/c_estaciones.csv …

Usage:
    uv run scripts/catalogos/extract.py
    uv run scripts/catalogos/extract.py --catalog output/catalog.csv --header-style snake
    uv run scripts/catalogos/extract.py --xls-dir hf/xls --csv-dir hf/csv
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import re
import sys
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Sequence

import xlrd

# ── constants ──────────────────────────────────────────────────────────────────

HF_XLS_DIR    = Path("hf/xls")
HF_CSV_DIR    = Path("hf/csv")
CATALOG_STATE = Path("catalog_state.csv")  # committed to git, outside hf/

CATALOG_HEADER_PATTERN = re.compile(r"^[cC]_\w+")
PART_SUFFIX_PATTERN    = re.compile(r"(?:_Parte)?_\d+$", re.IGNORECASE)


# ── data types ─────────────────────────────────────────────────────────────────


@dataclass
class SheetExtraction:
    headers: List[str]
    rows: List[List[str]]
    metadata: Dict[str, str]
    description: str


# ── string helpers ─────────────────────────────────────────────────────────────


def normalize_token(value: str) -> str:
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip().lower()


SPECIAL_SNAKE_HEADERS = {
    "fechainiciovigencia": "fecha_de_inicio_de_vigencia",
    "fechafinvigencia": "fecha_de_fin_de_vigencia",
}


def slugify(value: str) -> str:
    token = normalize_token(value)
    token = re.sub(r"[^0-9a-z]+", "_", token)
    return token.strip("_")


def strip_catalog_prefix(value: str) -> str:
    return re.sub(r"^(?:c|C)_", "", value.strip())


def split_compound_words(token: str) -> str:
    step1 = re.sub(r"([a-z])([A-Z])", r"\1 \2", token)
    step2 = re.sub(r"([A-Za-z])([0-9])", r"\1 \2", step1)
    step3 = re.sub(r"([0-9])([A-Za-z])", r"\1 \2", step2)
    return step3


def alnum_token(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^0-9a-z]+", "", normalized.lower())


def to_snake_ascii(value: str) -> str:
    normalized_key = alnum_token(value)
    if normalized_key in SPECIAL_SNAKE_HEADERS:
        return SPECIAL_SNAKE_HEADERS[normalized_key]
    ascii_value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    spaced = split_compound_words(ascii_value)
    ascii_value = re.sub(r"[^0-9A-Za-z]+", "_", spaced)
    ascii_value = re.sub(r"_+", "_", ascii_value)
    return ascii_value.strip("_").lower()


def first_non_empty(strings: Sequence[str]) -> str:
    for item in strings:
        if item:
            return item
    return ""


# ── XLS parsing ────────────────────────────────────────────────────────────────


def sheet_is_empty(row_values: Iterable[str]) -> bool:
    return not any(cell for cell in row_values)


def is_continuation_row(row_values: Sequence[str]) -> bool:
    first = first_non_empty(row_values)
    if not first:
        return False
    return normalize_token(first).startswith("continua en")


def format_number(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return ("%s" % value).rstrip("0").rstrip(".")


def format_cell(
    workbook: xlrd.book.Book, sheet: xlrd.sheet.Sheet, row_idx: int, col_idx: int
) -> str:
    cell = sheet.cell(row_idx, col_idx)
    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return ""
    if cell.ctype == xlrd.XL_CELL_DATE:
        dt = xlrd.xldate_as_datetime(cell.value, workbook.datemode)
        if dt.time() == datetime.min.time():
            return dt.date().isoformat()
        return dt.isoformat()
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        xf = workbook.xf_list[cell.xf_index]
        fmt = workbook.format_map.get(xf.format_key)
        fmt_str = fmt.format_str.split(";")[0] if fmt else ""
        fmt_str = fmt_str.strip()
        if fmt_str and re.fullmatch(r"0+", fmt_str):
            width = len(fmt_str)
            return f"{int(round(cell.value)):0{width}d}"
        return format_number(cell.value)
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return "TRUE" if cell.value else "FALSE"
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return "#ERROR"
    return str(cell.value).strip()


def detect_header_row(sheet: xlrd.sheet.Sheet) -> int:
    """Find the header row index.

    Primary strategy: look for a row where any cell starts with 'c_' (Anexo 20 style).
    Fallback: find the last empty-row separator in the first 10 rows, then return
    the next non-empty row after it (Carta Porte / complement-style sheets).
    """
    for row_idx in range(sheet.nrows):
        row_cells = [str(cell.value).strip() for cell in sheet.row(row_idx)]
        if sheet_is_empty(row_cells):
            continue
        if any(CATALOG_HEADER_PATTERN.match(value) for value in row_cells if value):
            return row_idx

    # Fallback: header follows the last empty separator in the metadata block
    last_empty = None
    for row_idx in range(min(10, sheet.nrows)):
        row_cells = [str(cell.value).strip() for cell in sheet.row(row_idx)]
        if sheet_is_empty(row_cells):
            last_empty = row_idx
    if last_empty is not None:
        for row_idx in range(last_empty + 1, sheet.nrows):
            row_cells = [str(cell.value).strip() for cell in sheet.row(row_idx)]
            if not sheet_is_empty(row_cells):
                return row_idx

    raise ValueError(f"Header row not found in sheet {sheet.name!r}")


def gather_header_rows(sheet: xlrd.sheet.Sheet, start_idx: int) -> List[int]:
    rows = [start_idx]
    row_idx = start_idx + 1
    while row_idx < sheet.nrows:
        row_values = [str(cell.value).strip() for cell in sheet.row(row_idx)]
        if sheet_is_empty(row_values):
            row_idx += 1
            continue
        first_cell = str(sheet.cell(row_idx, 0).value).strip() if sheet.ncols else ""
        if not first_cell:
            rows.append(row_idx)
            row_idx += 1
            continue
        break
    return rows


def combine_headers(
    workbook: xlrd.book.Book, sheet: xlrd.sheet.Sheet, header_rows: List[int]
) -> List[tuple[int, str]]:
    width = max(sheet.row_len(idx) for idx in header_rows)
    headers: List[tuple[int, str]] = []
    for col_idx in range(width):
        parts: List[str] = []
        for row_idx in header_rows:
            if col_idx >= sheet.row_len(row_idx):
                continue
            value = format_cell(workbook, sheet, row_idx, col_idx)
            if value:
                parts.append(value)
        if not parts:
            continue
        header_value = " - ".join(parts).strip() or f"column_{col_idx + 1}"
        headers.append((col_idx, header_value))
    return headers


def find_description(sheet: xlrd.sheet.Sheet, workbook: xlrd.book.Book) -> str:
    for row_idx in range(min(6, sheet.nrows)):
        values = [
            format_cell(workbook, sheet, row_idx, col_idx)
            for col_idx in range(sheet.row_len(row_idx))
        ]
        cleaned = [val for val in values if val]
        if cleaned and any("catalogo" in normalize_token(val) for val in cleaned):
            return " | ".join(cleaned)
    return ""


_METADATA_KEY_ALIASES: Dict[str, str] = {
    # version
    "version_catalogo":                       "version",
    "version_cfdi":                           "version",
    # revision
    "revision_catalogo":                      "revision",
    # fecha publicacion
    "fecha_publicacion_de_catalogo":          "fecha_publicacion",
    "fecha_de_publicacion":                   "fecha_publicacion",
    # fecha inicio vigencia
    "fecha_inicio_de_vigencia_del_catalogo":  "fecha_inicio_vigencia",
    "fecha_inicio_de_vigencia":               "fecha_inicio_vigencia",
    "fecha_de_inicio_de_vigencia":            "fecha_inicio_vigencia",
    "fechainiciodevigencia":                  "fecha_inicio_vigencia",
    "fechainiciovigencia":                    "fecha_inicio_vigencia",
    # fecha fin vigencia
    "fecha_fin_de_vigencia_del_catalogo":     "fecha_fin_vigencia",
    "fecha_fin_de_vigencia":                  "fecha_fin_vigencia",
    "fecha_de_fin_de_vigencia":               "fecha_fin_vigencia",
    "fechafindevigencia":                     "fecha_fin_vigencia",
    "fechafinvigencia":                       "fecha_fin_vigencia",
}


def extract_metadata(sheet: xlrd.sheet.Sheet, workbook: xlrd.book.Book) -> Dict[str, str]:
    key_row_idx = None
    for row_idx in range(min(10, sheet.nrows)):
        for col_idx in range(sheet.row_len(row_idx)):
            if "version" in normalize_token(str(sheet.cell(row_idx, col_idx).value)):
                key_row_idx = row_idx
                break
        if key_row_idx is not None:
            break
    if key_row_idx is None:
        return {}
    value_row_idx = key_row_idx + 1
    while value_row_idx < sheet.nrows and sheet.row_len(value_row_idx) == 0:
        value_row_idx += 1
    if value_row_idx >= sheet.nrows:
        return {}
    metadata: Dict[str, str] = {}
    for col_idx in range(sheet.row_len(key_row_idx)):
        key_raw = str(sheet.cell(key_row_idx, col_idx).value).strip()
        if not key_raw:
            continue
        key = slugify(key_raw)
        if not key:
            continue
        key = _METADATA_KEY_ALIASES.get(key, key)
        value = format_cell(workbook, sheet, value_row_idx, col_idx)
        if "fecha" in key and value == "0":
            value = ""
        metadata[key] = value
    return metadata


def extract_data_rows(
    workbook: xlrd.book.Book,
    sheet: xlrd.sheet.Sheet,
    header_rows: List[int],
    header_columns: List[tuple[int, str]],
) -> List[List[str]]:
    first_data_row = max(header_rows) + 1
    column_indices = [col_idx for col_idx, _ in header_columns]
    rows: List[List[str]] = []
    row_idx = first_data_row
    while row_idx < sheet.nrows:
        row_values = [format_cell(workbook, sheet, row_idx, col_idx) for col_idx in column_indices]
        if sheet_is_empty(row_values):
            row_idx += 1
            continue
        if is_continuation_row(row_values):
            row_idx += 1
            continue
        rows.append(row_values)
        row_idx += 1
    return _merge_orphan_rows(rows)


def _merge_orphan_rows(rows: List[List[str]]) -> List[List[str]]:
    """Merge rows with an empty first column into the previous row.

    When the XLS has a multi-value cell (e.g. N/Nómina with valor_maximo=NS on
    one row and the numeric value on the next), the continuation row has an empty
    clave. We merge it back by replacing any cell in the previous row that is
    empty or equals 'NS' with the non-empty value from the continuation row.
    """
    merged: List[List[str]] = []
    for row in rows:
        if row and not row[0] and merged:
            prev = merged[-1]
            for i, val in enumerate(row):
                if i < len(prev) and val and (not prev[i] or prev[i] == "NS"):
                    prev[i] = val
        else:
            merged.append(row)
    return merged


def parse_sheet(workbook: xlrd.book.Book, sheet: xlrd.sheet.Sheet) -> SheetExtraction:
    header_start = detect_header_row(sheet)
    header_rows = gather_header_rows(sheet, header_start)
    header_columns = combine_headers(workbook, sheet, header_rows)
    headers = [header for _, header in header_columns]
    rows = extract_data_rows(workbook, sheet, header_rows, header_columns)
    metadata = extract_metadata(sheet, workbook)
    description = find_description(sheet, workbook)
    return SheetExtraction(headers=headers, rows=rows, metadata=metadata, description=description)


def base_catalog_name(sheet_name: str) -> str:
    name = PART_SUFFIX_PATTERN.sub("", sheet_name)
    if name.startswith("C_"):
        name = "c_" + name[2:]
    return name


def transform_headers_for_output(headers: Sequence[str], catalog: str, style: str) -> List[str]:
    if style == "original":
        return list(headers)
    if style == "snake":
        catalog_snake = to_snake_ascii(catalog)
        catalog_snake_stripped = to_snake_ascii(strip_catalog_prefix(catalog))
        transformed: List[str] = []
        for idx, header in enumerate(headers, start=1):
            snake = to_snake_ascii(strip_catalog_prefix(header)) or f"column_{idx}"
            if snake in {catalog_snake, catalog_snake_stripped}:
                snake = "clave"
            transformed.append(snake)
        return transformed
    raise ValueError(f"Unsupported header style: {style}")


def write_csv(path: Path, headers: Sequence[str], rows: Sequence[Sequence[str]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        writer.writerows(rows)


# ── per-workbook extraction ────────────────────────────────────────────────────


def extract_workbook(xls_path: Path, output_dir: Path, header_style: str) -> tuple[List[str], List[Dict]]:
    """Extract all c_* sheets from xls_path into output_dir.

    Returns (list of written csv paths, list of metadata dicts).
    Metadata is not written here — caller accumulates across workbooks and writes once per dir.
    """
    workbook = xlrd.open_workbook(str(xls_path), formatting_info=True)
    catalogs: Dict[str, Dict] = defaultdict(
        lambda: {"headers": None, "rows": [], "metadata": {}, "description": "", "sheets": []}
    )
    for sheet_name in workbook.sheet_names():
        if not sheet_name.lower().startswith("c_"):
            continue
        sheet = workbook.sheet_by_name(sheet_name)
        try:
            parsed = parse_sheet(workbook, sheet)
        except ValueError as exc:
            print(f"  SKIP {sheet_name!r}: {exc}", file=sys.stderr)
            continue
        base_name = base_catalog_name(sheet_name)
        record = catalogs[base_name]
        if record["headers"] is None:
            record["headers"] = parsed.headers
        elif record["headers"] != parsed.headers:
            print(f"  WARN header mismatch for {base_name!r} — skipping extra sheet", file=sys.stderr)
            continue
        record["rows"].extend(parsed.rows)
        if not record["metadata"] and parsed.metadata:
            record["metadata"] = parsed.metadata
        if not record["description"] and parsed.description:
            record["description"] = parsed.description
        record["sheets"].append(sheet_name)

    output_dir.mkdir(parents=True, exist_ok=True)
    written: List[str] = []
    metadata_rows: List[Dict[str, str]] = []

    for catalog, payload in catalogs.items():
        if payload["headers"] is None:
            continue
        out_headers = transform_headers_for_output(payload["headers"], catalog, header_style) + ["row_hash"]
        dest = output_dir / f"{catalog}.csv"
        hashed_rows = [
            row + [hashlib.sha256("|".join(row).encode()).hexdigest()]
            for row in payload["rows"]
        ]
        write_csv(dest, out_headers, hashed_rows)
        print(f"  → {dest}  ({len(payload['rows'])} rows)", file=sys.stderr)
        written.append(str(dest))

        entry: Dict[str, str] = {
            "catalogo": catalog,
            "source_xls": str(xls_path),
            "xls_hash": hashlib.sha256(xls_path.read_bytes()).hexdigest(),
            "descripcion": payload["description"],
            "sheets": ",".join(payload["sheets"]),
            "numero_columnas": str(len(out_headers)),
            "nombres_columnas": "|".join(out_headers),
            "file_hash": hashlib.sha256(dest.read_bytes()).hexdigest(),
        }
        entry.update(payload["metadata"])
        metadata_rows.append(entry)

    return written, metadata_rows


# ── discovery ──────────────────────────────────────────────────────────────────


def discover_xls(xls_dir: Path) -> List[Path]:
    """Return all XLS paths under xls_dir, skipping ignored files."""
    def _skip(p: Path) -> bool:
        return "matriz" in p.name.lower()

    xls_exts = {".xls", ".xlsx"}
    return sorted(p for p in xls_dir.rglob("*") if p.suffix.lower() in xls_exts and not _skip(p))


# ── CLI ────────────────────────────────────────────────────────────────────────


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--xls-dir",
        type=Path,
        default=HF_XLS_DIR,
        help=f"Root directory of downloaded XLS files (default: {HF_XLS_DIR})",
    )
    parser.add_argument(
        "--csv-dir",
        type=Path,
        default=HF_CSV_DIR,
        help=f"Root output directory for extracted CSVs (default: {HF_CSV_DIR})",
    )
    parser.add_argument(
        "--state-file",
        type=Path,
        default=CATALOG_STATE,
        help=f"Path to catalog state CSV (default: {CATALOG_STATE})",
    )
    parser.add_argument(
        "--header-style",
        choices=["original", "snake"],
        default="snake",
        help="Column name format: 'original' keeps SAT names, 'snake' converts to snake_case (default: snake)",
    )
    parser.add_argument("--force", action="store_true", help="Re-extract all XLS even if unchanged")
    parser.add_argument(
        "--sections", nargs="+", metavar="SECTION",
        help=(
            "Only extract XLS files under these section paths. "
            "Examples: anexo20/factura  anexo20/retenciones  complementos"
        ),
    )
    return parser.parse_args(argv)


def _load_catalog_state(state_file: Path) -> dict[str, dict]:
    """Load catalog_state.csv into a dict keyed by (source_xls, catalogo)."""
    state: dict[str, dict] = {}
    if not state_file.exists():
        return state
    with state_file.open(newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            key = (row.get("source_xls", ""), row.get("catalogo", ""))
            state[key] = dict(row)
    return state


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])

    xls_files = discover_xls(args.xls_dir)
    if args.sections:
        sections_filter = [s.strip("/") for s in args.sections]
        xls_files = [
            p for p in xls_files
            if any(
                str(p.relative_to(args.xls_dir)).startswith(s + "/") or
                str(p.relative_to(args.xls_dir)).startswith(s + "\\")
                for s in sections_filter
            )
        ]
        print(f"Filtering to sections: {sections_filter}", file=sys.stderr)
    if not xls_files:
        print("No XLS files found — skipping extract.", file=sys.stderr)
        return 0

    # Load existing catalog_state.csv for skip-if-unchanged and merge
    existing_state = _load_catalog_state(args.state_file)
    stored_xls_hashes: dict[str, str] = {}
    for (src, _cat), row in existing_state.items():
        h = row.get("xls_hash", "")
        if src and h and src not in stored_xls_hashes:
            stored_xls_hashes[src] = h

    print(f"Found {len(xls_files)} XLS file(s):", file=sys.stderr)
    total_written = 0
    # Accumulate new metadata rows keyed by (source_xls, catalogo)
    new_state: dict[tuple, dict] = {}
    # Accumulate per output dir so multiple XLS in the same folder merge correctly
    metadata_by_dir: dict[Path, List[Dict]] = defaultdict(list)

    for xls_path in xls_files:
        if not xls_path.exists():
            print(f"\n[{xls_path}] not downloaded (version unchanged) — skipping", file=sys.stderr)
            continue
        current_hash = hashlib.sha256(xls_path.read_bytes()).hexdigest()
        if not args.force and stored_xls_hashes.get(str(xls_path)) == current_hash:
            print(f"\n[{xls_path}] unchanged — skipping", file=sys.stderr)
            continue
        print(f"\n[{xls_path}]", file=sys.stderr)
        try:
            rel_parent = xls_path.parent.relative_to(args.xls_dir)
        except ValueError:
            rel_parent = xls_path.parent
        output_dir = args.csv_dir / rel_parent
        written, meta_rows = extract_workbook(xls_path, output_dir, args.header_style)
        total_written += len(written)
        metadata_by_dir[output_dir].extend(meta_rows)

    # Build new index rows from freshly processed files, adding section + version from path
    for output_dir, meta_rows in metadata_by_dir.items():
        try:
            rel = output_dir.relative_to(args.csv_dir)
        except ValueError:
            rel = output_dir
        parts = rel.parts
        folder_version = parts[-1] if parts else ""
        section = "/".join(parts[:-1]) if len(parts) > 1 else str(rel)
        for row in meta_rows:
            entry = {"section": section, "folder_version": folder_version, **row}
            key = (entry.get("source_xls", ""), entry.get("catalogo", ""))
            new_state[key] = entry

    # Merge: start from existing state, overwrite with any freshly processed entries
    merged: dict[tuple, dict] = {**existing_state, **new_state}

    # Determine column order (union of all keys, stable)
    index_keys: List[str] = ["section", "folder_version"]
    for row in merged.values():
        for k in row:
            if k not in index_keys:
                index_keys.append(k)

    merged_rows = list(merged.values())
    if merged_rows:
        args.state_file.parent.mkdir(parents=True, exist_ok=True)
        write_csv(args.state_file, index_keys, [[r.get(k, "") for k in index_keys] for r in merged_rows])
        updated = len(new_state)
        preserved = len(merged_rows) - updated
        print(
            f"\nState   → {args.state_file}  ({len(merged_rows)} catalogs, {updated} updated, {preserved} preserved)",
            file=sys.stderr,
        )

    print(f"Done. Wrote {total_written} CSV file(s) to {args.csv_dir}/", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
