#!/usr/bin/env python3
"""Import UNSPSC CSVs into the HF extra catalog index.

Copies UNSPSC CSVs from `static/unspsc/` into `hf/csv/extra/unspsc/`
and upserts entries into `catalog_state.csv` so they are included in
the HF dataset as `extra/unspsc`.

Run after the UNSPSC CSVs have been refreshed in `static/unspsc/`.

Usage:
  uv run scripts/catalogos/import_unspsc.py
  uv run scripts/catalogos/import_unspsc.py --source static/unspsc
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import shutil
from pathlib import Path

CATALOG_STATE = Path("catalog_state.csv")
_SECTION      = "extra"
_FOLDER       = "unspsc"
_DEST_DIR     = Path("hf/csv/extra/unspsc")
_DEFAULT_SRC  = Path("static/unspsc")

# Human-readable descriptions per catalog stem
_DESCRIPTIONS: dict[str, str] = {
    "clases":    "Jerarquía SAT PyS: clases de productos y servicios (tipo/división/grupo/clase).",
    "divisiones":"Jerarquía SAT PyS: divisiones de productos y servicios.",
    "grupos":    "Jerarquía SAT PyS: grupos de productos y servicios.",
    "tipos":     "Jerarquía SAT PyS: tipos (Producto / Servicio).",
}

# CSVs to skip (not catalog data)
_SKIP = {"metadata"}


def _ver_key(v: str) -> tuple:
    parts = []
    for p in v.split("-"):
        try:
            parts.append(int(p))
        except ValueError:
            parts.append(p)
    return tuple(parts)


def _load_state(state_file: Path) -> tuple[list[str], list[dict]]:
    if not state_file.exists():
        return [], []
    with state_file.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = list(reader.fieldnames or [])
        rows = [dict(r) for r in reader]
    return fieldnames, rows


def _upsert_state(state_file: Path, new_entry: dict) -> None:
    fieldnames, rows = _load_state(state_file)

    section  = new_entry["section"]
    catalogo = new_entry["catalogo"]
    replaced = False
    new_rows: list[dict] = []
    for row in rows:
        if row.get("section") == section and row.get("catalogo") == catalogo:
            if not replaced:
                new_rows.append(new_entry)
                replaced = True
        else:
            new_rows.append(row)
    if not replaced:
        new_rows.append(new_entry)

    for k in new_entry:
        if k not in fieldnames:
            fieldnames.append(k)

    new_rows.sort(key=lambda r: (
        r.get("section", ""),
        _ver_key(r.get("folder_version", "")),
        r.get("catalogo", ""),
    ))

    state_file.parent.mkdir(parents=True, exist_ok=True)
    with state_file.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(new_rows)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--source", type=Path, default=_DEFAULT_SRC,
        help=f"Directory containing sat-pys-hierarchy CSVs (default: {_DEFAULT_SRC})",
    )
    parser.add_argument(
        "--state-file", type=Path, default=CATALOG_STATE,
        help=f"Path to catalog state CSV (default: {CATALOG_STATE})",
    )
    args = parser.parse_args()

    if not args.source.exists():
        print(f"Source directory not found: {args.source} — skipping.")
        return 0

    _DEST_DIR.mkdir(parents=True, exist_ok=True)

    for src_csv in sorted(args.source.glob("*.csv")):
        stem = src_csv.stem
        if stem in _SKIP:
            continue

        dest = _DEST_DIR / src_csv.name
        shutil.copy2(src_csv, dest)

        # Read header row for column metadata
        with dest.open(newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            headers = next(reader, [])
        num_cols = len(headers)
        col_names = "|".join(headers)
        file_hash = hashlib.sha256(dest.read_bytes()).hexdigest()

        entry = {
            "section":               _SECTION,
            "folder_version":        _FOLDER,
            "catalogo":              stem,
            "source_xls":            "",
            "xls_hash":              "",
            "descripcion":           _DESCRIPTIONS.get(stem, ""),
            "sheets":                "",
            "numero_columnas":       str(num_cols),
            "nombres_columnas":      col_names,
            "file_hash":             file_hash,
            "version":               "",
            "revision":              "",
            "fecha_publicacion":     "",
            "fecha_inicio_vigencia": "",
            "fecha_fin_vigencia":    "",
        }
        _upsert_state(args.state_file, entry)
        print(f"  {stem}  →  {dest}  ({num_cols} cols)")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
