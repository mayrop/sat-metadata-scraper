#!/usr/bin/env python3
"""Generate uso-regimen/persona rows derived from c_UsoCFDI.

Reads c_UsoCFDI.csv from the latest version folder under hf/csv/anexo20/cfdi/
and writes the derived table to hf/csv/anexo20/cfdi/{version}/c_UsoCFDI_Regimen.csv.

Also upserts the entry into catalog_state.csv so generate_hf.py includes it
in the HF dataset index automatically.

The version is derived from catalog_state.csv (authoritative) so that the
catalog_state.csv entry always reflects the correct version even when
c_UsoCFDI.csv is not locally available (e.g. extract.py skipped it because
the XLS was unchanged). The CSV is only re-generated when the input file exists.

Usage:
  uv run scripts/catalogos/generate_uso_regimen_table.py
  uv run scripts/catalogos/generate_uso_regimen_table.py --input hf/csv/anexo20/cfdi/4-0/c_UsoCFDI.csv
"""
from __future__ import annotations

import argparse
import csv
import hashlib
from pathlib import Path
from typing import Iterable

HF_CSV_DIR     = Path("hf/csv")
_FACTURA_DIR   = HF_CSV_DIR / "anexo20" / "cfdi"
CATALOG_STATE  = Path("catalog_state.csv")

_SECTION   = "anexo20/cfdi"
_CATALOGO  = "c_UsoCFDI_Regimen"
_SOURCE    = "c_UsoCFDI"

_DESCRIPCION = (
    "Tabla derivada: combinaciones válidas de uso de CFDI "
    "por tipo de persona (física/moral) y régimen fiscal receptor."
)


def _ver_key(v: str) -> tuple:
    try:
        return tuple(int(x) for x in v.split("-"))
    except ValueError:
        return (0,)


def _latest_version_from_state(state_file: Path) -> str | None:
    """Return the highest folder_version for c_UsoCFDI in section=anexo20/cfdi."""
    if not state_file.exists():
        return None
    with state_file.open(newline="", encoding="utf-8") as f:
        versions = [
            row.get("folder_version", "")
            for row in csv.DictReader(f)
            if row.get("section") == _SECTION and row.get("catalogo") == _SOURCE
        ]
    return max(versions, key=_ver_key) if versions else None


def _latest_version_from_fs() -> str | None:
    """Return the highest version directory name under _FACTURA_DIR."""
    dirs = [d for d in _FACTURA_DIR.iterdir() if d.is_dir()] if _FACTURA_DIR.exists() else []
    return max((d.name for d in dirs), key=_ver_key) if dirs else None


def _normalize_bool(value: str | None) -> bool:
    if value is None:
        return False
    return value.strip().lower() in {"si", "sí", "true", "1", "x", "s"}


def _split_regimens(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in value.split(",") if item.strip()]


def generate_rows(input_csv: Path) -> Iterable[tuple[str, str, str]]:
    with input_csv.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            clave = row.get("clave") or row.get("codigo")
            if not clave:
                continue
            regimens = _split_regimens(row.get("regimen_fiscal_receptor"))
            if not regimens:
                continue
            if _normalize_bool(
                row.get("aplica_fisica") or row.get("aplica_para_tipo_persona_fisica")
            ):
                for regimen in regimens:
                    yield (clave.strip(), "fisica", regimen)
            if _normalize_bool(row.get("aplica_moral") or row.get("moral")):
                for regimen in regimens:
                    yield (clave.strip(), "moral", regimen)


def _load_state(state_file: Path) -> tuple[list[str], list[dict]]:
    """Return (fieldnames, rows) from catalog_state.csv."""
    if not state_file.exists():
        return [], []
    with state_file.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = list(reader.fieldnames or [])
        rows = [dict(r) for r in reader]
    return fieldnames, rows


def _upsert_state(state_file: Path, new_entry: dict) -> None:
    """Insert or replace the entry for (section, catalogo) in catalog_state.csv.

    Matches on (section, catalogo) only so that a version change (e.g. 3-3 → 4-0)
    replaces the old row rather than leaving a stale entry alongside the new one.
    """
    fieldnames, rows = _load_state(state_file)

    section  = new_entry.get("section", "")
    catalogo = new_entry.get("catalogo", "")
    replaced = False
    new_rows = []
    for row in rows:
        if row.get("section") == section and row.get("catalogo") == catalogo:
            if not replaced:
                new_rows.append(new_entry)
                replaced = True
            # drop any additional stale entries for the same (section, catalogo)
        else:
            new_rows.append(row)
    if not replaced:
        new_rows.append(new_entry)

    for k in new_entry:
        if k not in fieldnames:
            fieldnames.append(k)

    new_rows.sort(key=lambda r: (r.get("section", ""), _ver_key(r.get("folder_version", "")), r.get("catalogo", "")))
    state_file.parent.mkdir(parents=True, exist_ok=True)
    with state_file.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(new_rows)

    action = "updated" if replaced else "added"
    print(f"State   → {state_file}  ({action} {_CATALOGO})")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help="Path to c_UsoCFDI CSV (default: auto-detected from catalog_state.csv)",
    )
    parser.add_argument(
        "--state-file",
        type=Path,
        default=CATALOG_STATE,
        help=f"Path to catalog state CSV (default: {CATALOG_STATE})",
    )
    args = parser.parse_args()

    # Determine the authoritative version: prefer catalog_state.csv, fall back to filesystem.
    version = _latest_version_from_state(args.state_file) or _latest_version_from_fs()
    if not version:
        print(f"Skipping: no {_SOURCE} version found in {args.state_file} or {_FACTURA_DIR}.")
        return 0

    input_path = args.input or (_FACTURA_DIR / version / f"{_SOURCE}.csv")

    headers = ["uso_clave", "tipo_persona", "regimen_fiscal"]
    file_hash = ""
    if input_path.exists():
        output = input_path.parent / f"{_CATALOGO}.csv"
        rows = list(generate_rows(input_path))
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            writer.writerow(headers)
            writer.writerows(rows)
        print(f"Wrote {output} ({len(rows)} rows)")
        file_hash = hashlib.sha256(output.read_bytes()).hexdigest()
    else:
        print(f"Skipping CSV generation: {input_path} not available locally.")
        # Preserve any existing file_hash so we don't lose it on runs where
        # the source CSV is not locally available.
        _, existing_rows = _load_state(args.state_file)
        for row in existing_rows:
            if row.get("section") == _SECTION and row.get("catalogo") == _CATALOGO:
                file_hash = row.get("file_hash", "") or ""
                break

    entry: dict = {
        "section":           _SECTION,
        "folder_version":    version,
        "catalogo":          _CATALOGO,
        "source_xls":        "",
        "xls_hash":          "",
        "descripcion":       _DESCRIPCION,
        "sheets":            "",
        "numero_columnas":   str(len(headers)),
        "nombres_columnas":  "|".join(headers),
        "file_hash":         file_hash,
        "version":           "",
        "revision":          "",
        "fecha_publicacion": "",
        "fecha_inicio_vigencia": "",
        "fecha_fin_vigencia": "",
    }
    _upsert_state(args.state_file, entry)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
