#!/usr/bin/env python3
"""Generate uso-regimen/persona rows derived from c_UsoCFDI.

Reads c_UsoCFDI.csv from the latest version folder under hf/csv/anexo20/factura/
and writes the derived table to hf/csv/anexo20/factura/{version}/c_UsoCFDI_Regimen.csv.

Also upserts the entry into catalog_state.csv so generate_hf.py includes it
in the HF dataset index automatically.

Usage:
  uv run scripts/catalogos/generate_uso_regimen_table.py
  uv run scripts/catalogos/generate_uso_regimen_table.py --input hf/csv/anexo20/factura/4-0/c_UsoCFDI.csv
"""
from __future__ import annotations

import argparse
import csv
import hashlib
from pathlib import Path
from typing import Iterable

HF_CSV_DIR     = Path("hf/csv")
_FACTURA_DIR   = HF_CSV_DIR / "anexo20" / "factura"
CATALOG_STATE  = Path("catalog_state.csv")

_DESCRIPCION = (
    "Tabla derivada: combinaciones válidas de uso de CFDI "
    "por tipo de persona (física/moral) y régimen fiscal receptor."
)


def _latest_version_dir(base: Path) -> Path | None:
    """Return the subdirectory with the highest version number under base."""
    dirs = [d for d in base.iterdir() if d.is_dir()] if base.exists() else []
    if not dirs:
        return None
    def _ver_key(d: Path) -> tuple:
        try:
            return tuple(int(x) for x in d.name.split("-"))
        except ValueError:
            return (0,)
    return max(dirs, key=_ver_key)


def _default_input() -> Path:
    d = _latest_version_dir(_FACTURA_DIR)
    return (d / "c_UsoCFDI.csv") if d else Path("hf/csv/anexo20/factura/c_UsoCFDI.csv")


def _normalize_bool(value: str | None) -> bool:
    if value is None:
        return False
    normalized = value.strip().lower()
    return normalized in {"si", "sí", "true", "1", "x", "s"}


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
    """Insert or replace the entry for (section, folder_version, catalogo) in catalog_state.csv."""
    fieldnames, rows = _load_state(state_file)

    key = (new_entry.get("section", ""), new_entry.get("folder_version", ""), new_entry.get("catalogo", ""))
    replaced = False
    for i, row in enumerate(rows):
        rkey = (row.get("section", ""), row.get("folder_version", ""), row.get("catalogo", ""))
        if rkey == key:
            rows[i] = new_entry
            replaced = True
            break
    if not replaced:
        rows.append(new_entry)

    # Merge fieldnames: preserve existing order, append any new keys
    for k in new_entry:
        if k not in fieldnames:
            fieldnames.append(k)

    state_file.parent.mkdir(parents=True, exist_ok=True)
    with state_file.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)

    action = "updated" if replaced else "added"
    print(f"State   → {state_file}  ({action} c_UsoCFDI_Regimen)")


def main() -> int:
    default_input = _default_input()

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        default=default_input,
        help=f"Path to c_UsoCFDI CSV (default: {default_input})",
    )
    parser.add_argument(
        "--state-file",
        type=Path,
        default=CATALOG_STATE,
        help=f"Path to catalog state CSV (default: {CATALOG_STATE})",
    )
    args = parser.parse_args()

    if not args.input.exists():
        print(f"ERROR: {args.input} not found. Run scripts/catalogos/extract.py first.")
        return 1

    # Derive version and section from the input path
    try:
        rel = args.input.parent.relative_to(HF_CSV_DIR)
        parts = rel.parts
        folder_version = parts[-1] if parts else ""
        section = "/".join(parts[:-1]) if len(parts) > 1 else str(rel)
    except ValueError:
        folder_version = ""
        section = ""

    # Compute output path: same directory as input
    output = args.input.parent / "c_UsoCFDI_Regimen.csv"

    rows = list(generate_rows(args.input))
    headers = ["uso_clave", "tipo_persona", "regimen_fiscal"]

    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        writer.writerows(rows)
    print(f"Wrote {output} ({len(rows)} rows)")

    file_hash = hashlib.sha256(output.read_bytes()).hexdigest()
    entry: dict = {
        "section":          section,
        "folder_version":   folder_version,
        "catalogo":         "c_UsoCFDI_Regimen",
        "source_xls":       "",
        "xls_hash":         "",
        "descripcion":      _DESCRIPCION,
        "sheets":           "",
        "numero_columnas":  str(len(headers)),
        "nombres_columnas": "|".join(headers),
        "file_hash":        file_hash,
        "version":          "",
        "revision":         "",
        "fecha_publicacion": "",
        "fecha_inicio_vigencia": "",
        "fecha_fin_vigencia": "",
    }
    _upsert_state(args.state_file, entry)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
