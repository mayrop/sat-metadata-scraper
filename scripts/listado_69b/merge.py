#!/usr/bin/env python3
"""Clean, transform, and merge SAT 69-B and 69-B Bis listado CSV files.

Reads the raw CSV files downloaded by scrape_listado_69b.py, strips the
SAT disclaimer/title preamble rows, re-encodes from Latin-1 to UTF-8,
transforms columns (oficio splitting, date normalization, situacion enum),
and writes clean individual files plus a merged file.

Output:
  hf/csv/listado-69b/69b.csv          — clean 69-B
  hf/csv/listado-69b/69b-bis.csv      — clean 69-B Bis
  hf/csv/listado-69b/listado-69b.csv  — merged (69-B + 69-B Bis)

Usage:
  uv run scripts/listado_69b/merge.py
  uv run scripts/listado_69b/merge.py --manifest output/listado-69b-manifest.json
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import unicodedata
from pathlib import Path

# ── constants ──────────────────────────────────────────────────────────────────

MANIFEST      = Path("output/listado-69b-manifest.json")
HF_CSV_DIR    = Path("hf/csv/listado-69b")
OUTPUT_MERGED = HF_CSV_DIR / "listado-69b.csv"

_CLEAN_NAME: dict[str, str] = {
    "69b":     "69b.csv",
    "69b-bis": "69b-bis.csv",
}

# Rename 69-B Bis columns to align with 69-B before transformation
_BIS_RENAME: dict[str, str] = {
    "No.":                                                    "No",
    "Número y fecha de oficio global definitivo SAT":         "Número y fecha de oficio global de definitivos SAT",
    "Publicación página SAT definitivo":                      "Publicación página SAT definitivos",
    "Número y fecha de oficio global definitivo DOF":         "Número y fecha de oficio global de definitivos DOF",
    "Publicación DOF definitivo":                             "Publicación DOF definitivos",
}

# (phase, source_col, code_col, original_col)
_OFICIO_SPLIT: list[tuple[str, str, str, str]] = [
    ("presuncion",          "Número y fecha de oficio global de presunción SAT",                       "presuncion_oficio_sat",          "presuncion_oficio_sat_original"),
    ("presuncion",          "Número y fecha de oficio global de presunción DOF",                       "presuncion_oficio_dof",          "presuncion_oficio_dof_original"),
    ("desvirtuados",        "Número y fecha de oficio global de contribuyentes que desvirtuaron SAT",  "desvirtuados_oficio_sat",        "desvirtuados_oficio_sat_original"),
    ("desvirtuados",        "Número y fecha de oficio global de contribuyentes que desvirtuaron DOF",  "desvirtuados_oficio_dof",        "desvirtuados_oficio_dof_original"),
    ("definitivos",         "Número y fecha de oficio global de definitivos SAT",                      "definitivos_oficio_sat",         "definitivos_oficio_sat_original"),
    ("definitivos",         "Número y fecha de oficio global de definitivos DOF",                      "definitivos_oficio_dof",         "definitivos_oficio_dof_original"),
    ("sentencia_favorable", "Número y fecha de oficio global de sentencia favorable SAT",              "sentencia_favorable_oficio_sat", "sentencia_favorable_oficio_sat_original"),
    ("sentencia_favorable", "Número y fecha de oficio global de sentencia favorable DOF",              "sentencia_favorable_oficio_dof", "sentencia_favorable_oficio_dof_original"),
]

# (phase, source_col, dest_col)
_DATE_COLS: list[tuple[str, str, str]] = [
    ("presuncion",          "Publicación página SAT presuntos",           "presuncion_fecha_sat"),
    ("presuncion",          "Publicación DOF presuntos",                  "presuncion_fecha_dof"),
    ("desvirtuados",        "Publicación página SAT desvirtuados",        "desvirtuados_fecha_sat"),
    ("desvirtuados",        "Publicación DOF desvirtuados",               "desvirtuados_fecha_dof"),
    ("definitivos",         "Publicación página SAT definitivos",         "definitivos_fecha_sat"),
    ("definitivos",         "Publicación DOF definitivos",                "definitivos_fecha_dof"),
    ("sentencia_favorable", "Publicación página SAT sentencia favorable", "sentencia_favorable_fecha_sat"),
    ("sentencia_favorable", "Publicación DOF sentencia favorable",        "sentencia_favorable_fecha_dof"),
]

OUTPUT_FIELDS: list[str] = [
    "articulo", "id", "numero", "rfc", "nombre_contribuyente", "situacion_contribuyente",
    "presuncion_oficio_sat",          "presuncion_oficio_sat_original",          "presuncion_fecha_sat",
    "presuncion_oficio_dof",          "presuncion_oficio_dof_original",          "presuncion_fecha_dof",
    "presuncion_tiene_extra",
    "desvirtuados_oficio_sat",        "desvirtuados_oficio_sat_original",        "desvirtuados_fecha_sat",
    "desvirtuados_oficio_dof",        "desvirtuados_oficio_dof_original",        "desvirtuados_fecha_dof",
    "desvirtuados_tiene_extra",
    "definitivos_oficio_sat",         "definitivos_oficio_sat_original",         "definitivos_fecha_sat",
    "definitivos_oficio_dof",         "definitivos_oficio_dof_original",         "definitivos_fecha_dof",
    "definitivos_tiene_extra",
    "sentencia_favorable_oficio_sat", "sentencia_favorable_oficio_sat_original", "sentencia_favorable_fecha_sat",
    "sentencia_favorable_oficio_dof", "sentencia_favorable_oficio_dof_original", "sentencia_favorable_fecha_dof",
    "sentencia_favorable_tiene_extra",
]

_SITUACION_MAP: dict[str, str] = {
    "sentencia favorable": "SENTENCIA_FAVORABLE",
    "desvirtuado":         "DESVIRTUADO",
    "definitivo":          "DEFINITIVO",
    "presunto":            "PRESUNTO",
}


# ── helpers ────────────────────────────────────────────────────────────────────


def _strip_preamble(data: bytes) -> str:
    """Decode Latin-1, skip SAT disclaimer rows, return text from header row onward."""
    text = data.decode("latin-1")
    lines = text.splitlines(keepends=True)
    for i, line in enumerate(lines):
        if re.match(r"No\.?,", line.lstrip("\ufeff").lstrip('"')):
            return "".join(lines[i:])
    return text


def _read_raw(path: Path, articulo: str, rename: dict[str, str] | None = None) -> list[dict]:
    """Read a raw SAT CSV (Latin-1, with preamble), strip preamble, inject articulo."""
    text = _strip_preamble(path.read_bytes())
    rows: list[dict] = []
    reader = csv.DictReader(text.splitlines())
    for row in reader:
        if rename:
            row = {rename.get(k, k): v for k, v in row.items()}
        row["articulo"] = articulo
        rows.append(row)
    return rows


def _first(text: str, sep: str) -> tuple[str, bool]:
    """Split on sep, return (first_value, had_multiple)."""
    parts = [p.strip() for p in text.split(sep) if p.strip()]
    return (parts[0] if parts else ""), len(parts) > 1


def _extract_oficio(text: str) -> tuple[str, bool]:
    """Extract first code from possibly multi-value oficio field (sep: ' // ')."""
    first, multi = _first(text, " // ")
    code = first.split()[0] if first.strip() else ""
    return code, multi


def _fmt_pub_date(text: str) -> tuple[str, bool]:
    """Convert first DD/MM/YYYY → YYYY-MM-DD from possibly multi-value field (sep: ' - ')."""
    first, multi = _first(text, " - ")
    if not first:
        return "", multi
    try:
        day, month, year = first.strip().split("/")
        return f"{year}-{month.zfill(2)}-{day.zfill(2)}", multi
    except Exception:
        return first, multi


def _normalize_situacion(text: str) -> str:
    return _SITUACION_MAP.get(text.strip().lower(), text.strip().upper().replace(" ", "_"))


def _transform(row: dict, row_id: int) -> dict:
    out: dict = {
        "articulo":                row.get("articulo", ""),
        "id":                      row_id,
        "numero":                  row.get("No", ""),
        "rfc":                     row.get("RFC", ""),
        "nombre_contribuyente":    row.get("Nombre del Contribuyente", ""),
        "situacion_contribuyente": _normalize_situacion(row.get("Situación del contribuyente", "")),
    }
    phase_extra: dict[str, bool] = {}
    for phase, src, code_col, original_col in _OFICIO_SPLIT:
        val = row.get(src, "")
        out[original_col] = val
        code, multi = _extract_oficio(val)
        out[code_col] = code
        phase_extra[phase] = phase_extra.get(phase, False) or multi
    for phase, src, dest in _DATE_COLS:
        date, multi = _fmt_pub_date(row.get(src, ""))
        out[dest] = date
        phase_extra[phase] = phase_extra.get(phase, False) or multi
    out["presuncion_tiene_extra"]          = int(phase_extra.get("presuncion", False))
    out["desvirtuados_tiene_extra"]        = int(phase_extra.get("desvirtuados", False))
    out["definitivos_tiene_extra"]         = int(phase_extra.get("definitivos", False))
    out["sentencia_favorable_tiene_extra"] = int(phase_extra.get("sentencia_favorable", False))
    return out


# ── main ───────────────────────────────────────────────────────────────────────


def merge(manifest_path: Path, output_dir: Path, output_merged: Path) -> int:
    if not manifest_path.exists():
        print(f"ERROR: {manifest_path} not found. Run scripts/listado_69b/scrape.py first.", file=sys.stderr)
        return 1

    manifest  = json.loads(manifest_path.read_text(encoding="utf-8"))
    articles  = {a["key"]: a for a in manifest.get("articles", [])}

    output_dir.mkdir(parents=True, exist_ok=True)

    all_rows_a: list[dict] = []
    all_rows_b: list[dict] = []
    id_counter = 1

    for key, clean_name in _CLEAN_NAME.items():
        article = articles.get(key)
        if not article:
            print(f"  WARN: no manifest entry for {key}", file=sys.stderr)
            continue

        raw_path = Path(article["local_file"])
        if not raw_path.exists():
            print(f"  ERROR: raw file not found: {raw_path}", file=sys.stderr)
            return 1

        articulo = article["name"].replace("Artículo ", "")
        rename   = _BIS_RENAME if key == "69b-bis" else None
        raw_rows = _read_raw(raw_path, articulo=articulo, rename=rename)

        # Transform and assign IDs
        transformed: list[dict] = []
        for row in raw_rows:
            transformed.append(_transform(row, id_counter))
            id_counter += 1

        # Write clean individual file (no articulo column)
        dest = output_dir / clean_name
        ind_fields = [f for f in OUTPUT_FIELDS if f != "articulo"]
        with dest.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=ind_fields, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(transformed)
        print(f"  {key}  →  {dest}  ({len(transformed)} rows)", file=sys.stderr)

        if key == "69b":
            all_rows_a = transformed
        else:
            all_rows_b = transformed

    if not all_rows_a and not all_rows_b:
        print("No data to merge.", file=sys.stderr)
        return 1

    all_rows = all_rows_a + all_rows_b
    with output_merged.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_FIELDS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(all_rows)
    print(
        f"  merged  →  {output_merged}  "
        f"({len(all_rows_a)} 69-B + {len(all_rows_b)} 69-B Bis = {len(all_rows)} rows)",
        file=sys.stderr,
    )
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, default=MANIFEST)
    parser.add_argument("--output",   type=Path, default=HF_CSV_DIR)
    parser.add_argument("--merged",   type=Path, default=OUTPUT_MERGED)
    args = parser.parse_args()
    return merge(args.manifest, args.output, args.merged)


if __name__ == "__main__":
    raise SystemExit(main())
