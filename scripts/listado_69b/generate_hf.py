#!/usr/bin/env python3
"""Generate Hugging Face dataset folder for SAT contribuyentes (69-B / 69-B Bis).

Copies the merged CSV and generates a metadata CSV + README.md ready to upload to HF.

─── Configuration ────────────────────────────────────────────────────────────
Edit the constants below (or pass CLI flags) to adapt this script to a
different repository without touching any other code.
──────────────────────────────────────────────────────────────────────────────

Output layout:
    hf/dataset/listado-69b/
    ├── README.md
    ├── listado-69.csv    ← merged 69-B + 69-B Bis
    └── metadata.csv       ← source info (urls, dates, hashes)

Usage:
  uv run scripts/listado_69b/generate_hf.py
  uv run scripts/listado_69b/generate_hf.py --output hf/dataset-listado-69b
"""
from __future__ import annotations

import argparse
import csv
import json
import shutil
import sys
from datetime import timezone
from email.utils import parsedate_to_datetime
from pathlib import Path

# ── configuration ──────────────────────────────────────────────────────────────

CSV_DIR        = Path("hf/csv/listado-69b")
MANIFEST_PATH  = Path("output/listado-69b-manifest.json")
OUTPUT_DIR     = Path("hf/dataset/listado-69b")

DATASET_TITLE       = "SAT Contribuyentes 69-B / 69-B Bis"
DATASET_DESCRIPTION = (
    "Listados oficiales del SAT (Servicio de Administración Tributaria) de contribuyentes "
    "publicados bajo los Artículos 69-B y 69-B Bis del Código Fiscal de la Federación.\n"
)
SOURCE_URL = (
    "http://omawww.sat.gob.mx/cifras_sat/Paginas/DatosAbiertos/contribuyentes_publicados.html"
)

# ── README ─────────────────────────────────────────────────────────────────────


def _build_readme(repo_id: str, articles: list[dict] | None = None) -> str:
    lines = [
        "---",
        "configs:",
        "",
        "- config_name: listado_69b",
        "  data_files:",
        "    - split: train",
        "      path: listado-69b.csv",
        "",
        "- config_name: metadata",
        "  data_files:",
        "    - split: train",
        "      path: metadata.csv",
        "",
        "---",
        "",
        f"# {DATASET_TITLE}",
        "",
        DATASET_DESCRIPTION,
        f"Fuente: [{SOURCE_URL}]({SOURCE_URL})",
        "",
        *([
            "## Última actualización",
            "",
            *[f"- **{a['name']}**: {_fmt_date(a.get('last_modified', ''))} (información al {a.get('info_date', '')})"
              for a in (articles or [])],
            "",
        ] if articles else []),
        "## Uso",
        "",
        "```python",
        "from datasets import load_dataset",
        "",
        "# Listado combinado (69-B + 69-B Bis)",
        f'ds = load_dataset("{repo_id}", "listado_69b")',
        'df = ds["train"].to_pandas()',
        "```",
        "",
        "## Columnas",
        "",
        "| Columna | Descripción |",
        "|---------|-------------|",
        "| `articulo` | Fuente: `69-B` o `69-B Bis` |",
        "| `id` | ID único del registro |",
        "| `numero` | Número de registro en el listado original |",
        "| `rfc` | RFC del contribuyente |",
        "| `nombre_contribuyente` | Razón social |",
        "| `situacion_contribuyente` | `PRESUNTO` / `DESVIRTUADO` / `DEFINITIVO` / `SENTENCIA_FAVORABLE` |",
        "| `presuncion_oficio_sat` | Clave del oficio de presunción SAT _(solo 69-B)_ |",
        "| `presuncion_oficio_sat_original` | Texto completo del oficio de presunción SAT _(solo 69-B)_ |",
        "| `presuncion_fecha_sat` | Fecha de publicación en página SAT (YYYY-MM-DD) _(solo 69-B)_ |",
        "| `presuncion_oficio_dof` | Clave del oficio de presunción DOF _(solo 69-B)_ |",
        "| `presuncion_oficio_dof_original` | Texto completo del oficio de presunción DOF _(solo 69-B)_ |",
        "| `presuncion_fecha_dof` | Fecha de publicación en DOF (YYYY-MM-DD) _(solo 69-B)_ |",
        "| `presuncion_tiene_extra` | `1` si hay más de un oficio/fecha en la etapa de presunción _(solo 69-B)_ |",
        "| `desvirtuados_oficio_sat` | Clave del oficio de desvirtuados SAT _(solo 69-B)_ |",
        "| `desvirtuados_oficio_sat_original` | Texto completo _(solo 69-B)_ |",
        "| `desvirtuados_fecha_sat` | Fecha de publicación SAT (YYYY-MM-DD) _(solo 69-B)_ |",
        "| `desvirtuados_oficio_dof` | Clave del oficio de desvirtuados DOF _(solo 69-B)_ |",
        "| `desvirtuados_oficio_dof_original` | Texto completo _(solo 69-B)_ |",
        "| `desvirtuados_fecha_dof` | Fecha de publicación DOF (YYYY-MM-DD) _(solo 69-B)_ |",
        "| `desvirtuados_tiene_extra` | `1` si hay más de un oficio/fecha en la etapa de desvirtuados _(solo 69-B)_ |",
        "| `definitivos_oficio_sat` | Clave del oficio definitivo SAT |",
        "| `definitivos_oficio_sat_original` | Texto completo |",
        "| `definitivos_fecha_sat` | Fecha de publicación SAT (YYYY-MM-DD) |",
        "| `definitivos_oficio_dof` | Clave del oficio definitivo DOF |",
        "| `definitivos_oficio_dof_original` | Texto completo |",
        "| `definitivos_fecha_dof` | Fecha de publicación DOF (YYYY-MM-DD) |",
        "| `definitivos_tiene_extra` | `1` si hay más de un oficio/fecha en la etapa de definitivos |",
        "| `sentencia_favorable_oficio_sat` | Clave del oficio de sentencia favorable SAT |",
        "| `sentencia_favorable_oficio_sat_original` | Texto completo |",
        "| `sentencia_favorable_fecha_sat` | Fecha de publicación SAT (YYYY-MM-DD) |",
        "| `sentencia_favorable_oficio_dof` | Clave del oficio de sentencia favorable DOF |",
        "| `sentencia_favorable_oficio_dof_original` | Texto completo |",
        "| `sentencia_favorable_fecha_dof` | Fecha de publicación DOF (YYYY-MM-DD) |",
        "| `sentencia_favorable_tiene_extra` | `1` si hay más de un oficio/fecha en la etapa de sentencia favorable |",
        "",
    ]
    return "\n".join(lines)


# ── metadata CSV ───────────────────────────────────────────────────────────────

_METADATA_FIELDS = ["key", "name", "url", "info_date", "last_modified", "size"]


def _fmt_date(value: str) -> str:
    """'Thu, 19 Mar 2026 00:10:07 GMT' → '2026-03-19 00:10:07'."""
    if not value:
        return value
    try:
        dt = parsedate_to_datetime(value)
        return dt.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")
    except Exception:
        return value


def _write_metadata(manifest_path: Path, dest: Path) -> None:
    manifest = {}
    if manifest_path.exists():
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))

    articles = manifest.get("articles", [])
    dest.parent.mkdir(parents=True, exist_ok=True)
    with dest.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=_METADATA_FIELDS, extrasaction="ignore")
        writer.writeheader()
        for article in articles:
            row = dict(article)
            row["last_modified"] = _fmt_date(row.get("last_modified", ""))
            writer.writerow(row)


# ── main ───────────────────────────────────────────────────────────────────────


def generate(csv_dir: Path, manifest_path: Path, output_dir: Path, repo_id: str) -> int:
    merged_src = csv_dir / "listado-69b.csv"
    if not merged_src.exists():
        print(f"ERROR: {merged_src} not found. Run scripts/listado_69b/merge.py first.", file=sys.stderr)
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)

    # Merged data file
    shutil.copy2(merged_src, output_dir / "listado-69b.csv")
    print(f"  listado-69b  →  {output_dir / 'listado-69b.csv'}", file=sys.stderr)

    # Metadata CSV
    meta_dest = output_dir / "metadata.csv"
    _write_metadata(manifest_path, meta_dest)
    print(f"  metadata     →  {meta_dest}", file=sys.stderr)

    # .huggingface_ignore
    (output_dir / ".huggingface_ignore").write_text(
        ".DS_Store\n__pycache__/\n*.pyc\n*.pyo\n", encoding="utf-8"
    )

    # README
    manifest = json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path.exists() else {}
    readme = output_dir / "README.md"
    readme.write_text(_build_readme(repo_id, articles=manifest.get("articles", [])), encoding="utf-8")
    print(f"\nREADME  → {readme}", file=sys.stderr)

    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--csv-dir", type=Path, default=CSV_DIR,
        help=f"Directory containing the cleaned CSVs (default: {CSV_DIR})",
    )
    parser.add_argument(
        "--manifest", type=Path, default=MANIFEST_PATH,
        help=f"Path to contribuyentes_manifest.json (default: {MANIFEST_PATH})",
    )
    parser.add_argument(
        "--output", type=Path, default=OUTPUT_DIR,
        help=f"Output directory for the HF dataset (default: {OUTPUT_DIR})",
    )
    parser.add_argument(
        "--repo-id", default="mayrop/sat-listado-69b",
        help="HF repo id used in the README usage example",
    )
    args = parser.parse_args()
    return generate(args.csv_dir, args.manifest, args.output, args.repo_id)


if __name__ == "__main__":
    raise SystemExit(main())
