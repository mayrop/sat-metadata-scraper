#!/usr/bin/env python3
"""Generate the Hugging Face dataset folder for TIGIE."""

from __future__ import annotations

import argparse
import csv
import json
import shutil
import sys
from pathlib import Path

CSV_PATH = Path("hf/csv/extra/tigie/tigie.csv")
RAW_DIR = Path("hf/extra/tigie")
CHAPTERS_CSV = Path("static/tigie/secciones_capitulos.csv")
METADATA_JSON = Path("output/tigie-metadata.json")
MANIFEST_JSON = Path("output/tigie-manifest.json")
OUTPUT_DIR = Path("hf/dataset/tigie")

DATASET_TITLE = "TIGIE"
DATASET_DESCRIPTION = (
    "Tarifa de la Ley de los Impuestos Generales de Importacion y de Exportacion "
    "(TIGIE) derivada del archivo oficial de SNICE, unificada con capitulos, "
    "partidas, subpartidas, fracciones y NICOs.\n"
)


def _read_json(path: Path) -> dict:
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def _write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _build_metadata_row(metadata: dict, manifest: dict, csv_path: Path, chapters_csv: Path) -> dict[str, object]:
    tipos = metadata.get("tipos", {})
    counts = manifest.get("counts", {})
    files = manifest.get("files", {})
    return {
        "dataset": metadata.get("dataset", "tigie"),
        "catalogo": "tigie",
        "source_url": metadata.get("source_url", ""),
        "source_file": metadata.get("source_file", ""),
        "source_filename": Path(metadata.get("source_file", "")).name,
        "chapters_csv": "secciones_capitulos.csv",
        "chapters_csv_filename": chapters_csv.name if chapters_csv else "",
        "raw_hash": metadata.get("raw_hash", ""),
        "csv_path": str(csv_path),
        "csv_sha256": _sha256_file(csv_path) if csv_path.exists() else "",
        "rows": metadata.get("rows", ""),
        "leaf_rows": metadata.get("leaf_rows", ""),
        "max_level": metadata.get("max_level", ""),
        "capitulos": tipos.get("capitulo", 0),
        "partidas": tipos.get("partida", 0),
        "intermedios": tipos.get("intermedio", 0),
        "subpartidas": tipos.get("subpartida", 0),
        "fracciones": tipos.get("fraccion", 0),
        "nicos": tipos.get("nico", 0),
        "combined_rows": counts.get("combined_rows", 0),
        "unified_rows": counts.get("unified_rows", metadata.get("rows", 0)),
        "fecha_extraccion": manifest.get("fecha_extraccion", ""),
        "raw_xlsx": files.get("raw_xlsx", ""),
        "raw_consolidated": files.get("raw_consolidated", ""),
        "raw_consolidated_xlsx": files.get("raw_consolidated_xlsx", ""),
        "raw_metadata": files.get("raw_metadata", ""),
        "hash": metadata.get("hash", ""),
    }


def _sha256_file(path: Path) -> str:
    import hashlib
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _build_readme(repo_id: str, metadata_row: dict[str, object], raw_entries: list[dict[str, object]]) -> str:
    source_url = str(metadata_row.get("source_url", ""))
    lines = [
        "---",
        "configs:",
        "",
        "- config_name: tigie",
        "  data_files:",
        "    - split: train",
        "      path: tigie.csv",
        "",
        "- config_name: metadata",
        "  data_files:",
        "    - split: train",
        "      path: metadata.csv",
        "",
        "- config_name: secciones_capitulos",
        "  data_files:",
        "    - split: train",
        "      path: secciones_capitulos.csv",
        "",
        "- config_name: raw_files",
        "  data_files:",
        "    - split: train",
        "      path: raw_files.csv",
        "",
        "---",
        "",
        f"# {DATASET_TITLE}",
        "",
        DATASET_DESCRIPTION,
        f"Fuente: [{source_url}]({source_url})" if source_url else "Fuente: SNICE",
        "",
        "## Uso",
        "",
        "```python",
        "from datasets import load_dataset",
        "",
        f'ds = load_dataset("{repo_id}", "tigie")',
        'df = ds["train"].to_pandas()',
        "",
        f'ds_cap = load_dataset("{repo_id}", "secciones_capitulos")',
        'df_cap = ds_cap["train"].to_pandas()',
        "```",
        "",
        "## Resumen",
        "",
        f"- Filas: `{metadata_row.get('rows', '')}`",
        f"- Capitulos: `{metadata_row.get('capitulos', '')}`",
        f"- Partidas: `{metadata_row.get('partidas', '')}`",
        f"- Intermedios: `{metadata_row.get('intermedios', '')}`",
        f"- Subpartidas: `{metadata_row.get('subpartidas', '')}`",
        f"- Fracciones: `{metadata_row.get('fracciones', '')}`",
        f"- NICOs: `{metadata_row.get('nicos', '')}`",
        "",
        "## Archivos incluidos",
        "",
        "- `tigie.csv`: version unificada lista para analisis",
        "- `secciones_capitulos.csv`: capitulos oficiales usados para poblar los nombres de capitulo",
        "- `metadata.csv`: resumen del proceso y conteos",
        "- `raw_files.csv`: indice de archivos fuente exportados",
        "",
    ]
    if raw_entries:
        lines.append("## Archivos fuente")
        lines.append("")
        for raw in raw_entries:
            lines.append(f"- `{raw['path']}`")
        lines.append("")
    return "\n".join(lines)


def generate(
    csv_path: Path,
    raw_dir: Path,
    chapters_csv: Path,
    metadata_json: Path,
    manifest_json: Path,
    output_dir: Path,
    repo_id: str,
) -> int:
    if not csv_path.exists():
        print(f"Missing TIGIE CSV: {csv_path}", file=sys.stderr)
        return 1

    metadata = _read_json(metadata_json)
    manifest = _read_json(manifest_json)

    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / ".huggingface_ignore").write_text(
        ".DS_Store\n__pycache__/\n*.pyc\n*.pyo\n", encoding="utf-8"
    )

    csv_dest = output_dir / "tigie.csv"
    shutil.copy2(csv_path, csv_dest)
    print(f"  tigie      →  {csv_dest}", file=sys.stderr)

    if chapters_csv.exists():
        chapters_dest = output_dir / "secciones_capitulos.csv"
        shutil.copy2(chapters_csv, chapters_dest)
        print(f"  capitulos  →  {chapters_dest}", file=sys.stderr)

    raw_entries: list[dict[str, object]] = []
    raw_output_dir = output_dir / "raw"
    if raw_output_dir.exists():
        shutil.rmtree(raw_output_dir)
    if raw_dir.exists():
        for src in sorted(raw_dir.iterdir()):
            if not src.is_file():
                continue
            dest = raw_output_dir / src.name
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
            raw_entries.append(
                {
                    "path": f"raw/{src.name}",
                    "source_path": str(src),
                    "size_bytes": src.stat().st_size,
                    "sha256": _sha256_file(src),
                }
            )
            print(f"  raw        →  {dest}", file=sys.stderr)

    metadata_row = _build_metadata_row(metadata, manifest, csv_path, chapters_csv)
    metadata_fields = list(metadata_row.keys())
    _write_csv(output_dir / "metadata.csv", metadata_fields, [metadata_row])
    print(f"  metadata   →  {output_dir / 'metadata.csv'}", file=sys.stderr)

    _write_csv(
        output_dir / "raw_files.csv",
        ["path", "source_path", "size_bytes", "sha256"],
        raw_entries,
    )
    print(f"  raw_files  →  {output_dir / 'raw_files.csv'}", file=sys.stderr)

    readme_path = output_dir / "README.md"
    readme_path.write_text(_build_readme(repo_id, metadata_row, raw_entries), encoding="utf-8")
    print(f"  README     →  {readme_path}", file=sys.stderr)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--csv", type=Path, default=CSV_PATH, help=f"TIGIE CSV path (default: {CSV_PATH})")
    parser.add_argument("--raw-dir", type=Path, default=RAW_DIR, help=f"TIGIE raw files directory (default: {RAW_DIR})")
    parser.add_argument("--chapters-csv", type=Path, default=CHAPTERS_CSV, help=f"TIGIE chapters CSV path (default: {CHAPTERS_CSV})")
    parser.add_argument("--metadata-json", type=Path, default=METADATA_JSON, help=f"TIGIE metadata JSON path (default: {METADATA_JSON})")
    parser.add_argument("--manifest", type=Path, default=MANIFEST_JSON, help=f"TIGIE manifest JSON path (default: {MANIFEST_JSON})")
    parser.add_argument("--output", type=Path, default=OUTPUT_DIR, help=f"Output directory for the HF dataset (default: {OUTPUT_DIR})")
    parser.add_argument("--repo-id", default="mayrop/tigie", help="HF repo id used in the README usage example")
    args = parser.parse_args()
    return generate(args.csv, args.raw_dir, args.chapters_csv, args.metadata_json, args.manifest, args.output, args.repo_id)


if __name__ == "__main__":
    raise SystemExit(main())
