#!/usr/bin/env python3
"""Generate a Hugging Face dataset index for SAT CFDI catalogs.

Uses catalog_state.csv as the authoritative catalog list so that the README
and metadata always reflect every known catalog — even when local hf/csv/ or
hf/xls/ files are absent (they live on HF from a prior upload).

For each catalog in the state, copies the local CSV into hf/dataset/catalogos/ if it
exists; otherwise leaves the entry in README/metadata so HF keeps serving the
previously-uploaded file.

Output layout:
    hf/dataset/catalogos/
    ├── README.md               ← YAML front-matter + HF configs
    ├── metadata/
    │   └── catalogos.csv
    ├── anexo20/
    │   ├── c_uso_cfdi.csv
    │   └── ...
    ├── retenciones/
    │   └── ...
    └── complementos/
        ├── carta_porte/
        ├── nomina/
        ├── comercio_exterior/
        └── recepcion_pagos/

Usage:
    uv run scripts/catalogos/generate_hf.py
    uv run scripts/catalogos/generate_hf.py --csv-dir hf/csv --output hf/dataset/catalogos
"""
from __future__ import annotations

import argparse
import csv
import re
import shutil
import sys
import unicodedata
from pathlib import Path

# ── constants ──────────────────────────────────────────────────────────────────

HF_XLS_DIR     = Path("hf/xls")
HF_CSV_DIR     = Path("hf/csv")
HF_DATASET_DIR = Path("hf/dataset/catalogos")
CATALOG_STATE  = Path("catalog_state.csv")
CATALOG_CSV    = Path("output/catalog.csv")

# Override: map section_rel path → friendly slug used in config names and output subdir.
_SECTION_SLUG_OVERRIDES: dict[str, str] = {
    "anexo20/cfdi":                          "anexo20__cfdi",
    "anexo20/retenciones":                   "anexo20__retenciones",
    "complementos/recibo-de-pago-de-nomina": "complemento_nomina",
    "complementos/carta-porte":              "complemento_carta_porte",
    "complementos/recepcion-de-pagos":       "complemento_recepcion_pagos",
    "complementos/comercio-exterior":        "complemento_comercio_exterior",
}

# Map slug → flat subdirectory path within the dataset
_SLUG_TO_SUBDIR: dict[str, str] = {
    "anexo20__cfdi":                 "anexo20/cfdi",
    "anexo20__retenciones":          "anexo20/retenciones",
    "complemento_nomina":            "complementos/nomina",
    "complemento_carta_porte":       "complementos/carta_porte",
    "complemento_recepcion_pagos":   "complementos/recepcion_pagos",
    "complemento_comercio_exterior": "complementos/comercio_exterior",
}


# ── helpers ────────────────────────────────────────────────────────────────────


def _slugify(value: str) -> str:
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    # Insert underscore between camelCase transitions: c_TipoHoras → c_Tipo_Horas
    value = re.sub(r"([a-z])([A-Z])", r"\1_\2", value)
    value = re.sub(r"[^0-9a-z]+", "_", value.lower())
    return value.strip("_")


def _ver_key(name: str) -> tuple:
    """Sort key for version folder names like '3-3', '4-0', '1-2-e'."""
    parts = []
    for p in name.split("-"):
        try:
            parts.append(int(p))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def _clean_stem(stem: str) -> str:
    """Strip version numbers, dates, and hashes from a catalog stem.

    Mirrors scrape.py _clean_stem so config names stay stable across SAT file renames.
    """
    s = stem
    s = re.sub(r"_[Vv]_?\d+[\d.]*", "", s)                        # _V_4, _v17
    s = re.sub(r"_\d{6,8}(?=_|$)", "", s)                         # _20231219
    s = re.sub(r"\d*_[0-9a-f]{8,}$", "", s, flags=re.IGNORECASE)  # _8ca5655de2
    s = re.sub(r"_[rR][A-Za-z0-9]{1,2}$", "", s)                  # _rA, _rB
    s = re.sub(r"_\d+$", "", s)                                    # _1, _2
    s = re.sub(r"(?<=[A-Za-z])\d{1,4}$", "", s)                   # Moneda20, CFDI40
    return s.strip("_") or stem


def _catalog_slug(csv_name: str) -> str:
    """c_UsoCFDI.csv → c_uso_cfdi  (version suffixes stripped)"""
    return _slugify(_clean_stem(Path(csv_name).stem))


def _section_slug(section_rel: str) -> str:
    """Return the namespace slug for a section path like 'complementos/carta-porte'."""
    if section_rel in _SECTION_SLUG_OVERRIDES:
        return _SECTION_SLUG_OVERRIDES[section_rel]
    parts = section_rel.split("/")
    if parts[0] == "complementos" and len(parts) == 2:
        return "complemento_" + _slugify(parts[1])
    return _slugify(section_rel.replace("/", "_"))


def _is_version_folder(name: str) -> bool:
    """True when the folder name is a version string (starts with digit), not a section slug."""
    return bool(name and re.match(r"^\d", name))


# ── state loading ──────────────────────────────────────────────────────────────


def _load_state_rows(state_file: Path) -> list[dict]:
    """Return every row from catalog_state.csv."""
    if not state_file.exists():
        return []
    with state_file.open(newline="", encoding="utf-8") as f:
        return [dict(row) for row in csv.DictReader(f)]


def _load_latest_map(catalog_csv: Path) -> dict[str, str]:
    """Return {local_file: latest} from output/catalog.csv."""
    if not catalog_csv.exists():
        return {}
    result: dict[str, str] = {}
    with catalog_csv.open(newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            lf = row.get("local_file", "")
            if lf:
                result[lf] = row.get("latest", "")
    return result


# ── README generation ──────────────────────────────────────────────────────────


def _build_readme(entries: list[dict]) -> str:
    """Build README.md content with YAML front-matter HF configs."""
    lines = ["---", "configs:"]
    for e in entries:
        lines.append(f"\n- config_name: {e['config_name']}")
        lines.append(  "  data_files:")
        lines.append(  "    - split: train")
        lines.append(f"      path: {e['path']}")
    lines.append("\n- config_name: metadata")
    lines.append(  "  data_files:")
    lines.append(  "    - split: train")
    lines.append(  "      path: metadata/catalogos.csv")
    lines.append("\n---\n")
    lines.append("# SAT CFDI Catálogos\n")
    lines.append(
        "Catálogos oficiales del SAT (Servicio de Administración Tributaria) para CFDI.\n"
    )
    lines.append("Incluye los catálogos del Anexo 20 (factura electrónica y retenciones)")
    lines.append("y los complementos de carta porte, nómina, comercio exterior y recepción de pagos.\n")
    lines.append("## Uso\n")
    # Pick a concrete example config: prefer the latest c_uso_cfdi, fall back to first entry
    uso_cfdi = sorted(
        (e for e in entries if "c_uso_cfdi" in e["config_name"] and "regimen" not in e["config_name"]),
        key=lambda e: _ver_key(e.get("source_version", "")),
    )
    example_config = (
        uso_cfdi[-1]["config_name"] if uso_cfdi
        else (entries[0]["config_name"] if entries else "anexo20__cfdi__4_0__c_uso_cfdi")
    )
    lines.append("```python")
    lines.append("from datasets import load_dataset, get_dataset_config_names\n")
    lines.append("# Cargar un catálogo específico")
    lines.append(f'ds = load_dataset("mayrop/sat-catalogos", "{example_config}")')
    lines.append('df = ds["train"].to_pandas()\n')
    lines.append("# Cargar todos los catálogos de una vez")
    lines.append('configs = get_dataset_config_names("mayrop/sat-catalogos")')
    lines.append('all_data = {c: load_dataset("mayrop/sat-catalogos", c)["train"].to_pandas() for c in configs}')
    lines.append("```\n")
    lines.append("## Fuentes\n")
    lines.append("| Sección | Fuente oficial |")
    lines.append("|---------|---------------|")
    lines.append("| Anexo 20 — Factura electrónica | [omawww.sat.gob.mx/tramitesyservicios/Paginas/anexo_20.htm](http://omawww.sat.gob.mx/tramitesyservicios/Paginas/anexo_20.htm) |")
    lines.append("| Anexo 20 — Retenciones e información de pagos | [omawww.sat.gob.mx/tramitesyservicios/Paginas/CFDI_retenciones.htm](http://omawww.sat.gob.mx/tramitesyservicios/Paginas/CFDI_retenciones.htm) |")
    lines.append("| Complementos | [sat.gob.mx/portal/public/tramites/complementos-de-factura](https://www.sat.gob.mx/portal/public/tramites/complementos-de-factura) |")
    lines.append("")
    lines.append("## Catálogos disponibles\n")
    # Build ordered list of unique (namespace, version) groups
    groups_seen: list[tuple[str, str]] = []
    for e in entries:
        g = (e["namespace"], e.get("source_version", ""))
        if g not in groups_seen:
            groups_seen.append(g)
    # Quick index
    for ns, ver in groups_seen:
        label = f"{ns} — {ver}" if ver else ns
        anchor = re.sub(r"[^\w\s-]", "", label.lower())  # drop em-dash and other punctuation
        anchor = re.sub(r" ", "-", anchor)              # each space → hyphen
        lines.append(f"- [{label}](#{anchor})")
    lines.append("")
    # Full listing
    current_group: tuple | None = None
    for e in entries:
        ns  = e["namespace"]
        ver = e.get("source_version", "")
        group = (ns, ver)
        if group != current_group:
            heading = f"### {ns} — {ver}" if ver else f"### {ns}"
            lines.append(f"\n{heading}\n")
            current_group = group
        desc = e.get("descripcion", "") or e.get("description", "")
        desc_str = f" — {desc}" if desc else ""
        lines.append(f"- `{e['config_name']}`{desc_str}")
    lines.append("")
    return "\n".join(lines)


# ── main ───────────────────────────────────────────────────────────────────────


def generate(csv_dir: Path, state_file: Path, output_dir: Path, xls_dir: Path | None = None, catalog_csv: Path | None = None) -> int:
    state_rows = _load_state_rows(state_file)
    latest_map = _load_latest_map(catalog_csv or CATALOG_CSV)
    if not state_rows:
        print(f"No catalog state found at {state_file}. Run scripts/catalogos/extract.py first.", file=sys.stderr)
        return 1

    # Collect all rows, annotating each with section_rel and is_versioned.
    # section_rel: for versioned dirs (folder_version starts with a digit) = section field;
    #              for flat dirs (no numeric version) = section/folder_version.
    # All versions are exported; versioned configs include the version in their name
    # (e.g. anexo20__cfdi__4_0__c_uso_cfdi) while non-versioned configs use the plain name.
    all_rows: list[dict] = []
    seen_nonversioned: set[tuple[str, str]] = set()
    for row in state_rows:
        section        = row.get("section", "")
        folder_version = row.get("folder_version", "")
        catalogo       = row.get("catalogo", "")
        if not section or not catalogo:
            continue
        is_versioned = _is_version_folder(folder_version)
        section_rel  = section if is_versioned else f"{section}/{folder_version}".strip("/")
        # Non-versioned entries are unique by (section_rel, catalogo) — skip duplicates
        if not is_versioned:
            key = (section_rel, catalogo)
            if key in seen_nonversioned:
                continue
            seen_nonversioned.add(key)
        all_rows.append({**row, "_section_rel": section_rel, "_is_versioned": is_versioned,
                         "_folder_version": folder_version})

    # Sort: by (namespace slug, version, catalog_id) so each section+version is a contiguous block
    all_rows.sort(key=lambda r: (
        _section_slug(r["_section_rel"]),
        _ver_key(r["_folder_version"]),
        _catalog_slug(f"{r.get('catalogo', '')}.csv"),
    ))

    entries: list[dict] = []
    copied = 0
    missing_locally = 0

    for row in all_rows:
        folder_version = row["_folder_version"]
        is_versioned   = row["_is_versioned"]
        section_rel    = row["_section_rel"]
        section_field  = row.get("section", "")
        catalogo       = row.get("catalogo", "")

        slug         = _section_slug(section_rel)
        section_path = "/".join(_slugify(p) for p in section_rel.split("/"))
        catalog_slug = _catalog_slug(f"{catalogo}.csv")

        if is_versioned:
            version_slug = _slugify(folder_version)
            config_name  = f"{slug}__{version_slug}__{catalog_slug}"
            dest_rel     = f"{section_path}/{folder_version}/{catalog_slug}.csv"
        else:
            config_name  = f"{slug}__{catalog_slug}"
            dest_rel     = f"{section_path}/{catalog_slug}.csv"

        dest = output_dir / dest_rel

        # CSV lives at csv_dir / section / folder_version / catalogo.csv
        src = csv_dir / section_field / folder_version / f"{catalogo}.csv"
        if src.exists():
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
            copied += 1
            print(f"  {config_name}  →  {dest_rel}", file=sys.stderr)
        else:
            missing_locally += 1
            print(f"  {config_name}  →  {dest_rel}  (not local — kept from HF)", file=sys.stderr)

        source_xls = row.get("source_xls", "")
        latest_val = latest_map.get(source_xls, "")
        entry = {
            "config_name":    config_name,
            "namespace":      slug,
            "catalog_id":     catalog_slug,
            "path":           dest_rel,
            "source_version": folder_version,
            "latest":         latest_val,
        }
        for k, v in row.items():
            if k not in entry and not k.startswith("_"):
                entry[k] = v
        entries.append(entry)

    if not entries:
        print(f"No catalogs found in {state_file}.", file=sys.stderr)
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)

    # Write .huggingface_ignore
    (output_dir / ".huggingface_ignore").write_text(
        ".DS_Store\n__pycache__/\n*.pyc\n*.pyo\n", encoding="utf-8"
    )

    # Write README.md
    readme_path = output_dir / "README.md"
    readme_path.write_text(_build_readme(entries), encoding="utf-8")
    print(
        f"\nREADME   → {readme_path}  ({len(entries)} configs, {copied} copied, {missing_locally} from HF)",
        file=sys.stderr,
    )

    # Write metadata/catalogos.csv
    meta_dir = output_dir / "metadata"
    meta_dir.mkdir(parents=True, exist_ok=True)
    meta_path = meta_dir / "catalogos.csv"
    _base_fields = ["config_name", "namespace", "catalog_id", "path", "source_version", "latest"]
    fieldnames: list[str] = list(_base_fields)
    for e in entries:
        for k in e:
            if k not in fieldnames:
                fieldnames.append(k)
    with meta_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(entries)
    print(f"Metadata → {meta_path}  ({len(entries)} rows)", file=sys.stderr)

    # Copy any locally-available XLS files into output_dir/xls/
    if xls_dir and xls_dir.exists():
        xls_count = 0
        for src in sorted(xls_dir.rglob("*")):
            if src.suffix.lower() not in (".xls", ".xlsx"):
                continue
            rel = src.relative_to(xls_dir)
            dest = output_dir / "xls" / rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
            xls_count += 1
        print(f"XLS      → {output_dir}/xls/  ({xls_count} files)", file=sys.stderr)

    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--csv-dir",
        type=Path,
        default=HF_CSV_DIR,
        help=f"Root directory of extracted CSVs (default: {HF_CSV_DIR})",
    )
    parser.add_argument(
        "--state-file",
        type=Path,
        default=CATALOG_STATE,
        help=f"Path to catalog state CSV (default: {CATALOG_STATE})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=HF_DATASET_DIR,
        help=f"Output directory for the HF dataset (default: {HF_DATASET_DIR})",
    )
    parser.add_argument(
        "--xls-dir",
        type=Path,
        default=HF_XLS_DIR,
        help=f"Root directory of XLS source files to include (default: {HF_XLS_DIR})",
    )
    parser.add_argument(
        "--catalog-file",
        type=Path,
        default=CATALOG_CSV,
        help=f"Path to scraped catalog CSV for 'latest' lookup (default: {CATALOG_CSV})",
    )
    args = parser.parse_args()
    return generate(args.csv_dir, args.state_file, args.output, args.xls_dir, args.catalog_file)


if __name__ == "__main__":
    raise SystemExit(main())
