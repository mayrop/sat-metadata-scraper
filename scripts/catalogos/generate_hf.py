#!/usr/bin/env python3
"""Generate a Hugging Face dataset index for SAT CFDI catalogs.

Uses catalog_state.csv as the authoritative catalog list so that the README
and metadata always reflect every known catalog — even when local hf/csv/ or
hf/raw/catalogos/ source files are absent (they live on HF from a prior upload).

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

HF_XLS_DIR     = Path("hf/raw/catalogos")
OUTPUT_FILES_DIR = Path("hf/raw/catalogos")
HF_CSV_DIR     = Path("hf/csv")
HF_DATASET_DIR = Path("hf/dataset/catalogos")
CATALOG_STATE  = Path("catalog_state.csv")
MATRIX_STATE   = Path("matrix_state.csv")
CATALOG_CSV    = Path("output/catalog.csv")

# Override: map section_rel path → friendly slug used in config names and output subdir.
_SECTION_SLUG_OVERRIDES: dict[str, str] = {
    "anexo20/cfdi":                          "anexo20__cfdi",
    "anexo20/retenciones":                   "anexo20__retenciones",
    "extra/unspsc":                          "extra_unspsc",
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

_SECTION_PATH_OVERRIDES: dict[str, str] = {
    "complementos-concepto/hidrocarburos-y-petroliferos": "complementos-concepto/hidrocarburos-y-petroliferos",
    "complementos_concepto/hidrocarburos-y-petroliferos": "complementos-concepto/hidrocarburos-y-petroliferos",
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


def _sha256_file(path: Path) -> str:
    import hashlib
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _should_include_raw_file(path: Path) -> bool:
    excluded_names = {"Thumbs.db", "desktop.ini"}
    for part in path.parts:
        if part.startswith("."):
            return False
        if part in excluded_names:
            return False
        if part.startswith("~$"):
            return False
    return True


def _should_include_output_files_rel(path: Path) -> bool:
    if not path.parts:
        return False
    return path.parts[0] in {
        "anexo-20",
        "complementos",
        "complementos-concepto",
        "complementos-retenciones",
    }


def _normalize_output_files_rel(path: Path) -> Path:
    parts = list(path.parts)
    if parts and parts[0] == "anexo-20":
        parts[0] = "anexo20"
    if len(parts) >= 2 and parts[0] == "anexo20":
        if parts[1] == "formato-de-factura":
            parts[1] = "cfdi"
        elif parts[1] == "factura-de-retenciones-e-informacion-de-pagos":
            parts[1] = "retenciones"
    for idx, part in enumerate(parts):
        if part.startswith("version-"):
            parts[idx] = part[len("version-") :]
        parts[idx] = parts[idx].replace("-revision-", "-")
        if parts[idx].startswith("revision-"):
            parts[idx] = parts[idx][len("revision-") :]
    return Path(*parts)


def _catalog_slug(csv_name: str) -> str:
    """c_UsoCFDI.csv → c_uso_cfdi  (version suffixes stripped)"""
    return _slugify(_clean_stem(Path(csv_name).stem))


def _entry_catalog_slug(row: dict) -> str:
    catalogo = f"{row.get('catalogo', '')}.csv"
    if row.get("file_type") == "matriz":
        return _slugify(Path(catalogo).stem)
    return _catalog_slug(catalogo)


def _entry_catalog_name(row: dict) -> str:
    if row.get("file_type") == "matriz":
        return "matriz_de_errores"
    return row.get("catalogo", "")


def _absolute_source_path(path_str: str) -> str:
    if not path_str:
        return ""
    path = Path(path_str)
    if path.is_absolute():
        return str(path)
    return str(path.resolve())


def _lookup_source_meta(catalog_index: dict[str, dict[str, str]], source_xls: str) -> dict[str, str]:
    if not source_xls:
        return {}
    if source_xls in catalog_index:
        return catalog_index[source_xls]
    for prefix in ("output/files/", "hf/raw/catalogos/"):
        if source_xls.startswith(prefix):
            return catalog_index.get(source_xls[len(prefix) :], {})
    return {}


def _section_slug(section_rel: str) -> str:
    """Return the namespace slug for a section path like 'complementos/carta-porte'."""
    if section_rel in _SECTION_SLUG_OVERRIDES:
        return _SECTION_SLUG_OVERRIDES[section_rel]
    parts = section_rel.split("/")
    if parts[0] == "complementos" and len(parts) == 2:
        return "complemento_" + _slugify(parts[1])
    return _slugify(section_rel.replace("/", "_"))


def _section_path(section_rel: str) -> str:
    if section_rel in _SECTION_PATH_OVERRIDES:
        return _SECTION_PATH_OVERRIDES[section_rel]
    return "/".join(_slugify(p) for p in section_rel.split("/"))


def _is_version_folder(name: str) -> bool:
    """True when the folder name is a version string (starts with digit), not a section slug."""
    return bool(name and re.match(r"^\d", name))


# ── state loading ──────────────────────────────────────────────────────────────


def _load_state_rows(state_files: list[Path]) -> list[dict]:
    """Return every row from the configured state CSV files."""
    rows: list[dict] = []
    for state_file in state_files:
        if not state_file.exists():
            continue
        with state_file.open(newline="", encoding="utf-8") as f:
            rows.extend(dict(row) for row in csv.DictReader(f))
    return rows


def _load_catalog_index(catalog_csv: Path) -> dict[str, dict[str, str]]:
    """Return {local_file: row} from output/catalog.csv for source metadata lookup."""
    if not catalog_csv.exists():
        return {}
    result: dict[str, dict[str, str]] = {}
    with catalog_csv.open(newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            lf = row.get("local_file", "")
            if lf:
                result[lf] = dict(row)
    return result


def _build_source_path_override_map(state_rows: list[dict]) -> dict[str, Path]:
    """Map local source file paths to their normalized raw export location.

    This lets raw source exports follow authoritative section/folder_version
    values from state files instead of inheriting legacy filesystem branches
    like hf/raw/catalogos/complementos/... when the logical section is now
    complementos-concepto/...
    """
    mapping: dict[str, Path] = {}
    for row in state_rows:
        source_xls = row.get("source_xls", "")
        section = row.get("section", "").strip("/")
        folder_version = row.get("folder_version", "").strip("/")
        if not source_xls or not section:
            continue
        filename = Path(source_xls).name
        target = Path(section)
        if folder_version:
            target /= folder_version
        target /= filename
        mapping[source_xls] = target
    return mapping


# ── README generation ──────────────────────────────────────────────────────────


def _build_readme(entries: list[dict], raw_entries: list[dict]) -> str:
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
    if raw_entries:
        lines.append("\n- config_name: raw_files")
        lines.append(  "  data_files:")
        lines.append(  "    - split: train")
        lines.append(  "      path: metadata/raw_files.csv")
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
    lines.append("| UNSPSC — Catálogo SAT PyS | [pys.sat.gob.mx/PyS/catPyS.aspx](http://pys.sat.gob.mx/PyS/catPyS.aspx) |")
    lines.append("| UNSPSC — Estándar internacional | [en.wikipedia.org/wiki/UNSPSC](https://en.wikipedia.org/wiki/UNSPSC) · [undp.org/unspsc](https://www.undp.org/unspsc) |")
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
    if raw_entries:
        lines.append("## Archivos fuente incluidos\n")
        lines.append("Los archivos fuente descargados localmente se exportan en `raw/`.")
        lines.append("")
        grouped_raw: dict[str, list[str]] = {}
        for raw in raw_entries:
            raw_path = Path(raw["path"])
            folder = str(raw_path.parent)
            grouped_raw.setdefault(folder, []).append(raw_path.name)
        for folder in sorted(grouped_raw):
            folder_path = Path(folder)
            label_parts = []
            for part in folder_path.parts[1:]:
                label_parts.append(part.replace("-", " ").replace("_", " ").title())
            label = " / ".join(label_parts) if label_parts else folder
            files = ", ".join(f"`{name}`" for name in sorted(grouped_raw[folder]))
            lines.append(f"- **{label}**: {files}")
        lines.append("")
    return "\n".join(lines)


# ── main ───────────────────────────────────────────────────────────────────────


def generate(
    csv_dir: Path,
    state_files: list[Path],
    output_dir: Path,
    xls_dir: Path | None = None,
    catalog_csv: Path | None = None,
    files_dir: Path | None = None,
) -> int:
    state_rows = _load_state_rows(state_files)
    catalog_index = _load_catalog_index(catalog_csv or CATALOG_CSV)
    source_path_overrides = _build_source_path_override_map(state_rows)
    if not state_rows:
        joined = ", ".join(str(p) for p in state_files)
        print(f"No catalog state found at {joined}. Run scripts/catalogos/extract.py first.", file=sys.stderr)
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
        _entry_catalog_slug(r),
    ))

    entries: list[dict] = []
    raw_entries: list[dict] = []
    copied = 0
    missing_locally = 0

    for row in all_rows:
        folder_version = row["_folder_version"]
        is_versioned   = row["_is_versioned"]
        section_rel    = row["_section_rel"]
        section_field  = row.get("section", "")
        catalogo       = row.get("catalogo", "")

        slug         = _section_slug(section_rel)
        section_path = _section_path(section_rel)
        catalog_slug = _entry_catalog_slug(row)

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
        source_meta = _lookup_source_meta(catalog_index, source_xls)
        latest_val = source_meta.get("latest", "")
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
        entry["catalogo"] = _entry_catalog_name(row)
        entry["source_xls"] = source_meta.get("url", "") or _absolute_source_path(row.get("source_xls", ""))
        entries.append(entry)

    if not entries:
        print(f"No catalogs found in {state_file}.", file=sys.stderr)
        return 1

    if output_dir.exists():
        for child in output_dir.iterdir():
            if child.name == ".git":
                continue
            if child.is_dir():
                shutil.rmtree(child)
            else:
                child.unlink()
    output_dir.mkdir(parents=True, exist_ok=True)

    # Write .huggingface_ignore
    (output_dir / ".huggingface_ignore").write_text(
        ".DS_Store\n__pycache__/\n*.pyc\n*.pyo\n", encoding="utf-8"
    )

    # Write README.md
    meta_dir = output_dir / "metadata"
    meta_dir.mkdir(parents=True, exist_ok=True)

    # Copy locally-available source files into output_dir/raw/ and index them.
    raw_dir = output_dir / "raw"
    if raw_dir.exists():
        shutil.rmtree(raw_dir)
    raw_count = 0
    seen_raw_paths: set[str] = set()

    if xls_dir and xls_dir.exists():
        for src in sorted(xls_dir.rglob("*")):
            if not src.is_file():
                continue
            rel = src.relative_to(xls_dir)
            if not _should_include_raw_file(rel):
                continue
            override_rel = source_path_overrides.get(str(src))
            raw_rel = Path("raw") / (override_rel or rel)
            dest = output_dir / raw_rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
            raw_entries.append(
                {
                    "path": str(raw_rel),
                    "source_path": str(src),
                    "size_bytes": src.stat().st_size,
                    "sha256": _sha256_file(src),
                }
            )
            seen_raw_paths.add(str(raw_rel))
            raw_count += 1

    if files_dir and files_dir.exists() and files_dir != xls_dir:
        for src in sorted(files_dir.rglob("*")):
            if not src.is_file():
                continue
            rel = src.relative_to(files_dir)
            if not _should_include_raw_file(rel):
                continue
            if not _should_include_output_files_rel(rel):
                continue
            raw_rel = Path("raw") / _normalize_output_files_rel(rel)
            if str(raw_rel) in seen_raw_paths:
                continue
            dest = output_dir / raw_rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
            raw_entries.append(
                {
                    "path": str(raw_rel),
                    "source_path": str(src),
                    "size_bytes": src.stat().st_size,
                    "sha256": _sha256_file(src),
                }
            )
            seen_raw_paths.add(str(raw_rel))
            raw_count += 1

    print(f"Raw      → {output_dir}/raw/  ({raw_count} files)", file=sys.stderr)

    readme_path = output_dir / "README.md"
    readme_path.write_text(_build_readme(entries, raw_entries), encoding="utf-8")
    print(
        f"\nREADME   → {readme_path}  ({len(entries)} configs, {copied} copied, {missing_locally} from HF)",
        file=sys.stderr,
    )

    # Write metadata/catalogos.csv
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

    raw_meta_path = meta_dir / "raw_files.csv"
    with raw_meta_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["path", "source_path", "size_bytes", "sha256"],
            extrasaction="ignore",
        )
        writer.writeheader()
        writer.writerows(raw_entries)
    print(f"RawMeta  → {raw_meta_path}  ({len(raw_entries)} rows)", file=sys.stderr)

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
        "--matrix-state-file",
        type=Path,
        default=MATRIX_STATE,
        help=f"Path to matrix state CSV to include in the main dataset (default: {MATRIX_STATE})",
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
    parser.add_argument(
        "--files-dir",
        type=Path,
        default=OUTPUT_FILES_DIR,
        help=f"Root directory of extra source files to include (default: {OUTPUT_FILES_DIR})",
    )
    args = parser.parse_args()
    state_files = [args.state_file]
    if args.matrix_state_file not in state_files:
        state_files.append(args.matrix_state_file)
    return generate(args.csv_dir, state_files, args.output, args.xls_dir, args.catalog_file, args.files_dir)


if __name__ == "__main__":
    raise SystemExit(main())
