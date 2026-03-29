#!/usr/bin/env python3
"""Compare comercio exterior c_FraccionArancelaria against TIGIE raw_consolidated."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import unicodedata
from datetime import UTC, datetime
from io import StringIO
from pathlib import Path

DEFAULT_CATALOG_PATH = Path("hf/csv/complementos/comercio-exterior/2-0/c_FraccionArancelaria_20240513.csv")
DEFAULT_TIGIE_PATH = Path("hf/csv/extra/tigie/tigie.csv")
OUT_DIR = Path("hf/csv/tigie-vs-fraccion-arancelaria")
MANIFEST = Path("output/tigie-vs-fraccion-arancelaria-manifest.json")


def _read_csv(path: Path) -> list[dict]:
    with path.open(encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def _resolve_catalog_path(path: Path) -> Path:
    if path.exists():
        return path
    candidates = sorted(
        Path("hf/csv/complementos/comercio-exterior/2-0").glob("c_FraccionArancelaria*.csv")
    )
    if candidates:
        return candidates[-1]
    dataset_path = Path("hf/dataset/catalogos/complementos/comercio_exterior/2-0/c_fraccion_arancelaria.csv")
    if dataset_path.exists():
        return dataset_path
    return path


def _sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _write_text_if_changed(path: Path, text: str) -> bool:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() and path.read_text(encoding="utf-8") == text:
        return False
    path.write_text(text, encoding="utf-8")
    return True


def _write_csv_if_changed(path: Path, fieldnames: list[str], rows: list[dict]) -> bool:
    buf = StringIO()
    writer = csv.DictWriter(buf, fieldnames=fieldnames, extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return _write_text_if_changed(path, buf.getvalue())


def _normalize_clave(clave: str) -> str:
    return re.sub(r"[^0-9]", "", clave or "")


def _normalize_text(text: str) -> str:
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^\w\s]", "", text.lower())
    return " ".join(text.split())


def _row_hash(row: dict, fields: list[str]) -> str:
    return _sha256_text("|".join(str(row.get(field, "")) for field in fields))


def _catalog_clave(row: dict) -> str:
    return row.get("clave", "") or row.get("fraccion_arancelaria", "")


def _tigie_leaf_rows(rows: list[dict]) -> list[dict]:
    rows_by_clave = {row.get("clave", ""): row for row in rows}

    def build_chain(clave: str) -> list[dict]:
        chain: list[dict] = []
        seen: set[str] = set()
        current = clave
        while current and current not in seen and current in rows_by_clave:
            seen.add(current)
            row = rows_by_clave[current]
            chain.append(
                {
                    "clave": row.get("clave", ""),
                    "nombre": row.get("nombre", ""),
                    "nivel": row.get("nivel", ""),
                }
            )
            current = row.get("parent_clave", "")
        chain.reverse()
        return chain

    out = []
    for row in rows:
        clave = _normalize_clave(row.get("clave", ""))
        if len(clave) != 10:
            continue
        chain = build_chain(row.get("clave", ""))
        item = {
            "clave": clave,
            "clave_tigie": row.get("clave", ""),
            "descripcion_tigie": row.get("nombre", ""),
            "nivel_tigie": row.get("nivel", ""),
            "source_file_tigie": row.get("source_file", ""),
        }
        for idx, level in enumerate(chain, start=1):
            item[f"nivel_{idx}_clave"] = level["clave"]
            item[f"nivel_{idx}_nombre"] = level["nombre"]
        out.append(item)
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--catalog-path", default=str(DEFAULT_CATALOG_PATH))
    parser.add_argument("--tigie-path", default=str(DEFAULT_TIGIE_PATH))
    args = parser.parse_args()

    catalog_path = _resolve_catalog_path(Path(args.catalog_path))
    tigie_path = Path(args.tigie_path)

    catalog_rows = _read_csv(catalog_path)
    tigie_rows = _tigie_leaf_rows(_read_csv(tigie_path))

    catalog_by_clave = {
        _normalize_clave(_catalog_clave(row)): {
            "clave_catalogo": _catalog_clave(row),
            "descripcion_catalogo": row["descripcion"],
            "umt": row.get("umt", ""),
            "fecha_de_inicio_de_vigencia": row.get("fecha_de_inicio_de_vigencia", ""),
            "fecha_de_fin_de_vigencia": row.get("fecha_de_fin_de_vigencia", ""),
        }
        for row in catalog_rows
        if _catalog_clave(row)
    }
    tigie_by_clave = {row["clave"]: row for row in tigie_rows}

    comparison_rows: list[dict] = []
    for clave in sorted(set(catalog_by_clave) | set(tigie_by_clave)):
        cat = catalog_by_clave.get(clave, {})
        tig = tigie_by_clave.get(clave, {})
        row = {
            "clave_normalizada": clave,
            "existe_en_catalogo": str(clave in catalog_by_clave).lower(),
            "existe_en_tigie": str(clave in tigie_by_clave).lower(),
            "clave_catalogo": cat.get("clave_catalogo", ""),
            "clave_tigie": tig.get("clave_tigie", ""),
            "descripcion_catalogo": cat.get("descripcion_catalogo", ""),
            "descripcion_tigie": tig.get("descripcion_tigie", ""),
            "coincide_descripcion_exacta": str(
                bool(cat.get("descripcion_catalogo") and cat.get("descripcion_catalogo") == tig.get("descripcion_tigie"))
            ).lower(),
            "coincide_descripcion_normalizada": str(
                bool(cat.get("descripcion_catalogo"))
                and _normalize_text(cat.get("descripcion_catalogo", "")) == _normalize_text(tig.get("descripcion_tigie", ""))
            ).lower(),
            "umt": cat.get("umt", ""),
            "fecha_de_inicio_de_vigencia": cat.get("fecha_de_inicio_de_vigencia", ""),
            "fecha_de_fin_de_vigencia": cat.get("fecha_de_fin_de_vigencia", ""),
            "source_file_tigie": tig.get("source_file_tigie", ""),
        }
        for idx in range(1, 7):
            row[f"nivel_{idx}_clave"] = tig.get(f"nivel_{idx}_clave", "")
            row[f"nivel_{idx}_nombre"] = tig.get(f"nivel_{idx}_nombre", "")
        row["row_hash"] = _row_hash(
            row,
            [
                "clave_normalizada",
                "existe_en_catalogo",
                "existe_en_tigie",
                "descripcion_catalogo",
                "descripcion_tigie",
                "coincide_descripcion_exacta",
                "coincide_descripcion_normalizada",
                "nivel_1_clave",
                "nivel_2_clave",
                "nivel_3_clave",
                "nivel_4_clave",
                "nivel_5_clave",
                "nivel_6_clave",
            ],
        )
        comparison_rows.append(row)

    mismatches = [
        row
        for row in comparison_rows
        if row["existe_en_catalogo"] != row["existe_en_tigie"]
        or row["coincide_descripcion_normalizada"] != "true"
    ]

    summary = [
        {
            "catalog_path": str(catalog_path),
            "tigie_path": str(tigie_path),
            "total_claves": len(comparison_rows),
            "solo_catalogo": sum(1 for row in comparison_rows if row["existe_en_catalogo"] == "true" and row["existe_en_tigie"] == "false"),
            "solo_tigie": sum(1 for row in comparison_rows if row["existe_en_catalogo"] == "false" and row["existe_en_tigie"] == "true"),
            "en_ambos": sum(1 for row in comparison_rows if row["existe_en_catalogo"] == "true" and row["existe_en_tigie"] == "true"),
            "descripcion_exacta": sum(1 for row in comparison_rows if row["coincide_descripcion_exacta"] == "true"),
            "descripcion_normalizada": sum(1 for row in comparison_rows if row["coincide_descripcion_normalizada"] == "true"),
        }
    ]

    fieldnames = [
        "clave_normalizada",
        "existe_en_catalogo",
        "existe_en_tigie",
        "clave_catalogo",
        "clave_tigie",
        "descripcion_catalogo",
        "descripcion_tigie",
        "coincide_descripcion_exacta",
        "coincide_descripcion_normalizada",
        "umt",
        "fecha_de_inicio_de_vigencia",
        "fecha_de_fin_de_vigencia",
        "source_file_tigie",
        "nivel_1_clave",
        "nivel_1_nombre",
        "nivel_2_clave",
        "nivel_2_nombre",
        "nivel_3_clave",
        "nivel_3_nombre",
        "nivel_4_clave",
        "nivel_4_nombre",
        "nivel_5_clave",
        "nivel_5_nombre",
        "nivel_6_clave",
        "nivel_6_nombre",
        "row_hash",
    ]
    _write_csv_if_changed(OUT_DIR / "comparacion.csv", fieldnames, comparison_rows)
    _write_csv_if_changed(OUT_DIR / "diferencias.csv", fieldnames, mismatches)
    _write_csv_if_changed(
        OUT_DIR / "resumen.csv",
        [
            "catalog_path",
            "tigie_path",
            "total_claves",
            "solo_catalogo",
            "solo_tigie",
            "en_ambos",
            "descripcion_exacta",
            "descripcion_normalizada",
        ],
        summary,
    )

    manifest = {
        "fecha_extraccion": datetime.now(UTC).isoformat(),
        "catalog_path": str(catalog_path),
        "tigie_path": str(tigie_path),
        "summary": summary[0],
        "files": {
            "comparacion": str(OUT_DIR / "comparacion.csv"),
            "diferencias": str(OUT_DIR / "diferencias.csv"),
            "resumen": str(OUT_DIR / "resumen.csv"),
        },
    }
    _write_text_if_changed(MANIFEST, json.dumps(manifest, ensure_ascii=False, indent=2))
    print(json.dumps(summary[0], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
