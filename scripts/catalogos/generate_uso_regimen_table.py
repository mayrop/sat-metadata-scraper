#!/usr/bin/env python3
"""Generate uso-regimen/persona rows derived from c_UsoCFDI.

Reads c_UsoCFDI.csv from the latest version folder under hf/csv/anexo20/factura/
and writes the derived table to hf/derived/anexo20/factura/{version}/c_UsoCFDI_Regimen.csv.
"""
from __future__ import annotations

import argparse
import csv
from pathlib import Path
from typing import Iterable

HF_CSV_DIR     = Path("hf/csv")
HF_DERIVED_DIR = Path("hf/derived")
_FACTURA_DIR   = HF_CSV_DIR / "anexo20" / "factura"


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


def _default_output(input_path: Path) -> Path:
    """Mirror the input version folder under hf/derived/."""
    try:
        rel = input_path.parent.relative_to(HF_CSV_DIR)
        return HF_DERIVED_DIR / rel / "c_UsoCFDI_Regimen.csv"
    except ValueError:
        return HF_DERIVED_DIR / "c_UsoCFDI_Regimen.csv"


def _normalize_bool(value: str | None) -> bool:
    if value is None:
        return False
    normalized = value.strip().lower()
    if normalized in {"si", "sí", "true", "1", "x", "s"}:
        return True
    return False


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
        "--output",
        type=Path,
        default=None,
        help="Where to write the derived CSV (default: hf/derived/anexo20/factura/{version}/c_UsoCFDI_Regimen.csv)",
    )
    args = parser.parse_args()

    output = args.output or _default_output(args.input)

    rows = list(generate_rows(args.input))
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["uso_clave", "tipo_persona", "regimen_fiscal"])
        writer.writerows(rows)
    print(f"Wrote {output} ({len(rows)} rows)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
