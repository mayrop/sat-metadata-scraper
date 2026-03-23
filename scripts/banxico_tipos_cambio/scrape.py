#!/usr/bin/env python3
"""Scrape Banxico daily exchange rates for SAT USD use and reference currencies."""

from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import re
import ssl
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from datetime import UTC, date, datetime, timedelta
from html.parser import HTMLParser
from io import BytesIO, StringIO
from pathlib import Path

RAW_DIR = Path("output/banxico-tipos-cambio/raw")
MANIFEST = Path("output/banxico-tipos-cambio-manifest.json")
HF_DIR = Path("hf/csv/banxico-tipos-cambio")

TIPCAM_URL = "https://www.banxico.org.mx/tipcamb/tipCamIHAction.do"
TIPCAM_SOURCE_URL = "https://www.banxico.org.mx/tipcamb/tipCamMIAction.do?idioma=es"
CF307_FORM_URL = "https://www.banxico.org.mx/SieInternet/consultarDirectorioInternetAction.do?accion=consultarSeries"
CF307_SOURCE_URL = (
    "https://www.banxico.org.mx/SieInternet/consultarDirectorioInternetAction.do"
    "?accion=consultarCuadroAnalitico&idCuadro=CF307"
)

SERIES = [
    ("SF46405", "USD", "Dolar EUA"),
    ("SF46410", "EUR", "Euro"),
    ("SF46406", "JPY", "Yen japones"),
    ("SF46407", "GBP", "Libra esterlina"),
    ("SF290383", "CNY", "Yuan chino"),
    ("SF46411", "DEG", "DEG"),
]

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)


def _ssl_context() -> ssl.SSLContext:
    # Banxico presents a certificate chain that does not validate in this environment.
    return ssl._create_unverified_context()


def _fetch(url: str, *, data: bytes | None = None, headers: dict[str, str] | None = None) -> bytes:
    req = urllib.request.Request(url, data=data, headers=headers or {})
    with urllib.request.urlopen(req, timeout=60, context=_ssl_context()) as response:
        return response.read()


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _write_text_if_changed(path: Path, text: str) -> bool:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() and path.read_text(encoding="utf-8") == text:
        return False
    path.write_text(text, encoding="utf-8")
    return True


def _write_bytes_if_changed(path: Path, data: bytes) -> bool:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() and path.read_bytes() == data:
        return False
    path.write_bytes(data)
    return True


def _write_csv_if_changed(path: Path, fieldnames: list[str], rows: list[dict]) -> bool:
    path.parent.mkdir(parents=True, exist_ok=True)
    buf = StringIO()
    writer = csv.DictWriter(buf, fieldnames=fieldnames, extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return _write_text_if_changed(path, buf.getvalue())


class TipCambParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self._in_data = False
        self._cell_text: list[str] = []
        self.data_cells: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr = dict(attrs)
        classes = set((attr.get("class") or "").split())
        if tag == "td" and ("renglonPar" in classes or "renglonNon" in classes):
            self._in_data = True
            self._cell_text = []

    def handle_endtag(self, tag: str) -> None:
        if tag == "td" and self._in_data:
            self._in_data = False
            value = " ".join("".join(self._cell_text).split())
            self.data_cells.append(html.unescape(value))
            self._cell_text = []

    def handle_data(self, data: str) -> None:
        if self._in_data:
            self._cell_text.append(data)


def _fetch_usd_official_sat_year(year: int) -> tuple[bytes, list[dict]]:
    params = urllib.parse.urlencode(
        {
            "idioma": "sp",
            "fechaInicial": f"01/01/{year}",
            "fechaFinal": f"31/12/{year}",
            "salida": "HTML",
        }
    ).encode()
    raw = _fetch(
        TIPCAM_URL,
        data=params,
        headers={"Content-Type": "application/x-www-form-urlencoded", "User-Agent": USER_AGENT},
    )
    parser = TipCambParser()
    parser.feed(raw.decode("utf-8", errors="replace"))
    cells = [c for c in parser.data_cells if c]
    if len(cells) % 4 != 0:
        raise ValueError(f"Unexpected Banxico tipcamb layout for {year}: {len(cells)} cells")
    n = len(cells) // 4
    fechas = cells[:n]
    determinacion = cells[n : n * 2]
    publicacion = cells[n * 2 : n * 3]
    pagos = cells[n * 3 : n * 4]
    rows: list[dict] = []
    for fecha, fix, dof, para_pagos in zip(fechas, determinacion, publicacion, pagos):
        row = {
            "fecha": datetime.strptime(fecha, "%d/%m/%Y").date().isoformat(),
            "moneda": "USD",
            "fix": fix,
            "publicacion_dof": dof,
            "para_pagos": para_pagos,
            "fuente_url": TIPCAM_SOURCE_URL,
        }
        row["row_hash"] = _sha256_text(
            "|".join(row[k] for k in ["fecha", "moneda", "fix", "publicacion_dof", "para_pagos"])
        )
        rows.append(row)
    rows.sort(key=lambda r: r["fecha"])
    return raw, rows


def _xlsx_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for si in root.findall("a:si", ns):
        strings.append("".join((t.text or "") for t in si.iterfind(".//a:t", ns)))
    return strings


def _xlsx_rows(data: bytes) -> list[list[str]]:
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(BytesIO(data)) as zf:
        sst = _xlsx_shared_strings(zf)
        ws = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        rows: list[list[str]] = []
        for row in ws.findall("a:sheetData/a:row", ns):
            row_values: dict[int, str] = {}
            for cell in row.findall("a:c", ns):
                ref = cell.attrib.get("r", "")
                col_letters = "".join(ch for ch in ref if ch.isalpha())
                col_idx = 0
                for ch in col_letters:
                    col_idx = col_idx * 26 + (ord(ch.upper()) - ord("A") + 1)
                col_idx -= 1
                value = ""
                cell_type = cell.attrib.get("t")
                v = cell.find("a:v", ns)
                if v is not None and v.text is not None:
                    value = v.text
                    if cell_type == "s":
                        value = sst[int(value)]
                row_values[col_idx] = value
            if not row_values:
                rows.append([])
                continue
            width = max(row_values) + 1
            rows.append([row_values.get(i, "") for i in range(width)])
        return rows


def _excel_serial_to_date(value: str) -> str:
    serial = int(float(value))
    return (date(1899, 12, 30) + timedelta(days=serial)).isoformat()


def _fetch_cf307_year(year: int) -> tuple[bytes, list[dict]]:
    params = [
        ("idCuadro", "CF307"),
        ("sector", "6"),
        ("version", "3"),
        ("locale", "es"),
        ("formatoHorizontal", "false"),
        ("metadatosWeb", "true"),
    ]
    for series_id, _, _ in SERIES:
        params.append(("series", series_id))
    params.extend(
        [
            ("anoInicial", str(year)),
            ("anoFinal", str(year)),
            ("tipoInformacion", "4,1"),
            ("formatoXLS.x", "10"),
            ("formatoXLS.y", "10"),
        ]
    )
    raw = _fetch(
        CF307_FORM_URL,
        data=urllib.parse.urlencode(params).encode(),
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Referer": CF307_SOURCE_URL,
            "User-Agent": USER_AGENT,
        },
    )
    out: list[dict] = []
    for row in _xlsx_rows(raw):
        if len(row) < len(SERIES) + 1:
            continue
        if not row[0] or not re.fullmatch(r"\d+(?:\.\d+)?", row[0]):
            continue
        fecha = _excel_serial_to_date(row[0])
        if not fecha.startswith(str(year)):
            continue
        for idx, (series_id, moneda, serie_nombre) in enumerate(SERIES, start=1):
            valor = row[idx].strip()
            if not valor:
                continue
            item = {
                "fecha": fecha,
                "moneda": moneda,
                "valor_mxn": valor,
                "tipo": "referencia",
                "caracter": "informativo",
                "serie_id": series_id,
                "serie_nombre": serie_nombre,
                "fuente_url": CF307_SOURCE_URL,
            }
            item["row_hash"] = _sha256_text(
                "|".join(item[k] for k in ["fecha", "moneda", "valor_mxn", "serie_id"])
            )
            out.append(item)
    out.sort(key=lambda r: (r["fecha"], r["moneda"]))
    return raw, out


def _series_hash(rows: list[dict], fields: list[str]) -> str:
    normalized = "\n".join("|".join(str(row.get(field, "")) for field in fields) for row in rows)
    return _sha256_text(normalized)


def _build_metadata(usd_rows: list[dict], ref_rows: list[dict]) -> list[dict]:
    metadata: list[dict] = []
    if usd_rows:
        metadata.append(
            {
                "grupo": "usd_oficial_sat",
                "moneda": "USD",
                "serie_id": "banxico_tipcamb_usd",
                "serie_nombre": "Tipo de cambio FIX / Publicacion DOF / Para solventar obligaciones",
                "frecuencia": "diaria",
                "caracter": "oficial",
                "fecha_inicio": min(r["fecha"] for r in usd_rows),
                "fecha_fin": max(r["fecha"] for r in usd_rows),
                "fuente_url": TIPCAM_SOURCE_URL,
                "hash": _series_hash(usd_rows, ["fecha", "fix", "publicacion_dof", "para_pagos"]),
            }
        )
    ref_by_series: dict[str, list[dict]] = {}
    for row in ref_rows:
        ref_by_series.setdefault(row["serie_id"], []).append(row)
    for series_id, moneda, serie_nombre in SERIES:
        rows = ref_by_series.get(series_id, [])
        if not rows:
            continue
        metadata.append(
            {
                "grupo": "divisas_referencia",
                "moneda": moneda,
                "serie_id": series_id,
                "serie_nombre": serie_nombre,
                "frecuencia": "diaria",
                "caracter": "informativo",
                "fecha_inicio": min(r["fecha"] for r in rows),
                "fecha_fin": max(r["fecha"] for r in rows),
                "fuente_url": CF307_SOURCE_URL,
                "hash": _series_hash(rows, ["fecha", "valor_mxn"]),
            }
        )
    return metadata


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--since", default="2020-01-01", help="Start date in YYYY-MM-DD format")
    parser.add_argument("--usd-only", action="store_true", help="Only fetch USD official SAT data")
    args = parser.parse_args()

    since = datetime.strptime(args.since, "%Y-%m-%d").date()
    start_year = since.year
    end_year = datetime.now(UTC).date().year

    usd_rows: list[dict] = []
    ref_rows: list[dict] = []
    usd_hashes: dict[int, str] = {}
    ref_hashes: dict[int, str] = {}
    sources: list[dict] = []

    for year in range(start_year, end_year + 1):
        raw_usd, year_usd_rows = _fetch_usd_official_sat_year(year)
        raw_usd_path = RAW_DIR / f"usd_oficial_sat_{year}.html"
        _write_bytes_if_changed(raw_usd_path, raw_usd)
        usd_hashes[year] = _sha256_bytes(raw_usd)
        usd_rows.extend(r for r in year_usd_rows if r["fecha"] >= since.isoformat())
        sources.append({"kind": "usd_oficial_sat", "year": year, "path": str(raw_usd_path), "hash": usd_hashes[year]})

        if args.usd_only:
            print(f"Fetched {year}: {len(year_usd_rows)} USD rows", file=sys.stderr)
            continue

        raw_ref, year_ref_rows = _fetch_cf307_year(year)
        raw_ref_path = RAW_DIR / f"cf307_{year}.xlsx"
        _write_bytes_if_changed(raw_ref_path, raw_ref)
        ref_hashes[year] = _sha256_bytes(raw_ref)
        ref_rows.extend(r for r in year_ref_rows if r["fecha"] >= since.isoformat())
        sources.append({"kind": "cf307", "year": year, "path": str(raw_ref_path), "hash": ref_hashes[year]})
        print(f"Fetched {year}: {len(year_usd_rows)} USD rows, {len(year_ref_rows)} reference rows", file=sys.stderr)

    usd_rows.sort(key=lambda r: r["fecha"])
    ref_rows.sort(key=lambda r: (r["fecha"], r["moneda"]))
    metadata_rows = _build_metadata(usd_rows, ref_rows)

    _write_csv_if_changed(
        HF_DIR / "usd_oficial_sat.csv",
        ["fecha", "moneda", "fix", "publicacion_dof", "para_pagos", "fuente_url", "row_hash"],
        usd_rows,
    )
    if not args.usd_only:
        _write_csv_if_changed(
            HF_DIR / "divisas_referencia.csv",
            ["fecha", "moneda", "valor_mxn", "tipo", "caracter", "serie_id", "serie_nombre", "fuente_url", "row_hash"],
            ref_rows,
        )
    _write_csv_if_changed(
        HF_DIR / "metadata.csv",
        ["grupo", "moneda", "serie_id", "serie_nombre", "frecuencia", "caracter", "fecha_inicio", "fecha_fin", "fuente_url", "hash"],
        metadata_rows,
    )

    manifest = {
        "fecha_extraccion": datetime.now(UTC).isoformat(),
        "url_fuente_usd_oficial_sat": TIPCAM_SOURCE_URL,
        "url_fuente_divisas_referencia": CF307_SOURCE_URL,
        "since": since.isoformat(),
        "sources": sources,
        "files": {
            "usd_oficial_sat": str(HF_DIR / "usd_oficial_sat.csv"),
            "divisas_referencia": str(HF_DIR / "divisas_referencia.csv"),
            "metadata": str(HF_DIR / "metadata.csv"),
        },
        "counts": {
            "usd_oficial_sat": len(usd_rows),
            "divisas_referencia": len(ref_rows),
            "metadata": len(metadata_rows),
        },
    }
    _write_text_if_changed(MANIFEST, json.dumps(manifest, ensure_ascii=False, indent=2))
    print(f"Wrote {len(usd_rows)} USD rows, {len(ref_rows)} reference rows, {len(metadata_rows)} metadata rows", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
