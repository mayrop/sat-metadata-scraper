#!/usr/bin/env python3
"""Build SAT PyS hierarchy (tipo/division/grupo/clase) from pys.sat.gob.mx."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import sys
from datetime import UTC, datetime
from html import unescape
from io import StringIO
from pathlib import Path
from typing import Any

from playwright.sync_api import sync_playwright

HF_DIR = Path("hf/csv/extra/unspsc")
MANIFEST = Path("output/sat-unspsc-manifest.json")
SOURCE_URL = "http://pys.sat.gob.mx/PyS/catPyS.aspx"


def _sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _write_text_if_changed(path: Path, text: str) -> bool:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() and path.read_text(encoding="utf-8") == text:
        return False
    path.write_text(text, encoding="utf-8")
    return True


def _write_csv_if_changed(path: Path, fieldnames: list[str], rows: list[dict[str, Any]]) -> bool:
    buf = StringIO()
    writer = csv.DictWriter(buf, fieldnames=fieldnames, extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return _write_text_if_changed(path, buf.getvalue())


def _row_hash(row: dict[str, Any], fields: list[str]) -> str:
    return _sha256_text("|".join(str(row.get(field, "")) for field in fields))


JS_SETUP = r"""
() => {
  function extractRecord(text, kind, name) {
    const marker = `|${kind}|${name}|`;
    const markerIndex = text.indexOf(marker);
    if (markerIndex === -1) {
      return null;
    }
    const lengthStart = text.lastIndexOf('|', markerIndex - 1) + 1;
    const length = parseInt(text.slice(lengthStart, markerIndex), 10);
    if (Number.isNaN(length)) {
      return null;
    }
    const valueStart = markerIndex + marker.length;
    return text.slice(valueStart, valueStart + length);
  }

  function extractOptions(html) {
    const wrapper = document.createElement('div');
    wrapper.innerHTML = html;
    return Array.from(wrapper.querySelectorAll('option')).map((option) => ({
      value: option.value || '',
      text: option.textContent.trim(),
    }));
  }

  async function postback(state, panelId, eventTarget, fields) {
    const form = new URLSearchParams();
    form.set('myScript', `${panelId}|${eventTarget}`);
    form.set('__LASTFOCUS', '');
    form.set('__EVENTTARGET', eventTarget);
    form.set('__EVENTARGUMENT', '');
    form.set('myTree_ExpandState', '');
    form.set('myTree_SelectedNode', '');
    form.set('myTree_PopulateLog', '');
    form.set('__VIEWSTATE', state.__VIEWSTATE);
    form.set('__EVENTVALIDATION', state.__EVENTVALIDATION);
    for (const [key, value] of Object.entries(fields)) {
      form.set(key, value);
    }
    form.set('__ASYNCPOST', 'true');
    const response = await fetch(window.location.href, {
      method: 'POST',
      headers: {
        Accept: '*/*',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'X-MicrosoftAjax': 'Delta=true',
        'X-Requested-With': 'XMLHttpRequest',
      },
      body: form.toString(),
      credentials: 'include',
    });
    const text = await response.text();
    if (!response.ok) {
      throw new Error(`Unexpected status ${response.status}`);
    }
    const nextState = { ...state };
    const panels = {};
    const options = {};
    for (const hiddenFieldName of ['__VIEWSTATE', '__EVENTVALIDATION']) {
      const hiddenValue = extractRecord(text, 'hiddenField', hiddenFieldName);
      if (hiddenValue !== null) {
        nextState[hiddenFieldName] = hiddenValue;
      }
    }
    for (const updatePanelName of ['pnlTipo', 'pnlSegmento', 'pnlFamilia', 'pnlClase']) {
      const panelValue = extractRecord(text, 'updatePanel', updatePanelName);
      if (panelValue !== null) {
        panels[updatePanelName] = panelValue;
        options[updatePanelName] = extractOptions(panelValue);
      }
    }
    return { state: nextState, panels, options };
  }

  function initialState() {
    return {
      __VIEWSTATE: document.getElementById('__VIEWSTATE').value,
      __EVENTVALIDATION: document.getElementById('__EVENTVALIDATION').value,
    };
  }

  function initialTipoOptions() {
    return extractOptions(document.getElementById('pnlTipo').innerHTML);
  }

  async function scrapeAll() {
    let state = initialState();
    const tipos = [];
    const divisiones = [];
    const grupos = [];
    const clases = [];

    for (const tipoOption of initialTipoOptions().filter((option) => option.value && option.value !== '0')) {
      const tipo = tipoOption.text.replace(/s$/, '');
      tipos.push({ tipo });

      const tipoResult = await postback(state, 'pnlTipo', 'cmbTipo', {
        cmbTipo: tipoOption.value,
        txtBuscar: '',
      });
      state = tipoResult.state;
      const divisionesOptions = (tipoResult.options.pnlSegmento || []).filter((option) => option.value && option.value !== '0');

      for (const divisionOption of divisionesOptions) {
        divisiones.push({
          tipo,
          codigo_division: divisionOption.value,
          nombre_division: divisionOption.text,
        });

        const divisionResult = await postback(state, 'pnlSegmento', 'cmbSegmento', {
          cmbTipo: tipoOption.value,
          cmbSegmento: divisionOption.value,
          txtBuscar: '',
        });
        state = divisionResult.state;
        const gruposOptions = (divisionResult.options.pnlFamilia || []).filter((option) => option.value && option.value !== '0');

        for (const grupoOption of gruposOptions) {
          grupos.push({
            tipo,
            codigo_division: divisionOption.value,
            nombre_division: divisionOption.text,
            codigo_grupo: grupoOption.value,
            nombre_grupo: grupoOption.text,
          });

          const grupoResult = await postback(state, 'pnlFamilia', 'cmbFamilia', {
            cmbTipo: tipoOption.value,
            cmbSegmento: divisionOption.value,
            cmbFamilia: grupoOption.value,
            txtBuscar: '',
          });
          state = grupoResult.state;
          const clasesOptions = (grupoResult.options.pnlClase || []).filter((option) => option.value && option.value !== '0');

          for (const claseOption of clasesOptions) {
            clases.push({
              tipo,
              codigo_division: divisionOption.value,
              nombre_division: divisionOption.text,
              codigo_grupo: grupoOption.value,
              nombre_grupo: grupoOption.text,
              codigo_clase: claseOption.value,
              nombre_clase: claseOption.text,
            });
          }
        }
      }
    }

    return { tipos, divisiones, grupos, clases };
  }

  window.__satPys = { scrapeAll };
}
"""


def _clean_options(options: list[dict[str, str]]) -> list[dict[str, str]]:
    return [
        {"value": option["value"].strip(), "text": unescape(option["text"]).strip()}
        for option in options
        if option["value"].strip() and option["value"].strip() != "0" and option["text"].strip()
    ]


def build() -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], dict[str, int]]:
    tipos: list[dict[str, Any]] = []
    divisiones: list[dict[str, Any]] = []
    grupos: list[dict[str, Any]] = []
    clases: list[dict[str, Any]] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(SOURCE_URL, wait_until="load", timeout=120000)
        page.evaluate(JS_SETUP)
        data = page.evaluate("() => window.__satPys.scrapeAll()")
        tipos = data["tipos"]
        divisiones = data["divisiones"]
        grupos = data["grupos"]
        clases = data["clases"]

        browser.close()

    for tipo_row in tipos:
        tipo_row["row_hash"] = _row_hash(tipo_row, ["tipo"])
    for division_row in divisiones:
        division_row["row_hash"] = _row_hash(
            division_row,
            ["tipo", "codigo_division", "nombre_division"],
        )
    for grupo_row in grupos:
        grupo_row["row_hash"] = _row_hash(
            grupo_row,
            ["tipo", "codigo_division", "nombre_division", "codigo_grupo", "nombre_grupo"],
        )
    for clase_row in clases:
        clase_row["row_hash"] = _row_hash(
            clase_row,
            [
                "tipo",
                "codigo_division",
                "nombre_division",
                "codigo_grupo",
                "nombre_grupo",
                "codigo_clase",
                "nombre_clase",
            ],
        )

    counts = {
        "tipos": len(tipos),
        "divisiones": len(divisiones),
        "grupos": len(grupos),
        "clases": len(clases),
    }
    return tipos, divisiones, grupos, clases, counts


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--verify", action="store_true", help="Placeholder for CLI symmetry.")
    args = parser.parse_args()
    _ = args.verify

    tipos, divisiones, grupos, clases, counts = build()

    _write_csv_if_changed(HF_DIR / "tipos.csv", ["tipo", "row_hash"], tipos)
    _write_csv_if_changed(
        HF_DIR / "divisiones.csv",
        ["tipo", "codigo_division", "nombre_division", "row_hash"],
        divisiones,
    )
    _write_csv_if_changed(
        HF_DIR / "grupos.csv",
        ["tipo", "codigo_division", "nombre_division", "codigo_grupo", "nombre_grupo", "row_hash"],
        grupos,
    )
    _write_csv_if_changed(
        HF_DIR / "clases.csv",
        [
            "tipo",
            "codigo_division",
            "nombre_division",
            "codigo_grupo",
            "nombre_grupo",
            "codigo_clase",
            "nombre_clase",
            "row_hash",
        ],
        clases,
    )

    metadata = [
        {
            "dataset": "sat_pys_hierarchy",
            "source_url": SOURCE_URL,
            **counts,
            "hash": _sha256_text("".join(row["row_hash"] for row in clases)),
        }
    ]
    _write_csv_if_changed(
        HF_DIR / "metadata.csv",
        ["dataset", "source_url", "tipos", "divisiones", "grupos", "clases", "hash"],
        metadata,
    )

    manifest = {
        "fecha_extraccion": datetime.now(UTC).isoformat(),
        "source_url": SOURCE_URL,
        "counts": counts,
        "files": {
            "tipos": str(HF_DIR / "tipos.csv"),
            "divisiones": str(HF_DIR / "divisiones.csv"),
            "grupos": str(HF_DIR / "grupos.csv"),
            "clases": str(HF_DIR / "clases.csv"),
            "metadata": str(HF_DIR / "metadata.csv"),
        },
    }
    _write_text_if_changed(MANIFEST, json.dumps(manifest, ensure_ascii=False, indent=2))
    print(
        f"Wrote {counts['tipos']} tipos, {counts['divisiones']} divisiones, {counts['grupos']} grupos, {counts['clases']} clases",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
