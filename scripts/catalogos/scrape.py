#!/usr/bin/env python3
"""SAT CFDI Complementos scraper.

Scrapes the "Complementos" section of:
  https://www.sat.gob.mx/portal/public/tramites/complementos-de-factura

For each complemento, extracts estándar URL, XSD, XSLT, and last-modified dates.
For complementos with a dedicated detail page, also extracts per-version data.

Writes a dated manifest.json + catalog.csv under output/YYYY-MM-DD/ only when
content has changed since the last run. Downloads all referenced files.

Usage:
    uv run scripts/catalogos/scrape.py
    uv run scripts/catalogos/scrape.py --debug          # saves page HTML + intermediate JSON
    uv run scripts/catalogos/scrape.py --show-browser   # show the browser window
    uv run scripts/catalogos/scrape.py --dry-run        # scrape only, don't write files
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import html.parser
import json
import re
import sys
import urllib.request
from datetime import datetime
from pathlib import Path
from urllib.error import HTTPError, URLError
from dataclasses import dataclass, field
from urllib.parse import urljoin

from playwright.sync_api import Page, TimeoutError as PWTimeout, sync_playwright

try:
    import xlrd as _xlrd
except ImportError:
    _xlrd = None

MAIN_URL = "https://www.sat.gob.mx/portal/public/tramites/complementos-de-factura"
BASE_URL = "https://www.sat.gob.mx"
# (tab label, folder/category slug)
SECTIONS = [
    ("Complementos",                                       "complementos"),
    ("Complementos concepto",                              "complementos-concepto"),
    ("Complementos de retenciones e información de pagos", "complementos-retenciones"),
]
OUTPUT_DIR = Path("output")
HF_XLS_DIR = Path("hf/xls")
NAV_TIMEOUT = 30_000
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# Extensions that are direct file downloads (not HTML pages to follow)
DIRECT_EXTS = {".xls", ".xlsx", ".zip", ".pdf", ".xsd", ".xslt", ".xml", ".rar"}

# Static HTML pages to scrape alongside the complementos
STATIC_PAGES = [
    (
        "Formato de factura (Anexo 20)",
        "http://omawww.sat.gob.mx/tramitesyservicios/Paginas/anexo_20.htm",
    ),
    (
        "Factura de retenciones e información de pagos (Anexo 20)",
        "http://omawww.sat.gob.mx/tramitesyservicios/Paginas/CFDI_retenciones.htm",
    ),
]

_STATIC_FILE_EXTS = {".xsd", ".xslt", ".xls", ".xlsx", ".pdf", ".xml", ".zip"}


# ── helpers ────────────────────────────────────────────────────────────────────


def _parse_date(raw: str | None) -> str | None:
    """Convert DD/MM/YYYY → YYYY-MM-DD. Returns None if unparseable."""
    if not raw:
        return None
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", raw.strip())
    return f"{m.group(3)}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}" if m else raw


def _ver_slug(ver: dict) -> str:
    """Build a directory-safe slug from a version dict's version + revision + sub."""
    parts = []
    if ver.get("version"):
        parts += ["version", ver["version"]]
    if ver.get("revision"):
        parts += ["revision", ver["revision"]]
    if ver.get("sub"):
        parts.append(ver["sub"])
    return slugify(" ".join(parts)) if parts else "files"


def _hf_ver_slug(ver: dict) -> str:
    """Like _ver_slug but without the 'version'/'revision' words — for filenames."""
    parts = []
    if ver.get("version"):
        parts.append(slugify(ver["version"]))
    if ver.get("revision"):
        parts.append(slugify(ver["revision"]))
    if ver.get("sub"):
        parts.append(ver["sub"])
    return "-".join(p for p in parts if p) or "files"


def slugify(text: str) -> str:
    for src, dst in [
        ("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"), ("ñ", "n"),
        ("Á", "a"), ("É", "e"), ("Í", "i"), ("Ó", "o"), ("Ú", "u"), ("Ñ", "n"),
    ]:
        text = text.replace(src, dst)
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")


def _clean_stem(stem: str) -> str:
    """Strip version numbers, dates, and hashes from a filename stem.

    Examples:
        catCFDI_V_4_20260313      → catCFDI
        Catalogos_Carta_Porte31_2db736cf0c → Catalogos_Carta_Porte
        cat_Pagos_8ca5655de2      → cat_Pagos
        c_CodigoPostal_V_4_20231219 → c_CodigoPostal
        c_FraccionArancelaria_v17_rA → c_FraccionArancelaria
        c_INCOTERM20              → c_INCOTERM
        catCFDI_Retenciones_1     → catCFDI_Retenciones

    The original name is preserved in the manifest's 'url' field.
    """
    s = stem
    s = re.sub(r"_[Vv]_?\d+[\d.]*", "", s)          # _V_4, _V_33, _v17
    s = re.sub(r"_\d{6,8}(?=_|$)", "", s)             # _20260313, _31032023
    s = re.sub(r"\d*_[0-9a-f]{8,}$", "", s, flags=re.IGNORECASE)  # 31_2db736cf0c, _8ca5655de2
    s = re.sub(r"_[rR][A-Za-z0-9]{1,2}$", "", s)       # _rA, _rB short revision suffixes
    s = re.sub(r"_\d+$", "", s)                        # _1, _2 trailing counters
    s = re.sub(r"(?<=[A-Za-z])\d{1,4}$", "", s)       # version-style suffix digits e.g. INCOTERM20, CFDI40
    s = s.strip("_")
    return s or stem  # fallback to original if everything was stripped


def abs_url(href: str) -> str:
    if not href:
        return ""
    return href if href.startswith("http") else urljoin(BASE_URL, href)


def _filename(url: str) -> str:
    """Return a safe filename for a URL.

    Satellite blob URLs all resolve to 'Satellite' — we append the blobwhere
    param as a unique suffix and add .pdf (SAT's blob server serves PDFs).
    """
    path_part = url.split("?")[0]
    name = Path(path_part).name or "file"
    if name.lower() == "satellite":
        m = re.search(r"blobwhere=(\d+)", url)
        suffix = m.group(1) if m else re.sub(r"[^a-z0-9]", "", url[-12:])
        name = f"Satellite_{suffix}.pdf"
    return name


def _is_html_url(url: str) -> bool:
    """True if the URL looks like an HTML page rather than a direct download."""
    if "blobwhere=" in url:
        return False  # Satellite blob URLs are binary files
    low = url.lower().split("?")[0]
    ext = Path(low).suffix
    return ext not in DIRECT_EXTS


def _resolve_catalog_url(url: str) -> list[dict]:
    """If url is an HTML page, fetch it and return version-grouped download links.

    Returns list of {"version": str | None, "files": list[str]}.
    Headings with "versión X.Y" in the HTML split files into sub-groups.
    Non-HTML URLs return [{"version": None, "files": [url]}].
    """
    if not _is_html_url(url):
        return [{"version": None, "files": [url]}]
    try:
        req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        with urllib.request.urlopen(req, timeout=20) as r:
            html = r.read().decode("utf-8", errors="replace")

        base_url = url.rsplit("/", 1)[0] + "/"
        file_ext_re = re.compile(r"\.(xls|xlsx|zip|pdf|xsd|xslt|xml)$", re.IGNORECASE)
        version_re = re.compile(r"versi[oó]n\s*([\d][.\d]*(?:\s+[Rr]evisión\s+\w+)?)", re.IGNORECASE)

        # Walk headings and anchors in document order using a proper HTML parser
        # so that mismatched or nested tags don't produce wrong version groupings.
        class _CatalogParser(html.parser.HTMLParser):
            def __init__(self) -> None:
                super().__init__()
                self.groups: list[dict] = []
                self._cur_version: str | None = None
                self._cur_files: list[str] = []
                self._in_heading = False
                self._heading_text: list[str] = []

            def handle_starttag(self, tag: str, attrs: list) -> None:
                if tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
                    self._in_heading = True
                    self._heading_text = []
                elif tag == "a":
                    href = dict(attrs).get("href", "") or ""
                    path_part = href.split("?")[0]
                    if file_ext_re.search(path_part):
                        full = href if href.startswith("http") else urljoin(base_url, href)
                        self._cur_files.append(full)

            def handle_endtag(self, tag: str) -> None:
                if tag in {"h1", "h2", "h3", "h4", "h5", "h6"} and self._in_heading:
                    self._in_heading = False
                    text = "".join(self._heading_text).strip()
                    vm = version_re.search(text)
                    if vm:
                        if self._cur_files or self._cur_version is not None:
                            self.groups.append({"version": self._cur_version, "files": self._cur_files})
                        self._cur_version = f"versión {vm.group(1)}"
                        self._cur_files = []

            def handle_data(self, data: str) -> None:
                if self._in_heading:
                    self._heading_text.append(data)

            def close(self) -> None:
                super().close()
                if self._cur_files:
                    self.groups.append({"version": self._cur_version, "files": self._cur_files})

        parser = _CatalogParser()
        parser.feed(html)
        parser.close()
        groups = parser.groups

        return groups if groups else [{"version": None, "files": [url]}]
    except Exception as exc:
        print(f"    WARN: could not resolve catalog page {url}: {exc}", file=sys.stderr)
        return [{"version": None, "files": [url]}]


def _sha256(path: Path) -> str | None:
    if not path.exists():
        return None
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _fetch_bytes(url: str) -> bytes | None:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.read()
    except (URLError, HTTPError, Exception) as exc:
        print(f"    WARN: could not download {url}: {exc}", file=sys.stderr)
        return None


@dataclass
class _PrevState:
    """Previous-manifest state for one catalog entry — used to skip unchanged downloads."""
    cat_fp: str | None = None
    files: list = field(default_factory=list)
    lm: dict[str, tuple[str | None, str]] = field(default_factory=dict)  # {url: (fingerprint, hash)}


def download(
    url: str,
    dest: Path,
    fingerprint: str | None = None,
    stored_fingerprint: str | None = None,
    stored_hash: str | None = None,
    force: bool = False,
    verify: bool = False,
) -> str:
    """Download url to dest, replacing only if content changed.

    Returns one of:
      "written"   — file fetched and saved
      "unchanged" — file fetched, hash identical, not written
      "skipped"   — fingerprint matched, fetch skipped entirely
      "error"     — fetch failed
    """
    dest.parent.mkdir(parents=True, exist_ok=True)
    # Fingerprint fast-path: skip fetch entirely when version/date unchanged.
    if not force and not verify:
        if fingerprint and stored_fingerprint and fingerprint == stored_fingerprint:
            if dest.exists() or stored_hash:
                return "skipped"
    data = _fetch_bytes(url)
    if data is None:
        return "error"
    new_hash = hashlib.sha256(data).hexdigest()
    if not force:
        prior_hash = stored_hash or _sha256(dest)
        if new_hash == prior_hash and (verify or dest.exists()):
            return "unchanged"
    dest.write_bytes(data)
    return "written"


def latest_manifest() -> dict | None:
    m = OUTPUT_DIR / "catalogos-manifest.json"
    if m.exists():
        return json.loads(m.read_text(encoding="utf-8"))
    return None


def _scrape_summary(complementos: list[dict]) -> list:
    """Extract version-identity keys only — used for change detection."""
    return [
        {
            "name": c["name"],
            "detail_url": c.get("detail_url"),
            "versions": [
                {"version": v.get("version"), "revision": v.get("revision")}
                for v in c.get("versions", [])
            ],
        }
        for c in complementos
    ]


def data_changed(complementos: list[dict]) -> bool:
    prev = latest_manifest()
    if prev is None:
        return True
    normalize = lambda d: json.dumps(d, sort_keys=True, ensure_ascii=False)
    return normalize(_scrape_summary(prev.get("complementos", []))) != normalize(
        _scrape_summary(complementos)
    )


# ── scraping ───────────────────────────────────────────────────────────────────


def _dismiss_modal(page: Page) -> None:
    try:
        btn = page.query_selector("button.close-modal")
        if btn:
            btn.click()
            page.wait_for_timeout(300)
    except Exception:
        pass


def _click_tab(page: Page, label: str) -> bool:
    found = page.evaluate(f"""() => {{
        const spans = Array.from(document.querySelectorAll(".tab-label span"));
        const tab = spans.find(s => s.textContent.trim().toLowerCase() === "{label.lower()}");
        if (!tab) return false;
        (tab.closest(".tab-container") || tab.parentElement).click();
        return true;
    }}""")
    if found:
        try:
            page.wait_for_selector(".tab-content", timeout=5_000)
        except Exception:
            page.wait_for_timeout(500)
        _dismiss_modal(page)
    return found


# JS helper embedded in page.evaluate calls — extracts last-modified date
# from the <p> sibling inside the same <li> as an anchor.
_LAST_MOD_JS = r"""
function lastModified(a) {
    let li = a;
    while (li && li.tagName !== 'LI') li = li.parentElement;
    if (!li) return null;
    for (const p of li.querySelectorAll('p')) {
        const m = p.textContent.match(/[ÚU]ltima modificaci[oó]n el\s*([\d\/]+)/);
        if (m) return m[1];
    }
    return null;
}
"""


def scrape_detail_page(page: Page, url: str, debug_dir: Path | None) -> list[dict]:
    """Scrape a complemento detail page → list of version dicts."""
    print(f"    -> {url}", file=sys.stderr)
    page.goto(url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
    try:
        page.wait_for_selector(".tab-content, .tab-label, .information-procedures", timeout=10_000)
    except Exception:
        pass

    clicked = _click_tab(page, "Información especializada")
    if not clicked:
        page.evaluate("""() => {
            const first = document.querySelector(".tab-container");
            if (first) first.click();
        }""")
        page.wait_for_timeout(1500)
        _dismiss_modal(page)

    if debug_dir:
        slug = url.rstrip("/").split("/")[-1]
        (debug_dir / f"detail_{slug}.html").write_text(page.content(), encoding="utf-8")

    # Version/revision come from the tab content text (shared across all panels)
    ver_raw = page.evaluate(rf"""() => {{
        const content = document.querySelector(".tab-content");
        if (!content) return null;
        const desc = content.textContent || "";
        const vm = desc.match(/[Vv]ersi[oó]n\s*([\d][\d.]*)(?:\s+[Rr]evisión\s+(\w+))?/);
        return vm ? {{ version: vm[1], revision: vm[2] || null }} : {{ version: null, revision: null }};
    }}""")

    if ver_raw is None:
        return []

    # Check for accordion panels (e.g. Hidrocarburos has multiple sub-complementos)
    btn_count = page.evaluate(
        "() => document.querySelectorAll('button.btn-modality').length"
    )

    _EXTRACT_JS = f"""(panelIdx) => {{
        {_LAST_MOD_JS}
        const container = panelIdx === null
            ? document.querySelector(".tab-content")
            : document.querySelectorAll(".information-procedures")[panelIdx];
        if (!container) return null;
        const entry = {{ estandar: null, xsd: null, xslt: null, catalogos: [] }};
        for (const a of container.querySelectorAll("a[href]")) {{
            const href = a.getAttribute("href") || "";
            const text = (a.textContent || "").toLowerCase().trim();
            const hlow = href.toLowerCase();
            const lm = lastModified(a);
            if (text.includes("catálogo") || text.includes("catalogo"))
                entry.catalogos.push({{ name: a.textContent.trim(), url: href, last_modified: lm }});
            else if (text.includes("estándar") || text.includes("estandar"))
                entry.estandar = {{ url: href, last_modified: lm }};
            else if (hlow.includes(".xslt"))
                entry.xslt = {{ url: href, last_modified: lm }};
            else if (hlow.includes(".xsd") && !entry.xsd)
                entry.xsd = {{ url: href, last_modified: lm }};
        }}
        return entry;
    }}"""

    def _finfo(raw_entry: dict | None) -> dict:
        if not raw_entry:
            return {"url": None, "last_modified": None, "local_file": None, "hash": None}
        return {
            "url": abs_url(raw_entry.get("url", "") or "") or None,
            "last_modified": _parse_date(raw_entry.get("last_modified")),
            "local_file": None,
            "hash": None,
        }

    def _build_version(e: dict, sub: str | None = None) -> dict:
        v = {
            "version": ver_raw["version"],
            "revision": ver_raw["revision"],
            "files": {
                "estandar": _finfo(e.get("estandar")),
                "xsd":      _finfo(e.get("xsd")),
                "xslt":     _finfo(e.get("xslt")),
            },
            "catalogos": [
                {
                    "name": c["name"],
                    "url": abs_url(c["url"]),
                    "last_modified": _parse_date(c.get("last_modified")),
                    "files": [],
                }
                for c in e["catalogos"]
            ],
        }
        if sub:
            v["sub"] = sub
        return v

    if btn_count > 0:
        versions = []
        for i in range(btn_count):
            page.evaluate(
                f"() => {{ document.querySelectorAll('button.btn-modality')[{i}].click(); }}"
            )
            try:
                page.wait_for_selector(".information-procedures", timeout=3_000)
            except Exception:
                page.wait_for_timeout(300)
            panel_title = page.evaluate(f"""() => {{
                const btn = document.querySelectorAll('button.btn-modality')[{i}];
                return btn?.querySelector('span.title-portal')?.textContent?.trim() || '';
            }}""")
            e = page.evaluate(_EXTRACT_JS, i)
            if e:
                # Extract meaningful short name: "Complemento para Gastos..." → "gastos"
                m = re.search(r"para\s+(\w+)", panel_title, re.IGNORECASE)
                sub = slugify(m.group(1)) if m else str(i)
                versions.append(_build_version(e, sub))
        raw_for_debug = versions
    else:
        e = page.evaluate(_EXTRACT_JS, None)
        raw_for_debug = e
        versions = [_build_version(e)] if e else []

    if debug_dir:
        slug = url.rstrip("/").split("/")[-1]
        (debug_dir / f"detail_{slug}_versions.json").write_text(
            json.dumps(raw_for_debug, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    return versions


def scrape_static_page(name: str, url: str) -> dict:
    """Scrape a static SAT HTML technical page into a complemento-style dict.

    Parses version sections delimited by "Documentación técnica - versión X.Y"
    heading paragraphs, plus a top catalog table that may appear before them.
    """
    print(f"  Fetching static page: {url}", file=sys.stderr)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        with urllib.request.urlopen(req, timeout=30) as r:
            html = r.read().decode("utf-8", errors="replace")
    except Exception as exc:
        print(f"  WARN: could not fetch {url}: {exc}", file=sys.stderr)
        return {"name": name, "detail_url": url, "versions": []}

    base_url = url.rsplit("/", 1)[0] + "/"
    version_re = re.compile(r"versi[oó]n\s*([\d][.\d]*)", re.IGNORECASE)
    date_re = re.compile(r"\b(\d{1,2}/\d{1,2}/\d{4})\b")

    # Locate version section boundaries: <p> tags containing "documentaci…versión X.Y"
    boundaries: list[tuple[int, str]] = []
    for m in re.finditer(r"<p\b[^>]*>(.*?)</p>", html, re.IGNORECASE | re.DOTALL):
        text = re.sub(r"<[^>]+>", "", m.group(1)).strip()
        if "documentaci" not in text.lower():
            continue
        vm = version_re.search(text)
        if vm:
            boundaries.append((m.end(), vm.group(1)))

    if not boundaries:
        # Single version, detect from page body
        vm = version_re.search(re.sub(r"<[^>]+>", " ", html))
        boundaries = [(0, vm.group(1) if vm else None)]

    # Build per-version buckets
    versions_data: dict[str | None, dict] = {}
    for _, ver in boundaries:
        if ver not in versions_data:
            versions_data[ver] = {"files": {}, "catalogos": []}

    def _process_rows(section_html: str, target_ver: str | None) -> None:
        if target_ver not in versions_data:
            return
        vd = versions_data[target_ver]
        for tr_m in re.finditer(r"<tr\b[^>]*>(.*?)</tr>", section_html, re.IGNORECASE | re.DOTALL):
            row_html = tr_m.group(1)
            a_m = re.search(
                r'<a\b[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>',
                row_html, re.IGNORECASE | re.DOTALL,
            )
            if not a_m:
                continue
            href = a_m.group(1).strip()
            link_text = re.sub(r"<[^>]+>", "", a_m.group(2)).strip()

            # Only direct downloadable files
            href_path = href.lower().split("?")[0]
            if Path(href_path).suffix not in _STATIC_FILE_EXTS:
                continue
            if href.startswith("javascript:") or href.startswith("#"):
                continue

            full_url = href if href.startswith("http") else urljoin(base_url, href)
            row_text = re.sub(r"<[^>]+>", " ", row_html)
            date_m = date_re.search(row_text)
            date_str = _parse_date(date_m.group(1)) if date_m else None

            result = _classify_static_link(link_text, href)
            if result is None:
                continue
            cat, type_key = result

            if cat == "files":
                if type_key not in vd["files"]:
                    vd["files"][type_key] = {
                        "url": full_url, "last_modified": date_str,
                        "local_file": None, "hash": None,
                    }
            else:
                if not any(c["url"] == full_url for c in vd["catalogos"]):
                    vd["catalogos"].append({
                        "name": link_text, "url": full_url,
                        "last_modified": date_str, "files": [],
                    })

    # Process each version section
    for i, (start, ver) in enumerate(boundaries):
        end = boundaries[i + 1][0] if i + 1 < len(boundaries) else len(html)
        _process_rows(html[start:end], ver)

    # Process the top section (before first boundary); infer version from <th> text
    top_end = boundaries[0][0] if boundaries else len(html)
    for tr_m in re.finditer(r"<tr\b[^>]*>(.*?)</tr>", html[:top_end], re.IGNORECASE | re.DOTALL):
        row_html = tr_m.group(1)
        a_m = re.search(
            r'<a\b[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>',
            row_html, re.IGNORECASE | re.DOTALL,
        )
        if not a_m:
            continue
        href = a_m.group(1).strip()
        link_text = re.sub(r"<[^>]+>", "", a_m.group(2)).strip()

        href_path = href.lower().split("?")[0]
        if Path(href_path).suffix not in _STATIC_FILE_EXTS:
            continue

        full_url = href if href.startswith("http") else urljoin(base_url, href)
        row_text = re.sub(r"<[^>]+>", " ", row_html)
        date_m = date_re.search(row_text)
        date_str = _parse_date(date_m.group(1)) if date_m else None

        # Determine target version from <th> cell, fall back to first boundary version
        th_m = re.search(r"<th\b[^>]*>(.*?)</th>", row_html, re.IGNORECASE | re.DOTALL)
        row_ver = None
        if th_m:
            vm = version_re.search(re.sub(r"<[^>]+>", "", th_m.group(1)))
            if vm:
                row_ver = vm.group(1)
        target_ver = row_ver if row_ver in versions_data else (boundaries[0][1] if boundaries else None)

        result = _classify_static_link(link_text, href)
        if result is None or target_ver not in versions_data:
            continue
        cat, type_key = result
        vd = versions_data[target_ver]

        if cat == "files":
            if type_key not in vd["files"]:
                vd["files"][type_key] = {
                    "url": full_url, "last_modified": date_str,
                    "local_file": None, "hash": None,
                }
        else:
            if not any(c["url"] == full_url for c in vd["catalogos"]):
                vd["catalogos"].append({
                    "name": link_text, "url": full_url,
                    "last_modified": date_str, "files": [],
                })

    versions = [
        {"version": ver, "revision": None, "files": vd["files"], "catalogos": vd["catalogos"]}
        for ver, vd in versions_data.items()
        if vd["files"] or vd["catalogos"]
    ]
    print(f"    {name}: {len(versions)} version(s)", file=sys.stderr)
    return {"name": name, "detail_url": url, "versions": versions, "_section": "anexo-20"}


_EXTRACT_MODALS_JS = f"""() => {{
    const BASE = "{BASE_URL}";
    const results = [];
    for (const modal of document.querySelectorAll("dialog.pacs-modal")) {{
        const name = (modal.querySelector("h1.title-portal")?.textContent || "").trim();
        if (!name) continue;
        let card = modal.nextElementSibling;
        while (card && !card.classList.contains("card-body"))
            card = card.nextElementSibling;
        const detailAnchor = card?.querySelector("a.link-procedures");
        let detail_url = null;
        if (detailAnchor) {{
            const href = detailAnchor.getAttribute("href") || "";
            if (href.startsWith("http")) detail_url = href;
            else if (href.startsWith("/")) detail_url = BASE + href;
            else detail_url = BASE + "/portal/public/tramites/" + href.replace(/^\\.\\//,"");
        }}
        const info = modal.querySelector(".info-modal");
        const infoText = info?.textContent || "";
        const vm = infoText.match(/[Vv]ersi[oó]n\s*([\d][.\d]*)(?:\s+[Rr]evisión\s+(\w+))?/);
        const entry = {{ name, detail_url,
                         version: vm ? vm[1] : null, revision: vm ? (vm[2] || null) : null,
                         estandar_url: null, xsd_url: null, xslt_url: null, catalogos: [] }};
        for (const a of (info?.querySelectorAll("a[href]") || [])) {{
            const href = a.getAttribute("href") || "";
            const text = (a.textContent || "").toLowerCase().trim();
            const hlow = href.toLowerCase();
            if (text.includes("catálogo") || text.includes("catalogo"))
                entry.catalogos.push({{ name: a.textContent.trim(), url: href, last_modified: null }});
            else if (text.includes("estándar") || text.includes("estandar"))
                entry.estandar_url = href;
            else if (hlow.includes(".xslt"))
                entry.xslt_url = href;
            else if (hlow.includes(".xsd"))
                entry.xsd_url = entry.xsd_url || href;
        }}
        results.push(entry);
    }}
    return results;
}}"""


def scrape_complementos_section(
    page: Page, debug_dir: Path | None, prev_comps: dict[str, dict] | None = None, force: bool = False
) -> list[dict]:
    """Scrape all sections of the main page → list of complemento dicts.

    prev_comps — {name: comp_dict} from previous manifest; detail pages whose
    name+detail_url already appear there are reused without re-navigating.
    """
    print(f"Loading {MAIN_URL}", file=sys.stderr)
    page.goto(MAIN_URL, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
    # Wait for the tab UI to render instead of waiting for all network activity
    try:
        page.wait_for_selector(".tab-label", timeout=15_000)
    except Exception:
        pass  # proceed anyway; JS extraction will report missing tabs

    def _finfo_url(url: str | None) -> dict:
        return {"url": abs_url(url or "") or None, "last_modified": None, "local_file": None, "hash": None}

    complementos: list[dict] = []

    # Extract all sections before navigating to any detail page
    for section_label, section_slug in SECTIONS:
        clicked = page.evaluate(f"""() => {{
            const spans = Array.from(document.querySelectorAll(".tab-label span"));
            const tab = spans.find(s => s.textContent.trim() === "{section_label}");
            if (!tab) return false;
            (tab.closest(".tab-container") || tab.parentElement).click();
            return true;
        }}""")
        if not clicked:
            print(f"  WARN: tab '{section_label}' not found", file=sys.stderr)
            continue

        try:
            page.wait_for_function("() => document.querySelectorAll('dialog.pacs-modal').length > 0", timeout=5_000)
        except Exception:
            page.wait_for_timeout(500)  # fallback brief wait
        _dismiss_modal(page)

        raw = page.evaluate(_EXTRACT_MODALS_JS)
        print(f"  [{section_label}] {len(raw)} card(s)", file=sys.stderr)

        if debug_dir:
            (debug_dir / f"section_{section_slug}.json").write_text(
                json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8"
            )

        for entry in raw:
            catalogos = [
                {"name": c["name"], "url": abs_url(c["url"]), "last_modified": None, "files": []}
                for c in entry["catalogos"]
            ]
            if not entry["detail_url"]:
                versions = [{
                    "version": entry.get("version"),
                    "revision": entry.get("revision"),
                    "files": {
                        "estandar": _finfo_url(entry.get("estandar_url")),
                        "xsd":      _finfo_url(entry.get("xsd_url")),
                        "xslt":     _finfo_url(entry.get("xslt_url")),
                    },
                    "catalogos": catalogos,
                }]
            else:
                versions = []

            complementos.append({
                "name": entry["name"],
                "detail_url": entry["detail_url"],
                "versions": versions,
                "_section": section_slug,
            })

    print(f"  Total: {len(complementos)} complemento(s)", file=sys.stderr)

    # Now scrape detail pages (navigates away from main page, order doesn't matter)
    for comp in complementos:
        if not comp.get("detail_url"):
            continue
        # Reuse previous versions if the detail URL hasn't changed and we have data
        if not force and prev_comps:
            prev = prev_comps.get(comp["name"])
            if prev and prev.get("detail_url") == comp["detail_url"] and prev.get("versions"):
                comp["versions"] = prev["versions"]
                comp["_cached"] = True  # already went through flatten/group; skip those steps
                print(f"    {comp['name']}: reused {len(comp['versions'])} version(s) from cache", file=sys.stderr)
                continue
        try:
            comp["versions"] = scrape_detail_page(page, comp["detail_url"], debug_dir)
            print(f"    {comp['name']}: {len(comp['versions'])} version(s)", file=sys.stderr)
        except Exception as exc:
            print(f"    ERROR scraping {comp['detail_url']}: {exc}", file=sys.stderr)

    return complementos


# ── downloads ──────────────────────────────────────────────────────────────────


def _catalog_type(cat: dict) -> str:
    """Determine file type key for a catalog entry (xls, xsd, pdf, …)."""
    url = cat.get("url", "")
    ext = Path(url.lower().split("?")[0]).suffix.lstrip(".")
    if ext in ("xls", "xlsx", "xsd", "xslt", "xml", "zip", "pdf", "rar"):
        return "xls" if ext == "xlsx" else ext
    # HTML page or blob — fall back to name hint
    name = cat.get("name", "").lower()
    for t in ("xlsx", "xls", "xsd", "xslt", "pdf", "xml", "zip"):
        if t in name:
            return "xls" if t == "xlsx" else t
    return "other"


def _classify_static_link(text: str, href: str) -> tuple[str, str] | None:
    """Classify a link from a static SAT page. Returns (category, type_key) or None to skip."""
    text_low = text.lower()
    ext = Path(href.lower().split("?")[0]).suffix.lstrip(".")

    # Skip non-technical documents (guides, calendars, FAQs, etc.)
    skip_kw = ("guía", "guia", "calendario", "preguntas", "material", "buscador", "histórico", "historico")
    if any(kw in text_low for kw in skip_kw):
        return None

    # Catalog files
    if "catálogo de datos" in text_low or "catalogo de datos" in text_low:
        return "catalogos", "xsd"
    if "catálogos" in text_low or "catalogos" in text_low:
        return "catalogos", "xls" if ext in ("xls", "xlsx") else "xsd"

    # Technical standard files
    if "esquema" in text_low:
        return "files", "xsd"
    if "estándar" in text_low or "estandar" in text_low:
        return "files", "estandar"
    if "xslt" in text_low or "cadena original" in text_low:
        return "files", "xslt"
    if "matriz" in text_low:
        return "files", "matriz"
    if "patrón" in text_low or "patron" in text_low:
        return "files", "patron"

    return None



def _flatten_catalogos(complementos: list[dict]) -> None:
    """After downloads, reshape each catalog wrapper dict to {source_url, last_modified, files}."""
    for comp in complementos:
        for ver in comp.get("versions", []):
            ver["catalogos"] = {
                cat_type: {
                    "source_url": cat.get("url"),
                    "last_modified": cat.get("last_modified"),
                    "files": cat.get("files", []),
                }
                for cat_type, cat in ver.get("catalogos", {}).items()
            }


def _group_all_catalogos(complementos: list[dict]) -> None:
    """Convert each version's catalogos from a list to a dict.

    Key is ``{type}`` when no catalog_version, or ``{type}.{catalog_version_slug}``
    when the entry came from a versioned HTML section.  Values are single dicts.
    """
    for comp in complementos:
        for ver in comp.get("versions", []):
            grouped: dict[str, dict] = {}
            for cat in ver.get("catalogos", []):
                t = _catalog_type(cat)
                cv = cat.get("catalog_version")
                key = f"{t}.{slugify(cv)}" if cv else t
                grouped[key] = {k: v for k, v in cat.items() if k not in ("name", "catalog_version")}
            ver["catalogos"] = grouped


def _expand_html_catalogs(complementos: list[dict]) -> None:
    """Promote catalog version groups into top-level version entries.

    When an HTML catalog page has sections for multiple versions (e.g. v2.0 and
    v1.1), the group matching the current complemento version stays in place and
    each other group becomes a new entry in comp["versions"] with empty files.
    """
    for comp in complementos:
        extra_versions: list[dict] = []

        for ver in comp.get("versions", []):
            new_cats: list[dict] = []
            for cat in ver.get("catalogos", []):
                url = cat.get("url", "")
                if not url or not _is_html_url(url):
                    new_cats.append(cat)
                    continue
                groups = _resolve_catalog_url(url)
                if len(groups) <= 1:
                    new_cats.append(cat)
                    continue

                for group in groups:
                    if not group["version"]:
                        # Files appearing before any versioned heading — skip to avoid
                        # creating a spurious version=None entry in the manifest.
                        continue
                    # Extract bare version number from heading string ("versión 2.0" → "2.0")
                    vm = re.search(r"(\d[\d.]*)(?:\s+[Rr]evisión\s+(\w+))?", group["version"])
                    gver = vm.group(1) if vm else None
                    grev = vm.group(2) if vm else None

                    # catalog_version is kept on the entry so _download_catalog
                    # can select the right group from the HTML page; it is
                    # stripped from the manifest in _group_all_catalogos.
                    if gver == ver.get("version") and grev == ver.get("revision"):
                        new_cats.append({**cat, "catalog_version": group["version"], "files": []})
                    else:
                        existing = next(
                            (v for v in extra_versions
                             if v.get("version") == gver and v.get("revision") == grev),
                            None,
                        )
                        if existing is None:
                            existing = {
                                "version": gver,
                                "revision": grev,
                                "files": {
                                    "estandar": {"url": None, "last_modified": None, "local_file": None},
                                    "xsd":      {"url": None, "last_modified": None, "local_file": None},
                                    "xslt":     {"url": None, "last_modified": None, "local_file": None},
                                },
                                "catalogos": [],
                            }
                            extra_versions.append(existing)
                        existing["catalogos"].append({**cat, "catalog_version": group["version"], "files": []})

            ver["catalogos"] = new_cats

        comp["versions"].extend(extra_versions)


def _xls_version_slug(path: Path) -> str:
    """Read version/revision from XLS metadata cells and return a slug like '1-0'.

    Finds the metadata header row, then reads 'Versión catálogo' and 'Revisión catálogo'
    values from the row below. Returns empty string if not found.
    """
    if _xlrd is None or not path.exists():
        return ""

    def _norm(s: str) -> str:
        import unicodedata
        s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
        return re.sub(r"\s+", "", s).lower()

    # Keys we want (normalized, no spaces)
    VERSION_KEYS  = {"versioncatalogo", "versioncfdi"}  # prefer catalog ver
    REVISION_KEYS = {"revisioncatalogo", "revision"}

    try:
        wb = _xlrd.open_workbook(str(path), formatting_info=False)
        sheet = wb.sheet_by_index(0)
        for row_idx in range(min(10, sheet.nrows)):
            row_cells = [str(sheet.cell(row_idx, ci).value).strip() for ci in range(sheet.row_len(row_idx))]
            norms = [_norm(c) for c in row_cells]
            if not any("version" in n for n in norms):
                continue
            val_row = row_idx + 1
            if val_row >= sheet.nrows:
                break

            def _fmt(v: str) -> str:
                try:
                    fv = float(v)
                    return str(int(fv)) if fv.is_integer() else v
                except ValueError:
                    return v

            found: dict[str, str] = {}
            for ci, norm in enumerate(norms):
                if ci >= sheet.row_len(val_row):
                    continue
                val = str(sheet.cell(val_row, ci).value).strip()
                if not val or val in ("0.0", "0"):
                    continue
                found[norm] = _fmt(val)

            version  = found.get("versioncatalogo") or found.get("versioncfdi", "")
            revision = found.get("revisioncatalogo") or found.get("revision", "")
            slug = "-".join(slugify(p) for p in [version, revision] if p)
            return slug
    except Exception:
        pass
    return ""


def _download_catalog(
    cat: dict,
    files_dir: Path,
    base_slug: str,
    subdir: str,
    hf_xls_dir: Path | None = None,
    hf_subpath: str = "",
    hf_subpath_has_version: bool = True,
    ver_fp: str | None = None,
    prev: _PrevState | None = None,
    verify: bool = False,
    force: bool = False,
) -> None:
    """Download a catalog entry's files into base_slug/subdir/.

    XLS/XLSX files are routed to hf_xls_dir/hf_subpath/ when hf_xls_dir is given.
    The filename is the cleaned original name; the version lives in the folder.
    Multiple XLS files from the same catalog get a -{n} counter suffix.
    """
    url = cat.get("url", "")
    if not url:
        return

    cat_fp = cat.get("last_modified") or ver_fp or None

    # Fast path: if fingerprint unchanged and all previous files have stored hashes, reuse them.
    # This avoids both _resolve_catalog_url (HTTP) and download() calls entirely.
    # Skip entirely if fingerprint unchanged and all previous files have stored hashes.
    # --verify bypasses this to always fetch and compare hash in memory.
    # _prev.cat_fp is what was stored in the manifest; cat_fp is what was just scraped.
    _prev = prev or _PrevState()
    if not force and not verify and cat_fp and _prev.files and all(fe.get("hash") for fe in _prev.files):
        if _prev.cat_fp is None or cat_fp == _prev.cat_fp:
            cat["files"] = _prev.files
            for fe in _prev.files:
                print(f"    ✓ skip (version {cat_fp}) {fe.get('local_file', fe.get('url', ''))}", file=sys.stderr)
            return
        # else: fingerprint changed — fall through to re-download

    cat_version = cat.get("catalog_version")
    groups = _resolve_catalog_url(url)
    if cat_version:
        files = next((g["files"] for g in groups if g["version"] == cat_version), [])
    else:
        files = groups[0]["files"] if groups else []

    file_entries = []
    hf_idx = 0
    for file_url in files:
        fname = _filename(file_url)
        ext = Path(fname.lower()).suffix
        if hf_xls_dir and ext in (".xls", ".xlsx"):
            original_stem = _clean_stem(Path(fname).stem)
            idx_suffix = f"-{hf_idx}" if hf_idx > 0 else ""
            hf_fname = f"{original_stem}{idx_suffix}{ext}"
            hf_idx += 1
            # Download to flat path first, then move into version subfolder if found
            dest = hf_xls_dir / hf_subpath / hf_fname
            rel = str(dest)
        else:
            rel = f"{base_slug}/{subdir}/{fname}"
            dest = files_dir / rel
        stored_fp, stored_hash = _prev.lm.get(file_url, (None, None))
        status = download(file_url, dest, fingerprint=cat_fp, stored_fingerprint=stored_fp, stored_hash=stored_hash, verify=verify, force=force)
        if status == "written":
            print(f"    ↓ {rel}", file=sys.stderr)
        elif status == "skipped":
            print(f"    ✓ skip (fingerprint) {rel}", file=sys.stderr)
        elif status == "unchanged":
            print(f"    ✓ skip (hash match)  {rel}", file=sys.stderr)
        elif status == "error":
            print(f"    ✗ error fetching     {rel}", file=sys.stderr)
        # For XLS with no version in the path (hf_subpath == hf_ver_subpath),
        # try reading version from the file itself
        if dest.exists() and hf_xls_dir and ext in (".xls", ".xlsx") and not hf_subpath_has_version:
            ver_slug = _xls_version_slug(dest)
            if ver_slug:
                versioned_dest = dest.parent / ver_slug / dest.name
                versioned_dest.parent.mkdir(parents=True, exist_ok=True)
                dest.rename(versioned_dest)
                dest = versioned_dest
                rel = str(dest)
                print(f"    → moved to versioned folder: {rel}", file=sys.stderr)
        if dest.exists():
            file_entries.append({"url": file_url, "local_file": rel, "hash": _sha256(dest)})
        elif status in ("skipped", "unchanged") and stored_hash:
            # File not on disk (hf/ gitignored) but content unchanged — preserve manifest entry
            file_entries.append({"url": file_url, "local_file": rel, "hash": stored_hash})

    cat["files"] = file_entries


def _comp_hf_section(comp: dict) -> str:
    """Return the HF section path for a comp entry (mirrors download_complemento routing)."""
    prefix = comp.get("_section", "complementos")
    if prefix == "anexo-20":
        return "anexo20/retenciones" if "retenciones" in comp["name"].lower() else "anexo20/factura"
    name_slug = slugify(re.sub(r"\s*\(.*?\)", "", comp["name"]).strip())
    return f"complementos/{name_slug}"


def _comp_matches_sections(comp: dict, sections: list[str]) -> bool:
    section = _comp_hf_section(comp)
    return any(section == s or section.startswith(s.rstrip("/") + "/") for s in sections)


def download_complemento(
    comp: dict,
    files_dir: Path,
    prefix: str = "complementos",
    prev_lm: dict[str, tuple[str | None, str]] | None = None,
    prev_manifest: dict | None = None,
    verify: bool = False,
    force: bool = False,
) -> None:
    print(f"\n  {comp['name']}", file=sys.stderr)
    name_slug = slugify(re.sub(r"\s*\(.*?\)", "", comp["name"]).strip())
    base = f"{prefix}/{name_slug}"

    # HF XLS subpath routing:
    #   anexo-20 retenciones → hf/xls/anexo20/retenciones/
    #   anexo-20 factura     → hf/xls/anexo20/factura/
    #   complementos/*       → hf/xls/complementos/{slug}/
    hf_subpath = _comp_hf_section(comp)

    def get(url: str, subdir: str) -> str | None:
        if not url:
            return None
        fname = _filename(url)
        rel = f"{base}/{subdir}/{fname}"
        dest = files_dir / rel
        status = download(url, dest, force=force)
        if status == "written":
            print(f"    ↓ {rel}", file=sys.stderr)
        elif status == "unchanged":
            print(f"    ✓ skip (hash match)  {rel}", file=sys.stderr)
        elif status == "error":
            print(f"    ✗ error fetching     {rel}", file=sys.stderr)
            if not dest.exists():
                return None
        return rel

    # Build {ver_fp: {cat_type: (stored_cat_fp, [file_entries])}} from previous manifest
    prev_ver_files: dict[str, dict[str, tuple]] = {}
    if prev_manifest:
        for pc in prev_manifest.get("complementos", []):
            if pc["name"] == comp["name"]:
                for pv in pc.get("versions", []):
                    pvfp = "-".join(str(v) for v in [pv.get("version"), pv.get("revision")] if v) or ""
                    prev_ver_files[pvfp] = {
                        ct: (cd.get("last_modified") or pvfp, cd.get("files", []))
                        for ct, cd in pv.get("catalogos", {}).items()
                    }
                break

    for ver in comp.get("versions", []):
        vslug = _ver_slug(ver)
        hf_vslug = _hf_ver_slug(ver)
        hf_ver_subpath = f"{hf_subpath}/{hf_vslug}" if hf_vslug != "files" else hf_subpath
        ver_fp = "-".join(str(v) for v in [ver.get("version"), ver.get("revision")] if v) or None
        for finfo in ver.get("files", {}).values():
            if finfo.get("url"):
                local = get(finfo["url"], vslug)
                finfo["local_file"] = local
                finfo["hash"] = _sha256(files_dir / local) if local else None
        for cat_type, cat in ver.get("catalogos", {}).items():
            pv_entry = prev_ver_files.get(ver_fp or "", {}).get(cat_type) if ver_fp else None
            prev_cat_fp, prev_files = pv_entry if pv_entry else (None, [])
            _download_catalog(
                cat, files_dir, base, vslug,
                hf_xls_dir=HF_XLS_DIR, hf_subpath=hf_ver_subpath,
                hf_subpath_has_version=(hf_vslug != "files"),
                ver_fp=ver_fp,
                prev=_PrevState(cat_fp=prev_cat_fp, files=prev_files, lm=prev_lm or {}),
                verify=verify,
                force=force,
            )


# ── CSV export ─────────────────────────────────────────────────────────────────

CSV_FIELDS = [
    "scraped_at", "category", "name", "slug", "detail_url", "version", "revision",
    "latest", "file_type", "url", "local_file", "size", "hash", "last_modified",
]


def _ver_tuple(v: str | None) -> tuple:
    """Parse a version string into a comparable tuple, e.g. "3.1" → (3, 1)."""
    if not v:
        return (0,)
    try:
        return tuple(int(x) for x in v.split("."))
    except ValueError:
        return (0,)


def write_csv(manifest: dict, csv_path: Path) -> None:
    """Write a flat CSV with one row per file across all complementos/versions."""
    scraped_at = manifest["scraped_at"]
    files_dir = OUTPUT_DIR / "files"
    rows: list[dict] = []

    for comp in manifest["complementos"]:
        versions = comp.get("versions", [])
        latest_ver = max((v.get("version") for v in versions), key=_ver_tuple, default=None)

        category = comp.get("_section", "complementos")
        name = comp["name"]
        slug = slugify(re.sub(r"\s*\(.*?\)", "", name).strip())
        base = {"scraped_at": scraped_at, "category": category, "name": name, "slug": slug, "detail_url": comp.get("detail_url") or ""}

        def row(version: str, revision: str, ftype: str, url: str | None, local: str | None, h: str | None, lm: str | None) -> None:
            if not url:
                return
            if local:
                p = Path(local) if local.startswith("hf/") else files_dir / local
            else:
                p = None
            size = p.stat().st_size if p and p.exists() else ""
            rows.append({**base, "version": version, "revision": revision,
                          "latest": "true" if (version or None) == latest_ver else "false",
                          "file_type": ftype,
                          "url": url, "local_file": local or "", "size": size,
                          "hash": h or "", "last_modified": lm or ""})

        for ver in versions:
            v = ver.get("version") or ""
            r = ver.get("revision") or ""
            for ftype, finfo in ver.get("files", {}).items():
                row(v, r, ftype, finfo.get("url"), finfo.get("local_file"), finfo.get("hash"), finfo.get("last_modified"))
            for cat_key, cat in ver.get("catalogos", {}).items():
                ftype = f"catalogo.{cat_key}"
                lm = cat.get("last_modified")
                files = cat.get("files") or []
                if files:
                    for fe in files:
                        row(v, r, ftype, fe.get("url"), fe.get("local_file"), fe.get("hash"), lm)
                else:
                    row(v, r, ftype, cat.get("source_url"), None, None, lm)

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        w.writeheader()
        w.writerows(rows)

    print(f"CSV     → {csv_path}  ({len(rows)} rows)", file=sys.stderr)


# ── main ───────────────────────────────────────────────────────────────────────


def redownload_hf(force: bool = False) -> int:
    """Re-download HF XLS catalog files from the existing manifest with current naming rules.

    Reads output/catalogos-manifest.json, downloads only XLS catalog files to hf/xls/ using
    the new _clean_stem + version suffix naming, then rewrites the manifest and CSV.
    Skips the Playwright scraping step entirely.
    """
    m = latest_manifest()
    if m is None:
        print("No manifest found. Run scripts/catalogos/scrape.py first.", file=sys.stderr)
        return 1

    print(f"Re-downloading HF XLS files from manifest ({m['scraped_at']})...", file=sys.stderr)

    for comp in m["complementos"]:
        hf_subpath = _comp_hf_section(comp)

        for ver in comp.get("versions", []):
            hf_vslug = _hf_ver_slug(ver)
            hf_ver_subpath = f"{hf_subpath}/{hf_vslug}" if hf_vslug != "files" else hf_subpath

            hf_idx: dict[str, int] = {}  # per-catalog counter for multi-file catalogs
            for cat in ver.get("catalogos", {}).values():
                new_files = []
                for fe in cat.get("files", []):
                    file_url = fe.get("url", "")
                    if not file_url:
                        continue
                    orig_fname = _filename(file_url)
                    ext = Path(orig_fname.lower()).suffix
                    if ext not in (".xls", ".xlsx"):
                        continue
                    original_stem = _clean_stem(Path(orig_fname).stem)
                    idx = hf_idx.get(original_stem, 0)
                    hf_idx[original_stem] = idx + 1
                    idx_suffix = f"-{idx}" if idx > 0 else ""
                    hf_fname = f"{original_stem}{idx_suffix}{ext}"
                    dest = HF_XLS_DIR / hf_ver_subpath / hf_fname
                    rel = str(dest)
                    status = download(file_url, dest, stored_hash=fe.get("hash"), force=force)
                    if status == "written":
                        print(f"  ↓ {rel}", file=sys.stderr)
                    elif status == "unchanged":
                        print(f"  ✓ skip (hash match)  {rel}", file=sys.stderr)
                    elif status == "skipped":
                        print(f"  ✓ skip (fingerprint) {rel}", file=sys.stderr)
                    elif status == "error":
                        print(f"  ✗ error fetching     {rel}", file=sys.stderr)
                    if dest.exists():
                        # Only try to derive version from XLS content when the manifest
                        # had no version info (hf_vslug == "files")
                        if hf_vslug == "files":
                            ver_slug = _xls_version_slug(dest)
                            if ver_slug:
                                versioned_dest = dest.parent / ver_slug / dest.name
                                versioned_dest.parent.mkdir(parents=True, exist_ok=True)
                                dest.rename(versioned_dest)
                                dest = versioned_dest
                                rel = str(dest)
                                print(f"  → moved to versioned folder: {rel}", file=sys.stderr)
                        new_files.append({"url": file_url, "local_file": rel, "hash": _sha256(dest)})
                if new_files:
                    cat["files"] = new_files

    manifest_path = OUTPUT_DIR / "catalogos-manifest.json"
    manifest_path.write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nManifest → {manifest_path}", file=sys.stderr)
    write_csv(m, OUTPUT_DIR / "catalog.csv")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--debug", action="store_true", help="Save intermediate HTML/JSON")
    parser.add_argument("--show-browser", action="store_true", help="Show browser window")
    parser.add_argument("--dry-run", action="store_true", help="Scrape only, don't write files")
    parser.add_argument("--redownload-hf", action="store_true",
                        help="Re-download HF XLS files from existing manifest (skips scraping)")
    parser.add_argument(
        "--sections", nargs="+", metavar="SECTION",
        help=(
            "Only download XLS for these sections. "
            "Shortcuts: anexo20, complementos, all. "
            "Can also use full paths: anexo20/factura complementos/carta-porte"
        ),
    )
    parser.add_argument(
        "--verify", action="store_true",
        help="Fetch every file and compare hash in memory; write only if content changed (slower than default, safer than --force)",
    )
    parser.add_argument(
        "--force", action="store_true",
        help="Re-download and overwrite all files regardless of fingerprint or hash",
    )
    args = parser.parse_args()

    if args.redownload_hf:
        return redownload_hf(force=args.force)

    run_dir = OUTPUT_DIR
    files_dir = run_dir / "files"
    debug_dir = (run_dir / "debug") if args.debug else None

    if debug_dir:
        debug_dir.mkdir(parents=True, exist_ok=True)

    prev_manifest = latest_manifest()
    prev_comps: dict[str, dict] = (
        {c["name"]: c for c in prev_manifest.get("complementos", [])}
        if prev_manifest else {}
    )

    # Determine which sources are needed based on --sections
    _requested = set(args.sections) if args.sections else {"all"}
    need_playwright = not args.sections or any(
        s in ("all", "complementos") or s.startswith("complementos/")
        for s in _requested
    )
    need_static = not args.sections or any(
        s in ("all", "anexo20") or s.startswith("anexo20/")
        for s in _requested
    )

    complementos: list[dict] = []
    cached_comps: list[dict] = []

    if need_playwright:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=not args.show_browser)
            ctx = browser.new_context(ignore_https_errors=True, user_agent=USER_AGENT)
            page = ctx.new_page()
            # Block resources that don't affect DOM content to speed up page loads
            page.route(
                "**/*",
                lambda route: route.abort()
                if route.request.resource_type in ("image", "media", "font", "stylesheet")
                or any(kw in route.request.url for kw in ("analytics", "gtm", "google-tag", "hotjar", "clarity", "doubleclick"))
                else route.continue_(),
            )
            try:
                complementos = scrape_complementos_section(page, debug_dir, prev_comps=prev_comps, force=args.force)
            finally:
                ctx.close()
                browser.close()

        # Separate cached entries (versions reused from previous manifest) from
        # fresh entries that need pipeline processing.
        fresh_from_playwright: list[dict] = []
        for _c in complementos:
            if _c.pop("_cached", False):
                cached_comps.append(_c)
            else:
                fresh_from_playwright.append(_c)
        complementos = fresh_from_playwright

        if not complementos and not cached_comps:
            print("No complementos found — run with --debug to inspect the page.", file=sys.stderr)
            return 1

    if need_static:
        print("\nScraping static pages...", file=sys.stderr)
        for page_name, page_url in STATIC_PAGES:
            comp = scrape_static_page(page_name, page_url)
            if comp["versions"]:
                complementos.append(comp)

    # Merge cached entries back for logging and change detection
    all_complementos = complementos + cached_comps

    _expand_html_catalogs(complementos)
    _group_all_catalogos(complementos)

    print(f"\nFound {len(all_complementos)} complemento(s):", file=sys.stderr)
    for c in all_complementos:
        print(f"  {c['name']} ({len(c.get('versions', []))} version(s))", file=sys.stderr)

    if args.dry_run:
        print(json.dumps(all_complementos, ensure_ascii=False, indent=2))
        return 0

    if not args.force and not data_changed(all_complementos):
        print("\nNo changes since last run — skipping write.", file=sys.stderr)
        return 0

    run_dir.mkdir(parents=True, exist_ok=True)

    # Build {url: fingerprint} from previous manifest for skip-if-unchanged logic.
    # --force clears this so every file is re-downloaded.
    # Fingerprint = last_modified if available, else "version-revision" from the version entry.
    prev_lm: dict[str, tuple[str | None, str]] = {}
    if prev_manifest:
        for c in prev_manifest.get("complementos", []):
            for ver in c.get("versions", []):
                ver_fp = "-".join(str(v) for v in [ver.get("version"), ver.get("revision")] if v)
                for cat in ver.get("catalogos", {}).values():
                    fp = cat.get("last_modified") or ver_fp
                    for fe in cat.get("files", []):
                        if fe.get("url"):
                            prev_lm[fe["url"]] = (fp, fe.get("hash") or "")
    if args.force:
        prev_lm.clear()

    _SECTION_SHORTCUTS = {
        "anexo20":      ["anexo20/factura", "anexo20/retenciones"],
        "complementos": ["complementos"],
        "all":          [],  # empty = no filter
    }
    sections_filter: list[str] | None = None
    if args.sections:
        expanded: list[str] = []
        for s in args.sections:
            if s in _SECTION_SHORTCUTS:
                expanded.extend(_SECTION_SHORTCUTS[s])
            else:
                expanded.append(s)
        sections_filter = expanded or None  # 'all' expands to [] → None = no filter

    for comp in complementos:
        if sections_filter and not _comp_matches_sections(comp, sections_filter):
            continue
        download_complemento(comp, files_dir, comp.get("_section", "complementos"), prev_lm=prev_lm, prev_manifest=prev_manifest, verify=args.verify, force=args.force)

    _flatten_catalogos(complementos)

    manifest = {"scraped_at": datetime.now().isoformat(), "complementos": all_complementos}

    manifest_path = run_dir / "catalogos-manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nManifest → {manifest_path}", file=sys.stderr)

    write_csv(manifest, run_dir / "catalog.csv")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
