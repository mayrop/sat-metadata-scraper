#!/usr/bin/env python3
"""Scrape SAT Deducciones Personales minisite into normalized CSV outputs.

Source:
  https://www.sat.gob.mx/minisitio/DeduccionesPersonales/index.html

Outputs:
  output/deducciones-personales-manifest.json
  output/deducciones-personales/pages/<slug>.html
  hf/csv/deducciones-personales/deducciones.csv
  hf/csv/deducciones-personales/catalogo_productos_servicios.csv

Usage:
  uv run scripts/deducciones_personales/scrape.py
  uv run scripts/deducciones_personales/scrape.py --force
  uv run scripts/deducciones_personales/scrape.py --dry-run
  uv run scripts/deducciones_personales/scrape.py --slug colegiaturas
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import html.parser
import json
import re
import ssl
import sys
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin, urlsplit

INDEX_URL = "https://www.sat.gob.mx/minisitio/DeduccionesPersonales/index.html"
BASE_URL = "https://www.sat.gob.mx"

OUTPUT_DIR = Path("output/deducciones-personales")
PAGES_DIR = OUTPUT_DIR / "pages"
MANIFEST = Path("output/deducciones-personales-manifest.json")
HF_DIR = Path("hf/csv/deducciones-personales")

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

SECTION_HEADINGS = [
    "¿Qué gastos son considerados deducibles?",
    "¿Qué debes revisar en tu factura para que sea deducible?",
    "¿Cuánto puedes deducir?",
    "Consideraciones para deducir correctamente tus gastos",
    "¿Qué no puedes deducir?",
    "Material adicional",
    "Contenidos relacionado",
    "Contenidos relacionados",
    "Documentos relacionados",
    "Catálogo de productos y servicios para la facturación",
]

NAV_LINES = {
    "Deducciones personales",
    "Toggle navigation",
    "Inicio",
    "Médicos",
    "Seguros",
    "Colegiaturas",
    "Transporte escolar",
    "Aportaciones",
    "Cuentas especiales",
    "Créditos hipotecarios",
    "Funerarios",
    "Donativos",
}

SECTION_KEYS = {
    "¿Qué gastos son considerados deducibles?": "deductible_items",
    "¿Qué debes revisar en tu factura para que sea deducible?": "factura_requirements",
    "¿Cuánto puedes deducir?": "limit_text",
    "Consideraciones para deducir correctamente tus gastos": "considerations",
    "¿Qué no puedes deducir?": "non_deductible_items",
    "Material adicional": "material_adicional",
    "Contenidos relacionado": "contenidos_relacionados",
    "Contenidos relacionados": "contenidos_relacionados",
    "Documentos relacionados": "documentos_relacionados",
    "Catálogo de productos y servicios para la facturación": "catalogo",
}

CATALOG_HEADER_LINES = {
    "clave de productos y servicios",
    "clave de productos y servicio",
    "descripcion",
    "descripción",
    "palabras similares",
}

CATALOG_TITLE_ALIASES = {
    "gastos-medicos-y-hospitalarios": [
        "Honorarios médicos, dentales y gastos hospitalarios",
        "Gastos médicos y hospitalarios",
    ],
    "primas-por-seguros-de-gastos-medicos": [
        "Primas por seguros",
        "Primas por seguros de gastos médicos",
    ],
    "colegiaturas": [
        "Pagos por servicios educativos",
        "Colegiaturas",
    ],
    "transporte-escolar": [
        "Transporte escolar",
    ],
    "aportaciones-complementarias-para-el-retiro": [
        "Aportaciones voluntarias y complementarias de retiro",
        "Aportaciones complementarias para el retiro",
    ],
    "depositos-en-cuentas-personales-especiales-para-el-ahorro": [
        "Depósitos en cuentas especiales para el ahorro",
        "Depósitos en cuentas personales especiales para el ahorro",
    ],
    "creditos-hipotecarios": [
        "Créditos hipotecarios",
        "Intereses reales efectivamente pagados por créditos hipotecarios",
    ],
    "gastos-funerarios": [
        "Gastos funerarios",
    ],
    "donativos": [
        "Donativos",
    ],
}

LOOKUP_LINK_LABELS = {
    "Consulta la lista aquí",
}

GENERIC_LINK_LABELS = {
    "aquí",
    "Consultar guía",
}

EXCLUDED_EXTRA_LINK_URLS = {
    "https://ptscvisorpredepro.clouda.sat.gob.mx/Resumen",
}


def slugify(text: str) -> str:
    for src, dst in [
        ("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"), ("ñ", "n"),
        ("Á", "a"), ("É", "e"), ("Í", "i"), ("Ó", "o"), ("Ú", "u"), ("Ñ", "n"),
    ]:
        text = text.replace(src, dst)
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")


def normalize_text(text: str) -> str:
    text = re.sub(r"\s+", " ", text or "").strip()
    return text


def normalize_key(text: str) -> str:
    return slugify(normalize_text(text))


def json_hash(data: object) -> str:
    raw = json.dumps(data, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def fetch_bytes(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=60) as response:
            return response.read()
    except URLError as exc:
        reason = getattr(exc, "reason", None)
        if isinstance(reason, ssl.SSLCertVerificationError):
            insecure = ssl._create_unverified_context()
            with urllib.request.urlopen(req, timeout=60, context=insecure) as response:
                return response.read()
        raise


def load_manifest() -> dict:
    if MANIFEST.exists():
        return json.loads(MANIFEST.read_text(encoding="utf-8"))
    return {}


def stable_manifest_view(manifest: dict) -> dict:
    pages = []
    for page in manifest.get("paginas", []):
        pages.append({k: v for k, v in page.items() if k != "ultima_revision"})
    return {
        "url_fuente": manifest.get("url_fuente"),
        "indice": manifest.get("indice", {}),
        "paginas": sorted(pages, key=lambda item: item["slug"]),
    }


class LineHTMLParser(html.parser.HTMLParser):
    """Convert HTML into normalized text lines and in-order links."""

    BLOCK_TAGS = {
        "p", "div", "section", "article", "main", "header", "footer", "aside",
        "h1", "h2", "h3", "h4", "h5", "h6", "li", "ul", "ol",
        "table", "thead", "tbody", "tr", "td", "th",
        "button", "summary",
    }

    def __init__(self, base_url: str) -> None:
        super().__init__(convert_charrefs=True)
        self.base_url = base_url
        self._skip_depth = 0
        self._parts: list[str] = []
        self._current_link: str | None = None
        self._current_link_text: list[str] = []
        self.links: list[dict[str, str]] = []
        self.title: str | None = None
        self._in_title = False

    def _emit_break(self) -> None:
        if not self._parts or self._parts[-1] == "\n":
            return
        self._parts.append("\n")

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"script", "style", "noscript"}:
            self._skip_depth += 1
            return
        if self._skip_depth:
            return
        if tag in self.BLOCK_TAGS or tag == "br":
            self._emit_break()
        if tag == "a":
            href = dict(attrs).get("href")
            self._current_link = urljoin(self.base_url, href) if href else None
            self._current_link_text = []
        if tag == "title":
            self._in_title = True

    def handle_endtag(self, tag: str) -> None:
        if tag in {"script", "style", "noscript"}:
            self._skip_depth = max(0, self._skip_depth - 1)
            return
        if self._skip_depth:
            return
        if tag in self.BLOCK_TAGS:
            self._emit_break()
        if tag == "a":
            text = normalize_text(" ".join(self._current_link_text))
            if self._current_link and text and text.lower() != "image":
                self.links.append({"label": text, "url": self._current_link})
            self._current_link = None
            self._current_link_text = []
        if tag == "title":
            self._in_title = False

    def handle_data(self, data: str) -> None:
        if self._skip_depth:
            return
        text = normalize_text(data)
        if not text:
            return
        if self._in_title and not self.title:
            self.title = text
        self._parts.append(text)
        if self._current_link is not None:
            self._current_link_text.append(text)

    def lines(self) -> list[str]:
        raw = "".join(self._parts)
        out: list[str] = []
        for line in raw.splitlines():
            text = normalize_text(line)
            if text:
                out.append(text)
        return out


class IndexParser(html.parser.HTMLParser):
    """Extract deduction tiles from the landing page."""

    def __init__(self, base_url: str) -> None:
        super().__init__(convert_charrefs=True)
        self.base_url = base_url
        self._current_href: str | None = None
        self._current_text: list[str] = []
        self.entries: list[dict[str, str]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag != "a":
            return
        href = dict(attrs).get("href")
        if not href:
            return
        if href.startswith("#"):
            return
        absolute = urljoin(self.base_url, href)
        if "/minisitio/DeduccionesPersonales/" not in absolute:
            return
        parts = urlsplit(absolute)
        if parts.path.endswith("/index.html"):
            return
        self._current_href = absolute
        self._current_text = []

    def handle_data(self, data: str) -> None:
        if self._current_href is not None:
            text = normalize_text(data)
            if text and text.lower() != "image":
                self._current_text.append(text)

    def handle_endtag(self, tag: str) -> None:
        if tag != "a" or self._current_href is None:
            return
        label = normalize_text(" ".join(self._current_text))
        if label:
            slug = slugify(label)
            entry = {"slug": slug, "titulo": label, "url": self._current_href}
            if entry not in self.entries:
                self.entries.append(entry)
        self._current_href = None
        self._current_text = []


def parse_html_document(html_text: str) -> tuple[list[str], list[dict[str, str]], str | None]:
    parser = LineHTMLParser(INDEX_URL)
    parser.feed(html_text)
    parser.close()
    enlaces = [{"etiqueta": link["label"], "url": link["url"]} for link in dedupe_links(parser.links)]
    return parser.lines(), enlaces, parser.title


def dedupe_links(links: list[dict[str, str]]) -> list[dict[str, str]]:
    seen: set[tuple[str, str]] = set()
    out: list[dict[str, str]] = []
    for link in links:
        label = normalize_text(link.get("label", ""))
        url = normalize_text(link.get("url", ""))
        if not label or not url:
            continue
        if label.lower() in {"inicio", "toggle navigation", "image"}:
            continue
        key = (label, url)
        if key in seen:
            continue
        seen.add(key)
        out.append({"label": label, "url": url})
    return out


def link_label_score(label: str) -> tuple[int, int]:
    clean = normalize_text(label)
    return (0 if clean in GENERIC_LINK_LABELS else 1, len(clean))


def parse_index(html_text: str) -> dict:
    lines, links, _ = parse_html_document(html_text)
    title = lines[0] if lines else "Deducciones personales"
    intro: list[str] = []
    deducciones: list[dict[str, str]] = []
    seen_urls: set[str] = set()

    for match in re.finditer(r'href="([^"]+\.html)"[^>]*>(.*?)</a>', html_text, flags=re.S | re.I):
        href, inner = match.groups()
        absolute = urljoin(INDEX_URL, href)
        parts = urlsplit(absolute)
        if parts.path.endswith("/index.html"):
            continue
        label = normalize_text(re.sub(r"<[^>]+>", " ", inner))
        if not label:
            continue
        if absolute in seen_urls:
            continue
        seen_urls.add(absolute)
        deducciones.append({
            "slug": slugify(label),
            "titulo": label,
            "url": absolute,
        })

    titles = {entry["titulo"] for entry in deducciones}
    for line in lines[1:]:
        if line in {"Toggle navigation", "Inicio", "Deducciones personales"}:
            continue
        if line in titles:
            break
        intro.append(line)

    return {
        "titulo": title,
        "introduccion": intro,
        "deducciones": deducciones,
        "enlaces": links,
    }


def heading_index(lines: list[str]) -> int:
    for idx, line in enumerate(lines):
        if line in SECTION_KEYS:
            return idx
    return len(lines)


def split_sections(lines: list[str], title_hint: str | None = None) -> tuple[str, list[str], dict[str, list[str]]]:
    title = title_hint or ""
    if not title:
        for line in lines:
            if line in SECTION_KEYS:
                break
            if "deducciones personales" not in line.lower():
                title = line
                break

    first_heading = heading_index(lines)
    start_idx = 0
    if title:
        positions = [i for i, line in enumerate(lines[:first_heading]) if line == title]
        if positions:
            start_idx = positions[-1] + 1

    intro: list[str] = []
    sections: dict[str, list[str]] = {v: [] for v in set(SECTION_KEYS.values())}
    current_key: str | None = None

    for line in lines[start_idx:]:
        if line in SECTION_KEYS:
            current_key = SECTION_KEYS[line]
            continue
        if current_key is None:
            if line == title or line in NAV_LINES or line.startswith("Inicio/"):
                continue
            intro.append(line)
        else:
            sections.setdefault(current_key, []).append(line)

    return title, intro, sections


def parse_limit_amounts(lines: list[str]) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for line in lines:
        m = re.match(r"(.+?):\s*\$([\d,]+(?:\.\d+)?)", line)
        if not m:
            continue
        amount = float(m.group(2).replace(",", ""))
        rows.append({
            "etiqueta": normalize_text(m.group(1)),
            "monto_mxn": int(amount) if amount.is_integer() else amount,
        })
    return rows


def parse_uma_limit(lines: list[str]) -> dict[str, object] | None:
    text = normalize_text(" ".join(lines))
    if not text or "uma" not in text.lower():
        return None
    compact = text.lower()

    word_numbers = {
        "uno": 1,
        "una": 1,
        "dos": 2,
        "tres": 3,
        "cuatro": 4,
        "cinco": 5,
        "seis": 6,
        "siete": 7,
        "ocho": 8,
        "nueve": 9,
        "diez": 10,
    }

    m = re.search(r"(\d+)\s*(?:veces\s+el\s+valor\s+anual\s+de\s+la\s+)?uma", compact, flags=re.I)
    if m:
        return {"valor": int(m.group(1)), "texto_fuente": text}

    for raw, value in word_numbers.items():
        if re.search(rf"{raw}\s+veces\s+el\s+valor\s+anual\s+de\s+la\s+uma", compact):
            return {"valor": value, "texto_fuente": text}
        if re.search(rf"{raw}\s+unidades?\s+de\s+medida\s+y\s+actualizaci[oó]n\s*\(uma\)", compact):
            return {"valor": value, "texto_fuente": text}
        if re.search(rf"entre\s*{raw}\s+veces?\s+el\s+valor\s+anual\s+de\s+la\s+uma", compact):
            return {"valor": value, "texto_fuente": text}
        if re.search(rf"entre\s*{raw}\s+unidades?\s+de\s+medida\s+y\s+actualizaci[oó]n", compact):
            return {"valor": value, "texto_fuente": text}

    return None


def parse_limit(lines: list[str]) -> dict[str, object] | None:
    raw = serialize_section(lines)
    if not raw:
        return None
    text_lines = [line.strip() for line in raw.splitlines() if line.strip()]

    percent_limits: list[float] = []
    percent_limits_isr: list[float] = []
    for match in re.finditer(r"(\d+(?:\.\d+)?)\s*%", raw):
        value = float(match.group(1))
        if value.is_integer():
            value = int(value)
        start = max(0, match.start() - 120)
        end = min(len(raw), match.end() + 120)
        context = raw[start:end].lower()
        target = (
            percent_limits_isr
            if any(token in context for token in ["isr", "impuesto sobre la renta", "ingresos acumulables"])
            else percent_limits
        )
        if value not in target:
            target.append(value)

    annual_absolute_limits = parse_limit_amounts(lines)
    uma_limit = parse_uma_limit(lines)

    return {
        "texto": text_lines,
        "umas": None if uma_limit is None else uma_limit["valor"],
        "porcentaje_anual": percent_limits,
        "porcentaje_anual_isr": percent_limits_isr,
        "monto_anual": annual_absolute_limits,
    }


def parse_cfdi_use_codes(lines: list[str]) -> list[str]:
    codes: list[str] = []
    for line in lines:
        for match in re.findall(r"[A-Z]\d{2}", line):
            if match not in codes:
                codes.append(match)
    return codes


def parse_cfdi_use_requirements(lines: list[str]) -> list[dict[str, str]]:
    requirements: list[dict[str, str]] = []
    for line in lines:
        if "cfdi" not in line.lower():
            continue
        for match in re.finditer(r"([A-Z]\d{2})\s*([^.]*)", line):
            code = match.group(1)
            label = normalize_text(match.group(2)).strip(' ."')
            if not label:
                continue
            item = {"clave": code, "descripcion": label, "linea_fuente": line}
            if item not in requirements:
                requirements.append(item)
    return requirements


def parse_product_service_requirement(
    lines: list[str],
    catalog_section_title: str | None,
    catalog_rows: list[dict[str, str]],
) -> dict | None:
    for line in lines:
        if "clave de producto y servicio" not in line.lower():
            continue
        return {
            "requerido": True,
            "linea_fuente": line,
            "titulo_seccion_catalogo": catalog_section_title,
            "cantidad_filas_catalogo": len(catalog_rows),
        }
    return None


def parse_lookup_source(url: str, label: str) -> dict:
    html_text: str | None = None
    page_title: str | None = None
    search_types: list[str] = []
    form_action: str | None = None
    query_input: str | None = None
    submit_control: str | None = None
    mode = "external_link"

    try:
        html_text = fetch_bytes(url).decode("utf-8", errors="replace")
    except Exception:
        html_text = None

    if html_text:
        title_match = re.search(r"<title>\s*(.*?)\s*</title>", html_text, flags=re.S | re.I)
        if title_match:
            page_title = normalize_text(re.sub(r"<[^>]+>", " ", title_match.group(1)))

        action_match = re.search(r'<form[^>]*action="([^"]+)"', html_text, flags=re.I)
        if action_match:
            form_action = urljoin(url, action_match.group(1))

        input_match = re.search(r'<input[^>]*name="([^"]+)"[^>]*id="txtSearch"', html_text, flags=re.I)
        if input_match:
            query_input = input_match.group(1)

        submit_match = re.search(r'<input[^>]*id="([^"]*SearchButton[^"]*)"', html_text, flags=re.I)
        if submit_match:
            submit_control = submit_match.group(1)

        radio_block = re.search(r'id="rblSearchType"[^>]*>(.*?)</span>', html_text, flags=re.S | re.I)
        if radio_block:
            search_types = [
                normalize_text(text)
                for text in re.findall(r"<label[^>]*>(.*?)</label>", radio_block.group(1), flags=re.S | re.I)
                if normalize_text(text)
            ]
        if query_input or search_types:
            mode = "search_form"

    return {
        "etiqueta": label,
        "url": url,
        "modo": mode,
        "titulo_pagina": page_title,
        "accion_formulario": form_action,
        "campo_busqueda": query_input,
        "control_envio": submit_control,
        "tipos_busqueda": search_types,
    }


def parse_catalog_sections(lines: list[str]) -> list[dict]:
    sections: list[dict] = []
    i = 0
    while i < len(lines):
        if lines[i] != "Catálogo de productos y servicios para la facturación":
            i += 1
            continue
        i += 1
        while i < len(lines) and normalize_key(lines[i]) in {normalize_key(h) for h in CATALOG_HEADER_LINES}:
            i += 1
        if i >= len(lines):
            break
        title = lines[i]
        i += 1
        while i < len(lines) and normalize_key(lines[i]) in {normalize_key(h) for h in CATALOG_HEADER_LINES}:
            i += 1

        rows: list[dict[str, str]] = []
        while i < len(lines):
            line = lines[i]
            if line == "Catálogo de productos y servicios para la facturación":
                break
            if line in SECTION_KEYS and line != "Catálogo de productos y servicios para la facturación":
                break
            if normalize_key(line) in {normalize_key(h) for h in CATALOG_HEADER_LINES}:
                i += 1
                continue
            if not re.fullmatch(r"\d{8}", line):
                i += 1
                continue
            clave = line
            descripcion = lines[i + 1] if i + 1 < len(lines) else ""
            palabras = ""
            if i + 2 < len(lines):
                candidate = lines[i + 2]
                if not re.fullmatch(r"\d{8}", candidate) and candidate not in SECTION_KEYS and candidate != "Catálogo de productos y servicios para la facturación":
                    palabras = candidate
            rows.append({
                "clave_prodserv": clave,
                "descripcion": descripcion,
                "palabras_similares": palabras,
            })
            i += 2 + (1 if palabras else 0)

        sections.append({"titulo": title, "filas": rows})
    return sections


def pick_catalog_section(slug: str, title: str, sections: list[dict]) -> dict | None:
    if not sections:
        return None

    wanted = [normalize_key(title), normalize_key(slug.replace("-", " "))]
    for alias in CATALOG_TITLE_ALIASES.get(slug, []):
        wanted.append(normalize_key(alias))

    for section in sections:
        section_key = normalize_key(section["titulo"])
        if section_key in wanted:
            return section

    for section in sections:
        section_key = normalize_key(section["titulo"])
        if any(section_key in candidate or candidate in section_key for candidate in wanted if candidate):
            return section

    return None


def serialize_section(lines: list[str]) -> str | None:
    return "\n".join(lines) if lines else None


def filter_page_links(links: list[dict[str, str]], source_url: str) -> list[dict[str, str]]:
    best_by_url: dict[str, dict[str, str]] = {}
    for link in links:
        label = link["etiqueta"]
        url = link["url"]
        if url == source_url:
            continue
        if label in SECTION_KEYS:
            continue
        if label in SECTION_HEADINGS:
            continue
        if label in {"Trámites y servicios", "Minisitios"}:
            continue
        parts = urlsplit(url)
        if (
            parts.netloc == "www.sat.gob.mx"
            and parts.path.startswith("/minisitio/DeduccionesPersonales/")
            and parts.path.endswith(".html")
        ):
            continue
        current = best_by_url.get(url)
        if current is None or link_label_score(label) > link_label_score(current["etiqueta"]):
            best_by_url[url] = {"etiqueta": label, "url": url}
    return list(best_by_url.values())


def extract_page(source_url: str, html_text: str, slug_hint: str | None = None, title_hint: str | None = None) -> dict:
    parser = LineHTMLParser(source_url)
    parser.feed(html_text)
    parser.close()
    links = [{"etiqueta": link["label"], "url": link["url"]} for link in dedupe_links(parser.links)]
    lines, title_tag = parser.lines(), parser.title
    title, intro, sections = split_sections(lines, title_hint=title_hint)
    title = title or title_hint or title_tag or slug_hint or ""
    slug = slug_hint or slugify(Path(source_url).stem)

    specific_catalog = pick_catalog_section(slug, title, parse_catalog_sections(lines))
    catalog_rows = specific_catalog["filas"] if specific_catalog else []

    page_links = filter_page_links(links, source_url)
    extra_links = [
        link for link in page_links
        if link["etiqueta"] not in {"Image", "Trámites y servicios", "Minisitios"}
        and link["url"] not in EXCLUDED_EXTRA_LINK_URLS
        and "/minisitio/DeduccionesPersonales/" not in link["url"]
    ]
    lookup_sources = [
        parse_lookup_source(link["url"], link["etiqueta"])
        for link in extra_links
        if link["etiqueta"] in LOOKUP_LINK_LABELS
    ]
    factura_requirements = sections.get("factura_requirements", [])
    cfdi_use_requirements = parse_cfdi_use_requirements(factura_requirements)
    product_service_requirement = parse_product_service_requirement(
        factura_requirements,
        specific_catalog["titulo"] if specific_catalog else None,
        catalog_rows,
    )

    payload = {
        "slug": slug,
        "titulo": title,
        "url_fuente": source_url,
        "resumen": serialize_section(intro),
        "quien_puede_aplicar": None,
        "requisitos_factura": factura_requirements,
        "gastos_deducibles": sections.get("deductible_items", []),
        "gastos_no_deducibles": sections.get("non_deductible_items", []),
        "consideraciones": sections.get("considerations", []),
        "limites": parse_limit(sections.get("limit_text", [])),
        "facturacion_json": {
            "claves_uso_cfdi": parse_cfdi_use_codes(factura_requirements),
            "requisitos_uso_cfdi": cfdi_use_requirements,
            "clave_prodserv": {
                **(product_service_requirement or {}),
                "filas_catalogo": catalog_rows,
            },
        },
        "enlaces_adicionales_json": extra_links,
        "fuentes_consulta_json": lookup_sources,
    }
    payload["hash_contenido"] = json_hash(payload)
    return payload


def write_csvs(manifest: dict) -> None:
    HF_DIR.mkdir(parents=True, exist_ok=True)
    scraped_at = manifest["fecha_extraccion"]
    pages = sorted(manifest.get("paginas", []), key=lambda item: item["slug"])

    deducciones_path = HF_DIR / "deducciones.csv"
    with deducciones_path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(
            fh,
            fieldnames=[
                "fecha_extraccion",
                "slug",
                "titulo",
                "url_fuente",
                "resumen",
                "quien_puede_aplicar",
                "requisitos_factura_json",
                "gastos_deducibles_json",
                "gastos_no_deducibles_json",
                "consideraciones_json",
                "limites",
                "facturacion_json",
                "fuentes_consulta_json",
                "enlaces_adicionales_json",
                "hash_contenido",
                "ultima_revision",
            ],
        )
        writer.writeheader()
        for page in pages:
            writer.writerow({
                "fecha_extraccion": scraped_at,
                "slug": page["slug"],
                "titulo": page["titulo"],
                "url_fuente": page["url_fuente"],
                "resumen": page.get("resumen") or "",
                "quien_puede_aplicar": page.get("quien_puede_aplicar") or "",
                "requisitos_factura_json": json.dumps(page.get("requisitos_factura", []), ensure_ascii=False),
                "gastos_deducibles_json": json.dumps(page.get("gastos_deducibles", []), ensure_ascii=False),
                "gastos_no_deducibles_json": json.dumps(page.get("gastos_no_deducibles", []), ensure_ascii=False),
                "consideraciones_json": json.dumps(page.get("consideraciones", []), ensure_ascii=False),
                "limites": json.dumps(page.get("limites"), ensure_ascii=False),
                "facturacion_json": json.dumps(page.get("facturacion_json"), ensure_ascii=False),
                "fuentes_consulta_json": json.dumps(page.get("fuentes_consulta_json", []), ensure_ascii=False),
                "enlaces_adicionales_json": json.dumps(page.get("enlaces_adicionales_json", []), ensure_ascii=False),
                "hash_contenido": page["hash_contenido"],
                "ultima_revision": page["ultima_revision"],
            })

    catalog_path = HF_DIR / "catalogo_productos_servicios.csv"
    with catalog_path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(
            fh,
            fieldnames=[
                "fecha_extraccion",
                "deduccion_slug",
                "deduccion_titulo",
                "titulo_seccion_catalogo",
                "clave_prodserv",
                "descripcion",
                "palabras_similares",
                "url_fuente",
            ],
        )
        writer.writeheader()
        for page in pages:
            facturacion = page.get("facturacion_json") or {}
            prodserv = facturacion.get("clave_prodserv") or {}
            for row in prodserv.get("filas_catalogo", []):
                writer.writerow({
                    "fecha_extraccion": scraped_at,
                    "deduccion_slug": page["slug"],
                    "deduccion_titulo": page["titulo"],
                    "titulo_seccion_catalogo": prodserv.get("titulo_seccion_catalogo") or "",
                    "clave_prodserv": row.get("clave_prodserv", ""),
                    "descripcion": row.get("descripcion", ""),
                    "palabras_similares": row.get("palabras_similares", ""),
                    "url_fuente": page["url_fuente"],
                })

    lookup_path = HF_DIR / "lookup_sources.csv"
    with lookup_path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(
            fh,
            fieldnames=[
                "fecha_extraccion",
                "deduccion_slug",
                "deduccion_titulo",
                "etiqueta",
                "url",
                "modo",
                "titulo_pagina",
                "accion_formulario",
                "campo_busqueda",
                "control_envio",
                "tipos_busqueda_json",
                "url_fuente",
            ],
        )
        writer.writeheader()
        for page in pages:
            for lookup in page.get("fuentes_consulta_json", []):
                writer.writerow({
                    "fecha_extraccion": scraped_at,
                    "deduccion_slug": page["slug"],
                    "deduccion_titulo": page["titulo"],
                    "etiqueta": lookup.get("etiqueta", ""),
                    "url": lookup.get("url", ""),
                    "modo": lookup.get("modo", ""),
                    "titulo_pagina": lookup.get("titulo_pagina", "") or "",
                    "accion_formulario": lookup.get("accion_formulario", "") or "",
                    "campo_busqueda": lookup.get("campo_busqueda", "") or "",
                    "control_envio": lookup.get("control_envio", "") or "",
                    "tipos_busqueda_json": json.dumps(lookup.get("tipos_busqueda", []), ensure_ascii=False),
                    "url_fuente": page["url_fuente"],
                })

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--force", action="store_true", help="Rewrite page HTML and manifest even if content is unchanged")
    parser.add_argument("--dry-run", action="store_true", help="Print the extracted records instead of writing files")
    parser.add_argument("--slug", help="Fetch and process a single deduction slug from the discovered index")
    args = parser.parse_args()

    try:
        index_bytes = fetch_bytes(INDEX_URL)
    except (URLError, HTTPError, TimeoutError) as exc:
        print(f"ERROR fetching {INDEX_URL}: {exc}", file=sys.stderr)
        return 1

    index_html = index_bytes.decode("utf-8", errors="replace")
    index_data = parse_index(index_html)
    discovered = index_data["deducciones"]
    if args.slug:
        discovered = [entry for entry in discovered if entry["slug"] == args.slug]
        if not discovered:
            print(f"No deduction with slug '{args.slug}' found in index.", file=sys.stderr)
            return 1

    prev_manifest = load_manifest()
    prev_pages = {page["slug"]: page for page in prev_manifest.get("paginas", [])}
    now = datetime.now(timezone.utc).isoformat()

    PAGES_DIR.mkdir(parents=True, exist_ok=True)
    pages_out: dict[str, dict] = {}
    changed = False

    if args.slug and prev_pages:
        pages_out.update({slug: page for slug, page in prev_pages.items() if slug != args.slug})

    for entry in discovered:
        slug = entry["slug"]
        url = entry["url"]
        print(f"Fetching {url}", file=sys.stderr)
        try:
            page_bytes = fetch_bytes(url)
        except (URLError, HTTPError, TimeoutError) as exc:
            if slug in prev_pages:
                print(f"  WARN: failed to fetch {slug}, preserving previous manifest data: {exc}", file=sys.stderr)
                pages_out[slug] = prev_pages[slug]
                continue
            print(f"  ERROR: failed to fetch {slug}: {exc}", file=sys.stderr)
            return 1

        html_text = page_bytes.decode("utf-8", errors="replace")
        extracted = extract_page(url, html_text, slug_hint=slug, title_hint=entry["titulo"])
        previous = prev_pages.get(slug)
        html_hash = hashlib.sha256(page_bytes).hexdigest()

        page_record = {
            **extracted,
            "hash_html": html_hash,
            "archivo_html": f"pages/{slug}.html",
            "ultima_revision": now,
        }

        if previous and previous.get("hash_contenido") == page_record["hash_contenido"] and not args.force:
            previous = dict(previous)
            previous["ultima_revision"] = now
            pages_out[slug] = previous
            print(f"  unchanged: {slug}", file=sys.stderr)
            continue

        changed = True
        (PAGES_DIR / f"{slug}.html").write_text(html_text, encoding="utf-8")
        pages_out[slug] = page_record
        print(f"  updated: {slug}", file=sys.stderr)

    if not args.slug:
        for slug, page in prev_pages.items():
            pages_out.setdefault(slug, page)

    pages = sorted(pages_out.values(), key=lambda item: item["slug"])
    manifest = {
        "fecha_extraccion": now,
        "url_fuente": INDEX_URL,
        "indice": index_data,
        "paginas": pages,
    }

    if args.dry_run:
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
        return 0

    old_hash = json_hash(stable_manifest_view(prev_manifest)) if prev_manifest else None
    new_hash = json_hash(stable_manifest_view(manifest))
    if not args.force and old_hash == new_hash:
        print("No content changes detected; skipping manifest and CSV rewrite.", file=sys.stderr)
        return 0

    MANIFEST.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csvs(manifest)
    print(f"Manifest -> {MANIFEST}", file=sys.stderr)
    print(f"Rows -> {len(pages)} deductions", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
