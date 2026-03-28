#!/usr/bin/env python3
"""Scrape SAT Artículo 69-B and 69-B Bis 'Listado completo' CSV files.

Source page:
  http://omawww.sat.gob.mx/cifras_sat/Paginas/DatosAbiertos/contribuyentes_publicados.html

Discovers download URLs from the page, uses Last-Modified + ETag as fingerprints,
and only re-downloads when content has changed (or --force).

Output:
  output/listado-69b-manifest.json                        — metadata + hashes
  output/files/listado-69b/Listado_completo_69-B.csv      — raw (original filename)
  output/files/listado-69b/Listado_69_B_Bis_Completo.csv  — raw (original filename)

Run scripts/listado_69b/merge.py afterwards to clean and merge into hf/csv/listado-69b/.

Usage:
  uv run scripts/listado_69b/scrape.py
  uv run scripts/listado_69b/scrape.py --force
  uv run scripts/listado_69b/scrape.py --verify
"""
from __future__ import annotations

import argparse
import hashlib
import html.parser
import json
import re
import sys
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from urllib.error import HTTPError, URLError

# ── constants ──────────────────────────────────────────────────────────────────

PAGE_URL   = "http://omawww.sat.gob.mx/cifras_sat/Paginas/DatosAbiertos/contribuyentes_publicados.html"
OUTPUT_DIR = Path("output")
RAW_DIR    = OUTPUT_DIR / "files" / "listado-69b"
MANIFEST   = OUTPUT_DIR / "listado-69b-manifest.json"

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# Sections we care about: section_id → (article key, display name)
_TARGET_SECTIONS: dict[str, tuple[str, str]] = {
    "collapseTwo2": ("69b",     "Artículo 69-B"),
    "collapseTwo3": ("69b-bis", "Artículo 69-B Bis"),
}

_MONTH_ES: dict[str, int] = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
}


# ── HTML parsing ───────────────────────────────────────────────────────────────


def _parse_info_date(text: str) -> str | None:
    """Extract 'Información actualizada al DD de MMMM de YYYY' → 'YYYY-MM-DD'."""
    m = re.search(
        r"[Ii]nformaci[oó]n\s+actualizada\s+al\s+(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})",
        text,
    )
    if not m:
        return None
    day, month_es, year = m.group(1), m.group(2).lower(), m.group(3)
    month = _MONTH_ES.get(month_es)
    if not month:
        return None
    return f"{year}-{month:02d}-{int(day):02d}"


class _PageParser(html.parser.HTMLParser):
    """Extract 'Listado completo' URL for each target section."""

    def __init__(self) -> None:
        super().__init__()
        self._current_section: str | None = None
        self._pending_href: str | None = None
        self.sections: dict[str, dict] = {sid: {} for sid in _TARGET_SECTIONS}

    def handle_starttag(self, tag: str, attrs: list) -> None:
        attr = dict(attrs)
        if tag == "div" and attr.get("id") in _TARGET_SECTIONS:
            self._current_section = attr["id"]
        if tag == "a" and self._current_section and "href" in attr:
            self._pending_href = attr["href"]
        else:
            self._pending_href = None

    def handle_data(self, data: str) -> None:
        if self._pending_href and self._current_section and "listado completo" in data.lower():
            self.sections[self._current_section]["url"] = self._pending_href


def _scrape_page() -> dict[str, dict]:
    """Fetch the source page and return {section_id: {url, info_date, name, key}}."""
    print(f"Fetching {PAGE_URL}", file=sys.stderr)
    req = urllib.request.Request(PAGE_URL, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            html_bytes = r.read()
    except (URLError, HTTPError) as exc:
        print(f"  ERROR fetching page: {exc}", file=sys.stderr)
        return {}

    html_text = html_bytes.decode("utf-8", errors="replace")

    parser = _PageParser()
    parser.feed(html_text)
    parser.close()

    # Extract info_date per section by searching a bounded snippet of the HTML.
    # This keeps each section's date isolated from the others.
    result = {}
    for sid, (key, name) in _TARGET_SECTIONS.items():
        entry = dict(parser.sections.get(sid, {}))
        entry["key"]  = key
        entry["name"] = name
        idx = html_text.find(f'id="{sid}"')
        if idx < 0:
            idx = html_text.find(f"id='{sid}'")
        if idx >= 0:
            d = _parse_info_date(html_text[idx: idx + 3000])
            if d:
                entry["info_date"] = d
        result[sid] = entry

    return result


# ── HTTP helpers ───────────────────────────────────────────────────────────────


def _sha256(path: Path) -> str | None:
    if not path.exists():
        return None
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _fetch_bytes(url: str) -> bytes | None:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.read()
    except (URLError, HTTPError, Exception) as exc:
        print(f"    ERROR fetching {url}: {exc}", file=sys.stderr)
        return None


def _head(url: str) -> dict[str, str]:
    """Return response headers for a HEAD request (lowercase keys)."""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT},
                                     method="HEAD")
        with urllib.request.urlopen(req, timeout=20) as r:
            return {k.lower(): v for k, v in r.headers.items()}
    except Exception:
        return {}


def _fingerprint(headers: dict[str, str]) -> str:
    """Stable fingerprint string from Last-Modified + ETag."""
    parts = [headers.get("last-modified", ""), headers.get("etag", "")]
    return "|".join(p for p in parts if p)


# ── download ───────────────────────────────────────────────────────────────────


def download(
    url: str,
    dest: Path,
    stored_fingerprint: str | None = None,
    stored_hash: str | None = None,
    force: bool = False,
    verify: bool = False,
) -> tuple[str, dict[str, str]]:
    """Download url to dest as-is, skip when unchanged.

    Returns (status, headers) where status is one of:
      'written'   — fetched and saved
      'unchanged' — fetched, hash identical, not written
      'skipped'   — fingerprint matched, fetch skipped entirely
      'error'     — fetch failed
    """
    dest.parent.mkdir(parents=True, exist_ok=True)

    # HEAD to get fingerprint (cheap; skip only in force mode)
    headers = _head(url) if not force else {}
    fp = _fingerprint(headers)

    # Fingerprint fast-path
    if not force and not verify and fp and stored_fingerprint and fp == stored_fingerprint:
        if dest.exists() or stored_hash:
            return "skipped", headers

    # Full download
    data = _fetch_bytes(url)
    if data is None:
        return "error", headers

    new_hash = hashlib.sha256(data).hexdigest()
    if not force:
        prior_hash = stored_hash or _sha256(dest)
        if new_hash == prior_hash and (verify or dest.exists()):
            return "unchanged", headers

    dest.write_bytes(data)
    return "written", headers


# ── manifest ───────────────────────────────────────────────────────────────────


def _load_manifest() -> dict:
    if MANIFEST.exists():
        try:
            return json.loads(MANIFEST.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _prev_article(manifest: dict, key: str) -> dict:
    for a in manifest.get("articles", []):
        if a.get("key") == key:
            return a
    return {}


# ── main ───────────────────────────────────────────────────────────────────────


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--force",  action="store_true",
                        help="Re-download and overwrite all files regardless of fingerprint or hash")
    parser.add_argument("--verify", action="store_true",
                        help="Fetch every file and compare hash; write only if changed")
    parser.add_argument("--dry-run", action="store_true",
                        help="Scrape page and print URLs, don't download")
    args = parser.parse_args()

    sections = _scrape_page()
    if not sections:
        print("Could not retrieve page. Exiting.", file=sys.stderr)
        return 1

    if args.dry_run:
        print(json.dumps(sections, ensure_ascii=False, indent=2))
        return 0

    prev_manifest = _load_manifest()
    articles_out: list[dict] = []

    for sid, info in sections.items():
        url = info.get("url", "")
        key  = info["key"]
        name = info["name"]

        print(f"\n{name}", file=sys.stderr)

        if not url:
            print(f"  ERROR: no 'Listado completo' URL found for {name}", file=sys.stderr)
            prev = _prev_article(prev_manifest, key)
            if prev:
                articles_out.append(prev)
            continue

        dest = RAW_DIR / Path(url).name
        prev = _prev_article(prev_manifest, key)
        stored_fp   = prev.get("fingerprint")
        stored_hash = prev.get("hash")

        status, headers = download(
            url, dest,
            stored_fingerprint=stored_fp,
            stored_hash=stored_hash,
            force=args.force,
            verify=args.verify,
        )

        rel = str(dest)
        if status == "written":
            print(f"  ↓  {rel}", file=sys.stderr)
        elif status == "skipped":
            print(f"  ✓  skip (fingerprint)  {rel}", file=sys.stderr)
        elif status == "unchanged":
            print(f"  ✓  skip (hash match)   {rel}", file=sys.stderr)
        elif status == "error":
            print(f"  ✗  error downloading   {rel}", file=sys.stderr)
            if prev:
                articles_out.append(prev)
            continue

        fp = _fingerprint(headers)
        entry: dict = {
            "_status":       status,
            "key":           key,
            "name":          name,
            "url":           url,
            "local_file":    rel,
            "fingerprint":   fp or stored_fp or "",
            "last_modified": headers.get("last-modified") or prev.get("last_modified") or "",
            "etag":          headers.get("etag") or prev.get("etag") or "",
            "info_date":     info.get("info_date") or prev.get("info_date") or "",
            "hash":          _sha256(dest) or stored_hash or "",
            "size":          headers.get("content-length") or prev.get("size") or "",
        }
        articles_out.append(entry)

    prev_hashes = {a["key"]: a.get("hash", "") for a in (prev_manifest or {}).get("articles", [])}
    any_changed = any(
        a.get("hash") and a.get("hash") != prev_hashes.get(a.get("key", ""))
        for a in articles_out
    )
    for a in articles_out:
        a.pop("_status", None)
    prev_scraped_at = prev_manifest.get("scraped_at") if prev_manifest else None
    scraped_at = datetime.now(timezone.utc).isoformat() if any_changed or not prev_scraped_at else prev_scraped_at

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    manifest = {
        "scraped_at": scraped_at,
        "source_url": PAGE_URL,
        "articles":   articles_out,
    }
    MANIFEST.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nManifest → {MANIFEST}", file=sys.stderr)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
