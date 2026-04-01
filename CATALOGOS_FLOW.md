# Catalogos Flow

This file documents the current `catalogos` pipeline and the source-of-truth files.

## Source Of Truth

- [output/catalog.csv](/Users/mayravaldes/Mayrop/cfdi/catalogos/output/catalog.csv)
  Scrape inventory. This is the source of truth for what SAT files were discovered and downloaded.
- [catalog_state.csv](/Users/mayravaldes/Mayrop/cfdi/catalogos/catalog_state.csv)
  Extracted catalog inventory. This is the source of truth for publishable catalog and matrix CSVs.

`output/catalogos-manifest.json` is no longer the operational source of truth. It is legacy/debug output and can be removed once the remaining code paths that still write it are cleaned up.

## Directory Roles

- [hf/raw/catalogos](/Users/mayravaldes/Mayrop/cfdi/catalogos/hf/raw/catalogos)
  Downloaded raw source files from SAT.
- [hf/csv](/Users/mayravaldes/Mayrop/cfdi/catalogos/hf/csv)
  Extracted CSV files.
- [hf/dataset/catalogos](/Users/mayravaldes/Mayrop/cfdi/catalogos/hf/dataset/catalogos)
  Final Hugging Face dataset layout.

## What Writes To What

- [scripts/catalogos/scrape.py](/Users/mayravaldes/Mayrop/cfdi/catalogos/scripts/catalogos/scrape.py)
  Writes:
  [output/catalog.csv](/Users/mayravaldes/Mayrop/cfdi/catalogos/output/catalog.csv)
  [output/catalogos-manifest.json](/Users/mayravaldes/Mayrop/cfdi/catalogos/output/catalogos-manifest.json)
  [hf/raw/catalogos](/Users/mayravaldes/Mayrop/cfdi/catalogos/hf/raw/catalogos)

- [scripts/catalogos/extract.py](/Users/mayravaldes/Mayrop/cfdi/catalogos/scripts/catalogos/extract.py)
  Reads:
  [output/catalog.csv](/Users/mayravaldes/Mayrop/cfdi/catalogos/output/catalog.csv)
  [hf/raw/catalogos](/Users/mayravaldes/Mayrop/cfdi/catalogos/hf/raw/catalogos)
  Writes:
  [hf/csv](/Users/mayravaldes/Mayrop/cfdi/catalogos/hf/csv)
  [catalog_state.csv](/Users/mayravaldes/Mayrop/cfdi/catalogos/catalog_state.csv)

- [scripts/catalogos/generate_hf.py](/Users/mayravaldes/Mayrop/cfdi/catalogos/scripts/catalogos/generate_hf.py)
  Reads:
  [catalog_state.csv](/Users/mayravaldes/Mayrop/cfdi/catalogos/catalog_state.csv)
  [hf/csv](/Users/mayravaldes/Mayrop/cfdi/catalogos/hf/csv)
  [hf/raw/catalogos](/Users/mayravaldes/Mayrop/cfdi/catalogos/hf/raw/catalogos)
  [output/catalog.csv](/Users/mayravaldes/Mayrop/cfdi/catalogos/output/catalog.csv)
  Writes:
  [hf/dataset/catalogos](/Users/mayravaldes/Mayrop/cfdi/catalogos/hf/dataset/catalogos)

## Execution Order

1. Run [scripts/catalogos/scrape.py](/Users/mayravaldes/Mayrop/cfdi/catalogos/scripts/catalogos/scrape.py)
2. Run [scripts/catalogos/extract.py](/Users/mayravaldes/Mayrop/cfdi/catalogos/scripts/catalogos/extract.py)
3. Run [scripts/catalogos/generate_hf.py](/Users/mayravaldes/Mayrop/cfdi/catalogos/scripts/catalogos/generate_hf.py)

## Notes

- `output/catalog.csv` now includes a `sub` column so entries like `hidrocarburos` `gastos` and `ingresos` do not collapse into the same visible key.
- The state files are keyed by logical identity, not by `source_xls`, to avoid stale-path duplication after raw path migrations.
- Matrices are now stored in the same state file as normal catalogs, using `file_type=matriz`.
