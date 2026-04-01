#!/usr/bin/env python3
"""Compatibility wrapper for the merged catalog extractor.

Matrix extraction is now handled by scripts/catalogos/extract.py, which writes both
catalog and matrix entries into catalog_state.csv using shared export rules.
"""
from __future__ import annotations

import sys

from scripts.catalogos.extract import main as extract_main


if __name__ == "__main__":
    print(
        "extract_matrices.py is deprecated; running scripts/catalogos/extract.py instead.",
        file=sys.stderr,
    )
    raise SystemExit(extract_main(sys.argv[1:]))
