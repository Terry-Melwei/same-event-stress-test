#!/usr/bin/env python3
"""Deprecated shim — use scripts/generate.py instead."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from generate import main

if __name__ == "__main__":
    sys.exit(main([
        "--config", str(ROOT / "examples" / "configs" / "default_illustrative.yaml"),
        "--out", str(ROOT / "templates" / "stress_template.xlsx"),
        *sys.argv[1:],
    ]))
