"""CI tests that need no LibreOffice.

1. generate.py runs on every example config (catches syntax/runtime breakage).
2. Freshly generated formulas match the shipped, recalculated workbooks.
3. check_workbook.py --skip-recalc PASSes on the shipped workbooks (they carry caches).
   Its live-input tests recalculate with LibreOffice, so this is skipped if `soffice`
   is not on PATH (the CI workflow installs it).
4. check_workbook.py FAILs on a fresh, uncached workbook (empty-cache guard).

The slower --regression mutation run is left to local use (see README).
"""
import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
CASES = {
    "default_illustrative": ROOT / "templates" / "stress_template.xlsx",
    "gold_miner": ROOT / "examples" / "workbooks" / "gold_miner.xlsx",
    "apparel_retailer": ROOT / "examples" / "workbooks" / "apparel_retailer.xlsx",
}


def _run(*args):
    return subprocess.run([sys.executable, *map(str, args)], cwd=ROOT,
                          capture_output=True, text=True)


def _config(name):
    return ROOT / "examples" / "configs" / f"{name}.yaml"


@pytest.fixture(scope="module")
def generated(tmp_path_factory):
    out = tmp_path_factory.mktemp("generated")
    paths = {}
    for name in CASES:
        path = out / f"{name}.xlsx"
        res = _run("scripts/generate.py", "--config", _config(name), "--out", path)
        assert res.returncode == 0, res.stdout + res.stderr
        paths[name] = path
    return paths


@pytest.mark.parametrize("name", list(CASES))
def test_generated_formulas_match_shipped(generated, name):
    fresh = openpyxl.load_workbook(generated[name])
    shipped = openpyxl.load_workbook(CASES[name])
    assert fresh.sheetnames == shipped.sheetnames
    diffs = []
    for ws in fresh.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    other = shipped[ws.title][cell.coordinate].value
                    if cell.value != other:
                        diffs.append(f"{ws.title}!{cell.coordinate}")
    assert not diffs, diffs[:20]


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice (soffice) not installed")
@pytest.mark.parametrize("name", list(CASES))
def test_checker_passes_on_shipped(name):
    res = _run("scripts/check_workbook.py", CASES[name], "--config", _config(name), "--skip-recalc")
    assert res.returncode == 0, res.stdout + res.stderr
    assert res.stdout.strip().splitlines()[-1] == "PASS"


def test_checker_rejects_uncached(generated):
    res = _run("scripts/check_workbook.py", generated["gold_miner"],
               "--config", _config("gold_miner"), "--skip-recalc")
    assert res.returncode != 0
    assert "Empty formula caches" in res.stdout + res.stderr
