#!/usr/bin/env python3
"""Automated workbook gate for same-event multi-channel stress tests.

Fails on: empty formula caches in the input file, pasted outputs, broken
FULL→tornado→envelope→football→chart-source→Summary links, event shock/sign misses,
pass-through invariant breaks, interaction rows that do not equal FULL minus the
standalone channels, mirror / reverse-stress grids that disagree with the independent
reference, a wrong reverse-stress result, a mis-ranked Summary, and (via --regression)
known broken mutations.
"""

from __future__ import annotations

import argparse
import zipfile
import xml.etree.ElementTree as ET
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

import yaml
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))
from reference_dcf import (
    format_not_reached,  # noqa: E402
    ladder_layout,
    reconcile_ladder,
    blended_mid,
    base_equity_perp,
    capex_components,
    configured_shock,
    full_at,
    load_inputs_from_workbook,
    reconcile_events,
    tariff_posture_equity,
)

# Blocks added to every event sheet (interaction, mirror, reverse grid, helpers)
EXTRA_RANGES = ["D6:G12", "C61:G63", "B67:C67", "C69:H76", "B80:I88", "B55"]
OUTPUT_RANGES = {
    "DCF": ["B19:B33", "B5:F16"],
    "Event_Rates": ["B23:B37"] + EXTRA_RANGES,
    "Event_Commodity": ["B43:B49", "B32:F32", "B37:F37", "B39:F41"] + EXTRA_RANGES,
    "Event_Tariff": ["B40:F40", "B41:B52"] + EXTRA_RANGES,
    "Event_Capex": ["B26:B27", "B30:F33", "B35:B49"] + EXTRA_RANGES,
    "Summary": ["B5:K7", "B9", "C14:K16"],
    "Tornado": ["B5:F7", "B8:B9", "B13:B15"],
    "Envelope": ["B5:C8", "B9:C10"],
    "Football": ["B5:D11", "F5:G11", "B12", "B14"],
    "PassThrough": ["B16:B19", "B22:B31", "B34:B37"],
}

EVENT_SHEETS = ("Event_Rates", "Event_Commodity", "Event_Tariff", "Event_Capex")
# Standalone channel rows summed by the interaction row (tariff uses its ladder form instead)
STANDALONE_ROWS = {
    "Event_Rates": [6, 8],
    "Event_Commodity": [6, 7, 8, 9],
    "Event_Capex": [6, 7, 8],
}
ROW_INT, ROW_CONF, ROW_MIRROR, ROW_ASYM = 12, 61, 62, 63
ROW_THRESHOLD, ROW_RESULT, ROW_GRID0, N_GRID = 66, 67, 69, 8
ROW_CALIB_SOURCE, ROW_CALIB_BASIS = 56, 57
COL_DROP = "D"  # yellow drop-% input next to the threshold (row 66)
INTERACTION_LABEL = "Cross-channel interaction (not size non-linearity)"
THRESHOLD_DEFAULT_FX = f'=IF({COL_DROP}{ROW_THRESHOLD}="","",DCF!$B$33*(1-{COL_DROP}{ROW_THRESHOLD}))'

# Required dependency chain cells (enumerated from workbook structure).
# second_event_sheet is resolved from the workbook (Event_Commodity or Event_Tariff).
CHAIN_TEMPLATE = {
    "rates_full": "Event_Rates!D11",
    "rates_pct": "Event_Rates!E11",
    "rates_proxy": "Event_Rates!F11",
    "rates_proxy_sh": "Event_Rates!G11",
    "tornado_rates_pct": "Tornado!B5",
    "tornado_rates_proxy": "Tornado!C5",
    "tornado_rates_sh": "Tornado!D5",
    "tornado_rates_delta": "Tornado!E5",
    "chart_rates_label": "Tornado!A13",
    "chart_rates_delta": "Tornado!B13",
    "envelope_rates": "Envelope!B6",
    "envelope_rates_sh": "Envelope!C6",
    "football_rates_low": "Football!B8",
    "football_rates_mid": "Football!C8",
    "football_rates_high": "Football!D8",
    "football_rates_base": "Football!F8",
    "football_rates_width": "Football!G8",
    "second_full": "{second}!D11",
    "second_pct": "{second}!E11",
    "second_proxy": "{second}!F11",
    "second_proxy_sh": "{second}!G11",
    "tornado_second_pct": "Tornado!B6",
    "tornado_second_proxy": "Tornado!C6",
    "tornado_second_sh": "Tornado!D6",
    "tornado_second_delta": "Tornado!E6",
    "chart_second_label": "Tornado!A14",
    "chart_second_delta": "Tornado!B14",
    "envelope_second": "Envelope!B7",
    "envelope_second_sh": "Envelope!C7",
    "football_second_low": "Football!B9",
    "football_second_mid": "Football!C9",
    "football_second_high": "Football!D9",
    "football_second_base": "Football!F9",
    "football_second_width": "Football!G9",
    "envelope_min": "Envelope!B9",
    "envelope_max": "Envelope!B10",
    "football_env_low": "Football!B10",
    "football_env_high": "Football!D10",
    "blend_mid": "DCF!B30",
    "blend_sh": "DCF!B33",
}


def parse_cell_ref(ref: str) -> tuple[str, str]:
    if "!" not in ref:
        raise ValueError(f"Expected Sheet!Cell, got {ref}")
    sh, cell = ref.split("!", 1)
    return sh, cell


def detect_second_event(path: Path) -> str:
    wb = load_workbook(path, data_only=False)
    sheets = wb.sheetnames
    wb.close()
    if "Event_Tariff" in sheets and "Event_Commodity" in sheets:
        # Prefer the one tornado actually links to
        wb = load_workbook(path, data_only=False)
        fx = str(wb["Tornado"]["B6"].value or "")
        wb.close()
        if "Event_Tariff" in fx:
            return "Event_Tariff"
        return "Event_Commodity"
    if "Event_Tariff" in sheets:
        return "Event_Tariff"
    if "Event_Commodity" in sheets:
        return "Event_Commodity"
    raise RuntimeError("Workbook has neither Event_Commodity nor Event_Tariff")


def has_capex(path: Path) -> bool:
    wb = load_workbook(path, read_only=True)
    out = "Event_Capex" in wb.sheetnames
    wb.close()
    return out


def event_sheets(path: Path) -> list[str]:
    """Enabled event sheets in Tornado order: rates, second event, optional capex."""
    sheets = ["Event_Rates", detect_second_event(path)]
    if has_capex(path):
        sheets.append("Event_Capex")
    return sheets


CAPEX_CHAIN = {
    "capex_full": "Event_Capex!D11",
    "capex_pct": "Event_Capex!E11",
    "capex_proxy_sh": "Event_Capex!G11",
    "tornado_capex_pct": "Tornado!B7",
    "tornado_capex_proxy": "Tornado!C7",
    "tornado_capex_sh": "Tornado!D7",
    "tornado_capex_delta": "Tornado!E7",
    "chart_capex_delta": "Tornado!B15",
    "envelope_capex": "Envelope!B8",
    "envelope_capex_sh": "Envelope!C8",
    "football_capex_low": "Football!B11",
    "football_capex_high": "Football!D11",
    "football_capex_width": "Football!G11",
}


def chain_refs(second: str, capex: bool = False) -> dict[str, str]:
    refs = {k: v.format(second=second) for k, v in CHAIN_TEMPLATE.items()}
    if capex:
        refs.update(CAPEX_CHAIN)
    return refs


def read_values(path: Path, refs: list[str]) -> dict[str, Any]:
    wb = load_workbook(path, data_only=True)
    out: dict[str, Any] = {}
    for ref in refs:
        sh, cell = parse_cell_ref(ref)
        out[ref] = wb[sh][cell].value
    wb.close()
    return out


def read_formulas(path: Path, refs: list[str]) -> dict[str, Any]:
    wb = load_workbook(path, data_only=False)
    out: dict[str, Any] = {}
    for ref in refs:
        sh, cell = parse_cell_ref(ref)
        out[ref] = wb[sh][cell].value
    wb.close()
    return out


def recalc_libreoffice(src: Path, out_dir: Path) -> Path:
    """Recalculate with LibreOffice into a separate out folder (never overwrite source)."""
    out_dir.mkdir(parents=True, exist_ok=True)
    in_dir = out_dir / "in"
    dest_dir = out_dir / "out"
    in_dir.mkdir(parents=True, exist_ok=True)
    dest_dir.mkdir(parents=True, exist_ok=True)
    inbound = in_dir / src.name
    shutil.copy2(src, inbound)
    for old in dest_dir.glob("*.xlsx"):
        old.unlink()
    proc = subprocess.run(
        [
            "soffice",
            "--headless",
            "--calc",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(dest_dir),
            str(inbound),
        ],
        capture_output=True,
        text=True,
    )
    result = dest_dir / src.name
    if proc.returncode != 0 or not result.exists():
        raise RuntimeError(
            f"LibreOffice recalc failed (code={proc.returncode}): {proc.stderr or proc.stdout}"
        )
    return result


def check_input_caches(path: Path) -> tuple[bool, str]:
    """FAIL when the file under test has empty formula caches."""
    required = ["DCF!B30", "DCF!B28", "Event_Rates!D11", "Tornado!C5", "Envelope!B9", "Football!B8"]
    for sheet in event_sheets(path)[1:]:
        required.append(f"{sheet}!D11")
    wb = load_workbook(path, read_only=True)
    if "Summary" in wb.sheetnames:
        required.append("Summary!E5")
    wb.close()
    vals = read_values(path, required)
    missing = [r for r, v in vals.items() if v is None]
    if missing:
        return False, (
            "Empty formula caches in input file: "
            + ", ".join(missing)
            + ". Recalculate into a separate outdir (e.g. build/recalculated/) and check THAT artifact."
        )
    return True, f"Input caches present (DCF!B30={vals['DCF!B30']})"


def scan_formula_errors(path: Path) -> list[str]:
    wb = load_workbook(path, data_only=True)
    errors = []
    for sh in wb.sheetnames:
        for row in wb[sh].iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("#"):
                    errors.append(f"{sh}!{cell.coordinate}={v}")
    wb.close()
    return errors


def _expand_range(ws, range_str: str):
    if ":" in range_str:
        cells = ws[range_str]
        if isinstance(cells, tuple):
            for row in cells:
                if isinstance(row, tuple):
                    yield from row
                else:
                    yield row
        else:
            yield cells
    else:
        yield ws[range_str]


def scan_pasted_constants(path: Path) -> list[str]:
    wb = load_workbook(path, data_only=False)
    hits = []
    for sh, ranges in OUTPUT_RANGES.items():
        if sh not in wb.sheetnames:
            continue
        ws = wb[sh]
        for rng in ranges:
            for cell in _expand_range(ws, rng):
                v = cell.value
                if v is None or v == "":
                    continue
                if isinstance(v, (int, float)):
                    hits.append(f"{sh}!{cell.coordinate} pasted number {v}")
                elif isinstance(v, str) and not v.startswith("="):
                    if sh == "PassThrough" and cell.coordinate == "B37":
                        continue
                    if re.fullmatch(r"-?\d+(\.\d+)?", v.strip()):
                        hits.append(f"{sh}!{cell.coordinate} pasted numeric text {v}")
    wb.close()
    return hits


def verify_dependency_chain(path: Path) -> tuple[list[str], list[str]]:
    """Trace FULL → tornado → envelope → football → chart-source; fail on constants/broken refs."""
    failures: list[str] = []
    details: list[str] = []
    second = detect_second_event(path)
    capex = has_capex(path)
    refs = chain_refs(second, capex)
    formulas = read_formulas(path, list(refs.values()))

    def fx(key: str) -> str:
        return str(formulas[refs[key]] or "")

    def must_be_formula(key: str, must_contain: list[str] | None = None) -> None:
        cell = refs[key]
        val = formulas[cell]
        if not isinstance(val, str) or not val.startswith("="):
            failures.append(f"Chain break: {cell} is not a formula ({val!r})")
            return
        if must_contain:
            for token in must_contain:
                if token not in val:
                    failures.append(f"Chain break: {cell}={val!r} missing ref {token}")

    # Rates chain
    must_be_formula("rates_full")
    must_be_formula("rates_pct", ["D11", "DCF"])
    must_be_formula("rates_proxy", ["DCF"])
    must_be_formula("tornado_rates_pct", ["Event_Rates!E11"])
    must_be_formula("tornado_rates_proxy", ["Event_Rates!F11"])
    must_be_formula("tornado_rates_sh", ["Event_Rates!G11"])
    must_be_formula("tornado_rates_delta", ["C5"])
    must_be_formula("chart_rates_delta", ["E5"])
    must_be_formula("envelope_rates", ["Tornado!C5"])
    must_be_formula("football_rates_low", ["Tornado!D5"])
    must_be_formula("football_rates_high", ["Tornado!D5"])
    must_be_formula("football_rates_width", ["D8", "B8"])

    # Second event chain
    must_be_formula("second_full")
    # FULL must not be a trivial alias of base equity (shock removed)
    sec_full = fx("second_full")
    if re.fullmatch(r"=?DCF!B28", sec_full.replace(" ", "")) or sec_full.strip() in (
        "=DCF!B28",
        "DCF!B28",
    ):
        failures.append(f"Chain break: {refs['second_full']} collapses to base equity ({sec_full})")
    # Heuristic: FULL formula should reference helpers on its own sheet (B…) or Inputs shock
    if second == "Event_Commodity" and "B47" not in sec_full and "B4" not in sec_full:
        # Allow =B47 (equity FULL cumulative)
        if not re.search(r"B\d+", sec_full):
            failures.append(f"Chain break: {refs['second_full']} looks disconnected: {sec_full}")
    must_be_formula("tornado_second_pct", [f"{second}!E11"])
    must_be_formula("tornado_second_proxy", [f"{second}!F11"])
    must_be_formula("tornado_second_sh", [f"{second}!G11"])
    must_be_formula("tornado_second_delta", ["C6"])
    must_be_formula("chart_second_delta", ["E6"])
    # Chart source must not be a pasted constant
    chart_delta = formulas[refs["chart_second_delta"]]
    if isinstance(chart_delta, (int, float)):
        failures.append(
            f"Chain break: {refs['chart_second_delta']} is constant {chart_delta} (expected =E6)"
        )
    must_be_formula("envelope_second", ["Tornado!C6"])
    must_be_formula("football_second_low", ["Tornado!D6"])
    must_be_formula("football_second_high", ["Tornado!D6"])
    must_be_formula("envelope_min", ["MIN"])
    must_be_formula("envelope_max", ["MAX"])
    must_be_formula("football_env_low", ["Envelope"])
    must_be_formula("football_env_high", ["Envelope"])

    # Tornado %Δ must come from FULL rows, not hard-coded zeros
    for key in ("tornado_rates_pct", "tornado_second_pct", "chart_rates_delta", "chart_second_delta"):
        val = formulas[refs[key]]
        if val == 0 or val == 0.0 or val == "0":
            failures.append(f"Chain break: {refs[key]} is constant 0")

    if capex:
        must_be_formula("capex_full")
        cx_full = fx("capex_full").replace(" ", "")
        if cx_full in ("=DCF!B28", "=B46"):
            failures.append(f"Chain break: Event_Capex!D11 = {cx_full} drops channels (financing mix / all)")
        must_be_formula("tornado_capex_pct", ["Event_Capex!E11"])
        must_be_formula("tornado_capex_proxy", ["Event_Capex!F11"])
        must_be_formula("tornado_capex_sh", ["Event_Capex!G11"])
        must_be_formula("tornado_capex_delta", ["C7"])
        must_be_formula("chart_capex_delta", ["E7"])
        must_be_formula("envelope_capex", ["Tornado!C7"])
        must_be_formula("envelope_capex_sh", ["Tornado!D7"])
        must_be_formula("football_capex_low", ["Tornado!D7"])
        must_be_formula("football_capex_high", ["Tornado!D7"])
        must_be_formula("football_capex_width", ["D11", "B11"])

    # Summary: ranked cells must be live LARGE/INDEX/MATCH formulas over a source block that links
    # to each event sheet (no pasted ranking).
    wbf = load_workbook(path)
    if "Summary" not in wbf.sheetnames:
        failures.append("Summary sheet missing")
    else:
        sm = wbf["Summary"]
        sheets = ["Event_Rates", second] + (["Event_Capex"] if capex else [])
        for i, sheet in enumerate(sheets):
            r = 14 + i
            for col, token in (("D", f"{sheet}!G11"), ("E", f"{sheet}!G{ROW_MIRROR}"), ("F", f"{sheet}!E11"),
                               ("H", f"{sheet}!G{ROW_INT}"), ("I", f"{sheet}!G{ROW_ASYM}"),
                               ("J", f"{sheet}!B{ROW_RESULT}")):
                v = sm[f"{col}{r}"].value
                if not (isinstance(v, str) and v.startswith("=") and token in v):
                    failures.append(f"Chain break: Summary!{col}{r}={v!r} must link to {token}")
        for k in range(len(sheets)):
            r = 5 + k
            for col in "BCDEFGHIJK":
                v = sm[f"{col}{r}"].value
                if not (isinstance(v, str) and v.startswith("=") and "INDEX(" in v and "LARGE(" in v):
                    failures.append(f"Chain break: Summary!{col}{r}={v!r} is not a live LARGE/INDEX ranking formula")
    wbf.close()

    details.append(f"Dependency chain traced for rates + {second}{' + Event_Capex' if capex else ''} ({len(refs)} cells + Summary)")
    if not failures:
        details.append("Dependency chain: OK")
    return failures, details


def check_passthrough(path: Path) -> tuple[bool, str]:
    wb = load_workbook(path, data_only=True)
    if "PassThrough" not in wb.sheetnames:
        wb.close()
        return False, "PassThrough sheet missing"
    ws = wb["PassThrough"]
    status, delta = ws["B37"].value, ws["B36"].value
    pt, elast, md, mit = ws["B10"].value, ws["B11"].value, ws["B12"].value, ws["B13"].value
    wb.close()
    try:
        posture = (
            abs(float(pt) - 1.0) < 1e-9
            and abs(float(elast)) < 1e-12
            and abs(float(md)) < 1e-12
            and abs(float(mit)) < 1e-12
        )
    except (TypeError, ValueError):
        return False, f"PassThrough inputs not numeric (pt={pt}, elast={elast})"
    if not posture:
        return False, f"PassThrough not in invariant posture (pt={pt}, elast={elast}, md={md}, mit={mit})"
    if status != "PASS":
        return False, f"PassThrough invariant status={status!r}, ΔGP={delta}"
    if delta is None or abs(float(delta)) >= 0.01:
        return False, f"PassThrough ΔGP={delta} not flat"
    return True, f"PassThrough invariant PASS (ΔGP={delta})"


def run_live_tests(
    src: Path, live_tests: list[dict], recalc_dir: Path, second: str
) -> tuple[list[str], list[str]]:
    failures, details = [], []
    base_path = recalc_libreoffice(src, recalc_dir / "base")
    all_refs = sorted({ref for t in live_tests for ref in t.get("expect_move", []) + t.get("expect_unchanged", [])})
    all_refs = sorted(
        set(all_refs)
        | {
            "Event_Rates!D11",
            f"{second}!D11",
            "Tornado!C5",
            "Tornado!C6",
            "Envelope!B9",
            "Football!B8",
            "Football!B9",
        }
    )
    baseline = read_values(base_path, all_refs)

    for test in live_tests:
        name = test["name"]
        work = recalc_dir / f"live_{name}"
        work.mkdir(parents=True, exist_ok=True)
        xlsx = work / src.name
        shutil.copy2(src, xlsx)
        wb = load_workbook(xlsx)
        wb[test["sheet"]][test["cell"]] = test["new_value"]
        wb.save(xlsx)
        wb.close()
        out = recalc_libreoffice(xlsx, work / "out")
        watch = test.get("expect_move", []) + test.get("expect_unchanged", [])
        after = read_values(out, watch)
        moved_ok = 0
        for ref in test.get("expect_move", []):
            b, a = baseline.get(ref), after.get(ref)
            ok = b is not None and a is not None and a != b
            moved_ok += int(ok)
            details.append(f"  [{name}] {ref}: {b} -> {a} {'MOVED' if ok else 'STALE'}")
        if moved_ok != len(test.get("expect_move", [])):
            failures.append(f"Live test {name}: moved {moved_ok}/{len(test['expect_move'])}")
        else:
            details.append(f"  [{name}] OK move {moved_ok}/{len(test['expect_move'])}")
        for ref in test.get("expect_unchanged", []):
            b, a = baseline.get(ref), after.get(ref)
            if b is None or a is None or a != b:
                failures.append(f"Live test {name}: {ref} should be unchanged ({b} -> {a})")
                details.append(f"  [{name}] {ref}: {b} -> {a} UNEXPECTED CHANGE")
            else:
                details.append(f"  [{name}] {ref}: unchanged OK")
    return failures, details


def run_shock_tests(
    src: Path, shock_tests: list[dict], recalc_dir: Path
) -> tuple[list[str], list[str]]:
    """Perturb each event's own shock input; require FULL move in declared direction."""
    failures, details = [], []
    if not shock_tests:
        failures.append("No shock_tests declared in config (required)")
        return failures, details

    base_path = recalc_libreoffice(src, recalc_dir / "shock_base")
    full_cells = sorted({t["full_cell"] for t in shock_tests})
    unchanged_all = sorted({u for t in shock_tests for u in t.get("expect_unchanged", [])})
    baseline = read_values(base_path, full_cells + unchanged_all + ["DCF!B28"])

    for test in shock_tests:
        name = test["name"]
        work = recalc_dir / f"shock_{name}"
        work.mkdir(parents=True, exist_ok=True)
        xlsx = work / src.name
        shutil.copy2(src, xlsx)
        wb = load_workbook(xlsx)
        wb[test["sheet"]][test["cell"]] = test["new_value"]
        wb.save(xlsx)
        wb.close()
        out = recalc_libreoffice(xlsx, work / "out")
        full = test["full_cell"]
        after = read_values(out, [full] + test.get("expect_unchanged", []) + ["DCF!B28"])
        b, a = baseline[full], after[full]
        details.append(f"  [shock:{name}] {full}: {b} -> {a} (expect {test['expected_direction']})")
        if b is None or a is None:
            failures.append(f"Shock test {name}: missing values {b} -> {a}")
            continue
        if a == b:
            failures.append(f"Shock test {name}: FULL {full} did not move")
            continue
        direction = test["expected_direction"]
        if direction == "decrease" and not (float(a) < float(b)):
            failures.append(f"Shock test {name}: expected decrease, got {b} -> {a}")
        elif direction == "increase" and not (float(a) > float(b)):
            failures.append(f"Shock test {name}: expected increase, got {b} -> {a}")
        else:
            details.append(f"  [shock:{name}] direction OK")
        for ref in test.get("expect_unchanged", []):
            ub, ua = baseline[ref], after[ref]
            if ub is None or ua is None or ua != ub:
                failures.append(f"Shock test {name}: unrelated {ref} moved ({ub} -> {ua})")
            else:
                details.append(f"  [shock:{name}] unrelated {ref} unchanged OK")
    return failures, details


CAPEX_SHOCK_TEST = {
    "name": "capex_shock_bigger",
    "sheet": "Event_Capex",
    "cell": "B21",
    "new_value": 0.45,
    "full_cell": "Event_Capex!D11",
    "expected_direction": "decrease",
    "expect_unchanged": ["Event_Rates!D11"],
}


def default_shock_tests(second: str, capex: bool = False) -> list[dict]:
    tests = [
        {
            "name": "rates_bp_up",
            "sheet": "Inputs",
            "cell": "B27",
            "new_value": 200,
            "full_cell": "Event_Rates!D11",
            "expected_direction": "decrease",
            "expect_unchanged": [f"{second}!D11"],
        }
    ]
    if second == "Event_Commodity":
        tests.append(
            {
                "name": "commodity_more_adverse",
                "sheet": "Inputs",
                "cell": "B18",
                "new_value": -0.35,
                "full_cell": "Event_Commodity!D11",
                "expected_direction": "decrease",
                "expect_unchanged": ["Event_Rates!D11"],
            }
        )
    else:
        tests.append(
            {
                "name": "tariff_cost_up",
                "sheet": "Event_Tariff",
                "cell": "B25",
                "new_value": 0.25,
                "full_cell": "Event_Tariff!D11",
                "expected_direction": "decrease",
                "expect_unchanged": ["Event_Rates!D11"],
            }
        )
    if capex:
        tests.append(dict(CAPEX_SHOCK_TEST))
    return tests


def default_live_tests(second: str, capex: bool = False) -> list[dict]:
    tests = [
        {
            "name": "revenue_growth",
            "sheet": "Inputs",
            "cell": "B6",
            "new_value": 0.06,
            "expect_move": [
                "DCF!B30",
                "Event_Rates!D11",
                f"{second}!D11",
                "Tornado!C5",
                "Envelope!B9",
                "Football!B8",
                "Summary!D14",
                "Summary!D15",
            ],
        },
        {
            "name": "risk_free_rate",
            "sheet": "Inputs",
            "cell": "B21",
            "new_value": 0.05,
            "expect_move": [
                "DCF!B30",
                "Event_Rates!D11",
                "Event_Rates!F11",
                "Tornado!C5",
                "Envelope!B9",
                "Summary!D14",
            ],
            "expect_unchanged": [],  # second FULL may move via WACC in DCF shared paths for tariff; leave empty
        },
    ]
    if capex:
        tests[0]["expect_move"] += ["Event_Capex!D11", "Tornado!C7", "Summary!D16"]
    return tests



def _norm_formula(s: str) -> str:
    return "".join(str(s).upper().split())


def verify_channel_bindings(path: Path, bindings: list[dict]) -> tuple[list[str], list[str]]:
    """Structural + economic sign checks for channel rows (catches D6 = 2*base - helper flips)."""
    failures, details = [], []
    if not bindings:
        failures.append("No channel_bindings declared in config (required)")
        return failures, details
    formulas = read_formulas(path, [b["cell"] for b in bindings] + ["DCF!B28"])
    values = read_values(path, [b["cell"] for b in bindings] + ["DCF!B28"] + [b.get("base_cell", "DCF!B28") for b in bindings])
    for b in bindings:
        cell = b["cell"]
        expected = b["must_equal_formula"]
        actual = formulas.get(cell)
        if not isinstance(actual, str) or not actual.startswith("="):
            failures.append(f"Channel binding: {cell} is not a formula ({actual!r})")
            continue
        if _norm_formula(actual) != _norm_formula(expected):
            failures.append(
                f"Channel binding/sign: {cell} formula {actual!r} != required {expected!r}"
            )
        else:
            details.append(f"Channel binding OK: {cell} = {actual}")
        if b.get("adverse_vs_base"):
            base_cell = b.get("base_cell", "DCF!B28")
            cv, bv = values.get(cell), values.get(base_cell)
            if cv is None or bv is None:
                failures.append(f"Channel sign: missing values for {cell} or {base_cell}")
            elif float(cv) >= float(bv):
                failures.append(
                    f"Channel sign: adverse channel {cell}={cv} should be < base {base_cell}={bv}"
                )
            else:
                details.append(f"Channel sign OK: {cell} < {base_cell}")
    return failures, details


def run_channel_tests(
    src: Path, channel_tests: list[dict], recalc_dir: Path
) -> tuple[list[str], list[str]]:
    """Perturb each channel input; require channel row AND FULL to move in declared directions.

    Coverage: every test with covers_full true (default) must move FULL — a FULL that omits
    that channel (e.g. D11=B42 while elasticity changes) FAILs.
    """
    failures, details = [], []
    if not channel_tests:
        failures.append("No channel_tests declared in config (required)")
        return failures, details

    refs = sorted(
        {
            r
            for t in channel_tests
            for r in (t["channel_cell"], t["full_cell"])
        }
    )
    base_path = recalc_libreoffice(src, recalc_dir / "channel_base")
    baseline = read_values(base_path, refs)

    for test in channel_tests:
        name = test["name"]
        work = recalc_dir / f"channel_{name}"
        work.mkdir(parents=True, exist_ok=True)
        xlsx = work / src.name
        shutil.copy2(src, xlsx)
        wb = load_workbook(xlsx)
        wb[test["sheet"]][test["cell"]] = test["new_value"]
        wb.save(xlsx)
        wb.close()
        out = recalc_libreoffice(xlsx, work / "out")
        after = read_values(out, [test["channel_cell"], test["full_cell"]])
        ch_b, ch_a = baseline[test["channel_cell"]], after[test["channel_cell"]]
        fu_b, fu_a = baseline[test["full_cell"]], after[test["full_cell"]]
        details.append(
            f"  [channel:{name}] {test['channel_cell']}: {ch_b} -> {ch_a}; "
            f"{test['full_cell']}: {fu_b} -> {fu_a}"
        )
        if ch_b is None or ch_a is None:
            failures.append(f"Channel test {name}: missing channel values")
            continue
        if ch_a == ch_b:
            failures.append(f"Channel test {name}: channel row {test['channel_cell']} did not move")
        else:
            d = test["expected_channel_direction"]
            ok = (d == "decrease" and float(ch_a) < float(ch_b)) or (
                d == "increase" and float(ch_a) > float(ch_b)
            )
            if not ok:
                failures.append(
                    f"Channel test {name}: channel direction expected {d}, got {ch_b} -> {ch_a}"
                )
            else:
                details.append(f"  [channel:{name}] channel direction OK")

        covers = test.get("covers_full", True)
        if covers:
            if fu_b is None or fu_a is None:
                failures.append(f"Channel coverage {name}: missing FULL values")
            elif fu_a == fu_b:
                failures.append(
                    f"Channel coverage {name}: FULL {test['full_cell']} did not move when "
                    f"{test['sheet']}!{test['cell']} changed — channel omitted from FULL?"
                )
            else:
                d = test["expected_full_direction"]
                ok = (d == "decrease" and float(fu_a) < float(fu_b)) or (
                    d == "increase" and float(fu_a) > float(fu_b)
                )
                if not ok:
                    failures.append(
                        f"Channel coverage {name}: FULL direction expected {d}, got {fu_b} -> {fu_a}"
                    )
                else:
                    details.append(f"  [channel:{name}] FULL coverage + direction OK")
    return failures, details


_CHART_NS = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def _normalize_ref(ref: str) -> str:
    r = ref.replace("$", "").replace("'", "").strip().upper()
    return r


def inspect_chart_series(path: Path, charts_cfg: dict) -> tuple[list[str], list[str]]:
    """Parse chart XML: tornado val/cat refs; football baseline+width refs; baseline noFill."""
    failures, details = [], []
    if not charts_cfg:
        failures.append("No checks.charts declared in config (required)")
        return failures, details

    with zipfile.ZipFile(path) as zf:
        chart_files = sorted(n for n in zf.namelist() if n.startswith("xl/charts/chart") and n.endswith(".xml"))
        if len(chart_files) < 2:
            failures.append(f"Expected >=2 chart XML parts, found {chart_files}")
            return failures, details

        parsed = []
        for n in chart_files:
            root = ET.fromstring(zf.read(n))
            series = []
            for ser in root.findall(".//c:ser", _CHART_NS):
                # title / cat / val formula refs
                fs = [f.text or "" for f in ser.findall(".//c:f", _CHART_NS)]
                # noFill on series spPr (shape), not just line
                spPr = ser.find("c:spPr", _CHART_NS)
                no_fill = False
                if spPr is not None:
                    no_fill = spPr.find("a:noFill", _CHART_NS) is not None
                series.append({"fs": fs, "noFill": no_fill, "raw_fs": fs})
            parsed.append({"file": n, "series": series})

        # Identify tornado vs football by refs content
        tornado_cfg = charts_cfg.get("tornado") or {}
        football_cfg = charts_cfg.get("football") or {}
        tornado_chart = None
        football_chart = None
        for ch in parsed:
            blob = " ".join(" ".join(s["fs"]) for s in ch["series"]).upper()
            if "TORNADO" in blob:
                tornado_chart = ch
            if "FOOTBALL" in blob:
                football_chart = ch

        if tornado_chart is None:
            failures.append("Could not locate Tornado chart XML series")
        else:
            details.append(f"Tornado chart: {tornado_chart['file']}")
            want_val = _normalize_ref(tornado_cfg.get("value_ref", "Tornado!$B$13:$B$14"))
            want_cat = _normalize_ref(tornado_cfg.get("cat_ref", "Tornado!$A$13:$A$14"))
            found_vals = [_normalize_ref(f) for s in tornado_chart["series"] for f in s["fs"]]
            if want_val not in found_vals:
                failures.append(f"Tornado chart missing value ref {want_val}; found {found_vals}")
            if want_cat not in found_vals:
                failures.append(f"Tornado chart missing cat ref {want_cat}; found {found_vals}")
            if want_val in found_vals and want_cat in found_vals:
                details.append(f"Tornado chart refs OK ({want_cat}, {want_val})")

        if football_chart is None:
            failures.append("Could not locate Football chart XML series")
        else:
            details.append(f"Football chart: {football_chart['file']}")
            if len(football_chart["series"]) < 2:
                failures.append(f"Football chart needs 2 series (baseline+width), found {len(football_chart['series'])}")
            else:
                base_ser, width_ser = football_chart["series"][0], football_chart["series"][1]
                want_base = _normalize_ref(football_cfg.get("baseline_value_ref", "Football!$F$5:$F$10"))
                want_width = _normalize_ref(football_cfg.get("width_value_ref", "Football!$G$5:$G$10"))
                want_cat = _normalize_ref(football_cfg.get("cat_ref", "Football!$A$5:$A$10"))
                base_fs = [_normalize_ref(f) for f in base_ser["fs"]]
                width_fs = [_normalize_ref(f) for f in width_ser["fs"]]
                if want_base not in base_fs:
                    failures.append(
                        f"Football baseline series missing {want_base}; found {base_fs}"
                    )
                if want_width not in width_fs:
                    failures.append(
                        f"Football width/visible series missing {want_width}; found {width_fs}"
                    )
                # Visible series must NOT be the Low column (B) or Mid (C)
                for bad in ("FOOTBALL!B5:B10", "FOOTBALL!C5:C10", "FOOTBALL!B5:B11", "FOOTBALL!C5:C11"):
                    if _normalize_ref(bad) in width_fs:
                        failures.append(
                            f"Football width series incorrectly references {_normalize_ref(bad)}"
                        )
                all_fs = base_fs + width_fs
                if want_cat not in all_fs:
                    failures.append(f"Football chart missing cat ref {want_cat}; found {all_fs}")
                require_nofill = football_cfg.get("baseline_nofill", True)
                if require_nofill and not base_ser["noFill"]:
                    failures.append("Football baseline series missing spPr/a:noFill (should be invisible)")
                elif require_nofill:
                    details.append("Football baseline noFill OK")
                if want_base in base_fs and want_width in width_fs:
                    details.append(f"Football chart refs OK (baseline {want_base}, width {want_width})")

    return failures, details


CAPEX_CHANNEL_TESTS = [
    {
        "name": "capex_revenue_elasticity",
        "sheet": "Event_Capex",
        "cell": "B22",
        "new_value": 1.0,
        "channel_cell": "Event_Capex!D6",
        "full_cell": "Event_Capex!D11",
        "expected_channel_direction": "increase",
        "expected_full_direction": "increase",
    },
    {
        "name": "capex_own_capex_response",
        "sheet": "Event_Capex",
        "cell": "B23",
        "new_value": 1.6,
        "channel_cell": "Event_Capex!D7",
        "full_cell": "Event_Capex!D11",
        "expected_channel_direction": "decrease",
        "expected_full_direction": "decrease",
    },
    {
        "name": "capex_equity_share",
        "sheet": "Event_Capex",
        "cell": "B24",
        "new_value": 0.9,
        "channel_cell": "Event_Capex!D8",
        "full_cell": "Event_Capex!D11",
        "expected_channel_direction": "decrease",
        "expected_full_direction": "decrease",
    },
    {
        "name": "capex_issue_price",
        "sheet": "Event_Capex",
        "cell": "B25",
        "new_value": 30.0,
        "channel_cell": "Event_Capex!D8",
        "full_cell": "Event_Capex!D11",
        "expected_channel_direction": "increase",
        "expected_full_direction": "increase",
    },
]

CAPEX_BINDINGS = [
    {"cell": "Event_Capex!D6", "must_equal_formula": "=B43"},
    {"cell": "Event_Capex!D7", "must_equal_formula": "=B44"},
    {"cell": "Event_Capex!D8", "must_equal_formula": "=B45"},
    {"cell": "Event_Capex!D11", "must_equal_formula": "=B47"},
    {"cell": "Event_Capex!B47", "must_equal_formula": "=(B46-B40)*Inputs!B39/B41"},
]


def default_channel_tests(second: str, capex: bool = False) -> list[dict]:
    extra = [dict(t) for t in CAPEX_CHANNEL_TESTS] if capex else []
    return _default_channel_tests(second) + extra


def _default_channel_tests(second: str) -> list[dict]:
    rates = [
        {
            "name": "rates_shock_bp",
            "sheet": "Inputs",
            "cell": "B27",
            "new_value": 200,
            "channel_cell": "Event_Rates!D6",
            "full_cell": "Event_Rates!D11",
            "expected_channel_direction": "decrease",
            "expected_full_direction": "decrease",
        },
        {
            "name": "rates_duration_mtm",
            "sheet": "Inputs",
            "cell": "B38",
            "new_value": 6.0,
            "channel_cell": "Event_Rates!D8",
            "full_cell": "Event_Rates!D11",
            "expected_channel_direction": "increase",
            "expected_full_direction": "increase",
        },
    ]
    if second == "Event_Tariff":
        return rates + [
            {
                "name": "tariff_unit_cost",
                "sheet": "Event_Tariff",
                "cell": "B25",
                "new_value": 0.25,
                "channel_cell": "Event_Tariff!D6",
                "full_cell": "Event_Tariff!D11",
                "expected_channel_direction": "decrease",
                "expected_full_direction": "decrease",
            },
            {
                "name": "tariff_pass_through",
                "sheet": "Event_Tariff",
                "cell": "B26",
                "new_value": 0.90,
                "channel_cell": "Event_Tariff!D7",
                "full_cell": "Event_Tariff!D11",
                "expected_channel_direction": "increase",
                "expected_full_direction": "increase",
            },
            {
                "name": "tariff_elasticity",
                "sheet": "Event_Tariff",
                "cell": "B27",
                "new_value": -0.60,
                "channel_cell": "Event_Tariff!D8",
                "full_cell": "Event_Tariff!D11",
                "expected_channel_direction": "decrease",
                "expected_full_direction": "decrease",
            },
            {
                "name": "tariff_markdown",
                "sheet": "Event_Tariff",
                "cell": "B28",
                "new_value": 0.05,
                "channel_cell": "Event_Tariff!D9",
                "full_cell": "Event_Tariff!D11",
                "expected_channel_direction": "decrease",
                "expected_full_direction": "decrease",
            },
            {
                "name": "tariff_mitigation",
                "sheet": "Event_Tariff",
                "cell": "B29",
                "new_value": 0.40,
                "channel_cell": "Event_Tariff!D11",
                "full_cell": "Event_Tariff!D11",
                "expected_channel_direction": "increase",
                "expected_full_direction": "increase",
            },
        ]
    # commodity
    return rates + [
        {
            "name": "commodity_price",
            "sheet": "Inputs",
            "cell": "B18",
            "new_value": -0.35,
            "channel_cell": "Event_Commodity!D6",
            "full_cell": "Event_Commodity!D11",
            "expected_channel_direction": "decrease",
            "expected_full_direction": "decrease",
        },
        {
            "name": "commodity_cost_elasticity",
            "sheet": "Inputs",
            "cell": "B13",
            "new_value": 0.60,
            "channel_cell": "Event_Commodity!D8",
            "full_cell": "Event_Commodity!D11",
            "expected_channel_direction": "increase",
            "expected_full_direction": "increase",
        },
        {
            "name": "commodity_capex_elasticity",
            "sheet": "Inputs",
            "cell": "B14",
            "new_value": 0.80,
            "channel_cell": "Event_Commodity!D9",
            "full_cell": "Event_Commodity!D11",
            "expected_channel_direction": "increase",
            "expected_full_direction": "increase",
        },
    ]


def default_channel_bindings(second: str, capex: bool = False) -> list[dict]:
    extra = [dict(b) for b in CAPEX_BINDINGS] if capex else []
    return _default_channel_bindings(second) + extra


def _default_channel_bindings(second: str) -> list[dict]:
    rates = [
        {"cell": "Event_Rates!D6", "must_equal_formula": "=B30", "adverse_vs_base": True},
        {"cell": "Event_Rates!D11", "must_equal_formula": "=B30+B35"},
    ]
    if second == "Event_Tariff":
        return rates + [
            {"cell": "Event_Tariff!D6", "must_equal_formula": "=B41", "adverse_vs_base": True},
            {"cell": "Event_Tariff!D7", "must_equal_formula": "=B42"},
            {"cell": "Event_Tariff!D8", "must_equal_formula": "=B43"},
            {"cell": "Event_Tariff!D9", "must_equal_formula": "=B44"},
            {"cell": "Event_Tariff!D11", "must_equal_formula": "=B45"},
        ]
    return rates + [
        {"cell": "Event_Commodity!D6", "must_equal_formula": "=B43"},
        {"cell": "Event_Commodity!D11", "must_equal_formula": "=B47"},
    ]


def default_charts_cfg(capex: bool = False) -> dict:
    t_last = 15 if capex else 14
    f_last = 11 if capex else 10
    return {
        "tornado": {
            "value_ref": f"Tornado!$B$13:$B${t_last}",
            "cat_ref": f"Tornado!$A$13:$A${t_last}",
        },
        "football": {
            "baseline_value_ref": f"Football!$F$5:$F${f_last}",
            "width_value_ref": f"Football!$G$5:$G${f_last}",
            "cat_ref": f"Football!$A$5:$A${f_last}",
            "baseline_nofill": True,
        },
    }




def _expand_a1_range(a1: str) -> set[str]:
    """Expand B5:B7 or B5 into a set of cell addresses (no sheet)."""
    a1 = a1.replace("$", "").strip().upper()
    if ":" not in a1:
        return {a1}
    from openpyxl.utils import range_boundaries, get_column_letter
    min_col, min_row, max_col, max_row = range_boundaries(a1)
    out = set()
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            out.add(f"{get_column_letter(c)}{r}")
    return out


def verify_envelope_formula_coverage(path: Path) -> tuple[list[str], list[str]]:
    """Parse Envelope MIN/MAX ranges; require base + every enabled event's FULL proxy."""
    failures, details = [], []
    fx = read_formulas(path, ["Envelope!B9", "Envelope!B10", "Envelope!C9", "Envelope!C10",
                              "Football!B10", "Football!D10"])
    last = 5 + len(event_sheets(path))
    rows = range(5, last + 1)
    required_by_cell = {
        "Envelope!B9": {f"B{r}" for r in rows},
        "Envelope!B10": {f"B{r}" for r in rows},
        "Envelope!C9": {f"C{r}" for r in rows},
        "Envelope!C10": {f"C{r}" for r in rows},
    }
    for cell, kind in (("Envelope!B9", "MIN"), ("Envelope!B10", "MAX"),
                       ("Envelope!C9", "MIN"), ("Envelope!C10", "MAX")):
        formula = str(fx.get(cell) or "")
        m = re.search(rf"{kind}\(([^)]+)\)", formula, flags=re.I)
        if not m:
            failures.append(f"Envelope coverage: {cell} missing {kind}(...) got {formula!r}")
            continue
        arg = m.group(1).strip()
        cells = _expand_a1_range(arg)
        required = required_by_cell[cell]
        if cells != required:
            failures.append(
                f"Envelope coverage: {cell} {kind} range {arg} → {sorted(cells)} "
                f"!= required {sorted(required)} (must include both FULL-linked tornado endpoints)"
            )
        else:
            details.append(f"Envelope coverage OK: {cell} {kind}({arg})")

    # Football band must formula-link to envelope $/sh
    if _norm_formula(str(fx.get("Football!B10"))) != _norm_formula("=Envelope!C9"):
        failures.append(
            f"Football!B10 formula {fx.get('Football!B10')!r} must be =Envelope!C9"
        )
    if _norm_formula(str(fx.get("Football!D10"))) != _norm_formula("=Envelope!C10"):
        failures.append(
            f"Football!D10 formula {fx.get('Football!D10')!r} must be =Envelope!C10"
        )
    if not any("Football!B10" in f or "Football!D10" in f for f in failures):
        details.append("Football envelope band formulas link to Envelope!C9/C10 OK")

    return failures, details


def _close(a: Any, b: Any, tol_rel: float = 1e-4, tol_abs: float = 0.01) -> bool:
    try:
        a, b = float(a), float(b)
    except (TypeError, ValueError):
        return False
    return abs(a - b) <= max(tol_abs, abs(b) * tol_rel)


def expected_reverse_result(grid: list[tuple[float, float]], threshold: Any, base_sh: float, unit: str = "bp") -> Any:
    """grid = [(shock size, FULL proxy $/sh)] in ascending-multiple order."""
    if threshold is None or threshold == "":
        return "n/a"
    thr = float(threshold)
    if base_sh < thr:
        return "base already below threshold"
    for shock, sh in grid:
        if sh < thr:
            return shock
    shock, sh = min(grid, key=lambda p: p[1])
    return format_not_reached(sh / base_sh - 1.0, shock, unit)


def verify_event_extras(path: Path) -> tuple[list[str], list[str], list[str], list[tuple[str, str, str]]]:
    """Interaction rows, up/down mirror, reverse-stress grid + result, calibration cells, Summary.

    Values are checked against the independent reference (reference_dcf.py), not the sheet formulas.
    Returns (failures, details, manual_review_items, calibration_to_fill) where calibration_to_fill
    lists (sheet, 'source'|'basis', cell value) for every calibration cell still marked TO FILL.
    """
    failures: list[str] = []
    details: list[str] = []
    manual: list[str] = []
    to_fill: list[tuple[str, str, str]] = []
    m = load_inputs_from_workbook(path)
    base = base_equity_perp(m)
    blend = blended_mid(m)
    base_sh = blend / m.diluted_shares
    wb = load_workbook(path, data_only=True)
    wbf = load_workbook(path)
    sheets = event_sheets(path)
    summary_src: list[dict[str, Any]] = []

    def proxy_sh(eq: float) -> float:
        return blend * (1.0 + (eq - base) / base) / m.diluted_shares

    for sheet in sheets:
        ws, wf = wb[sheet], wbf[sheet]
        unit = "bp" if sheet == "Event_Rates" else "%"
        # --- interaction row ---
        kind = str(ws[f"B{ROW_INT}"].value or "").lower()
        int_fx = wf[f"D{ROW_INT}"].value
        label = str(ws[f"C{ROW_INT}"].value or "")
        if not label.startswith(INTERACTION_LABEL):
            failures.append(f"Interaction: {sheet}!C{ROW_INT} label must start {INTERACTION_LABEL!r}, got {label!r}")
        if kind != "diagnostic":
            failures.append(f"Interaction: {sheet}!B{ROW_INT} kind must be 'diagnostic', got {kind!r}")
        if not (isinstance(int_fx, str) and int_fx.startswith("=")):
            failures.append(f"Interaction: {sheet}!D{ROW_INT} is not a live formula ({int_fx!r})")
        full_fx = str(wf["D11"].value or "")
        if re.search(rf"\b[D-G]{ROW_INT}\b", full_fx):
            failures.append(f"Interaction: {sheet}!D11 FULL references the interaction row ({full_fx})")
        d = {r: ws[f"D{r}"].value for r in range(6, 12)}
        if any(v is None for v in d.values()):
            failures.append(f"Interaction: {sheet} channel values missing")
            continue
        d = {r: float(v) for r, v in d.items()}
        if sheet == "Event_Tariff":
            mit_only = tariff_posture_equity(m, pass_through=0.0, elasticity=0.0, markdown=0.0,
                                             mitigation=float(m.tariff_mitigation))
            exp_int = d[11] - d[9] - (mit_only - d[6])
        else:
            exp_int = (d[11] - base) - sum(d[r] - base for r in STANDALONE_ROWS[sheet])
        if sheet == "Event_Capex":
            c = capex_components(m)
            ref_int = (c["full"] - base) - sum(c[k] - base for k in ("A", "B", "C"))
            if not _close(exp_int, ref_int):
                failures.append(f"Interaction: {sheet} channel rows imply {exp_int}, reference {ref_int}")
        wb_int = ws[f"D{ROW_INT}"].value
        if not _close(wb_int, exp_int):
            failures.append(f"Interaction: {sheet}!D{ROW_INT}={wb_int} != FULL − standalones {exp_int}")
        else:
            details.append(f"Interaction OK: {sheet}!D{ROW_INT}={float(wb_int):.4f}")

        # --- up / down mirror ---
        shock = float(configured_shock(m, sheet))
        ref_mirror = full_at(m, sheet, -shock)
        if not _close(ws[f"D{ROW_CONF}"].value, d[11]):
            failures.append(f"Mirror: {sheet}!D{ROW_CONF} must equal FULL D11")
        if not _close(ws[f"D{ROW_MIRROR}"].value, ref_mirror):
            failures.append(f"Mirror: {sheet}!D{ROW_MIRROR}={ws[f'D{ROW_MIRROR}'].value} != reference {ref_mirror}")
        else:
            details.append(f"Mirror OK: {sheet} FULL at −shock = {ref_mirror:.4f}")
        exp_asym = (d[11] - base) + (ref_mirror - base)
        if not _close(ws[f"D{ROW_ASYM}"].value, exp_asym):
            failures.append(f"Asymmetry: {sheet}!D{ROW_ASYM}={ws[f'D{ROW_ASYM}'].value} != {exp_asym}")

        # --- reverse stress grid + result ---
        grid: list[tuple[float, float]] = []
        prev_mult = 0.0
        grid_ok = True
        for i in range(N_GRID):
            r = ROW_GRID0 + i
            mult = ws[f"B{r}"].value
            if mult is None or float(mult) <= prev_mult:
                failures.append(f"Reverse grid: {sheet}!B{r} multiple {mult} must be > previous ({prev_mult})")
                grid_ok = False
                break
            prev_mult = float(mult)
            x = shock * prev_mult
            ref_eq = full_at(m, sheet, x)
            if not _close(ws[f"D{r}"].value, ref_eq):
                failures.append(f"Reverse grid: {sheet}!D{r}={ws[f'D{r}'].value} != reference {ref_eq} at shock {x}")
                grid_ok = False
            if not _close(ws[f"G{r}"].value, proxy_sh(ref_eq), tol_abs=1e-4):
                failures.append(f"Reverse grid: {sheet}!G{r} proxy $/sh mismatch")
                grid_ok = False
            grid.append((x, proxy_sh(ref_eq)))
        if grid_ok:
            details.append(f"Reverse grid OK: {sheet} {N_GRID} steps match reference")
        thr = ws[f"B{ROW_THRESHOLD}"].value
        thr_fx = wf[f"B{ROW_THRESHOLD}"].value
        if isinstance(thr_fx, str) and thr_fx.startswith("="):
            # Default threshold: live formula = base blended-mid $/sh × (1 − drop %), blank drop → blank → n/a
            if thr_fx.replace(" ", "") != THRESHOLD_DEFAULT_FX:
                failures.append(f"Reverse stress: {sheet}!B{ROW_THRESHOLD} formula {thr_fx!r} is not the default "
                                f"{THRESHOLD_DEFAULT_FX!r} (type a number to override instead)")
            drop = ws[f"{COL_DROP}{ROW_THRESHOLD}"].value
            if drop is None or drop == "":
                if thr not in (None, ""):
                    failures.append(f"Reverse stress: {sheet} blank drop % but threshold = {thr!r}")
            elif not _close(thr, base_sh * (1.0 - float(drop)), tol_abs=1e-9):
                failures.append(f"Reverse stress: {sheet}!B{ROW_THRESHOLD}={thr!r} != base × (1 − {drop}) "
                                f"= {base_sh * (1.0 - float(drop))}")
            else:
                details.append(f"Reverse stress threshold OK: {sheet} default = base × (1 − {float(drop):.0%}) = {float(thr):.4f}")
        exp_res = expected_reverse_result(grid, thr, base_sh, unit)
        got = ws[f"B{ROW_RESULT}"].value
        if isinstance(exp_res, float):
            ok = _close(got, exp_res, tol_abs=1e-9)
        else:
            ok = got == exp_res
        if not ok:
            failures.append(f"Reverse stress: {sheet}!B{ROW_RESULT}={got!r}, expected {exp_res!r} (threshold {thr!r})")
        else:
            details.append(f"Reverse stress OK: {sheet} threshold={thr!r} → {got!r}")

        # --- calibration disclosure ---
        for r, what in ((ROW_CALIB_SOURCE, "source"), (ROW_CALIB_BASIS, "basis")):
            v = ws[f"B{r}"].value
            if v is None or str(v).strip() == "":
                failures.append(f"Calibration: {sheet}!B{r} ({what}) is empty — write TO FILL or cite it")
            elif str(v).upper().startswith("TO FILL"):
                to_fill.append((sheet, what, str(v)))

        summary_src.append({
            "sheet": sheet,
            "conf_sh": ws["G11"].value,
            "mirror_sh": proxy_sh(ref_mirror),
            "delta_sh": proxy_sh(d[11]) - base_sh,
            "result": got,
        })

    # --- Summary ranking (independent sort of the event-sheet values) ---
    if "Summary" in wb.sheetnames:
        sm = wb["Summary"]
        ranked = sorted(enumerate(summary_src), key=lambda t: (-abs(t[1]["delta_sh"]), t[0]))
        sm_ok = True
        for k, (_i, ev) in enumerate(ranked):
            r = 5 + k
            if sm[f"C{r}"].value != ev["sheet"]:
                failures.append(f"Summary: rank {k + 1} is {sm[f'C{r}'].value!r}, expected {ev['sheet']}")
                sm_ok = False
                continue
            lo, hi = sorted([float(ev["conf_sh"]), ev["mirror_sh"]])
            for col, exp in (("E", lo), ("F", hi), ("H", ev["delta_sh"])):
                if not _close(sm[f"{col}{r}"].value, exp, tol_abs=1e-4):
                    failures.append(f"Summary!{col}{r}={sm[f'{col}{r}'].value} != {exp}")
                    sm_ok = False
            if sm[f"K{r}"].value != ev["result"]:
                failures.append(f"Summary!K{r}={sm[f'K{r}'].value!r} != event result {ev['result']!r}")
                sm_ok = False
        if sm_ok:
            details.append(f"Summary ranking OK ({' > '.join(e['sheet'] for _i, e in ranked)})")
    wb.close()
    wbf.close()
    return failures, details, manual, to_fill


def calibration_messages(to_fill: list[tuple[str, str, str]], strict: bool) -> tuple[list[str], list[str]]:
    """TO FILL calibration cells: WARNING by default, FAIL under --strict-calibration. Returns (warnings, failures)."""
    if not to_fill:
        return [], []
    by_sheet: dict[str, list[str]] = {}
    for sheet, what, _v in to_fill:
        by_sheet.setdefault(sheet, []).append(what)
    listing = "; ".join(f"{sheet} ({' + '.join(whats)})" for sheet, whats in by_sheet.items())
    msg = (f"Calibration still TO FILL on {len(by_sheet)} event(s): {listing}. "
           "Cite a source and basis (rows 56–57) before a real-issuer CLEAR.")
    if strict:
        return [], [msg + " [--strict-calibration]"]
    return [msg + " Use --strict-calibration to make this a FAIL."], []



def verify_ladder_structure(path: Path) -> tuple[list[str], list[str]]:
    """Ladder outputs must be live formulas, and each chart must plot FULL equity (D) against shock size (C)."""
    import zipfile
    from openpyxl.utils import get_column_letter
    failures: list[str] = []
    details: list[str] = []
    wb = load_workbook(path, data_only=False)
    if "Ladder" not in wb.sheetnames:
        wb.close()
        return ["Ladder sheet missing"], details
    ld = wb["Ladder"]
    sheets = event_sheets(path)
    lay = ladder_layout(len(sheets))
    for sh, blk in zip(sheets, lay["blocks"]):
        for r in blk["rows"]:
            for c in "CDEFGH":
                v = ld[f"{c}{r}"].value
                if not (isinstance(v, str) and v.startswith("=")):
                    failures.append(f"Ladder!{c}{r} ({sh}) is not a formula: {v!r}")
            d = ld[f"D{r}"].value
            if isinstance(d, str) and not d.startswith(f"={sh}!"):
                failures.append(f"Ladder!D{r} must link to {sh}'s helper row, got {d}")
    for rr in lay["rev_rows"]:
        for k in range(len(sheets)):
            v = ld.cell(rr, 2 + k).value
            if not (isinstance(v, str) and v.startswith("=")):
                failures.append(f"Ladder!{get_column_letter(2 + k)}{rr} reverse result is not a formula: {v!r}")
    wb.close()
    # chart series references
    with zipfile.ZipFile(path) as z:
        charts = sorted(n for n in z.namelist() if n.startswith("xl/charts/chart") and n.endswith(".xml"))
        found = []
        for n in charts:
            x = z.read(n).decode("utf-8", "ignore")
            xv = re.findall(r"<c:xVal>.*?<c:f>(.*?)</c:f>", x, re.S)
            yv = re.findall(r"<c:yVal>.*?<c:f>(.*?)</c:f>", x, re.S)
            if xv and yv and "Ladder" in xv[0]:
                found.append((xv[0].replace("'", ""), yv[0].replace("'", "")))
    want = []
    for blk in lay["blocks"]:
        a, b = blk["rows"][0], blk["rows"][-1]
        want.append((f"Ladder!$C${a}:$C${b}", f"Ladder!$D${a}:$D${b}"))
    missing = [w for w in want if w not in found]
    if missing:
        failures.append(f"Ladder charts: expected series not found {missing}; found {found}")
    if not failures:
        details.append(f"Ladder structure: rungs and reverse results live; {len(want)} charts plot FULL equity vs shock size")
    return failures, details



def _rewrite_chart_xml(src: Path, dst: Path, old: str, new: str) -> None:
    """Copy an xlsx, replacing a range string inside chart XML only (keeps cached values)."""
    import zipfile
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/charts/chart") and item.filename.endswith(".xml"):
                data = data.decode("utf-8").replace(old, new).encode("utf-8")
            zout.writestr(item, data)

def run_regression_mutations(good: Path, config: Path | None) -> int:
    """Reviewer mutations (13, or 16 with Event_Capex) — each must FAIL."""
    results = []

    def record(name, rc, reason):
        status = "FAIL" if rc != 0 else "PASS"
        results.append((name, status, reason))
        print(f"  >> {name}: {status} ({reason})")
        if rc == 0:
            print(f"REGRESSION FAIL: {name} incorrectly PASSED")
            return False
        return True

    print("=== Regression suite ===")
    with tempfile.TemporaryDirectory(prefix="stress-reg-") as td:
        td_path = Path(td)
        second = detect_second_event(good)

        # 1) empty cache
        print("=== 1) empty-cache ===")
        empty = td_path / "empty_cache.xlsx"
        wb = load_workbook(good)
        wb.save(empty)
        wb.close()
        rc = main_check(empty, config, skip_recalc=True, run_regression=False)
        if not record("empty-cache", rc, "DCF!B30 None / empty formula caches"):
            return 1

        # 2) FULL = DCF!B28
        print("=== 2) FULL collapsed to DCF!B28 ===")
        src = td_path / "mut_full_base.xlsx"
        shutil.copy2(good, src)
        wb = load_workbook(src)
        wb[second]["D11"] = "=DCF!B28"
        wb.save(src)
        wb.close()
        mut = recalc_libreoffice(src, td_path / "mut_full_recalc")
        rc = main_check(mut, config, skip_recalc=False, run_regression=False)
        if not record("FULL=DCF!B28", rc, "chain/coverage: FULL collapses to base"):
            return 1

        # 3) Tornado!B14 = 0
        print("=== 3) Tornado!B14 = 0 ===")
        src = td_path / "mut_b14.xlsx"
        shutil.copy2(good, src)
        wb = load_workbook(src)
        wb["Tornado"]["B14"] = 0
        wb.save(src)
        wb.close()
        mut = recalc_libreoffice(src, td_path / "mut_b14_recalc")
        rc = main_check(mut, config, skip_recalc=False, run_regression=False)
        if not record("Tornado!B14=0", rc, "chart-source constant 0"):
            return 1

        # 4) Football!B10 = 1+0
        print("=== 4) Football!B10 = 1+0 ===")
        src = td_path / "mut_fb10.xlsx"
        shutil.copy2(good, src)
        wb = load_workbook(src)
        wb["Football"]["B10"] = "=1+0"
        wb.save(src)
        wb.close()
        mut = recalc_libreoffice(src, td_path / "mut_fb10_recalc")
        rc = main_check(mut, config, skip_recalc=False, run_regression=False)
        if not record("Football!B10=1+0", rc, "envelope band low disconnected / chain break"):
            return 1

        # 5) Event_Tariff channel A sign flip (apparel) or commodity equivalent
        print("=== 5) channel A sign flip ===")
        src = td_path / "mut_sign.xlsx"
        shutil.copy2(good, src)
        wb = load_workbook(src)
        if second == "Event_Tariff":
            wb["Event_Tariff"]["D6"] = "=2*DCF!B28-B41"
            label = "Event_Tariff!D6=2*B28-B41"
        else:
            # mirror commodity channel A around base using its helper B43
            wb["Event_Commodity"]["D6"] = "=2*DCF!B28-B43"
            label = "Event_Commodity!D6=2*B28-B43"
        wb.save(src)
        wb.close()
        mut = recalc_libreoffice(src, td_path / "mut_sign_recalc")
        rc = main_check(mut, config, skip_recalc=False, run_regression=False)
        if not record(label, rc, "channel binding/sign flip"):
            return 1

        # 6) FULL drops a channel (tariff: D11=B42 omits elasticity; commodity: D11=B43 omits others)
        print("=== 6) FULL omits channel ===")
        src = td_path / "mut_omit.xlsx"
        shutil.copy2(good, src)
        wb = load_workbook(src)
        if second == "Event_Tariff":
            wb["Event_Tariff"]["D11"] = "=B42"
            label = "Event_Tariff!D11=B42"
        else:
            wb["Event_Commodity"]["D11"] = "=B43"
            label = "Event_Commodity!D11=B43"
        wb.save(src)
        wb.close()
        mut = recalc_libreoffice(src, td_path / "mut_omit_recalc")
        rc = main_check(mut, config, skip_recalc=False, run_regression=False)
        if not record(label, rc, "channel coverage: FULL omits elasticity/other channels"):
            return 1

        # 7) Football chart width series retargeted to B5:B10
        print("=== 7) Football chart series G→B ===")
        src = td_path / "mut_chartser.xlsx"
        shutil.copy2(good, src)
        wb = load_workbook(src)
        ch = wb["Football"]._charts[0]
        # series[1] is width/visible
        ch.series[1].val.numRef.f = "Football!$B$5:$B$10"
        wb.save(src)
        wb.close()
        mut = recalc_libreoffice(src, td_path / "mut_chartser_recalc")
        rc = main_check(mut, config, skip_recalc=False, run_regression=False)
        if not record("Football chart width→B5:B10", rc, "chart XML series ref mismatch"):
            return 1

        # 8) FULL helper double-counts a channel (commodity B47 / tariff B45 / rates via D11)
        print("=== 8) FULL double-count / omit-via-helper ===")
        src = td_path / "mut_full_dbl.xlsx"
        shutil.copy2(good, src)
        wb = load_workbook(src)
        if second == "Event_Commodity":
            # reviewer (d): double-count elasticity in the FULL equity helper
            wb["Event_Commodity"]["B47"] = "=B43+B44+B45+B45+B46-4*DCF!B28"
            label = "Event_Commodity!B47 double-count elast"
        elif second == "Event_Tariff":
            # analogue: FULL helper B45 becomes B42+B43 (double elast path) - use sum that drops mit and doubles elast
            wb["Event_Tariff"]["B45"] = "=B42+B43-DCF!B28"
            label = "Event_Tariff!B45 omits/distorts joint"
        else:
            wb["Event_Rates"]["D11"] = "=B30+B35+B33"
            label = "Event_Rates!D11 double-counts interest"
        wb.save(src)
        wb.close()
        mut = recalc_libreoffice(src, td_path / "mut_full_dbl_recalc")
        rc = main_check(mut, config, skip_recalc=False, run_regression=False)
        if not record(label, rc, "FULL reconciliation mismatch vs independent joint"):
            return 1

        # 9) Envelope skips a FULL proxy: MIN(B5:B6) instead of MIN(B5:B7)
        print("=== 9) Envelope MIN(B5:B6) skips second FULL ===")
        src = td_path / "mut_env.xlsx"
        shutil.copy2(good, src)
        wb = load_workbook(src)
        wb["Envelope"]["B9"] = "=MIN(B5:B6)"
        wb["Envelope"]["C9"] = "=MIN(C5:C6)"
        # keep MAX intact or also shrink — reviewer only mutated B9; also shrink C9 for $/sh consistency
        wb.save(src)
        wb.close()
        mut = recalc_libreoffice(src, td_path / "mut_env_recalc")
        rc = main_check(mut, config, skip_recalc=False, run_regression=False)
        if not record("Envelope!B9=MIN(B5:B6)", rc, "envelope coverage / reconciliation"):
            return 1

        def mutate(tag: str, edits: dict[str, Any]) -> int:
            src = td_path / f"mut_{tag}.xlsx"
            shutil.copy2(good, src)
            wb = load_workbook(src)
            for ref, val in edits.items():
                sh, cell = parse_cell_ref(ref)
                wb[sh][cell] = val
            wb.save(src)
            wb.close()
            mut = recalc_libreoffice(src, td_path / f"mut_{tag}_recalc")
            return main_check(mut, config, skip_recalc=False, run_regression=False)

        cached = read_values(good, ["Event_Rates!D12", "Summary!E5"])

        # 10) Interaction row pasted as a number (looks right today, goes stale tomorrow)
        print("=== 10) interaction row pasted ===")
        rc = mutate("int_paste", {"Event_Rates!D12": cached["Event_Rates!D12"] or 0.0})
        if not record("Event_Rates!D12 pasted value", rc, "pasted constant / interaction not live"):
            return 1

        # 11) Interaction row sums an INFORMATION row (B) as if it were a standalone channel
        print("=== 11) interaction includes information row ===")
        rc = mutate("int_info", {
            "Event_Rates!D12": "=(D11-DCF!$B$28)-((D6-DCF!$B$28)+(D7-DCF!$B$28)+(D8-DCF!$B$28))",
        })
        if not record("Event_Rates!D12 includes info row B", rc, "interaction != FULL − standalones"):
            return 1

        # 12) Summary ranking pasted
        print("=== 12) Summary ranked cell pasted ===")
        rc = mutate("sum_paste", {"Summary!E5": cached["Summary!E5"] or 0.0})
        if not record("Summary!E5 pasted value", rc, "Summary ranking not live"):
            return 1

        # 13) Reverse-stress result pasted as n/a while a threshold is set
        print("=== 13) reverse-stress result pasted ===")
        thr = read_values(good, ["DCF!B33"])["DCF!B33"]
        rc = mutate("rev_paste", {"Event_Rates!B66": round(float(thr) * 0.9, 2), "Event_Rates!B67": "n/a"})
        if not record("Event_Rates!B67 pasted 'n/a' with threshold", rc, "reverse-stress result wrong / not live"):
            return 1

        if has_capex(good):
            # 14) FULL omits the capex financing channel (operating-only A+B row)
            print("=== 14) Event_Capex FULL omits financing ===")
            rc = mutate("capex_nofin", {"Event_Capex!D11": "=B46"})
            if not record("Event_Capex!D11=B46 (no financing)", rc, "FULL omits financing mix channel"):
                return 1

            # 15) Financing counted twice (net-debt change subtracted twice)
            print("=== 15) Event_Capex financing double count ===")
            rc = mutate("capex_dbl", {"Event_Capex!B47": "=(B46-B40-B40)*Inputs!B39/B41"})
            if not record("Event_Capex!B47 double-counts financing", rc, "FULL reconciliation mismatch"):
                return 1

            # 16) Envelope skips the capex FULL proxy
            print("=== 16) Envelope MIN(B5:B7) skips capex ===")
            rc = mutate("env_capex", {"Envelope!B9": "=MIN(B5:B7)", "Envelope!C9": "=MIN(C5:C7)"})
            if not record("Envelope!B9=MIN(B5:B7) with capex", rc, "envelope coverage"):
                return 1


        # 17-19) Shock ladder
        lay = ladder_layout(len(event_sheets(good)))
        rates_rows = lay["blocks"][0]["rows"]
        r_plus1 = rates_rows[5]
        print("=== 17) Ladder rung links to the wrong helper (mirror rung) ===")
        rc = mutate("ladder_wrong", {f"Ladder!D{r_plus1}": f"=Event_Rates!G{91 + 1}"})
        if not record(f"Ladder!D{r_plus1} wrong rung", rc, "ladder rung != independent FULL"):
            return 1
        print("=== 18) Reverse-stress result pasted ===")
        rr0 = lay["rev_rows"][0]
        rc = mutate("ladder_rev_paste", {f"Ladder!B{rr0}": 100})
        if not record(f"Ladder!B{rr0} pasted 100", rc, "reverse result not live / mismatch"):
            return 1
        print("=== 19) Ladder chart plots % change instead of equity ===")
        src = td_path / "mut_ladder_chart.xlsx"
        a, b = rates_rows[0], rates_rows[-1]
        _rewrite_chart_xml(good, src, f"$D${a}:$D${b}", f"$F${a}:$F${b}")
        rc = main_check(src, config, skip_recalc=True, run_regression=False)
        if not record("Ladder chart y -> F (% change)", rc, "chart series wrong"):
            return 1

    print("\n=== Regression table ===")
    print(f"{'mutation':<42} {'result':<6} reason")
    for name, status, reason in results:
        print(f"{name:<42} {status:<6} {reason}")
    print("=== Regression mutations: all correctly FAILED ===")
    return 0


def main_check(
    workbook: Path,
    config: Path | None,
    skip_recalc: bool = False,
    run_regression: bool = False,
    strict_calibration: bool = False,
) -> int:
    manual: list[str] = []
    warnings: list[str] = []
    failures: list[str] = []
    details: list[str] = []

    print(f"Checking {workbook}")

    # 1) Input caches must be present (unless this IS the empty-cache regression probe with skip)
    ok_cache, msg_cache = check_input_caches(workbook)
    details.append(msg_cache)
    if not ok_cache:
        failures.append(msg_cache)
        # Still print and fail fast for empty cache — do not mask with temp recalc success
        _print_report(details, manual, failures)
        return 1

    second = detect_second_event(workbook)
    capex = has_capex(workbook)
    details.append(f"Event sheets: {', '.join(event_sheets(workbook))}")

    cfg: dict[str, Any] = {}
    if config:
        cfg = yaml.safe_load(config.read_text(encoding="utf-8"))
    checks_cfg = cfg.get("checks") or {}
    live_tests = checks_cfg.get("live_tests") or default_live_tests(second, capex)
    shock_tests = checks_cfg.get("shock_tests") or default_shock_tests(second, capex)
    channel_tests = checks_cfg.get("channel_tests") or default_channel_tests(second, capex)
    channel_bindings = checks_cfg.get("channel_bindings") or default_channel_bindings(second, capex)
    charts_cfg = checks_cfg.get("charts") or default_charts_cfg(capex)
    if capex:
        # A config that lists its own tests must still cover the capex event (financing channel included).
        if not any(t.get("full_cell") == "Event_Capex!D11" for t in shock_tests):
            shock_tests = list(shock_tests) + [dict(CAPEX_SHOCK_TEST)]
        if not any(t.get("full_cell") == "Event_Capex!D11" for t in channel_tests):
            channel_tests = list(channel_tests) + [dict(t) for t in CAPEX_CHANNEL_TESTS]
        if not any(b["cell"].startswith("Event_Capex!") for b in channel_bindings):
            channel_bindings = list(channel_bindings) + [dict(b) for b in CAPEX_BINDINGS]

    # Normalize second-event refs in live tests if config still says Event_Commodity
    if second == "Event_Tariff":
        for t in live_tests:
            t["expect_move"] = [r.replace("Event_Commodity!", "Event_Tariff!") for r in t.get("expect_move", [])]
            t["expect_unchanged"] = [
                r.replace("Event_Commodity!", "Event_Tariff!") for r in t.get("expect_unchanged", [])
            ]

    chain_fail, chain_detail = verify_dependency_chain(workbook)
    failures.extend(chain_fail)
    details.extend(chain_detail)

    bind_fail, bind_detail = verify_channel_bindings(workbook, channel_bindings)
    failures.extend(bind_fail)
    details.extend(bind_detail)

    chart_fail, chart_detail = inspect_chart_series(workbook, charts_cfg)
    failures.extend(chart_fail)
    details.extend(chart_detail)

    env_fail, env_detail = verify_envelope_formula_coverage(workbook)
    failures.extend(env_fail)
    details.extend(env_detail)

    pasted = scan_pasted_constants(workbook)
    if pasted:
        failures.append(f"Pasted constants in output ranges: {pasted[:8]}")
    else:
        details.append("Pasted-constant scan: clean")

    with tempfile.TemporaryDirectory(prefix="stress-check-") as td:
        td_path = Path(td)
        base = workbook if skip_recalc else recalc_libreoffice(workbook, td_path / "initial")
        if not skip_recalc:
            details.append("LibreOffice recalc (temp): ok")

        errors = scan_formula_errors(base)
        if errors:
            failures.append(f"Formula errors: {errors[:10]}")
        else:
            details.append("Formula errors: 0")

        ok_pt, msg_pt = check_passthrough(base)
        details.append(msg_pt)
        if not ok_pt:
            failures.append(msg_pt)

        # Independent FULL + envelope reconciliation (does not read channel/FULL formulas)
        recon_fail, recon_detail, _ind = reconcile_events(workbook, second)
        failures.extend(recon_fail)
        details.extend(recon_detail)

        # Interaction rows, mirror, reverse-stress grid/result, calibration, Summary ranking
        ex_fail, ex_detail, ex_manual, to_fill = verify_event_extras(base)
        failures.extend(ex_fail)
        details.extend(ex_detail)
        manual.extend(ex_manual)
        cal_warn, cal_fail = calibration_messages(to_fill, strict_calibration)
        warnings.extend(cal_warn)
        failures.extend(cal_fail)

        # Shock ladder + reverse stress test (independent reference; ±1× rung tied to mirror asymmetry)
        ls_fail, ls_detail = verify_ladder_structure(workbook)
        failures.extend(ls_fail)
        details.extend(ls_detail)
        ld_fail, ld_detail = reconcile_ladder(base)
        failures.extend(ld_fail)
        details.extend(ld_detail)

        # Tariff-specific: if Event_Tariff present, FULL equity < base at configured <100% PT
        if second == "Event_Tariff":
            vals = read_values(base, ["DCF!B28", "Event_Tariff!D11", "Event_Tariff!B26", "Event_Tariff!B48", "Event_Tariff!B49"])
            details.append(
                f"Tariff sign: base={vals['DCF!B28']} FULL={vals['Event_Tariff!D11']} "
                f"PT={vals['Event_Tariff!B26']} invariantΔGP={vals.get('Event_Tariff!B48')} status={vals.get('Event_Tariff!B49')}"
            )
            if vals["DCF!B28"] is not None and vals["Event_Tariff!D11"] is not None:
                if float(vals["Event_Tariff!D11"]) >= float(vals["DCF!B28"]):
                    failures.append(
                        f"Tariff FULL equity {vals['Event_Tariff!D11']} should be < base {vals['DCF!B28']} "
                        f"at partial pass-through"
                    )
                else:
                    details.append("Tariff sign: FULL < base OK")

        # WACC override
        wb = load_workbook(workbook)
        wb["Inputs"]["B28"] = 0.12
        override_path = td_path / "override.xlsx"
        wb.save(override_path)
        wb.close()
        # Override save clears caches — recalc required
        over = recalc_libreoffice(override_path, td_path / "override_out")
        # After mutation+recalc, caches exist on temp file
        vals = read_values(over, ["Inputs!B46", "Event_Rates!B25", "Inputs!B48", "DCF!B28", "Event_Rates!D6"])
        selected, stressed, delta = vals["Inputs!B46"], vals["Event_Rates!B25"], vals["Inputs!B48"]
        details.append(f"WACC override test: selected={selected}, stressed={stressed}, Δ={delta}")
        if selected is None or stressed is None or delta is None:
            failures.append("WACC override test: missing values")
        else:
            if abs(float(selected) - 0.12) > 1e-9:
                failures.append(f"Selected WACC expected 0.12, got {selected}")
            if abs(float(stressed) - (float(selected) + float(delta))) > 1e-9:
                failures.append(f"Stressed WACC {stressed} != selected+Δ {float(selected)+float(delta)}")
            base_eq, a_eq = vals["DCF!B28"], vals["Event_Rates!D6"]
            if base_eq is not None and a_eq is not None and float(a_eq) >= float(base_eq):
                failures.append(
                    f"With WACC override 12% and +rate shock, channel A equity {a_eq} should be < base {base_eq}"
                )

        live_fail, live_detail = run_live_tests(workbook, live_tests, td_path / "live", second)
        failures.extend(live_fail)
        details.extend(live_detail)

        shock_fail, shock_detail = run_shock_tests(workbook, shock_tests, td_path / "shock")
        failures.extend(shock_fail)
        details.extend(shock_detail)

        ch_fail, ch_detail = run_channel_tests(workbook, channel_tests, td_path / "channel")
        failures.extend(ch_fail)
        details.extend(ch_detail)

    # Structural rates check
    wb = load_workbook(workbook, data_only=False)
    full_fx = wb["Event_Rates"]["D11"].value
    kind_b = wb["Event_Rates"]["B7"].value
    kind_e = wb["Event_Rates"]["B10"].value
    wb.close()
    details.append(f"Event_Rates FULL formula: {full_fx}")
    if full_fx and "B33" in str(full_fx) and "B37" in str(full_fx):
        failures.append("Rates FULL still references both interest PV (B33) and buyback (B37)")
    if str(kind_b).lower() != "information" or str(kind_e).lower() != "information":
        failures.append(f"Rates B/E should be information rows, got B={kind_b}, E={kind_e}")

    _print_report(details, manual, failures, warnings)
    rc = 1 if failures else 0

    if run_regression and rc == 0:
        print("\n")
        return run_regression_mutations(workbook, config)
    return rc


def _print_report(details, manual, failures, warnings=()) -> None:
    print("\n--- Details ---")
    for line in details:
        print(line)
    if warnings:
        print("\n--- Warnings ---")
        for w in warnings:
            print(f"WARNING: {w}")
    print("\n--- Manual review list ---")
    if not manual:
        print("  (none automated; still run REVIEW_CHECKLIST.md for CLEAR/BLOCK)")
    for m in manual:
        print(f"  - {m}")
    print("  - Confirm tornado and football charts render after open in Excel/LibreOffice")
    print("  - Confirm ILLUSTRATIVE labels and draft disclaimer on Cover")
    print("  - Gate any real-issuer workbook with REVIEW_CHECKLIST.md")
    print("\n--- Verdict ---")
    if failures:
        print("FAIL")
        for f in failures:
            print(f"  - {f}")
    else:
        print("PASS")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Check a same-event stress-test workbook")
    parser.add_argument("workbook", type=Path, help="Path to .xlsx (must already have formula caches)")
    parser.add_argument("--config", type=Path, default=None)
    parser.add_argument(
        "--skip-recalc",
        action="store_true",
        help="Do not temp-recalc (still requires input caches unless used for empty-cache probe)",
    )
    parser.add_argument(
        "--regression",
        action="store_true",
        help="After PASS, mutate workbook two ways + empty-cache and require those checks to FAIL",
    )
    parser.add_argument(
        "--strict-calibration",
        action="store_true",
        help="Treat calibration source/basis cells still marked TO FILL as FAIL (use for a real-issuer CLEAR)",
    )
    args = parser.parse_args(argv)
    return main_check(args.workbook, args.config, skip_recalc=args.skip_recalc, run_regression=args.regression,
                      strict_calibration=args.strict_calibration)


if __name__ == "__main__":
    sys.exit(main())
