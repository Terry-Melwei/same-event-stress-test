#!/usr/bin/env python3
"""Automated workbook gate for same-event multi-channel stress tests.

Fails on: empty formula caches in the input file, pasted outputs, broken
FULL→tornado→envelope→football→chart-source links, event shock/sign misses,
pass-through invariant breaks, and (via --regression) known broken mutations.
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
from reference_dcf import reconcile_events  # noqa: E402

OUTPUT_RANGES = {
    "DCF": ["B19:B33", "B5:F16"],
    "Event_Rates": ["D6:G11", "B23:B37"],
    "Event_Commodity": ["D6:G11", "B43:B49", "B32:F32", "B37:F37", "B39:F41"],
    "Event_Tariff": ["D6:G11", "B40:B52"],
    "Tornado": ["B5:F6", "B8:B9", "B13:B14"],
    "Envelope": ["B5:C7", "B9:C10"],
    "Football": ["B5:D10", "F5:G10", "B12", "B14"],
    "PassThrough": ["B16:B19", "B22:B31", "B34:B37"],
}

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


def chain_refs(second: str) -> dict[str, str]:
    return {k: v.format(second=second) for k, v in CHAIN_TEMPLATE.items()}


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
    wb = load_workbook(path, data_only=True)
    required = ["DCF!B30", "DCF!B28", "Event_Rates!D11", "Tornado!C5", "Envelope!B9", "Football!B8"]
    second = None
    try:
        if "Event_Commodity" in wb.sheetnames:
            required.append("Event_Commodity!D11")
            second = "Event_Commodity"
        if "Event_Tariff" in wb.sheetnames:
            # If both exist, require the one tornado links
            pass
    finally:
        pass
    wb.close()
    second = detect_second_event(path)
    required.append(f"{second}!D11")
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
    refs = chain_refs(second)
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

    details.append(f"Dependency chain traced for rates + {second} ({len(refs)} cells)")
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


def default_shock_tests(second: str) -> list[dict]:
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
    return tests


def default_live_tests(second: str) -> list[dict]:
    return [
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
            ],
            "expect_unchanged": [],  # second FULL may move via WACC in DCF shared paths for tariff; leave empty
        },
    ]



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
                for bad in ("FOOTBALL!B5:B10", "FOOTBALL!$B$5:$B$10", "FOOTBALL!C5:C10"):
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


def default_channel_tests(second: str) -> list[dict]:
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


def default_channel_bindings(second: str) -> list[dict]:
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


def default_charts_cfg() -> dict:
    return {
        "tornado": {
            "value_ref": "Tornado!$B$13:$B$14",
            "cat_ref": "Tornado!$A$13:$A$14",
        },
        "football": {
            "baseline_value_ref": "Football!$F$5:$F$10",
            "width_value_ref": "Football!$G$5:$G$10",
            "cat_ref": "Football!$A$5:$A$10",
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
    """Parse Envelope MIN/MAX ranges; require {B5,B6,B7} (base + both FULL proxies)."""
    failures, details = [], []
    fx = read_formulas(path, ["Envelope!B9", "Envelope!B10", "Envelope!C9", "Envelope!C10",
                              "Envelope!B5", "Envelope!B6", "Envelope!B7",
                              "Football!B10", "Football!D10"])
    # $m column uses B5:B7; $/sh column uses C5:C7 — both must span base + rates + second FULL
    required_by_cell = {
        "Envelope!B9": {"B5", "B6", "B7"},
        "Envelope!B10": {"B5", "B6", "B7"},
        "Envelope!C9": {"C5", "C6", "C7"},
        "Envelope!C10": {"C5", "C6", "C7"},
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

    # Tornado endpoints must feed Envelope B6/B7
    # (structural: B6=Tornado!C5, B7=Tornado!C6 already in chain; reinforce here)
    return failures, details


def run_regression_mutations(good: Path, config: Path | None) -> int:
    """Reviewer mutations (9+) — each must FAIL."""
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
) -> int:
    manual: list[str] = []
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
    details.append(f"Second event sheet: {second}")

    cfg: dict[str, Any] = {}
    if config:
        cfg = yaml.safe_load(config.read_text(encoding="utf-8"))
    checks_cfg = cfg.get("checks") or {}
    live_tests = checks_cfg.get("live_tests") or default_live_tests(second)
    shock_tests = checks_cfg.get("shock_tests") or default_shock_tests(second)
    channel_tests = checks_cfg.get("channel_tests") or default_channel_tests(second)
    channel_bindings = checks_cfg.get("channel_bindings") or default_channel_bindings(second)
    charts_cfg = checks_cfg.get("charts") or default_charts_cfg()

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

    _print_report(details, manual, failures)
    rc = 1 if failures else 0

    if run_regression and rc == 0:
        print("\n")
        return run_regression_mutations(workbook, config)
    return rc


def _print_report(details, manual, failures) -> None:
    print("\n--- Details ---")
    for line in details:
        print(line)
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
    args = parser.parse_args(argv)
    return main_check(args.workbook, args.config, skip_recalc=args.skip_recalc, run_regression=args.regression)


if __name__ == "__main__":
    sys.exit(main())
