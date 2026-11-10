"""CI tests for the generator, the shipped workbooks and the checker.

No LibreOffice needed:
1. generate.py runs on every example config (catches syntax/runtime breakage).
2. Freshly generated formulas match the shipped, recalculated workbooks.
3. check_workbook.py FAILs on a fresh, uncached workbook (empty-cache guard).
4. The capex financing channel is paired once (new shares × issue price = net-debt change).

Need LibreOffice (`soffice` on PATH; the CI workflow installs it), skipped otherwise:
5. check_workbook.py PASSes on every shipped workbook, printing a WARNING that lists every event whose
   calibration source/basis is still TO FILL; --strict-calibration turns that into FAIL (and PASSes once
   the cells are filled); an empty calibration cell FAILs either way.
6. After recalculation, every interaction row equals FULL minus the standalone channel rows.
7. The capex FULL / standalones match the independent reference (reference_dcf.py).
8. The reverse-stress threshold defaults to base $/sh × (1 − drop %), drop % 25%; the result is the
   smallest tested adverse step whose FULL $/sh is below the threshold (and the previous step is
   at/above it); a blank drop % or a blank threshold gives 'n/a'.
9. The Summary sheet ranks events by absolute FULL change in $/sh.

The slower --regression mutation run is left to local use (see README).
"""
import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from reference_dcf import (  # noqa: E402
    base_equity_perp, capex_components, covenant_at, demand_components,
    format_not_reached, fx_components, load_inputs_from_workbook,
)

CASES = {
    "default_illustrative": ROOT / "templates" / "stress_template.xlsx",
    "gold_miner": ROOT / "examples" / "workbooks" / "gold_miner.xlsx",
    "apparel_retailer": ROOT / "examples" / "workbooks" / "apparel_retailer.xlsx",
    "optical_supplier": ROOT / "examples" / "workbooks" / "optical_supplier.xlsx",
}
STANDALONE_ROWS = {"Event_Rates": [6, 8], "Event_Commodity": [6, 7, 8, 9], "Event_Capex": [6, 7, 8],
                   "Event_FX": [6, 7, 8], "Event_Demand": [6, 7, 8, 9]}
GRID = range(69, 77)
INTERACTION_LABEL = "Cross-channel interaction (not size non-linearity)"
THRESHOLD_DEFAULT_FX = '=IF(D66="","",DCF!$B$33*(1-D66))'
needs_lo = pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice (soffice) not installed")


def _run(*args):
    return subprocess.run([sys.executable, *map(str, args)], cwd=ROOT,
                          capture_output=True, text=True)


def _config(name):
    return ROOT / "examples" / "configs" / f"{name}.yaml"


def _recalc(src: Path, out_dir: Path) -> Path:
    from check_workbook import recalc_libreoffice
    return recalc_libreoffice(src, out_dir)


def _event_sheets(wb):
    return [s for s in wb.sheetnames if s.startswith("Event_")]


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


@pytest.fixture(scope="module")
def recalculated(generated, tmp_path_factory):
    if shutil.which("soffice") is None:
        pytest.skip("LibreOffice (soffice) not installed")
    out = tmp_path_factory.mktemp("recalc")
    return {name: _recalc(path, out / name) for name, path in generated.items()}


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


def test_summary_sheet_follows_cover(generated):
    for path in generated.values():
        names = openpyxl.load_workbook(path).sheetnames
        assert names[:2] == ["Cover", "Summary"]


def test_capex_financing_paired_once():
    """Equity-funded cash lowers net debt by exactly new shares × issue price (one step, both sides)."""
    wb = openpyxl.load_workbook(CASES["optical_supplier"], data_only=True)
    cx = wb["Event_Capex"]
    new_shares, price, nd_change = cx["B39"].value, cx["B25"].value, cx["B40"].value
    assert new_shares > 0
    assert new_shares * price == pytest.approx(-nd_change)
    assert cx["B41"].value == pytest.approx(wb["Inputs"]["B39"].value + new_shares)
    # FULL must include the financing step: it differs from the operating-only information row
    assert cx["D11"].value != pytest.approx(cx["D9"].value)


@needs_lo
@pytest.mark.parametrize("name", list(CASES))
def test_checker_passes_on_shipped(name):
    res = _run("scripts/check_workbook.py", CASES[name], "--config", _config(name), "--skip-recalc")
    assert res.returncode == 0, res.stdout + res.stderr
    assert res.stdout.strip().splitlines()[-1] == "PASS"
    # Shipped calibration cells are TO FILL: one WARNING naming every event sheet (source + basis)
    warn = [ln for ln in res.stdout.splitlines() if ln.startswith("WARNING: Calibration still TO FILL")]
    assert len(warn) == 1, res.stdout
    sheets = _event_sheets(openpyxl.load_workbook(CASES[name], read_only=True))
    for sheet in sheets:
        assert f"{sheet} (source + basis + severity)" in warn[0]


def _check(path, name, *flags):
    res = _run("scripts/check_workbook.py", path, "--config", _config(name), "--skip-recalc", *flags)
    return res.returncode, res.stdout + res.stderr


@needs_lo
def test_strict_calibration(recalculated, tmp_path):
    name = "gold_miner"
    rc, out = _check(CASES[name], name, "--strict-calibration")
    assert rc != 0 and "[--strict-calibration]" in out, out
    assert "WARNING: Calibration" not in out
    # Filled source/basis on every event sheet -> strict PASS, no warning
    wb = openpyxl.load_workbook(recalculated[name])
    for sheet in _event_sheets(wb):
        wb[sheet]["B56"] = "ILLUSTRATIVE test source, 2026-11-10"
        wb[sheet]["B57"] = "ILLUSTRATIVE test basis"
        wb[sheet]["B58"] = "plausible"
    filled = tmp_path / "filled.xlsx"
    wb.save(filled)
    filled = _recalc(filled, tmp_path / "filled_out")
    rc, out = _check(filled, name, "--strict-calibration")
    assert rc == 0, out
    assert "WARNING: Calibration" not in out
    # Empty calibration cell is a FAIL even without --strict-calibration
    wb = openpyxl.load_workbook(recalculated[name])
    wb["Event_Rates"]["B56"] = None
    empty = tmp_path / "empty_cal.xlsx"
    wb.save(empty)
    empty = _recalc(empty, tmp_path / "empty_out")
    rc, out = _check(empty, name)
    assert rc != 0 and "Event_Rates!B56 (source) is empty" in out, out


def test_checker_rejects_uncached(generated):
    res = _run("scripts/check_workbook.py", generated["gold_miner"],
               "--config", _config("gold_miner"), "--skip-recalc")
    assert res.returncode != 0
    assert "Empty formula caches" in res.stdout + res.stderr


@needs_lo
@pytest.mark.parametrize("name", list(CASES))
def test_interaction_row_equals_full_minus_standalones(recalculated, name):
    wb = openpyxl.load_workbook(recalculated[name], data_only=True)
    wf = openpyxl.load_workbook(recalculated[name])
    base = wb["DCF"]["B28"].value
    for sheet in _event_sheets(wb):
        ws = wb[sheet]
        assert ws["B12"].value == "diagnostic"
        assert ws["C12"].value == INTERACTION_LABEL
        assert str(wf[sheet]["D12"].value).startswith("=")
        assert "12" not in str(wf[sheet]["D11"].value)  # FULL never reads the diagnostic
        d = {r: ws[f"D{r}"].value for r in range(6, 12)}
        if sheet == "Event_Tariff":
            expected = d[11] - d[9] - (ws["B47"].value - d[6])
        else:
            expected = (d[11] - base) - sum(d[r] - base for r in STANDALONE_ROWS[sheet])
        assert ws["D12"].value == pytest.approx(expected, abs=1e-6), sheet
        assert ws["G12"].value == pytest.approx(
            wb["DCF"]["B30"].value * expected / base / wb["Inputs"]["B39"].value, abs=1e-9)
    assert wb["Summary"]["I4"].value.startswith(INTERACTION_LABEL)
    # Blend lock: the blended midpoint still averages the two method rows only
    assert wf["DCF"]["B30"].value == "=(B28+B29)/2"


@needs_lo
def test_capex_matches_independent_reference(recalculated):
    path = recalculated["optical_supplier"]
    wb = openpyxl.load_workbook(path, data_only=True)
    cx = wb["Event_Capex"]
    m = load_inputs_from_workbook(path)
    ref = capex_components(m)
    assert base_equity_perp(m) == pytest.approx(wb["DCF"]["B28"].value, rel=1e-9)
    for row, key in ((6, "A"), (7, "B"), (8, "C"), (9, "ops"), (11, "full")):
        assert cx[f"D{row}"].value == pytest.approx(ref[key], rel=1e-9), key
    # Mirror (−shock) has no funding need, so the financing channel switches off: asymmetric by design
    mirror = capex_components(m, -m.capex_shock)
    assert mirror["need"] == 0
    assert cx["D62"].value == pytest.approx(mirror["full"], rel=1e-9)


def _grid(ws):
    return [(ws[f"C{r}"].value, ws[f"G{r}"].value) for r in GRID]


@needs_lo
@pytest.mark.parametrize("name", ["default_illustrative", "apparel_retailer", "optical_supplier"])
def test_reverse_stress_threshold(recalculated, tmp_path, name):
    shipped = openpyxl.load_workbook(recalculated[name], data_only=True)
    sheets = _event_sheets(shipped)
    formulas = openpyxl.load_workbook(recalculated[name])
    base_sh = shipped["DCF"]["B33"].value
    # 1) Default threshold = live formula base × (1 − drop %), drop % a yellow input defaulting to 25%;
    #    a config threshold_value_per_share is a typed override (optical Event_Capex).
    for sheet in sheets:
        ws = shipped[sheet]
        assert ws["D66"].value == pytest.approx(0.25)
        if formulas[sheet]["B66"].value == THRESHOLD_DEFAULT_FX:
            assert ws["B66"].value == pytest.approx(base_sh * 0.75)
        else:
            assert isinstance(formulas[sheet]["B66"].value, (int, float))
        grid = _grid(ws)
        below = [x for x, sh in grid if sh < ws["B66"].value]
        if below:
            assert ws["B67"].value == pytest.approx(below[0]), (sheet, grid)
        else:
            shock, shv = min(grid, key=lambda p: p[1])
            unit = "bp" if sheet == "Event_Rates" else "%"
            expected = format_not_reached(shv / base_sh - 1.0, shock, unit)
            assert ws["B67"].value == expected, (sheet, grid, ws["B67"].value)
    # 1b) Blank drop % (default formula) or blank threshold (typed override) -> 'n/a', also on Summary
    blank = tmp_path / f"{name}_blank.xlsx"
    wb = openpyxl.load_workbook(recalculated[name])
    for sheet in sheets:
        if formulas[sheet]["B66"].value == THRESHOLD_DEFAULT_FX:
            wb[sheet]["D66"] = None
        else:
            wb[sheet]["B66"] = None
    wb.save(blank)
    out = openpyxl.load_workbook(_recalc(blank, tmp_path / "blank"), data_only=True)
    for sheet in sheets:
        assert out[sheet]["B67"].value == "n/a", sheet
        assert out[sheet]["B66"].value in (None, ""), sheet
    assert [out["Summary"][f"K{5 + k}"].value for k in range(len(sheets))] == ["n/a"] * len(sheets)
    # 2) Threshold between grid steps 2 and 3 -> step 3 reported; step 2 at/above it.
    #    Also one threshold above base ("base already below") and one below every step ("none").
    thresholds = {}
    for sheet in sheets:
        g = [sh for _x, sh in _grid(shipped[sheet])]
        assert g[1] > g[2], f"{sheet}: grid not adverse in the configured direction"
        thresholds[sheet] = (g[1] + g[2]) / 2
    copy = tmp_path / f"{name}_thr.xlsx"
    wb = openpyxl.load_workbook(recalculated[name])
    for sheet, thr in thresholds.items():
        wb[sheet]["B66"] = thr
    wb.save(copy)
    out = openpyxl.load_workbook(_recalc(copy, tmp_path / "thr"), data_only=True)
    for sheet, thr in thresholds.items():
        ws = out[sheet]
        grid = _grid(ws)
        result = ws["B67"].value
        below = [i for i, (_x, sh) in enumerate(grid) if sh < thr]
        assert below, sheet
        k = below[0]
        assert result == pytest.approx(grid[k][0]), (sheet, result, grid)
        assert grid[k][1] < thr <= grid[k - 1][1] if k > 0 else out["DCF"]["B33"].value >= thr
        assert k == 2
        assert ws["C67"].value == pytest.approx(ws[f"B{GRID[k]}"].value)

    edge = tmp_path / f"{name}_edge.xlsx"
    wb = openpyxl.load_workbook(recalculated[name])
    wb["Event_Rates"]["B66"] = base_sh * 1.01
    wb[sheets[1]]["B66"] = min(sh for _x, sh in _grid(shipped[sheets[1]])) - 1.0
    wb.save(edge)
    out = openpyxl.load_workbook(_recalc(edge, tmp_path / "edge"), data_only=True)
    assert out["Event_Rates"]["B67"].value == "base already below threshold"
    grid = _grid(out[sheets[1]])
    shock, shv = min(grid, key=lambda p: p[1])
    unit = "bp" if sheets[1] == "Event_Rates" else "%"
    expected = format_not_reached(shv / base_sh - 1.0, shock, unit)
    assert out[sheets[1]]["B67"].value == expected


@needs_lo
def test_summary_ranks_by_abs_full_change(recalculated):
    for name, path in recalculated.items():
        wb = openpyxl.load_workbook(path, data_only=True)
        base_sh = wb["DCF"]["B33"].value
        sheets = _event_sheets(wb)
        deltas = {s: wb[s]["G11"].value - base_sh for s in sheets}
        order = sorted(sheets, key=lambda s: -abs(deltas[s]))
        sm = wb["Summary"]
        assert [sm[f"C{5 + k}"].value for k in range(len(sheets))] == order, name
        for k, s in enumerate(order):
            r = 5 + k
            pair = sorted([wb[s]["G11"].value, wb[s]["G62"].value])
            assert sm[f"E{r}"].value == pytest.approx(pair[0])
            assert sm[f"F{r}"].value == pytest.approx(pair[1])
            assert sm[f"H{r}"].value == pytest.approx(deltas[s])
            assert sm[f"I{r}"].value == pytest.approx(wb[s]["G12"].value)
            assert sm[f"K{r}"].value == wb[s]["B67"].value


@needs_lo
def test_ladder_plus_one_matches_mirror(recalculated):
    """Product manager's test: the ±1× rung's asymmetry equals the event-sheet mirror (and Summary)."""
    for name, path in recalculated.items():
        wb = openpyxl.load_workbook(path, data_only=True)
        wf = openpyxl.load_workbook(path)
        assert "Ladder" in wb.sheetnames, name
        sheets = _event_sheets(wb)
        # layout: first block starts at row 6, 12 rows per event, rungs at title+2 .. +8
        for k, sheet in enumerate(sheets):
            srow = 6 + 12 * k
            # multiples in column B, rungs start at srow+2
            rows = list(range(srow + 2, srow + 2 + 7))
            mults = [wb["Ladder"][f"B{r}"].value for r in rows]
            assert 1.0 in [float(m) for m in mults], (name, sheet, mults)
            r1 = rows[[float(m) for m in mults].index(1.0)]
            lad = wb["Ladder"][f"H{r1}"].value
            mir = wb[sheet]["G63"].value
            assert lad == pytest.approx(mir, abs=1e-4), (name, sheet, lad, mir)
            assert wf["Ladder"][f"I{srow + 2}"].value is not None
            assert str(wb["Ladder"][f"I{srow + 2}"].value) == "OK", (name, sheet)
        # a -25% target that sits outside the ladder still names the worst fall
        # (rates on default/gold/apparel often do not reach -25% by +1.5×)
        # Just require any non-numeric reverse result to start with 'Not reached:'
        from reference_dcf import ladder_layout
        lay = ladder_layout(len(sheets))
        for rr in lay["rev_rows"]:
            for k, sheet in enumerate(sheets):
                got = wb["Ladder"].cell(rr, 2 + k).value
                if not isinstance(got, (int, float)):
                    assert isinstance(got, str) and got.startswith("Not reached:"), (name, sheet, got)
        wb.close()
        wf.close()


@needs_lo
def test_fx_and_demand_match_independent_reference(recalculated):
    path = recalculated["apparel_retailer"]
    wb = openpyxl.load_workbook(path, data_only=True)
    m = load_inputs_from_workbook(path)
    fx = fx_components(m)
    dm = demand_components(m)
    assert wb["Event_FX"]["D6"].value == pytest.approx(fx["A"], rel=1e-9)
    assert wb["Event_FX"]["D7"].value == pytest.approx(fx["B"], rel=1e-9)
    assert wb["Event_FX"]["D8"].value == pytest.approx(fx["C"], rel=1e-9)
    assert wb["Event_FX"]["D11"].value == pytest.approx(fx["full"], rel=1e-9)
    assert wb["Event_Demand"]["D6"].value == pytest.approx(dm["A"], rel=1e-9)
    assert wb["Event_Demand"]["D7"].value == pytest.approx(dm["B"], rel=1e-9)
    assert wb["Event_Demand"]["D8"].value == pytest.approx(dm["C"], rel=1e-9)
    assert wb["Event_Demand"]["D9"].value == pytest.approx(dm["D"], rel=1e-9)
    assert wb["Event_Demand"]["D11"].value == pytest.approx(dm["full"], rel=1e-9)
    # FX is linear: the mirror of +x equals FULL at -x and the asymmetry is ~0
    assert wb["Event_FX"]["D62"].value == pytest.approx(fx_components(m, -m.fx_shock)["full"], rel=1e-9)
    assert wb["Event_FX"]["D63"].value == pytest.approx(0.0, abs=0.05)
    # Demand is not linear: the revenue product and downside inventory create a non-zero asymmetry
    assert wb["Event_Demand"]["D62"].value == pytest.approx(demand_components(m, -m.demand_shock)["full"], rel=1e-9)
    assert abs(wb["Event_Demand"]["D63"].value) > 0.05
    wb.close()


@needs_lo
def test_covenant_row_matches_reference(recalculated):
    for name, path in recalculated.items():
        wb = openpyxl.load_workbook(path, data_only=True)
        m = load_inputs_from_workbook(path)
        sheets = _event_sheets(wb)
        for sheet in sheets:
            shock = {"Event_Rates": m.rate_shock_bp, "Event_Commodity": m.commodity_shock,
                     "Event_Tariff": m.tariff_cost_shock, "Event_Capex": m.capex_shock,
                     "Event_FX": m.fx_shock, "Event_Demand": m.demand_shock}[sheet]
            conf = covenant_at(m, sheet, shock)
            mir = covenant_at(m, sheet, -shock)
            ws = wb[sheet]
            assert ws["I61"].value == pytest.approx(conf["nd"], rel=1e-6), (name, sheet, "ND")
            assert ws["J61"].value == pytest.approx(conf["ebitda"], rel=1e-6), (name, sheet, "EBITDA")
            if conf["ratio"] is None:
                assert ws["K61"].value == "n/m"
            else:
                assert ws["K61"].value == pytest.approx(conf["ratio"], rel=1e-6)
            assert ws["M61"].value == conf["flag"], (name, sheet, ws["M61"].value, conf)
            assert ws["M62"].value == mir["flag"], (name, sheet, "mirror flag")
            # Diagnostic: covenant cells must not be read by FULL
        # apparel: the 0.35x limit should be tight enough that demand breaches and the base-like rates case passes
        if name == "apparel_retailer":
            assert wb["Inputs"]["B52"].value == pytest.approx(0.35)
            assert wb["Event_Demand"]["M61"].value == "BREACH"
            assert wb["Event_Rates"]["M61"].value == "PASS"
            assert "BREACH" in [wb["Summary"][f"M{5 + k}"].value for k in range(len(sheets))]
        else:
            assert wb["Event_Rates"]["M61"].value == "n/a"
        wb.close()


@needs_lo
def test_checker_catches_broken_covenant_and_fx_full(recalculated, tmp_path):
    path = recalculated["apparel_retailer"]
    # pasted wrong covenant flag
    wb = openpyxl.load_workbook(path)
    wb["Event_Rates"]["M61"] = "zzz"
    broken = tmp_path / "cov.xlsx"
    wb.save(broken)
    wb.close()
    broken = _recalc(broken, tmp_path / "cov_out")
    rc, out = _check(broken, "apparel_retailer")
    assert rc != 0 and "Event_Rates!M61" in out, out
    # FULL omits FX debt channel
    wb = openpyxl.load_workbook(path)
    wb["Event_FX"]["D11"] = "=B39"
    broken = tmp_path / "fx.xlsx"
    wb.save(broken)
    wb.close()
    broken = _recalc(broken, tmp_path / "fx_out")
    rc, out = _check(broken, "apparel_retailer")
    assert rc != 0, out

def test_reverse_stress_label_falls_back_on_unknown_sheet():
    import reverse_stress
    assert reverse_stress.event_label("Event_Unknown", None) == "Event_Unknown"
    assert reverse_stress.event_label("Event_Unknown", "") == "Event_Unknown"
    assert reverse_stress.event_label("Event_Unknown", "   ") == "Event_Unknown"
    # A title is used when present; the sheet name is the fallback, never a KeyError.
    got = reverse_stress.event_label(
        "Event_FX", "EVENT: Foreign exchange (FX) shock — ONE event (standalone / information / FULL)"
    )
    assert "Foreign exchange (FX)" in got


def test_reverse_stress_script_covers_fx_and_demand():
    """The exact-solve script must exit 0 on the apparel example and name both new events."""
    res = _run("scripts/reverse_stress.py", CASES["apparel_retailer"])
    out = res.stdout + res.stderr
    assert res.returncode == 0, out
    assert "Event_FX" in out
    assert "Event_Demand" in out
    assert "Draft, not investment advice" in out

