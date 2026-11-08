#!/usr/bin/env python3
"""YAML config-driven generator for the same-event multi-channel stress-test workbook.

Regeneration writes formulas only; Excel/LibreOffice caches stay empty until you recalculate.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

import yaml
from jsonschema import Draft202012Validator
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
SCHEMA_PATH = ROOT / "schemas" / "config.schema.json"

BLUE = Font(name="Calibri", size=11, color="0000FF")
BLACK = Font(name="Calibri", size=11, color="000000")
BOLD = Font(name="Calibri", size=11, color="000000", bold=True)
TITLE = Font(name="Calibri", size=14, color="000000", bold=True)
HEADER = Font(name="Calibri", size=11, color="000000", bold=True)
MUTED = Font(name="Calibri", size=10, italic=True, color="666666")
YELLOW = PatternFill("solid", fgColor="FFFF99")
LIGHT = PatternFill("solid", fgColor="F2F2F2")
INFO_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def style_input(cell) -> None:
    cell.font = BLUE
    cell.border = THIN


def style_formula(cell) -> None:
    cell.font = BLACK
    cell.border = THIN


def style_judgment(cell) -> None:
    """Discretionary assumption: yellow fill + blue font."""
    cell.font = BLUE
    cell.fill = YELLOW
    cell.border = THIN


def set_widths(ws, widths) -> None:
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def load_config(path: Path) -> dict[str, Any]:
    with path.open(encoding="utf-8") as f:
        data = yaml.safe_load(f)
    schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
    Draft202012Validator(schema).validate(data)
    if data["inputs"].get("wacc_override", None) is None:
        data["inputs"]["wacc_override"] = None
    if "passthrough" not in data:
        data["passthrough"] = {
            "units": 1000.0,
            "price": 50.0,
            "unit_cost": 20.0,
            "fixed_cost": 5000.0,
            "cost_shock_pct": 0.15,
            "pass_through_pct": 1.0,
            "elasticity": 0.0,
            "markdown_pct": 0.0,
            "mitigation_pct": 0.0,
        }
    return data


def build_workbook(cfg: dict[str, Any]) -> Workbook:
    inp = cfg["inputs"]
    meta = cfg["meta"]
    pt = cfg["passthrough"]
    wb = Workbook()

    # ---------- Cover ----------
    cover = wb.active
    cover.title = "Cover"
    cover["A1"] = meta["title"]
    cover["A1"].font = TITLE
    cover["A2"] = meta["company_label"]
    cover["A3"] = "DRAFT, NOT INVESTMENT ADVICE"
    cover["A3"].font = Font(name="Calibri", size=12, bold=True, color="C00000")
    cover["A4"] = (
        "This workbook uses invented ILLUSTRATIVE numbers only. "
        "It teaches the same-event multi-channel stress-test method. "
        "It does not value any real company and is not a recommendation to buy, sell, or hold any security."
    )
    cover["A4"].alignment = Alignment(wrap_text=True)
    cover.merge_cells("A4:F6")
    cover["A8"] = "Colour conventions"
    cover["A8"].font = HEADER
    cover["A9"] = "Blue font"
    cover["B9"] = "Inputs (editable)"
    cover["A9"].font = BLUE
    cover["A10"] = "Black font"
    cover["B10"] = "Formulas (do not overwrite)"
    cover["A11"] = "Yellow fill + blue font"
    cover["B11"] = "Judgment / discretionary assumption"
    cover["A11"].font = BLUE
    cover["A11"].fill = YELLOW
    cover["A13"] = "Terms on this sheet: DCF = discounted cash flow. WACC = weighted average cost of capital. FCFF = free cash flow to the firm. EV = enterprise value. EPS = earnings per share. FULL = all valuation channels of one event applied together."
    cover["A13"].alignment = Alignment(wrap_text=True)
    cover.merge_cells("A13:F14")
    cover["A16"] = "Sheets: Inputs, DCF, Event_Rates, Event_Commodity, Tornado, Envelope, Football, PassThrough, How_Measured."
    cover["A17"] = "Workflow: build → run REVIEW_CHECKLIST.md → written CLEAR / BLOCK verdict."
    cover["A18"] = meta.get("notes") or ""
    # Raw WACC reference outside blend (disclosure)
    cover["A20"] = "Raw WACC (CAPM) — reference outside blend"
    cover["B20"] = "=Inputs!B45"
    cover["B20"].number_format = "0.00%"
    style_formula(cover["B20"])
    cover["C20"] = "CAPM = capital asset pricing model"
    cover["A21"] = "Selected WACC used in DCF"
    cover["B21"] = "=Inputs!B46"
    cover["B21"].number_format = "0.00%"
    style_formula(cover["B21"])
    set_widths(cover, [42, 18, 55, 12, 12, 12])

    # ---------- Inputs ----------
    # Cell map (stable for formulas and check_workbook):
    # B5 rev_y0, B6 growth, B7 opex_pct, B8 tax, B9 da, B10 capex, B11 nwc,
    # B12 royalty, B13 cost_elasticity (judgment), B14 capex_price_elasticity (judgment),
    # B17 commodity index, B18 commodity shock,
    # B21 rf, B22 ERP, B23 beta, B24 pre-tax kd, B25 wd, B26 we, B27 rate_shock_bp,
    # B28 wacc_override (judgment), B31 g, B32 exit mult,
    # B35 gross debt, B36 cash, B37 floating share, B38 duration, B39 shares, B40 buyback,
    # B43 ke, B44 after-tax kd, B45 raw WACC, B46 selected WACC, B47 net debt, B48 rate Δ decimal, B49 stressed commodity index
    ws = wb.create_sheet("Inputs")
    ws["A1"] = "INPUTS — ILLUSTRATIVE (blue = edit; yellow+blue = judgment)"
    ws["A1"].font = TITLE
    ws["A2"] = "Draft, not investment advice. Units: $ million unless noted. DCF = discounted cash flow."
    ws["A2"].font = MUTED

    def section(r, title):
        ws[f"A{r}"] = title
        ws[f"A{r}"].font = HEADER
        ws[f"A{r}"].fill = LIGHT

    def put(r, label, value, fmt, note, judgment=False):
        ws[f"A{r}"] = label
        ws[f"B{r}"] = value
        if fmt:
            ws[f"B{r}"].number_format = fmt
        if judgment:
            style_judgment(ws[f"B{r}"])
        else:
            style_input(ws[f"B{r}"])
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = MUTED

    section(4, "Operating")
    put(5, "Revenue year-0 ($m)", inp["revenue_y0"], "#,##0.0", "ILLUSTRATIVE base revenue")
    put(6, "Revenue growth Y1–Y5", inp["revenue_growth"], "0.0%", "Constant growth for template simplicity")
    put(7, "Opex % of revenue (ex-royalty)", inp["opex_pct_of_revenue"], "0.0%", "Operating expenses; held fixed when price moves unless cost elasticity applies")
    put(8, "Tax rate", inp["tax_rate"], "0.0%", "")
    put(9, "D&A % of revenue", inp["da_pct_of_revenue"], "0.0%", "D&A = depreciation and amortisation")
    put(10, "Capex % of revenue", inp["capex_pct_of_revenue"], "0.0%", "Capex = capital expenditure")
    put(11, "ΔNWC % of Δrevenue", inp["nwc_pct_of_delta_revenue"], "0.0%", "NWC = net working capital")
    put(12, "Royalty rate (% of revenue)", inp["royalty_rate"], "0.0%", "Historical calibration — do not overwrite with a shock")
    put(13, "Cost elasticity (opex %Δ / price %Δ)", inp["cost_elasticity"], "0.00", "Judgment: costs move only via this elasticity", True)
    put(14, "Capex–price elasticity", inp["capex_price_elasticity"], "0.00", "Judgment: capex % adjusts with price shock", True)

    section(16, "Commodity")
    put(17, "Commodity price index (base)", inp["commodity_price_index"], "0.0", "ILLUSTRATIVE index, not a vendor print")
    put(18, "Commodity shock (e.g. -0.20)", inp["commodity_shock"], "0.0%", "Event_Commodity uses this; one event only")

    section(20, "Capital structure / WACC")
    put(21, "Risk-free rate", inp["risk_free_rate"], "0.00%", "WACC = weighted average cost of capital")
    put(22, "Equity risk premium", inp["equity_risk_premium"], "0.00%", "")
    put(23, "Beta", inp["beta"], "0.00", "")
    put(24, "Pre-tax cost of debt", inp["pre_tax_cost_of_debt"], "0.00%", "")
    put(25, "Debt weight", inp["debt_weight"], "0.0%", "")
    put(26, "Equity weight", inp["equity_weight"], "0.0%", "Should equal 1 − debt weight")
    put(27, "Rate shock (bp)", inp["rate_shock_bp"], "0", "bp = basis points; ± for Event_Rates")
    override = inp.get("wacc_override")
    put(28, "Judgment WACC override (blank = formula)", override, "0.00%", "Yellow = judgment; blank uses raw CAPM WACC", True)

    section(30, "Terminal")
    put(31, "Terminal growth", inp["terminal_growth"], "0.0%", "Perpetual-growth method")
    put(32, "Exit EBITDA multiple", inp["exit_ebitda_multiple"], "0.0x", "EBITDA = earnings before interest, tax, depreciation and amortisation")

    section(34, "Balance sheet / shares")
    put(35, "Gross debt ($m)", inp["gross_debt"], "#,##0.0", "")
    put(36, "Cash ($m)", inp["cash"], "#,##0.0", "")
    put(37, "Floating / near-maturity share of gross debt", inp["floating_debt_share"], "0.0%", "INFORMATION channel for rates")
    put(38, "Fixed-debt duration (years, for MTM)", inp["fixed_debt_duration_years"], "0.0", "MTM = mark-to-market")
    put(39, "Diluted shares (million)", inp["diluted_shares_million"], "#,##0.0", "Always diluted")
    put(40, "Base buyback cash / year ($m)", inp["base_buyback_cash"], "#,##0.0", "INFORMATION channel for rates")

    section(42, "Derived (formulas — black)")
    ws["A43"] = "Cost of equity (CAPM)"
    ws["B43"] = "=B21+B23*B22"
    ws["B43"].number_format = "0.00%"
    style_formula(ws["B43"])
    ws["C43"] = "CAPM = capital asset pricing model; ke = rf + beta × ERP"

    ws["A44"] = "After-tax cost of debt"
    ws["B44"] = "=B24*(1-B8)"
    ws["B44"].number_format = "0.00%"
    style_formula(ws["B44"])

    ws["A45"] = "Raw WACC (CAPM)"
    ws["B45"] = "=B26*B43+B25*B44"
    ws["B45"].number_format = "0.00%"
    style_formula(ws["B45"])

    ws["A46"] = "Selected WACC used in DCF"
    ws["B46"] = '=IF(B28="",B45,B28)'
    ws["B46"].number_format = "0.00%"
    style_formula(ws["B46"])
    ws["C46"] = "Override in B28 if filled; Event_Rates stresses this selected WACC + Δ"

    ws["A47"] = "Net debt (gross debt − cash)"
    ws["B47"] = "=B35-B36"
    ws["B47"].number_format = "#,##0.0"
    style_formula(ws["B47"])
    ws["C47"] = "Recompute EV / net debt from raw inputs — do not trust a vendor EV field. EV = enterprise value."

    ws["A48"] = "Rate shock as decimal"
    ws["B48"] = "=B27/10000"
    ws["B48"].number_format = "0.00%"
    style_formula(ws["B48"])

    ws["A49"] = "Stressed commodity index"
    ws["B49"] = "=B17*(1+B18)"
    ws["B49"].number_format = "0.0"
    style_formula(ws["B49"])

    ws["A50"] = "Implied base EBIT margin (1 − opex% − royalty%)"
    ws["B50"] = "=1-B7-B12"
    ws["B50"].number_format = "0.0%"
    style_formula(ws["B50"])
    ws["C50"] = "EBIT = earnings before interest and tax; for display only"

    set_widths(ws, [52, 14, 72])

    # ---------- DCF ----------
    dcf = wb.create_sheet("DCF")
    dcf["A1"] = "DCF — ILLUSTRATIVE 5-year FCFF (operating-only)"
    dcf["A1"].font = TITLE
    dcf["A2"] = (
        "Draft, not investment advice. FCFF = free cash flow to the firm (operating-only). "
        "Financing enters the net-debt bridge once, not FCFF. Revenue and opex are separate lines."
    )
    dcf["A2"].font = MUTED

    dcf["A4"] = "Item"
    for i, yr in enumerate([1, 2, 3, 4, 5], start=2):
        dcf.cell(4, i, f"Y{yr}")
        dcf.cell(4, i).font = HEADER
        dcf.cell(4, i).fill = LIGHT

    dcf["A5"] = "Revenue"
    dcf["B5"] = "=Inputs!$B$5*(1+Inputs!$B$6)"
    for col in range(3, 7):
        dcf.cell(5, col, f"={get_column_letter(col-1)}5*(1+Inputs!$B$6)")
    for col in range(2, 7):
        dcf.cell(5, col).number_format = "#,##0.0"
        style_formula(dcf.cell(5, col))

    dcf["A6"] = "Opex (ex-royalty) — base path"
    for col in range(2, 7):
        cl = get_column_letter(col)
        dcf.cell(6, col, f"={cl}5*Inputs!$B$7")
        dcf.cell(6, col).number_format = "#,##0.0"
        style_formula(dcf.cell(6, col))

    dcf["A7"] = "Royalty (base calibration rate × revenue)"
    for col in range(2, 7):
        cl = get_column_letter(col)
        dcf.cell(7, col, f"={cl}5*Inputs!$B$12")
        dcf.cell(7, col).number_format = "#,##0.0"
        style_formula(dcf.cell(7, col))

    dcf["A8"] = "EBIT"
    for col in range(2, 7):
        cl = get_column_letter(col)
        dcf.cell(8, col, f"={cl}5-{cl}6-{cl}7")
        dcf.cell(8, col).number_format = "#,##0.0"
        style_formula(dcf.cell(8, col))

    dcf["A9"] = "Tax on EBIT"
    for col in range(2, 7):
        cl = get_column_letter(col)
        dcf.cell(9, col, f"={cl}8*Inputs!$B$8")
        dcf.cell(9, col).number_format = "#,##0.0"
        style_formula(dcf.cell(9, col))

    dcf["A10"] = "NOPAT"
    for col in range(2, 7):
        cl = get_column_letter(col)
        dcf.cell(10, col, f"={cl}8-{cl}9")
        dcf.cell(10, col).number_format = "#,##0.0"
        style_formula(dcf.cell(10, col))
    dcf["A10"].font = BOLD

    dcf["A11"] = "D&A"
    for col in range(2, 7):
        cl = get_column_letter(col)
        dcf.cell(11, col, f"={cl}5*Inputs!$B$9")
        dcf.cell(11, col).number_format = "#,##0.0"
        style_formula(dcf.cell(11, col))

    dcf["A12"] = "Capex"
    for col in range(2, 7):
        cl = get_column_letter(col)
        dcf.cell(12, col, f"=-{cl}5*Inputs!$B$10")
        dcf.cell(12, col).number_format = "#,##0.0"
        style_formula(dcf.cell(12, col))

    dcf["A13"] = "ΔNWC (use of cash)"
    dcf["B13"] = "=-(B5-Inputs!$B$5)*Inputs!$B$11"
    for col in range(3, 7):
        cl = get_column_letter(col)
        prev = get_column_letter(col - 1)
        dcf.cell(13, col, f"=-({cl}5-{prev}5)*Inputs!$B$11")
    for col in range(2, 7):
        dcf.cell(13, col).number_format = "#,##0.0"
        style_formula(dcf.cell(13, col))

    dcf["A14"] = "FCFF (operating-only)"
    for col in range(2, 7):
        cl = get_column_letter(col)
        dcf.cell(14, col, f"={cl}10+{cl}11+{cl}12+{cl}13")
        dcf.cell(14, col).number_format = "#,##0.0"
        style_formula(dcf.cell(14, col))
        dcf.cell(14, col).font = BOLD

    dcf["A15"] = "Discount factor"
    for col in range(2, 7):
        yr = col - 1
        dcf.cell(15, col, f"=1/(1+Inputs!$B$46)^{yr}")
        dcf.cell(15, col).number_format = "0.000"
        style_formula(dcf.cell(15, col))

    dcf["A16"] = "PV of FCFF"
    for col in range(2, 7):
        cl = get_column_letter(col)
        dcf.cell(16, col, f"={cl}14*{cl}15")
        dcf.cell(16, col).number_format = "#,##0.0"
        style_formula(dcf.cell(16, col))

    dcf["A18"] = "Valuation summary"
    dcf["A18"].font = HEADER
    dcf["A18"].fill = LIGHT

    rows_sum = [
        (19, "Sum PV explicit FCFF", "=SUM(B16:F16)"),
        (20, "Y5 EBITDA (EBIT + D&A)", "=F8+F11"),
        (21, "Terminal value (perpetual growth)", "=F14*(1+Inputs!$B$31)/(Inputs!$B$46-Inputs!$B$31)"),
        (22, "PV terminal (perpetual growth)", "=B21*F15"),
        (23, "EV — perpetual growth method", "=B19+B22"),
        (24, "Terminal value (exit multiple)", "=B20*Inputs!$B$32"),
        (25, "PV terminal (exit multiple)", "=B24*F15"),
        (26, "EV — exit multiple method", "=B19+B25"),
        (27, "Net debt (bridge)", "=Inputs!B47"),
        (28, "Equity value — perpetual growth", "=B23-B27"),
        (29, "Equity value — exit multiple", "=B26-B27"),
        (30, "Blended midpoint equity (methods only)", "=(B28+B29)/2"),
        (31, "Equity $/sh — perpetual (diluted)", "=B28/Inputs!B39"),
        (32, "Equity $/sh — exit (diluted)", "=B29/Inputs!B39"),
        (33, "Blended mid $/sh (diluted)", "=B30/Inputs!B39"),
    ]
    # Wait - I renumbered. Let me fix the summary to use correct FCFF row 14 and consistent cells.
    # Actually I have B28 as equity perp in the list above but DCF equity perp should be B28... 
    # Looking at my list: B28 = equity perp, B29 = equity exit, B30 = blend. Good.
    # But earlier event sheets referenced DCF!B27 for equity perp and DCF!B29 for blend.
    # I need STABLE cell refs matching check configs: DCF!B29 = blend mid equity $m, DCF!B27 was old equity perp.
    #
    # For backward compatibility with checks and tornado, keep:
    # B27 net debt, B28 eq perp? Old was: B26 net debt, B27 eq perp, B28 eq exit, B29 blend, B32 blend/sh
    #
    # Let me realign to preserve: B27=eq_perp, B28=eq_exit, B29=blend, B30-B32 per share variants.
    # So insert net debt at B26, and shift.

    # Clear and rewrite summary cleanly
    for r in range(18, 40):
        for c in range(1, 4):
            dcf.cell(r, c).value = None

    dcf["A18"] = "Valuation summary"
    dcf["A18"].font = HEADER
    dcf["A18"].fill = LIGHT

    dcf["A19"] = "Sum PV explicit FCFF"
    dcf["B19"] = "=SUM(B16:F16)"
    dcf["A20"] = "Y5 EBITDA (EBIT + D&A)"
    dcf["B20"] = "=F8+F11"
    dcf["A21"] = "Terminal value (perpetual growth)"
    dcf["B21"] = "=F14*(1+Inputs!$B$31)/(Inputs!$B$46-Inputs!$B$31)"
    dcf["A22"] = "PV terminal (perpetual growth)"
    dcf["B22"] = "=B21*F15"
    dcf["A23"] = "EV — perpetual growth method"
    dcf["B23"] = "=B19+B22"
    dcf["A24"] = "Terminal value (exit multiple)"
    dcf["B24"] = "=B20*Inputs!$B$32"
    dcf["A25"] = "PV terminal (exit multiple)"
    dcf["B25"] = "=B24*F15"
    dcf["A26"] = "EV — exit multiple method"
    dcf["B26"] = "=B19+B25"
    dcf["A27"] = "Net debt (bridge)"
    dcf["B27"] = "=Inputs!B47"
    dcf["A28"] = "Equity value — perpetual growth"
    dcf["B28"] = "=B23-B27"
    dcf["A29"] = "Equity value — exit multiple"
    dcf["B29"] = "=B26-B27"
    dcf["A30"] = "Blended midpoint equity (methods only)"
    dcf["B30"] = "=(B28+B29)/2"
    dcf["C30"] = "Blend lock: stress bands stay outside this average"
    dcf["A31"] = "Equity $/sh — perpetual (diluted)"
    dcf["B31"] = "=B28/Inputs!B39"
    dcf["A32"] = "Equity $/sh — exit (diluted)"
    dcf["B32"] = "=B29/Inputs!B39"
    dcf["A33"] = "Blended mid $/sh (diluted)"
    dcf["B33"] = "=B30/Inputs!B39"

    # Alias row labels for check scripts expecting B29 as blend — UPDATE checks to B30.
    # I'll use B28 eq_perp, B30 blend $m, B33 blend/sh and update YAML checks accordingly.

    for r in range(19, 34):
        style_formula(dcf[f"B{r}"])
        if r in (31, 32, 33):
            dcf[f"B{r}"].number_format = "0.00"
        else:
            dcf[f"B{r}"].number_format = "#,##0.0"
    dcf["B23"].font = BOLD
    dcf["B26"].font = BOLD
    dcf["B30"].font = BOLD
    dcf["B33"].font = BOLD

    dcf["A35"] = (
        "How_Measured: FCFF excludes interest. Revenue and opex are separate. "
        "Debt mark-to-market and rate information rows sit on Event_Rates."
    )
    dcf.merge_cells("A35:F36")
    dcf["A35"].alignment = Alignment(wrap_text=True)
    set_widths(dcf, [46, 12, 12, 12, 12, 12, 55])

    # Stable named refs used across sheets:
    # DCF!B28 = equity perp base
    # DCF!B30 = blended mid $m
    # DCF!B33 = blended mid $/sh

    events = cfg.get("events") or {}
    use_tariff = bool((events.get("tariff") or {}).get("enabled"))
    use_commodity = bool((events.get("commodity") or {}).get("enabled", not use_tariff))
    if use_tariff and use_commodity:
        # Apparel-style: prefer tariff as the second tornado event; still allow commodity disabled in config.
        use_commodity = False
    if not use_tariff and not use_commodity:
        use_commodity = True

    _build_event_rates(wb)
    second_sheet = "Event_Commodity"
    second_label = "Commodity FULL"
    if use_commodity:
        _build_event_commodity(wb)
        second_sheet = "Event_Commodity"
        second_label = (events.get("commodity") or {}).get("label") or "Commodity FULL"
    if use_tariff:
        tariff_cfg = cfg.get("tariff") or pt
        _build_event_tariff(wb, tariff_cfg)
        second_sheet = "Event_Tariff"
        second_label = (events.get("tariff") or {}).get("label") or "Tariff / input-cost FULL"
    _build_tornado(wb, second_sheet=second_sheet, second_label=second_label)
    _build_envelope(wb)
    _build_football(wb, second_sheet=second_sheet)
    _build_passthrough(wb, pt)
    _build_how_measured(wb, second_sheet=second_sheet)
    # Cover sheet list
    cover = wb["Cover"]
    cover["A16"] = (
        f"Sheets: Inputs, DCF, Event_Rates, {second_sheet}, Tornado, Envelope, Football, PassThrough, How_Measured."
    )
    return wb


def _equity_from_fcff_row(fcff_row: int, wacc_ref: str, g_ref: str = "Inputs!$B$31", nd_ref: str = "Inputs!$B$47") -> str:
    """PV of 5y FCFF row + perpetual TV on Y5, minus net debt. wacc_ref e.g. Inputs!$B$46 or $B$25."""
    parts = [f"{get_column_letter(c)}{fcff_row}/(1+{wacc_ref})^{c-1}" for c in range(2, 7)]
    y5 = f"F{fcff_row}"
    tv = f"({y5}*(1+{g_ref})/({wacc_ref}-{g_ref}))/(1+{wacc_ref})^5"
    return "=" + "+".join(parts) + f"+{tv}-{nd_ref}"


def _build_event_rates(wb: Workbook) -> None:
    er = wb.create_sheet("Event_Rates")
    er["A1"] = "EVENT: Interest rates shock — ONE event (standalone / information / FULL)"
    er["A1"].font = TITLE
    er["A2"] = (
        "Shock size (bp) = Inputs!B27. FCFF stays operating-only. Financing once: FULL = WACC channel + fixed-debt MTM in the bridge. "
        "Interest and buyback capacity are INFORMATION rows outside FULL. Draft, not investment advice."
    )
    er["A2"].font = MUTED
    er["A2"].alignment = Alignment(wrap_text=True)
    er.merge_cells("A2:G3")

    headers = [
        "Row",
        "Kind",
        "Channel",
        "Stressed equity ($m, perp. method)",
        "%Δ vs own method base",
        "Blended-mid sensitivity proxy ($m)",
        "Proxy $/sh diluted",
    ]
    for c, h in enumerate(headers, 1):
        er.cell(5, c, h)
        er.cell(5, c).font = HEADER
        er.cell(5, c).fill = LIGHT
        er.cell(5, c).border = THIN

    # Helpers
    er["A22"] = "Helper calculations (live formulas)"
    er["A22"].font = HEADER
    er["A23"] = "Selected WACC (base)"
    er["B23"] = "=Inputs!B46"
    er["B23"].number_format = "0.00%"
    style_formula(er["B23"])
    er["A24"] = "Rate Δ (decimal)"
    er["B24"] = "=Inputs!B48"
    er["B24"].number_format = "0.00%"
    style_formula(er["B24"])
    er["A25"] = "Stressed WACC = selected WACC + Δ"
    er["B25"] = "=B23+B24"
    er["B25"].number_format = "0.00%"
    style_formula(er["B25"])
    er["C25"] = "Respects judgment override on Inputs!B28"

    er["A26"] = "Sum PV FCFF @ stressed WACC"
    er["B26"] = (
        "=DCF!B14/(1+B25)^1+DCF!C14/(1+B25)^2+DCF!D14/(1+B25)^3"
        "+DCF!E14/(1+B25)^4+DCF!F14/(1+B25)^5"
    )
    er["B26"].number_format = "#,##0.0"
    style_formula(er["B26"])
    er["A27"] = "TV perp @ stressed WACC"
    er["B27"] = "=DCF!F14*(1+Inputs!B31)/(B25-Inputs!B31)"
    er["B27"].number_format = "#,##0.0"
    style_formula(er["B27"])
    er["A28"] = "PV TV @ stressed WACC"
    er["B28"] = "=B27/(1+B25)^5"
    er["B28"].number_format = "#,##0.0"
    style_formula(er["B28"])
    er["A29"] = "EV @ stressed WACC"
    er["B29"] = "=B26+B28"
    er["B29"].number_format = "#,##0.0"
    style_formula(er["B29"])
    er["A30"] = "Equity @ stressed WACC (channel A) — base net debt"
    er["B30"] = "=B29-Inputs!B47"
    er["B30"].number_format = "#,##0.0"
    style_formula(er["B30"])

    er["A31"] = "Floating debt ($m)"
    er["B31"] = "=Inputs!B35*Inputs!B37"
    er["B31"].number_format = "#,##0.0"
    style_formula(er["B31"])
    er["A32"] = "Extra annual pre-tax interest"
    er["B32"] = "=B31*Inputs!B48"
    er["B32"].number_format = "#,##0.0"
    style_formula(er["B32"])
    er["A33"] = "PV after-tax interest drain (INFORMATION — not in FULL)"
    er["B33"] = "=B32*(1-Inputs!B8)/Inputs!B46"
    er["B33"].number_format = "#,##0.0"
    style_formula(er["B33"])
    er["C33"] = "Same cash interest must not also reduce buyback in FULL"

    er["A34"] = "Fixed debt ($m)"
    er["B34"] = "=Inputs!B35*(1-Inputs!B37)"
    er["B34"].number_format = "#,##0.0"
    style_formula(er["B34"])
    er["A35"] = "MTM decrease in debt FV (channel C → bridge)"
    er["B35"] = "=Inputs!B38*B34*Inputs!B48"
    er["B35"].number_format = "#,##0.0"
    style_formula(er["B35"])
    er["C35"] = "Rates up → fixed debt FV down → economic net debt down"

    er["A36"] = "EPS hit from extra interest ($/sh) — INFORMATION"
    er["B36"] = "=-B32*(1-Inputs!B8)/Inputs!B39"
    er["B36"].number_format = "0.000"
    style_formula(er["B36"])

    er["A37"] = "Lost buyback cash — INFORMATION (not in FULL)"
    er["B37"] = "=MIN(Inputs!B40,B32*(1-Inputs!B8))"
    er["B37"].number_format = "#,##0.0"
    style_formula(er["B37"])

    # Channel table rows 6-11
    # A standalone WACC, B information interest, C standalone MTM, D information EPS, E information buyback, FULL = A+C
    channels = [
        (6, "A", "standalone", "WACC via selected WACC + Δ",
         "=B30",
         "Re-discount base FCFF at stressed WACC = selected WACC + rate Δ"),
        (7, "B", "information", "Refinancing interest on floating debt",
         "=DCF!B28-B33",
         "INFORMATION: PV of after-tax interest — outside FULL (financing shown once via WACC/MTM)"),
        (8, "C", "standalone", "Fixed-debt mark-to-market in net-debt bridge",
         "=DCF!B28+B35",
         "Bridge only: MTM debt relief; not inside FCFF"),
        (9, "D", "information", "Profit / EPS view",
         "=DCF!B28",
         "INFORMATION: EPS communication only; equity equals base"),
        (10, "E", "information", "Buyback capacity",
         "=DCF!B28-B37",
         "INFORMATION: outside FULL so interest cash is not subtracted twice"),
        (11, "FULL", "cumulative", "A + C (WACC + MTM bridge)",
         "=B30+B35",
         "Cumulative valuation channels only; B/D/E excluded to avoid double count"),
    ]
    # FULL = equity at stressed WACC with MTM-adjusted net debt = (EV_stressed - (net_debt - MTM)) = B30 + B35. Yes.

    for r, rowid, kind, channel, d_fx, note in channels:
        er.cell(r, 1, rowid)
        er.cell(r, 2, kind)
        er.cell(r, 3, channel)
        er.cell(r, 4, d_fx)
        er.cell(r, 5, f"=(D{r}-DCF!$B$28)/DCF!$B$28")
        er.cell(r, 6, f"=DCF!$B$30*(1+E{r})")
        er.cell(r, 7, f"=F{r}/Inputs!$B$39")
        for c in range(1, 8):
            er.cell(r, c).border = THIN
        for c in range(4, 8):
            style_formula(er.cell(r, c))
        er.cell(r, 4).number_format = "#,##0.0"
        er.cell(r, 5).number_format = "0.0%"
        er.cell(r, 6).number_format = "#,##0.0"
        er.cell(r, 7).number_format = "0.00"
        if kind == "information":
            er.cell(r, 2).fill = INFO_FILL
        if rowid == "FULL":
            for c in range(1, 8):
                er.cell(r, c).font = BOLD
        er.cell(r, 8, note)
        er.cell(r, 8).font = MUTED

    # FULL is on row 11 now (was 10). Update tornado to Event_Rates!E11 etc.
    # Store note on row 13
    er["A13"] = "Own method base for %Δ (perpetual-growth equity)"
    er["B13"] = "=DCF!B28"
    er["B13"].number_format = "#,##0.0"
    style_formula(er["B13"])
    er["A14"] = "Blended mid (reference)"
    er["B14"] = "=DCF!B30"
    er["B14"].number_format = "#,##0.0"
    style_formula(er["B14"])
    er["A16"] = "FULL row is sheet row 11. Information rows are excluded from FULL by design."
    er["A17"] = "Never stack a second event into this table — use Event_Commodity."
    set_widths(er, [8, 14, 48, 18, 14, 18, 14, 70])


def _build_event_commodity(wb: Workbook) -> None:
    ec = wb.create_sheet("Event_Commodity")
    ec["A1"] = "EVENT: Commodity price shock — ONE event (standalone / cumulative / diagnostic)"
    ec["A1"].font = TITLE
    ec["A2"] = (
        "Price moves revenue only. Opex moves only via Inputs!B13 cost elasticity (yellow judgment). "
        "Royalty uses the base calibration rate on Inputs!B12 (never re-fitted). Draft, not investment advice."
    )
    ec["A2"].font = MUTED
    ec["A2"].alignment = Alignment(wrap_text=True)
    ec.merge_cells("A2:G3")

    headers = [
        "Row",
        "Kind",
        "Channel",
        "Stressed equity ($m, perp. method)",
        "%Δ vs own method base",
        "Blended-mid sensitivity proxy ($m)",
        "Proxy $/sh diluted",
    ]
    for c, h in enumerate(headers, 1):
        ec.cell(5, c, h)
        ec.cell(5, c).font = HEADER
        ec.cell(5, c).fill = LIGHT
        ec.cell(5, c).border = THIN

    # Helpers / grids from row 22
    ec["A22"] = "Helper grids (live formulas)"
    ec["A22"].font = HEADER
    ec["A23"] = "Price factor (1 + shock)"
    ec["B23"] = "=1+Inputs!B18"
    ec["B23"].number_format = "0.00"
    style_formula(ec["B23"])

    ec["A24"] = "Item"
    for i, yr in enumerate([1, 2, 3, 4, 5], start=2):
        ec.cell(24, i, f"Y{yr}")
        ec.cell(24, i).font = HEADER

    # Base opex path reference (from DCF)
    # Stressed revenue
    ec["A25"] = "Stressed revenue (price × base volume path)"
    for col in range(2, 7):
        cl = get_column_letter(col)
        ec.cell(25, col, f"=DCF!{cl}5*$B$23")
        ec.cell(25, col).number_format = "#,##0.0"
        style_formula(ec.cell(25, col))

    # Opex fixed at base (for channel A)
    ec["A26"] = "Opex fixed at base (channel A)"
    for col in range(2, 7):
        cl = get_column_letter(col)
        ec.cell(26, col, f"=DCF!{cl}6")
        ec.cell(26, col).number_format = "#,##0.0"
        style_formula(ec.cell(26, col))

    # Opex with elasticity (channel C / FULL)
    # opex' = opex_base * (1 + cost_elasticity * commodity_shock)
    ec["A27"] = "Opex with cost elasticity (C / FULL)"
    for col in range(2, 7):
        cl = get_column_letter(col)
        ec.cell(27, col, f"=DCF!{cl}6*(1+Inputs!$B$13*Inputs!$B$18)")
        ec.cell(27, col).number_format = "#,##0.0"
        style_formula(ec.cell(27, col))

    # Royalty at base dollars (channel A isolates revenue)
    ec["A28"] = "Royalty at base dollars (A isolate)"
    for col in range(2, 7):
        cl = get_column_letter(col)
        ec.cell(28, col, f"=DCF!{cl}7")
        ec.cell(28, col).number_format = "#,##0.0"
        style_formula(ec.cell(28, col))

    # Royalty on stressed revenue at base rate (B / FULL)
    ec["A29"] = "Royalty on stressed revenue @ base rate"
    for col in range(2, 7):
        cl = get_column_letter(col)
        ec.cell(29, col, f"={cl}25*Inputs!$B$12")
        ec.cell(29, col).number_format = "#,##0.0"
        style_formula(ec.cell(29, col))

    ec["A30"] = "Capex % adjusted"
    ec["B30"] = "=Inputs!B10*(1+Inputs!B14*Inputs!B18)"
    ec["B30"].number_format = "0.00%"
    style_formula(ec["B30"])

    def write_fcff(row, rev_row, opex_row, roy_row, use_adj_capex: bool, da_capex_row=None, nwc_row=None):
        """Build an FCFF path. da_capex_row defaults to rev_row; for price-only channels use base volume row."""
        if da_capex_row is None:
            da_capex_row = rev_row
        if nwc_row is None:
            nwc_row = rev_row
        ec.cell(row, 1, f"FCFF row {row}")
        for col in range(2, 7):
            cl = get_column_letter(col)
            ebit = f"({cl}{rev_row}-{cl}{opex_row}-{cl}{roy_row})"
            nopat = f"{ebit}*(1-Inputs!$B$8)"
            da = f"{cl}{da_capex_row}*Inputs!$B$9"
            if use_adj_capex:
                capex = f"-{cl}{da_capex_row}*$B$30"
            else:
                capex = f"-{cl}{da_capex_row}*Inputs!$B$10"
            if col == 2:
                dnwc = f"-({cl}{nwc_row}-Inputs!$B$5)*Inputs!$B$11"
            else:
                prev = get_column_letter(col - 1)
                dnwc = f"-({cl}{nwc_row}-{prev}{nwc_row})*Inputs!$B$11"
            ec.cell(row, col, f"={nopat}+{da}+({capex})+({dnwc})")
            ec.cell(row, col).number_format = "#,##0.0"
            style_formula(ec.cell(row, col))

    # Base refs on this sheet
    ec["A31"] = "Base revenue (ref)"
    for col in range(2, 7):
        cl = get_column_letter(col)
        ec.cell(31, col, f"=DCF!{cl}5")
        ec.cell(31, col).number_format = "#,##0.0"
        style_formula(ec.cell(31, col))
    ec["A31"] = "Base revenue (ref)"

    ec["A33"] = "Royalty dollars for B-only (base rev × rate × price factor)"
    for col in range(2, 7):
        cl = get_column_letter(col)
        ec.cell(33, col, f"={cl}31*Inputs!$B$12*$B$23")
        ec.cell(33, col).number_format = "#,##0.0"
        style_formula(ec.cell(33, col))

    # Reuse row 26 as base opex already; add alias label
    ec["A34"] = "Base opex alias note: row 26 is base opex; row 27 is elastic opex"

    # FCFF A: stressed rev, base opex, base royalty $, DA/capex on base volume (price≠volume)
    write_fcff(32, 25, 26, 28, False, da_capex_row=31, nwc_row=25)
    ec["A32"] = "FCFF A (price→revenue only; opex fixed; royalty base $; DA/capex on base volume)"

    # FCFF B: base rev, base opex, stressed royalty dollars
    write_fcff(37, 31, 26, 33, False, da_capex_row=31, nwc_row=31)
    ec["A37"] = "FCFF B — royalty only"

    # Base royalty ref
    ec["A38"] = "Base royalty (ref)"
    for col in range(2, 7):
        cl = get_column_letter(col)
        ec.cell(38, col, f"=DCF!{cl}7")
        ec.cell(38, col).number_format = "#,##0.0"
        style_formula(ec.cell(38, col))

    # FCFF C: base rev, elastic opex, base royalty
    write_fcff(39, 31, 27, 38, False, da_capex_row=31, nwc_row=31)
    ec["A39"] = "FCFF C — cost elasticity only"

    # FCFF D: base rev/opex/royalty, adjusted capex %
    write_fcff(40, 31, 26, 38, True, da_capex_row=31, nwc_row=31)
    ec["A40"] = "FCFF D — capex pace only"

    # FCFF FULL: stressed rev, elastic opex, royalty on stressed rev, adj capex on base volume
    write_fcff(41, 25, 27, 29, True, da_capex_row=31, nwc_row=25)
    ec["A41"] = "FCFF FULL — A+B+C+D cumulative"

    # Equity valuations
    def eq_cell(out_row, fcff_row, label):
        ec[f"A{out_row}"] = label
        ec[f"B{out_row}"] = (
            f"=B{fcff_row}/(1+Inputs!$B$46)^1+C{fcff_row}/(1+Inputs!$B$46)^2"
            f"+D{fcff_row}/(1+Inputs!$B$46)^3+E{fcff_row}/(1+Inputs!$B$46)^4"
            f"+F{fcff_row}/(1+Inputs!$B$46)^5"
            f"+(F{fcff_row}*(1+Inputs!$B$31)/(Inputs!$B$46-Inputs!$B$31))/(1+Inputs!$B$46)^5"
            f"-Inputs!$B$47"
        )
        ec[f"B{out_row}"].number_format = "#,##0.0"
        style_formula(ec[f"B{out_row}"])

    eq_cell(43, 32, "Equity A standalone")
    eq_cell(44, 37, "Equity B standalone")
    eq_cell(45, 39, "Equity C standalone")
    eq_cell(46, 40, "Equity D standalone")
    eq_cell(47, 41, "Equity FULL cumulative A+B+C+D")

    # Diagnostic E: bridge view of cumulative 5y FCFF Δ (not in FULL)
    ec["A48"] = "Δ cumulative 5y FCFF FULL vs base (diagnostic)"
    ec["B48"] = "=(B41+C41+D41+E41+F41)-(DCF!B14+DCF!C14+DCF!D14+DCF!E14+DCF!F14)"
    ec["B48"].number_format = "#,##0.0"
    style_formula(ec["B48"])
    ec["A49"] = "Diagnostic bridge equity (base − (−ΔFCFF))"
    ec["B49"] = "=DCF!B28-(-B48)"
    ec["B49"].number_format = "#,##0.0"
    style_formula(ec["B49"])
    ec["C49"] = "DIAGNOSTIC only — excluded from FULL (PV path already capitalises cash)"

    # Channel display rows 6-11
    channels = [
        (6, "A", "standalone", "Revenue (price × volume; volume held; opex fixed)",
         "=B43"),
        (7, "B", "standalone", "Royalties / tax at base calibration rate",
         "=B44"),
        (8, "C", "standalone", "Cost co-move via yellow cost elasticity only",
         "=B45"),
        (9, "D", "standalone", "Capex / reserve pace via yellow capex–price elasticity",
         "=B46"),
        (10, "E", "diagnostic", "Net debt / dividends bridge view (out of FULL)",
         "=B49"),
        (11, "FULL", "cumulative", "A+B+C+D together (E diagnostic excluded)",
         "=B47"),
    ]
    for r, rowid, kind, channel, d_fx in channels:
        ec.cell(r, 1, rowid)
        ec.cell(r, 2, kind)
        ec.cell(r, 3, channel)
        ec.cell(r, 4, d_fx)
        ec.cell(r, 5, f"=(D{r}-DCF!$B$28)/DCF!$B$28")
        ec.cell(r, 6, f"=DCF!$B$30*(1+E{r})")
        ec.cell(r, 7, f"=F{r}/Inputs!$B$39")
        for c in range(1, 8):
            ec.cell(r, c).border = THIN
        for c in range(4, 8):
            style_formula(ec.cell(r, c))
        ec.cell(r, 4).number_format = "#,##0.0"
        ec.cell(r, 5).number_format = "0.0%"
        ec.cell(r, 6).number_format = "#,##0.0"
        ec.cell(r, 7).number_format = "0.00"
        if kind == "diagnostic":
            ec.cell(r, 2).fill = INFO_FILL
        if rowid == "FULL":
            for c in range(1, 8):
                ec.cell(r, c).font = BOLD

    ec["A13"] = "Own method base for %Δ"
    ec["B13"] = "=DCF!B28"
    ec["B13"].number_format = "#,##0.0"
    style_formula(ec["B13"])
    ec["A14"] = "Blended mid (reference)"
    ec["B14"] = "=DCF!B30"
    ec["B14"].number_format = "#,##0.0"
    style_formula(ec["B14"])
    ec["A16"] = "FULL is sheet row 11. Diagnostic row E is excluded from FULL."
    ec["A17"] = "Pitfall guard: never overwrite Inputs!B12 historical royalty calibration with the shocked rate."
    set_widths(ec, [8, 12, 58, 18, 14, 18, 14])



def _build_event_tariff(wb: Workbook, tariff: dict) -> None:
    """Units × price × unit-cost tariff / input-cost event. Fixed cost separate. Pass-through once."""
    et = wb.create_sheet("Event_Tariff")
    et["A1"] = "EVENT: Tariff / input-cost shock — ONE event (units × price × unit cost)"
    et["A1"].font = TITLE
    et["A2"] = (
        "Fixed costs stay separate. Pass-through counted once. At 100% pass-through and zero elasticity/"
        "markdowns/mitigation, ΔGP = 0. At partial pass-through, +cost lowers equity. Draft, not investment advice."
    )
    et["A2"].font = MUTED
    et["A2"].alignment = Alignment(wrap_text=True)
    et.merge_cells("A2:G3")

    headers = [
        "Row", "Kind", "Channel",
        "Stressed equity ($m, perp. method)", "%Δ vs own method base",
        "Blended-mid sensitivity proxy ($m)", "Proxy $/sh diluted",
    ]
    for c, h in enumerate(headers, 1):
        et.cell(5, c, h)
        et.cell(5, c).font = HEADER
        et.cell(5, c).fill = LIGHT
        et.cell(5, c).border = THIN

    # Inputs (blue); judgment yellow+blue for discretionary recovery assumptions
    et["A20"] = "Tariff bridge inputs"
    et["A20"].font = HEADER
    et["A20"].fill = LIGHT

    def put(r, label, val, fmt, judgment=False, note=""):
        et[f"A{r}"] = label
        et[f"B{r}"] = val
        if fmt:
            et[f"B{r}"].number_format = fmt
        if judgment:
            style_judgment(et[f"B{r}"])
        else:
            style_input(et[f"B{r}"])
        et[f"C{r}"] = note
        et[f"C{r}"].font = MUTED

    for c, h in enumerate(headers, 1):
        et.cell(5, c, h)
        et.cell(5, c).font = HEADER
        et.cell(5, c).fill = LIGHT
        et.cell(5, c).border = THIN

    put(21, "Units (Y0)", tariff.get("units", 20.0), "#,##0.0", note="Scale so units×price ≈ revenue year-0")
    put(22, "Price / unit", tariff.get("price", 40.0), "#,##0.00")
    put(23, "Unit cost", tariff.get("unit_cost", 18.0), "#,##0.00")
    put(24, "Fixed cost / year ($m equiv scale)", tariff.get("fixed_cost", 150.0), "#,##0.0", note="Separate from unit cost")
    put(25, "Unit-cost shock %", tariff.get("cost_shock_pct", 0.15), "0.0%")
    put(26, "Pass-through % of cost shock into price", tariff.get("pass_through_pct", 0.60), "0.0%", True, "Judgment; <100% ⇒ +cost lowers equity")
    put(27, "Demand elasticity (qty %Δ / price %Δ)", tariff.get("elasticity", -0.30), "0.00", True)
    put(28, "Markdown % of sales", tariff.get("markdown_pct", 0.0), "0.0%", True)
    put(29, "Sourcing mitigation % of cost shock", tariff.get("mitigation_pct", 0.0), "0.0%", True)

    # Year grid
    et["A31"] = "Year grid (live formulas)"
    et["A31"].font = HEADER
    et["A32"] = "Item"
    for i, yr in enumerate([1, 2, 3, 4, 5], start=2):
        et.cell(32, i, f"Y{yr}")
        et.cell(32, i).font = HEADER

    # Units path
    et["A33"] = "Base units"
    for col in range(2, 7):
        yr = col - 1
        et.cell(33, col, f"=$B$21*(1+Inputs!$B$6)^{yr}")
        et.cell(33, col).number_format = "#,##0.0"
        style_formula(et.cell(33, col))

    # Helper: compute delta_nopat for a posture defined by PT/elast/md/mit cell refs
    def write_delta_row(row, label, pt_ref, elast_ref, md_ref, mit_ref):
        et.cell(row, 1, label)
        for col in range(2, 7):
            cl = get_column_letter(col)
            # base
            # units already in row 33
            # eff shock, stressed uc, price, units, gp delta
            u = f"{cl}33"
            eff = f"($B$25*(1-{mit_ref}))"
            suc = f"($B$23*(1+{eff}))"
            pup = f"($B$23*{eff}*{pt_ref})"
            sp = f"($B$22+{pup})"
            ppct = f"IF($B$22=0,0,({sp}-$B$22)/$B$22)"
            su = f"({u}*(1+{elast_ref}*{ppct}))"
            srev = f"({su}*{sp}*(1-{md_ref}))"
            svar = f"({su}*{suc})"
            sgp = f"({srev}-{svar})"
            brev = f"({u}*$B$22)"
            bvar = f"({u}*$B$23)"
            bgp = f"({brev}-{bvar})"
            dgp = f"({sgp}-{bgp})"
            et.cell(row, col, f"={dgp}*(1-Inputs!$B$8)")
            et.cell(row, col).number_format = "#,##0.00"
            style_formula(et.cell(row, col))

    # Row 34: FULL posture delta nopat (configured)
    write_delta_row(34, "ΔNOPAT FULL posture", "$B$26", "$B$27", "$B$28", "$B$29")
    # Row 35: A only — unit cost, PT=0, elast=0, md=0, mit=0
    write_delta_row(35, "ΔNOPAT A (cost only)", "0", "0", "0", "0")
    # Row 36: B — cost + configured PT, elast=0, md=0, mit=0
    write_delta_row(36, "ΔNOPAT B (cost+PT)", "$B$26", "0", "0", "0")
    # Row 37: C — cost + PT + elast
    write_delta_row(37, "ΔNOPAT C (+elasticity)", "$B$26", "$B$27", "0", "0")
    # Row 38: D — +markdown
    write_delta_row(38, "ΔNOPAT D (+markdown)", "$B$26", "$B$27", "$B$28", "0")
    # Row 39: invariant posture 100% PT zero elast/md/mit
    write_delta_row(39, "ΔNOPAT invariant (PT=100%, elast=md=mit=0)", "1", "0", "0", "0")

    def equity_from_delta_row(out_r, delta_row, label):
        et[f"A{out_r}"] = label
        # FCFF' = DCF FCFF + delta_nopat; PV at selected WACC
        et[f"B{out_r}"] = (
            f"=(DCF!B14+B{delta_row})/(1+Inputs!$B$46)^1"
            f"+(DCF!C14+C{delta_row})/(1+Inputs!$B$46)^2"
            f"+(DCF!D14+D{delta_row})/(1+Inputs!$B$46)^3"
            f"+(DCF!E14+E{delta_row})/(1+Inputs!$B$46)^4"
            f"+(DCF!F14+F{delta_row})/(1+Inputs!$B$46)^5"
            f"+((DCF!F14+F{delta_row})*(1+Inputs!$B$31)/(Inputs!$B$46-Inputs!$B$31))/(1+Inputs!$B$46)^5"
            f"-Inputs!$B$47"
        )
        et[f"B{out_r}"].number_format = "#,##0.0"
        style_formula(et[f"B{out_r}"])

    equity_from_delta_row(41, 35, "Equity A standalone (unit cost only)")
    equity_from_delta_row(42, 36, "Equity B standalone (cost + pass-through)")
    equity_from_delta_row(43, 37, "Equity C (+ elasticity)")
    equity_from_delta_row(44, 38, "Equity D (+ markdown)")
    equity_from_delta_row(45, 34, "Equity FULL cumulative")
    equity_from_delta_row(46, 39, "Equity invariant posture (diagnostic)")

    # ΔGP invariant check on Y1 in configured vs 100% PT
    et["A48"] = "Y1 ΔGP at invariant posture (should be ~0)"
    et["B48"] = "=B39/(1-Inputs!$B$8)"  # ΔNOPAT/(1-t) = ΔGP when fixed cost unchanged
    et["B48"].number_format = "#,##0.00"
    style_formula(et["B48"])
    et["A49"] = "Invariant OK? (|ΔGP Y1|<0.01)"
    et["B49"] = '=IF(ABS(B48)<0.01,"PASS","FAIL")'
    style_formula(et["B49"])
    et["C49"] = "Diagnostic; FULL uses partial pass-through from B26"

    # Channel table rows 6-11
    channels = [
        (6, "A", "standalone", "Unit cost (no pass-through)", "=B41"),
        (7, "B", "standalone", "Price pass-through (with cost)", "=B42"),
        (8, "C", "standalone", "Demand elasticity", "=B43"),
        (9, "D", "standalone", "Markdowns / promotions", "=B44"),
        (10, "E", "diagnostic", "Invariant 100% PT posture equity (out of FULL)", "=B46"),
        (11, "FULL", "cumulative", "Cost + PT + elasticity + markdown + mitigation", "=B45"),
    ]
    for r, rowid, kind, channel, d_fx in channels:
        et.cell(r, 1, rowid)
        et.cell(r, 2, kind)
        et.cell(r, 3, channel)
        et.cell(r, 4, d_fx)
        et.cell(r, 5, f"=(D{r}-DCF!$B$28)/DCF!$B$28")
        et.cell(r, 6, f"=DCF!$B$30*(1+E{r})")
        et.cell(r, 7, f"=F{r}/Inputs!$B$39")
        for c in range(1, 8):
            et.cell(r, c).border = THIN
        for c in range(4, 8):
            style_formula(et.cell(r, c))
        et.cell(r, 4).number_format = "#,##0.0"
        et.cell(r, 5).number_format = "0.0%"
        et.cell(r, 6).number_format = "#,##0.0"
        et.cell(r, 7).number_format = "0.00"
        if kind == "diagnostic":
            et.cell(r, 2).fill = INFO_FILL
        if rowid == "FULL":
            for c in range(1, 8):
                et.cell(r, c).font = BOLD

    # Alias B9 as cost shock for checker default (Event_Tariff!B9) — but B9 is channel D row.
    # Checker shock test uses Event_Tariff!B9 — CONFLICT with channel table row 9.
    # Move channel table stays at 6-11; put cost shock at B25. Update checker to use B25.
    # Also expose B9 shock via note: checker uses B25.

    et["A13"] = "Own method base for %Δ"
    et["B13"] = "=DCF!B28"
    et["B13"].number_format = "#,##0.0"
    style_formula(et["B13"])
    et["A14"] = "Blended mid (reference)"
    et["B14"] = "=DCF!B30"
    et["B14"].number_format = "#,##0.0"
    style_formula(et["B14"])
    et["A16"] = "FULL is sheet row 11. Shock input for tests: B25 (unit-cost shock %). Pass-through B26."
    et["A17"] = "Never stack rates into this table — use Event_Rates."
    # Fixed cost not in unit bridge — shown for disclosure
    et["A18"] = "Fixed cost (annual, separate — not in unit cost)"
    et["B18"] = "=B24"
    style_formula(et["B18"])
    set_widths(et, [8, 12, 52, 18, 14, 18, 14])


def _build_tornado(wb: Workbook, second_sheet: str = "Event_Commodity", second_label: str = "Commodity FULL") -> None:
    tn = wb.create_sheet("Tornado")
    tn["A1"] = "Tornado — FULL-row blended-mid sensitivity proxies"
    tn["A1"].font = TITLE
    tn["A2"] = (
        "Each bar uses that event's FULL %Δ versus its own method base, applied to the blended mid. "
        "Draft, not investment advice. Chart is formula-driven from this table."
    )
    tn["A2"].font = MUTED

    for c, h in enumerate(
        ["Shock (FULL)", "Own-base %Δ", "Proxy equity ($m)", "Proxy $/sh", "Δ vs blended mid ($m)", "Abs Δ ($m)"],
        1,
    ):
        tn.cell(4, c, h)
        tn.cell(4, c).font = HEADER
        tn.cell(4, c).fill = LIGHT
        tn.cell(4, c).border = THIN

    # FULL rows are now row 11 on both event sheets
    tn["A5"] = "Rates FULL"
    tn["B5"] = "=Event_Rates!E11"
    tn["C5"] = "=Event_Rates!F11"
    tn["D5"] = "=Event_Rates!G11"
    tn["E5"] = "=C5-DCF!B30"
    tn["F5"] = "=ABS(E5)"

    tn["A6"] = second_label
    tn["B6"] = f"={second_sheet}!E11"
    tn["C6"] = f"={second_sheet}!F11"
    tn["D6"] = f"={second_sheet}!G11"
    tn["E6"] = "=C6-DCF!B30"
    tn["F6"] = "=ABS(E6)"

    for r in range(5, 7):
        for c in range(1, 7):
            tn.cell(r, c).border = THIN
        for c in range(2, 7):
            style_formula(tn.cell(r, c))
        tn.cell(r, 2).number_format = "0.0%"
        tn.cell(r, 3).number_format = "#,##0.0"
        tn.cell(r, 4).number_format = "0.00"
        tn.cell(r, 5).number_format = "#,##0.0"
        tn.cell(r, 6).number_format = "#,##0.0"

    tn["A8"] = "Base blended mid ($m)"
    tn["B8"] = "=DCF!B30"
    tn["B8"].number_format = "#,##0.0"
    style_formula(tn["B8"])
    tn["A9"] = "Base blended mid ($/sh)"
    tn["B9"] = "=DCF!B33"
    tn["B9"].number_format = "0.00"
    style_formula(tn["B9"])

    # Chart source block for signed Δ tornado
    tn["A11"] = "Chart source (signed Δ vs blended mid)"
    tn["A11"].font = HEADER
    tn["A12"] = "Shock"
    tn["B12"] = "Δ ($m)"
    tn["A13"] = "=A5"
    tn["B13"] = "=E5"
    tn["A14"] = "=A6"
    tn["B14"] = "=E6"
    for r in range(13, 15):
        style_formula(tn[f"A{r}"])
        style_formula(tn[f"B{r}"])
        tn[f"B{r}"].number_format = "#,##0.0"

    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = "Tornado — FULL blended-mid sensitivity proxy Δ"
    chart.y_axis.title = None
    chart.x_axis.title = "Δ vs blended mid ($m)"
    data = Reference(tn, min_col=2, min_row=12, max_row=14)
    cats = Reference(tn, min_col=1, min_row=13, max_row=14)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.legend = None
    tn.add_chart(chart, "D11")

    tn["A16"] = "Rule: never mix an exit-multiple FULL %Δ denominator with a perpetual-growth base in this tornado."
    set_widths(tn, [22, 14, 18, 12, 20, 14])


def _build_envelope(wb: Workbook) -> None:
    env = wb.create_sheet("Envelope")
    env["A1"] = "Envelope — MIN/MAX across FULL stressed equity proxies"
    env["A1"].font = TITLE
    env["A2"] = "Live MIN/MAX on tornado proxy equities plus base blended mid. Draft, not investment advice."
    env["A2"].font = MUTED

    env["A4"] = "Series"
    env["B4"] = "Equity ($m)"
    env["C4"] = "$/sh diluted"
    for c in range(1, 4):
        env.cell(4, c).font = HEADER
        env.cell(4, c).fill = LIGHT
        env.cell(4, c).border = THIN

    env["A5"] = "Base blended mid"
    env["B5"] = "=DCF!B30"
    env["C5"] = "=DCF!B33"
    env["A6"] = "Rates FULL proxy"
    env["B6"] = "=Tornado!C5"
    env["C6"] = "=Tornado!D5"
    env["A7"] = "Commodity FULL proxy"
    env["B7"] = "=Tornado!C6"
    env["C7"] = "=Tornado!D6"
    env["A9"] = "Envelope MIN"
    env["B9"] = "=MIN(B5:B7)"
    env["C9"] = "=MIN(C5:C7)"
    env["A10"] = "Envelope MAX"
    env["B10"] = "=MAX(B5:B7)"
    env["C10"] = "=MAX(C5:C7)"

    for r in range(5, 11):
        if r == 8:
            continue
        for c in range(1, 4):
            env.cell(r, c).border = THIN
        style_formula(env.cell(r, 2))
        style_formula(env.cell(r, 3))
        env.cell(r, 2).number_format = "#,##0.0"
        env.cell(r, 3).number_format = "0.00"
    env["A9"].font = BOLD
    env["A10"].font = BOLD
    set_widths(env, [28, 14, 14])


def _build_football(wb: Workbook, second_sheet: str = "Event_Commodity") -> None:
    fb = wb.create_sheet("Football")
    fb["A1"] = "Football field — methods + FULL stress bands"
    fb["A1"].font = TITLE
    fb["A2"] = (
        "Draft, not investment advice. ILLUSTRATIVE. Stress bands are blended-mid sensitivity proxies from FULL rows. "
        "Floating-bar chart uses Low and Width (=High−Low) formula columns."
    )
    fb["A2"].font = MUTED

    for c, h in enumerate(
        ["Method / band", "Low ($/sh)", "Mid ($/sh)", "High ($/sh)", "Notes", "Invisible base", "Float width"],
        1,
    ):
        fb.cell(4, c, h)
        fb.cell(4, c).font = HEADER
        fb.cell(4, c).fill = LIGHT
        fb.cell(4, c).border = THIN

    fb["A5"] = "Perpetual growth (method)"
    fb["B5"] = "=DCF!B31"
    fb["C5"] = "=DCF!B31"
    fb["D5"] = "=DCF!B31"
    fb["E5"] = "Own method point estimate"

    fb["A6"] = "Exit multiple (method)"
    fb["B6"] = "=DCF!B32"
    fb["C6"] = "=DCF!B32"
    fb["D6"] = "=DCF!B32"
    fb["E6"] = "Own method point estimate"

    fb["A7"] = "Blended midpoint (methods only)"
    fb["B7"] = "=DCF!B33"
    fb["C7"] = "=DCF!B33"
    fb["D7"] = "=DCF!B33"
    fb["E7"] = "Blend lock — no stress rows inside the average"

    fb["A8"] = "Rates FULL stress band (proxy)"
    fb["B8"] = "=MIN(DCF!B33,Tornado!D5)"
    fb["C8"] = "=DCF!B33"
    fb["D8"] = "=MAX(DCF!B33,Tornado!D5)"
    fb["E8"] = "Blended-mid sensitivity proxy from Event_Rates FULL"

    fb["A9"] = f"{second_sheet.replace("Event_", "")} FULL stress band (proxy)"
    fb["B9"] = "=MIN(DCF!B33,Tornado!D6)"
    fb["C9"] = "=DCF!B33"
    fb["D9"] = "=MAX(DCF!B33,Tornado!D6)"
    fb["E9"] = f"Blended-mid sensitivity proxy from {second_sheet} FULL"

    fb["A10"] = "Envelope band"
    fb["B10"] = "=Envelope!C9"
    fb["C10"] = "=DCF!B33"
    fb["D10"] = "=Envelope!C10"
    fb["E10"] = "MIN/MAX of base and FULL proxies"

    for r in range(5, 11):
        fb.cell(r, 6, f"=B{r}")  # invisible base = low
        fb.cell(r, 7, f"=D{r}-B{r}")  # width
        for c in range(1, 8):
            fb.cell(r, c).border = THIN
        for c in range(2, 5):
            style_formula(fb.cell(r, c))
            fb.cell(r, c).number_format = "0.00"
        style_formula(fb.cell(r, 6))
        style_formula(fb.cell(r, 7))
        fb.cell(r, 6).number_format = "0.00"
        fb.cell(r, 7).number_format = "0.00"

    fb["A12"] = "Raw WACC (reference, outside blend)"
    fb["B12"] = "=Inputs!B45"
    fb["B12"].number_format = "0.00%"
    style_formula(fb["B12"])
    fb["C12"] = "Shown live when a yellow judgment WACC override is used — outside the blend"
    fb["C12"].fill = YELLOW
    fb["C12"].font = BLUE

    fb["A14"] = "Diluted shares (million)"
    fb["B14"] = "=Inputs!B39"
    style_formula(fb["B14"])

    # Floating bar chart: stacked bar with base (invisible) + width
    chart = BarChart()
    chart.type = "bar"
    chart.grouping = "stacked"
    chart.title = "Football field ($/sh diluted)"
    chart.y_axis.title = None
    chart.x_axis.title = "$/sh"
    data = Reference(fb, min_col=6, min_row=4, max_col=7, max_row=10)
    cats = Reference(fb, min_col=1, min_row=5, max_row=10)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    # Make first series (invisible base) no-fill via graphical properties where supported
    try:
        from openpyxl.chart.shapes import GraphicalProperties as GP
        from openpyxl.drawing.line import LineProperties
        s0 = chart.series[0]
        s0.graphicalProperties = GP(noFill=True)
        s0.graphicalProperties.line = LineProperties(noFill=True)
    except Exception:
        pass
    fb.add_chart(chart, "A16")

    set_widths(fb, [40, 12, 12, 12, 55, 14, 12])


def _build_passthrough(wb: Workbook, pt: dict[str, Any]) -> None:
    ps = wb.create_sheet("PassThrough")
    ps["A1"] = "Pass-through fixture — units × price × unit cost (fixed cost separate)"
    ps["A1"].font = TITLE
    ps["A2"] = (
        "Invariant: at 100% pass-through and zero elasticity / markdowns / mitigation, "
        "gross profit (GP) dollars stay flat after a unit-cost shock. Draft, not investment advice."
    )
    ps["A2"].font = MUTED

    ps["A4"] = "Inputs"
    ps["A4"].font = HEADER
    ps["A4"].fill = LIGHT

    def put(r, label, val, fmt, judgment=False):
        ps[f"A{r}"] = label
        ps[f"B{r}"] = val
        if fmt:
            ps[f"B{r}"].number_format = fmt
        if judgment:
            style_judgment(ps[f"B{r}"])
        else:
            style_input(ps[f"B{r}"])

    put(5, "Units", pt["units"], "#,##0.0")
    put(6, "Price", pt["price"], "#,##0.00")
    put(7, "Unit cost", pt["unit_cost"], "#,##0.00")
    put(8, "Fixed cost (separate)", pt["fixed_cost"], "#,##0.0")
    put(9, "Unit-cost shock %", pt.get("cost_shock_pct", 0.15), "0.0%")
    put(10, "Pass-through % of cost shock into price", pt["pass_through_pct"], "0.0%", True)
    put(11, "Demand elasticity (qty %Δ / price %Δ)", pt["elasticity"], "0.00", True)
    put(12, "Markdown % of sales", pt["markdown_pct"], "0.0%", True)
    put(13, "Sourcing mitigation % of cost shock", pt["mitigation_pct"], "0.0%", True)

    ps["A15"] = "Base bridge"
    ps["A15"].font = HEADER
    ps["A15"].fill = LIGHT
    ps["A16"] = "Base revenue"
    ps["B16"] = "=B5*B6"
    ps["A17"] = "Base variable cost"
    ps["B17"] = "=B5*B7"
    ps["A18"] = "Base gross profit (before fixed)"
    ps["B18"] = "=B16-B17"
    ps["A19"] = "Base GP after fixed cost"
    ps["B19"] = "=B18-B8"

    ps["A21"] = "Stressed bridge"
    ps["A21"].font = HEADER
    ps["A21"].fill = LIGHT
    ps["A22"] = "Effective unit-cost shock after mitigation"
    ps["B22"] = "=B9*(1-B13)"
    ps["A23"] = "Stressed unit cost"
    ps["B23"] = "=B7*(1+B22)"
    ps["A24"] = "Price increase (pass-through × cost-shock dollars)"
    ps["B24"] = "=B7*B22*B10"
    ps["A25"] = "Stressed price"
    ps["B25"] = "=B6+B24"
    ps["A26"] = "Price %Δ"
    ps["B26"] = "=(B25-B6)/B6"
    ps["A27"] = "Stressed units"
    ps["B27"] = "=B5*(1+B11*B26)"
    ps["A28"] = "Stressed revenue (after markdown)"
    ps["B28"] = "=B27*B25*(1-B12)"
    ps["A29"] = "Stressed variable cost"
    ps["B29"] = "=B27*B23"
    ps["A30"] = "Stressed gross profit (before fixed)"
    ps["B30"] = "=B28-B29"
    ps["A31"] = "Stressed GP after fixed cost"
    ps["B31"] = "=B30-B8"

    ps["A33"] = "Invariant checks"
    ps["A33"].font = HEADER
    ps["A33"].fill = LIGHT
    ps["A34"] = "GP before fixed — base"
    ps["B34"] = "=B18"
    ps["A35"] = "GP before fixed — stressed"
    ps["B35"] = "=B30"
    ps["A36"] = "Δ GP before fixed"
    ps["B36"] = "=B35-B34"
    ps["A37"] = "Invariant OK? (|ΔGP|<0.01 when PT=100% and elast=md=mit=0)"
    ps["B37"] = '=IF(AND(ABS(B10-1)<0.000000001,ABS(B11)<0.000000000001,ABS(B12)<0.000000000001,ABS(B13)<0.000000000001),IF(ABS(B36)<0.01,"PASS","FAIL"),"N/A (not in invariant posture)")'

    for r in range(16, 32):
        if ps[f"B{r}"].value and isinstance(ps[f"B{r}"].value, str) and ps[f"B{r}"].value.startswith("="):
            style_formula(ps[f"B{r}"])
            ps[f"B{r}"].number_format = "#,##0.00"
    for r in (22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 34, 35, 36):
        style_formula(ps[f"B{r}"])
        if r == 26:
            ps[f"B{r}"].number_format = "0.00%"
        else:
            ps[f"B{r}"].number_format = "#,##0.00"
    style_formula(ps["B37"])

    ps["A39"] = (
        "Pass-through is counted once (B24). Fixed occupancy/cost stays on B8 and is not mixed into unit cost. "
        "check_workbook.py asserts PASS when the fixture is in the 100% / zero-elasticity posture."
    )
    ps["A39"].alignment = Alignment(wrap_text=True)
    ps.merge_cells("A39:D40")
    set_widths(ps, [64, 18, 40, 20])


def _build_how_measured(wb: Workbook, second_sheet: str = "Event_Commodity") -> None:
    hm = wb.create_sheet("How_Measured")
    hm["A1"] = "How_Measured — methods and stress channels"
    hm["A1"].font = TITLE
    hm["A2"] = "Draft, not investment advice. Acronyms spelled out on first use on this sheet."
    notes = [
        "",
        "DCF = discounted cash flow of FCFF = free cash flow to the firm. FCFF = NOPAT + D&A − capex − ΔNWC, with NOPAT from EBIT = earnings before interest and tax. Revenue and opex are separate lines.",
        "WACC = weighted average cost of capital. Cost of equity from CAPM = capital asset pricing model: risk-free + beta × equity risk premium. Selected WACC = override if yellow cell filled, else raw CAPM.",
        "Net-debt bridge: equity = EV − (gross debt − cash) ± mark-to-market. EV = enterprise value. Financing is not inside FCFF.",
        "Event_Rates: A standalone WACC (stressed WACC = selected WACC + Δ); B information floating interest; C standalone fixed-debt MTM; D information EPS; E information buyback; FULL cumulative = A+C only.",
        f"{second_sheet}: see channel table on that sheet (commodity price or tariff/input-cost).",
        "%Δ: each FULL versus its own perpetual-growth method base, then × blended mid → blended-mid sensitivity proxy.",
        "PassThrough: units × price × unit cost with fixed cost separate; invariant at 100% pass-through and zero elasticity/markdowns/mitigation.",
        "Review: run REVIEW_CHECKLIST.md and write CLEAR or BLOCK.",
    ]
    for i, line in enumerate(notes, start=3):
        hm[f"A{i}"] = line
        hm[f"A{i}"].alignment = Alignment(wrap_text=True)
    set_widths(hm, [130])
    for r in range(4, 11):
        hm.row_dimensions[r].height = 48


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Generate same-event stress-test workbook from YAML config")
    parser.add_argument(
        "--config",
        type=Path,
        default=ROOT / "examples" / "configs" / "default_illustrative.yaml",
        help="Path to YAML config (validated against schemas/config.schema.json)",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=ROOT / "templates" / "stress_template.xlsx",
        help="Output .xlsx path",
    )
    args = parser.parse_args(argv)
    cfg = load_config(args.config)
    wb = build_workbook(cfg)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.out)
    # Strip openpyxl creator metadata paths if any — set generic creator
    # openpyxl writes creator as openpyxl; fine for scrub
    print(f"Wrote {args.out}")
    print("Note: formula caches are empty until you recalculate (LibreOffice or Excel).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
