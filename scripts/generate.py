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
from openpyxl.chart.shapes import GraphicalProperties
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
    cover["A13"] = (
        "Terms on this sheet: DCF = discounted cash flow. WACC = weighted average cost of capital. FCFF = free cash flow "
        "to the firm. EV = enterprise value. EPS = earnings per share. bp = basis points. Capex = capital expenditure. "
        "$/sh = US dollars per diluted share. FULL = all valuation channels of one event applied together."
    )
    cover["A13"].alignment = Alignment(wrap_text=True)
    cover.merge_cells("A13:F14")
    cover["A16"] = "Sheets: Inputs, DCF, Event_Rates, Event_Commodity, Tornado, Envelope, Football, PassThrough, How_Measured."
    cover["A17"] = "Workflow: build → run REVIEW_CHECKLIST.md → written CLEAR / BLOCK verdict."
    cover["A18"] = meta.get("notes") or ""
    # Raw WACC reference outside blend (disclosure)
    cover["A20"] = "Raw WACC (capital asset pricing model, CAPM) — reference outside blend"
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
    put(7, "Opex % of revenue (ex-royalty; INCLUDES depreciation and amortisation, D&A)", inp["opex_pct_of_revenue"], "0.0%",
        "Operating expenses including D&A; free cash flow to the firm (FCFF) adds D&A back. "
        "Held fixed when price moves unless cost elasticity applies")
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

    section(20, "Capital structure / weighted average cost of capital (WACC)")
    put(21, "Risk-free rate", inp["risk_free_rate"], "0.00%", "WACC = weighted average cost of capital")
    put(22, "Equity risk premium", inp["equity_risk_premium"], "0.00%", "")
    put(23, "Beta", inp["beta"], "0.00", "")
    put(24, "Pre-tax cost of debt", inp["pre_tax_cost_of_debt"], "0.00%", "")
    put(25, "Debt weight", inp["debt_weight"], "0.0%", "")
    put(26, "Equity weight", inp["equity_weight"], "0.0%", "Should equal 1 − debt weight")
    put(27, "Rate shock (bp)", inp["rate_shock_bp"], "0", "bp = basis points; ± for Event_Rates")
    override = inp.get("wacc_override")
    put(28, "Judgment WACC override (blank = formula)", override, "0.00%", "Yellow = judgment; blank uses the raw capital asset pricing model (CAPM) WACC", True)

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
    ws["C43"] = "Cost of equity = risk-free rate + beta × equity risk premium"

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
    ws["C47"] = "Recompute enterprise value (EV) / net debt from raw inputs — do not trust a vendor EV field."

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
    dcf["A1"] = "Discounted cash flow (DCF) — ILLUSTRATIVE 5-year free cash flow to the firm (FCFF, operating-only)"
    dcf["A1"].font = TITLE
    dcf["A2"] = (
        "Draft, not investment advice. FCFF = free cash flow to the firm (operating-only). "
        "Financing enters the net-debt bridge once, not FCFF. Revenue and opex are separate lines."
    )
    dcf["A2"].font = MUTED
    dcf["A3"] = (
        "Terms: EBIT = earnings before interest and tax; NOPAT = net operating profit after tax; D&A = depreciation and "
        "amortisation (inside opex, added back); Capex = capital expenditure; ΔNWC = change in net working capital; "
        "PV = present value; EBITDA = EBIT + D&A; EV = enterprise value; $/sh = US dollars per diluted share."
    )
    dcf["A3"].font = MUTED

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
        # Only one second event is built; tariff wins when both are enabled.
        use_commodity = False
    if not use_tariff and not use_commodity:
        use_commodity = True
    capex_cfg = events.get("capex") or {}
    use_capex = bool(capex_cfg.get("enabled"))

    # (sheet, tornado label, event config, shock cell on that sheet's own terms)
    event_list: list[tuple[str, str, dict]] = []
    rates_cfg = events.get("rates") or {}
    _build_event_rates(wb)
    event_list.append(("Event_Rates", rates_cfg.get("label") or "Rates FULL", rates_cfg))
    if use_commodity:
        com_cfg = events.get("commodity") or {}
        _build_event_commodity(wb)
        event_list.append(("Event_Commodity", com_cfg.get("label") or "Commodity FULL", com_cfg))
    if use_tariff:
        tar_cfg = events.get("tariff") or {}
        tariff_cfg = cfg.get("tariff") or pt
        _build_event_tariff(wb, tariff_cfg)
        event_list.append(("Event_Tariff", tar_cfg.get("label") or "Tariff / input-cost FULL", tar_cfg))
    if use_capex:
        _build_event_capex(wb, capex_cfg)
        event_list.append(("Event_Capex", capex_cfg.get("label") or "Capex cycle FULL", capex_cfg))

    for sheet, _label, ev_cfg in event_list:
        _add_event_extras(wb[sheet], sheet, ev_cfg)

    _build_tornado(wb, event_list)
    _build_envelope(wb, event_list)
    _build_football(wb, event_list)
    _build_passthrough(wb, pt)
    _build_ladder(wb, event_list)
    _build_summary(wb, event_list)
    _build_how_measured(wb, event_list)
    cover = wb["Cover"]
    cover["A16"] = "Sheets: " + ", ".join(wb.sheetnames[1:]) + "."
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
        "Shock size in basis points (bp) = Inputs!B27. Free cash flow to the firm (FCFF) stays operating-only. Financing once: "
        "FULL = weighted average cost of capital (WACC) channel + fixed-debt mark-to-market (MTM) in the bridge. "
        "Interest and buyback capacity are INFORMATION rows outside FULL. Draft, not investment advice."
    )
    er["A2"].font = MUTED
    er["A2"].alignment = Alignment(wrap_text=True)
    er.merge_cells("A2:G3")
    er["A4"] = (
        "Terms: %Δ = percentage change; PV = present value; TV = terminal value; EV = enterprise value; FV = fair value; "
        "EPS = earnings per share; $/sh = US dollars per diluted share."
    )
    er["A4"].font = MUTED

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
    # FULL = equity at stressed WACC with MTM-adjusted net debt = (EV_stressed - (net_debt - MTM)) = B30 + B35.

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

    er["A13"] = "Own method base for %Δ (perpetual-growth equity)"
    er["B13"] = "=DCF!B28"
    er["B13"].number_format = "#,##0.0"
    style_formula(er["B13"])
    er["A14"] = "Blended mid (reference)"
    er["B14"] = "=DCF!B30"
    er["B14"].number_format = "#,##0.0"
    style_formula(er["B14"])
    er["A16"] = "FULL row is sheet row 11. Information rows and the row-12 interaction diagnostic are excluded from FULL by design."
    er["A17"] = "Never stack a second event into this table — each other event has its own Event_ sheet."
    set_widths(er, [8, 14, 48, 18, 14, 18, 14, 70])


def _build_event_commodity(wb: Workbook) -> None:
    ec = wb.create_sheet("Event_Commodity")
    ec["A1"] = "EVENT: Commodity price shock — ONE event (standalone / cumulative / diagnostic)"
    ec["A1"].font = TITLE
    ec["A2"] = (
        "Price moves revenue only. Opex (operating expenses) moves only via Inputs!B13 cost elasticity (yellow judgment). "
        "Royalty uses the base calibration rate on Inputs!B12 (never re-fitted). Draft, not investment advice."
    )
    ec["A2"].font = MUTED
    ec["A2"].alignment = Alignment(wrap_text=True)
    ec.merge_cells("A2:G3")
    ec["A4"] = (
        "Terms: FCFF = free cash flow to the firm; DA = depreciation and amortisation; capex = capital expenditure; "
        "%Δ = percentage change; $/sh = US dollars per diluted share."
    )
    ec["A4"].font = MUTED

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
    ec["A16"] = "FULL is sheet row 11. Diagnostic row E and the row-12 interaction diagnostic are excluded from FULL."
    ec["A17"] = "Pitfall guard: never overwrite Inputs!B12 historical royalty calibration with the shocked rate."
    set_widths(ec, [8, 12, 58, 18, 14, 18, 14])



def _build_event_tariff(wb: Workbook, tariff: dict) -> None:
    """Units × price × unit-cost tariff / input-cost event. Fixed cost separate. Pass-through once."""
    et = wb.create_sheet("Event_Tariff")
    et["A1"] = "EVENT: Tariff / input-cost shock — ONE event (units × price × unit cost)"
    et["A1"].font = TITLE
    et["A2"] = (
        "Fixed costs stay separate. Pass-through counted once. At 100% pass-through and zero elasticity/"
        "markdowns/mitigation, the change in gross profit (GP) is 0. At partial pass-through, +cost lowers equity. "
        "Draft, not investment advice."
    )
    et["A2"].font = MUTED
    et["A2"].alignment = Alignment(wrap_text=True)
    et.merge_cells("A2:G3")
    et["A4"] = (
        "Terms: PT = pass-through; ΔNOPAT = change in net operating profit after tax; FCFF = free cash flow to the firm; "
        "%Δ = percentage change; $/sh = US dollars per diluted share."
    )
    et["A4"].font = MUTED

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
    def write_delta_row(row, label, pt_ref, elast_ref, md_ref, mit_ref, shock_ref="$B$25"):
        et.cell(row, 1, label)
        for col in range(2, 7):
            cl = get_column_letter(col)
            # base
            # units already in row 33
            # eff shock, stressed uc, price, units, gp delta
            u = f"{cl}33"
            eff = f"({shock_ref}*(1-{mit_ref}))"
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
    # Row 40: mitigation-only posture (cost shock reduced by mitigation, no other response) — used by
    # the interaction diagnostic only.
    write_delta_row(40, "ΔNOPAT mitigation only (PT=elast=md=0)", "0", "0", "0", "$B$29")

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
    equity_from_delta_row(47, 40, "Equity mitigation-only posture (interaction diagnostic only)")

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
        (7, "B", "cumulative step", "Price pass-through (A + pass-through)", "=B42"),
        (8, "C", "cumulative step", "Demand elasticity (B + elasticity)", "=B43"),
        (9, "D", "cumulative step", "Markdowns / promotions (C + markdown)", "=B44"),
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

    et["A13"] = "Own method base for %Δ"
    et["B13"] = "=DCF!B28"
    et["B13"].number_format = "#,##0.0"
    style_formula(et["B13"])
    et["A14"] = "Blended mid (reference)"
    et["B14"] = "=DCF!B30"
    et["B14"].number_format = "#,##0.0"
    style_formula(et["B14"])
    et["A16"] = "FULL is sheet row 11 (row 12 interaction is a diagnostic). Shock input for tests: B25 (unit-cost shock %). Pass-through B26."
    et["A17"] = "Never stack rates into this table — use Event_Rates."
    # Fixed cost not in unit bridge — shown for disclosure
    et["A18"] = "Fixed cost (annual, separate — not in unit cost)"
    et["B18"] = "=B24"
    style_formula(et["B18"])
    set_widths(et, [8, 12, 52, 18, 14, 18, 14])


EVENT_SPECS: dict[str, dict[str, Any]] = {
    "Event_Rates": {
        "shock_ref": "Inputs!$B$27",
        "unit": "bp",
        "fmt": "0",
        "standalone_rows": [6, 8],
        "short": "Rates",
    },
    "Event_Commodity": {
        "shock_ref": "Inputs!$B$18",
        "unit": "%",
        "fmt": "0.0%",
        "standalone_rows": [6, 7, 8, 9],
        "short": "Commodity",
    },
    "Event_Tariff": {
        "shock_ref": "$B$25",
        "unit": "%",
        "fmt": "0.0%",
        # Rows B-D are a sequential ladder (each adds one response to the unit-cost shock), so their
        # increments sum to row D exactly; the only separable add-on left is sourcing mitigation (B47).
        "interaction_formula": "=D11-D9-(B47-D6)",
        "interaction_label": "FULL − (ladder A→D) − (mitigation-only − A)",
        "short": "Tariff",
    },
    "Event_Capex": {
        "shock_ref": "$B$21",
        "unit": "%",
        "fmt": "0.0%",
        "standalone_rows": [6, 7, 8],
        "short": "Capex",
    },
}

DEFAULT_REVERSE_MULTIPLES = [0.25, 0.5, 0.75, 1.0, 1.25, 1.5, 1.75, 2.0]
DEFAULT_THRESHOLD_DROP = 0.25  # default reverse-stress threshold = base $/sh × (1 − 25%)
INTERACTION_LABEL = "Cross-channel interaction (not size non-linearity)"

# Fixed layout of the blocks added to every event sheet (stable for check_workbook / tests).
ROW_INTERACTION = 12
ROW_CALIB_TITLE, ROW_CALIB_SHOCK, ROW_CALIB_SOURCE, ROW_CALIB_BASIS = 54, 55, 56, 57
ROW_MIRROR_TITLE, ROW_MIRROR_HDR, ROW_MIRROR_CONF, ROW_MIRROR_MIRROR, ROW_MIRROR_ASYM = 59, 60, 61, 62, 63
ROW_REV_TITLE, ROW_REV_THRESHOLD, ROW_REV_RESULT, ROW_REV_HDR = 65, 66, 67, 68
ROW_REV_FIRST = 69  # 8 grid rows: 69..76
N_REV_STEPS = 8
ROW_HELP_TITLE, ROW_HELP_HDR, ROW_HELP_MIRROR = 78, 79, 80
ROW_HELP_FIRST = 81  # grid helpers 81..88


def _pv_sum(row_cells: list[str], wacc: str) -> str:
    return "+".join(f"{c}/(1+{wacc})^{t}" for t, c in enumerate(row_cells, start=1))


def _equity_formula(fcff_cells: list[str], wacc: str = "Inputs!$B$46") -> str:
    """Perpetual-growth equity from five FCFF cells (no leading '=')."""
    y5 = fcff_cells[-1]
    tv = f"({y5}*(1+Inputs!$B$31)/({wacc}-Inputs!$B$31))/(1+{wacc})^5"
    return f"{_pv_sum(fcff_cells, wacc)}+{tv}-Inputs!$B$47"


def _dnwc(col: int, rev_expr_by_col) -> str:
    cl = get_column_letter(col)
    if col == 2:
        return f"-({rev_expr_by_col(cl)}-Inputs!$B$5)*Inputs!$B$11"
    prev = get_column_letter(col - 1)
    return f"-({rev_expr_by_col(cl)}-{rev_expr_by_col(prev)})*Inputs!$B$11"


def _helper_rates(ws, row: int, x: str) -> str:
    """FULL equity at rate shock x (bp): selected WACC + x/10000 on base FCFF, plus MTM bridge."""
    w = f"($B$23+{x}/10000)"
    cells = [f"DCF!{get_column_letter(c)}14" for c in range(2, 7)]
    ws[f"G{row}"] = "=" + _equity_formula(cells, w) + f"+Inputs!$B$38*$B$34*{x}/10000"
    ws[f"G{row}"].number_format = "#,##0.0"
    style_formula(ws[f"G{row}"])
    return f"G{row}"


def _helper_commodity(ws, row: int, x: str) -> str:
    """FULL A+B+C+D at commodity shock x (same construction as FCFF row 41)."""
    def rev(cl):
        return f"DCF!{cl}5*(1+{x})"
    for col in range(2, 7):
        cl = get_column_letter(col)
        ebit = f"({rev(cl)}-DCF!{cl}6*(1+Inputs!$B$13*{x})-{rev(cl)}*Inputs!$B$12)"
        capex = f"-DCF!{cl}5*Inputs!$B$10*(1+Inputs!$B$14*{x})"
        ws.cell(row, col, f"={ebit}*(1-Inputs!$B$8)+DCF!{cl}5*Inputs!$B$9+({capex})+({_dnwc(col, rev)})")
        ws.cell(row, col).number_format = "#,##0.0"
        style_formula(ws.cell(row, col))
    ws[f"G{row}"] = "=" + _equity_formula([f"{get_column_letter(c)}{row}" for c in range(2, 7)])
    ws[f"G{row}"].number_format = "#,##0.0"
    style_formula(ws[f"G{row}"])
    return f"G{row}"


def _helper_capex_fcff(ws, row: int, x: str, rev_on: bool = True, capex_on: bool = True) -> None:
    """Capex-cycle FCFF path at customer-capex shock x (B..F of `row`)."""
    def rev(cl):
        return f"DCF!{cl}5*(1+$B$22*{x})" if rev_on else f"DCF!{cl}5"
    for col in range(2, 7):
        cl = get_column_letter(col)
        nopat = f"({rev(cl)}*(1-Inputs!$B$7-Inputs!$B$12))*(1-Inputs!$B$8)"
        capex = f"DCF!{cl}12*(1+$B$23*{x})" if capex_on else f"DCF!{cl}12"
        ws.cell(row, col, f"={nopat}+DCF!{cl}11+({capex})+({_dnwc(col, rev)})")
        ws.cell(row, col).number_format = "#,##0.0"
        style_formula(ws.cell(row, col))


def _helper_capex(ws, row: int, x: str) -> str:
    """FULL (A+B+C) at shock x: operating path, funding need, paired net-debt / share-count change."""
    _helper_capex_fcff(ws, row, x)
    cells = [f"{get_column_letter(c)}{row}" for c in range(2, 7)]
    ws[f"H{row}"] = "=" + _equity_formula(cells)
    pv_inc = "+".join(
        f"-DCF!{get_column_letter(c)}12*$B$23*{x}/(1+Inputs!$B$46)^{c - 1}" for c in range(2, 7)
    )
    ws[f"I{row}"] = f"=MAX(0,{pv_inc})"
    ws[f"G{row}"] = (
        f"=(H{row}+$B$24*I{row})*Inputs!$B$39/(Inputs!$B$39+$B$24*I{row}/$B$25)"
    )
    for c in ("G", "H", "I"):
        ws[f"{c}{row}"].number_format = "#,##0.0"
        style_formula(ws[f"{c}{row}"])
    return f"G{row}"


def _helper_tariff(ws, row: int, x: str) -> str:
    """FULL posture ΔNOPAT at unit-cost shock x, then equity = PV(base FCFF + ΔNOPAT) − net debt."""
    for col in range(2, 7):
        cl = get_column_letter(col)
        u = f"{cl}33"
        eff = f"({x}*(1-$B$29))"
        suc = f"($B$23*(1+{eff}))"
        sp = f"($B$22+$B$23*{eff}*$B$26)"
        ppct = f"IF($B$22=0,0,({sp}-$B$22)/$B$22)"
        su = f"({u}*(1+$B$27*{ppct}))"
        dgp = f"(({su}*{sp}*(1-$B$28)-{su}*{suc})-({u}*$B$22-{u}*$B$23))"
        ws.cell(row, col, f"={dgp}*(1-Inputs!$B$8)")
        ws.cell(row, col).number_format = "#,##0.00"
        style_formula(ws.cell(row, col))
    cells = [f"(DCF!{get_column_letter(c)}14+{get_column_letter(c)}{row})" for c in range(2, 7)]
    ws[f"G{row}"] = "=" + _equity_formula(cells)
    ws[f"G{row}"].number_format = "#,##0.0"
    style_formula(ws[f"G{row}"])
    return f"G{row}"


HELPER_WRITERS = {
    "Event_Rates": _helper_rates,
    "Event_Commodity": _helper_commodity,
    "Event_Tariff": _helper_tariff,
    "Event_Capex": _helper_capex,
}


def _proxy_cols(ws, r: int) -> None:
    """E: %Δ vs own perpetual-growth base, F: blended-mid proxy ($m), G: proxy $/sh (diluted)."""
    ws[f"E{r}"] = f"=(D{r}-DCF!$B$28)/DCF!$B$28"
    ws[f"F{r}"] = f"=DCF!$B$30*(1+E{r})"
    ws[f"G{r}"] = f"=F{r}/Inputs!$B$39"
    for c, fmt in (("D", "#,##0.0"), ("E", "0.0%"), ("F", "#,##0.0"), ("G", "0.00")):
        ws[f"{c}{r}"].number_format = fmt
        style_formula(ws[f"{c}{r}"])
        ws[f"{c}{r}"].border = THIN


def _add_event_extras(ws, sheet: str, ev_cfg: dict[str, Any]) -> None:
    """Interaction row, calibration disclosure, up/down mirror and reverse stress grid."""
    spec = EVENT_SPECS[sheet]
    shock = spec["shock_ref"]
    unit = spec["unit"]
    writer = HELPER_WRITERS[sheet]

    # --- Cross-channel interaction diagnostic (row 12; outside FULL and outside the blend) ---
    r = ROW_INTERACTION
    if "interaction_formula" in spec:
        fx = spec["interaction_formula"]
        how = spec["interaction_label"]
    else:
        rows = spec["standalone_rows"]
        fx = "=(D11-DCF!$B$28)-(" + "+".join(f"(D{x}-DCF!$B$28)" for x in rows) + ")"
        how = "FULL − sum of standalone rows (" + ", ".join(ws[f"A{x}"].value for x in rows) + ")"
    ws[f"A{r}"] = "INT"
    ws[f"B{r}"] = "diagnostic"
    ws[f"B{r}"].fill = INFO_FILL
    ws[f"C{r}"] = INTERACTION_LABEL
    ws[f"D{r}"] = fx
    ws[f"E{r}"] = f"=D{r}/DCF!$B$28"
    ws[f"F{r}"] = f"=DCF!$B$30*E{r}"
    ws[f"G{r}"] = f"=F{r}/Inputs!$B$39"
    for c, fmt in (("D", "#,##0.0"), ("E", "0.0%"), ("F", "#,##0.0"), ("G", "0.00")):
        ws[f"{c}{r}"].number_format = fmt
        style_formula(ws[f"{c}{r}"])
    for c in range(1, 8):
        ws.cell(r, c).border = THIN
    ws[f"H{r}"] = (
        f"DIAGNOSTIC (a change, not a level) = {how}. How the channels combine at this one shock size; "
        "outside FULL and the blend. 0 does NOT mean value is linear in shock size: see the mirror and "
        f"asymmetry rows {ROW_MIRROR_CONF}–{ROW_MIRROR_ASYM} and the Ladder sheet."
    )
    ws[f"H{r}"].font = MUTED

    # --- Shock calibration disclosure ---
    ws[f"A{ROW_CALIB_TITLE}"] = "Shock calibration disclosure (required by REVIEW_CHECKLIST.md)"
    ws[f"A{ROW_CALIB_TITLE}"].font = HEADER
    ws[f"A{ROW_CALIB_TITLE}"].fill = LIGHT
    ws[f"A{ROW_CALIB_SHOCK}"] = f"Shock size used by FULL ({unit})"
    ws[f"B{ROW_CALIB_SHOCK}"] = f"={shock}"
    ws[f"B{ROW_CALIB_SHOCK}"].number_format = spec["fmt"]
    style_formula(ws[f"B{ROW_CALIB_SHOCK}"])
    ws[f"A{ROW_CALIB_SOURCE}"] = "Calibration source (data set, vendor or document, and date)"
    ws[f"B{ROW_CALIB_SOURCE}"] = ev_cfg.get("calibration_source") or "TO FILL: source"
    style_judgment(ws[f"B{ROW_CALIB_SOURCE}"])
    ws[f"A{ROW_CALIB_BASIS}"] = "Calibration basis (e.g. historical 1-year 95th percentile move)"
    ws[f"B{ROW_CALIB_BASIS}"] = ev_cfg.get("calibration_basis") or "TO FILL: basis"
    style_judgment(ws[f"B{ROW_CALIB_BASIS}"])
    ws[f"C{ROW_CALIB_SOURCE}"] = "Judgment (yellow). Cite the source; never invent a statistic. TO FILL blocks a real-issuer CLEAR."
    ws[f"C{ROW_CALIB_SOURCE}"].font = MUTED

    # --- Up / down mirror ---
    hdr = ["Case", "Multiple of shock", f"Shock size ({unit})", "FULL equity ($m, perp. method)",
           "%Δ vs own method base", "Blended-mid sensitivity proxy ($m)", "Proxy $/sh diluted"]
    ws[f"A{ROW_MIRROR_TITLE}"] = "Up / down mirror — the same event in both directions (diagnostic, outside the blend)"
    ws[f"A{ROW_MIRROR_TITLE}"].font = HEADER
    ws[f"A{ROW_MIRROR_TITLE}"].fill = LIGHT
    for c, h in enumerate(hdr, 1):
        ws.cell(ROW_MIRROR_HDR, c, h).font = HEADER
        ws.cell(ROW_MIRROR_HDR, c).border = THIN
    rc, rm, ra = ROW_MIRROR_CONF, ROW_MIRROR_MIRROR, ROW_MIRROR_ASYM
    ws[f"A{rc}"] = "Configured shock (FULL row 11)"
    ws[f"B{rc}"] = "1×"
    ws[f"C{rc}"] = f"={shock}"
    ws[f"D{rc}"] = "=D11"
    _proxy_cols(ws, rc)
    ws[f"A{rm}"] = "Mirrored shock (same size, opposite sign)"
    ws[f"B{rm}"] = "−1×"
    ws[f"C{rm}"] = f"=-{shock}"
    ws[f"D{rm}"] = "=" + writer(ws, ROW_HELP_MIRROR, f"$C${rm}")
    _proxy_cols(ws, rm)
    for rr in (rc, rm):
        ws[f"C{rr}"].number_format = spec["fmt"]
        style_formula(ws[f"C{rr}"])
    ws[f"A{ra}"] = "Asymmetry = (configured − base) + (mirrored − base)"
    ws[f"D{ra}"] = f"=(D{rc}-DCF!$B$28)+(D{rm}-DCF!$B$28)"
    ws[f"E{ra}"] = f"=E{rc}+E{rm}"
    ws[f"F{ra}"] = f"=(F{rc}-DCF!$B$30)+(F{rm}-DCF!$B$30)"
    ws[f"G{ra}"] = f"=(G{rc}-DCF!$B$33)+(G{rm}-DCF!$B$33)"
    for c, fmt in (("D", "#,##0.0"), ("E", "0.0%"), ("F", "#,##0.0"), ("G", "0.00")):
        ws[f"{c}{ra}"].number_format = fmt
        style_formula(ws[f"{c}{ra}"])
    ws[f"H{ra}"] = "0 = symmetric at this size. Negative = the bad direction hurts more than the good direction helps."
    ws[f"H{ra}"].font = MUTED

    # --- Reverse stress test (formula grid; no goal seek, no macros) ---
    ws[f"A{ROW_REV_TITLE}"] = "Reverse stress test — smallest tested shock that takes FULL value per share below a threshold"
    ws[f"A{ROW_REV_TITLE}"].font = HEADER
    ws[f"A{ROW_REV_TITLE}"].fill = LIGHT
    th = ROW_REV_THRESHOLD
    ws[f"A{th}"] = "Threshold value per share ($/sh, blended-mid proxy). Default = base × (1 − drop %); blank = n/a"
    drop = ev_cfg.get("threshold_drop_pct", DEFAULT_THRESHOLD_DROP)
    ws[f"C{th}"] = "Drop below base value per share →"
    ws[f"C{th}"].font = MUTED
    ws[f"D{th}"] = drop  # None (config null) leaves it blank → n/a
    ws[f"D{th}"].number_format = "0.0%"
    style_judgment(ws[f"D{th}"])
    override = ev_cfg.get("threshold_value_per_share")
    if override is None:
        ws[f"B{th}"] = f'=IF(D{th}="","",DCF!$B$33*(1-D{th}))'
        ws[f"E{th}"] = "Yellow judgment cells: change the drop % (D), or type a $/sh level over B. Blank drop % → n/a."
    else:
        ws[f"B{th}"] = float(override)
        ws[f"E{th}"] = "B is a typed threshold (config threshold_value_per_share); the drop % in D is not used."
    ws[f"E{th}"].font = MUTED
    ws[f"B{th}"].number_format = "0.00"
    style_judgment(ws[f"B{th}"])
    first, last = ROW_REV_FIRST, ROW_REV_FIRST + N_REV_STEPS - 1
    res = ROW_REV_RESULT
    ws[f"A{res}"] = f"Result: smallest tested shock ({unit}) with FULL proxy $/sh below threshold"
    idx = f'INDEX(C{first}:C{last},MATCH(MIN(E{first}:E{last}),E{first}:E{last},0))'
    if unit == "bp":
        # IF sign + TEXT, not "+0", because LibreOffice TEXT with a leading + renders negatives as "-+200"
        shock_txt = f'IF({idx}>=0,"+","")&TEXT({idx},"0")&" basis points"'
    else:
        shock_txt = f'IF({idx}>=0,"+","")&TEXT({idx},"0.0%")'
    miss = f'"Not reached: worst is "&TEXT(MIN(E{first}:E{last}),"0.0%")&" at "&{shock_txt}'
    ws[f"B{res}"] = (
        f'=IF(B{th}="","n/a",IF(DCF!$B$33<B{th},"base already below threshold",'
        f'IFERROR(INDEX(C{first}:C{last},MATCH(1,H{first}:H{last},0)),{miss})))'
    )
    ws[f"B{res}"].number_format = spec["fmt"]
    ws[f"C{res}"] = f'=IF(ISNUMBER(B{res}),INDEX(B{first}:B{last},MATCH(1,H{first}:H{last},0)),"")'
    ws[f"C{res}"].number_format = "0.00\"×\""
    style_formula(ws[f"B{res}"])
    style_formula(ws[f"C{res}"])
    ws[f"D{res}"] = "← multiple of the configured shock. Grid runs in the configured (adverse) direction; keep multiples ascending."
    ws[f"D{res}"].font = MUTED
    rhdr = ["Step", "Multiple of shock", f"Shock size ({unit})", "FULL equity ($m, perp. method)",
            "%Δ vs own method base", "Blended-mid sensitivity proxy ($m)", "Proxy $/sh diluted", "Below threshold? (1 = yes)"]
    for c, h in enumerate(rhdr, 1):
        ws.cell(ROW_REV_HDR, c, h).font = HEADER
        ws.cell(ROW_REV_HDR, c).border = THIN
    multiples = ev_cfg.get("reverse_stress_multiples") or DEFAULT_REVERSE_MULTIPLES
    for i in range(N_REV_STEPS):
        r = first + i
        ws[f"A{r}"] = i + 1
        ws[f"B{r}"] = float(multiples[i])
        ws[f"B{r}"].number_format = "0.00"
        style_input(ws[f"B{r}"])
        ws[f"C{r}"] = f"={shock}*B{r}"
        ws[f"C{r}"].number_format = spec["fmt"]
        style_formula(ws[f"C{r}"])
        ws[f"D{r}"] = "=" + writer(ws, ROW_HELP_FIRST + i, f"$C${r}")
        _proxy_cols(ws, r)
        ws[f"H{r}"] = f'=IF($B${th}="","",IF(G{r}<$B${th},1,0))'
        style_formula(ws[f"H{r}"])

    ws[f"A{ROW_HELP_TITLE}"] = "Helper rows for the mirror and grid (live formulas; same construction as the FULL row)"
    ws[f"A{ROW_HELP_TITLE}"].font = HEADER
    ws[f"A{ROW_HELP_HDR}"] = "Helper"
    for i, h in enumerate(["Y1", "Y2", "Y3", "Y4", "Y5", "FULL equity ($m)"], start=2):
        ws.cell(ROW_HELP_HDR, i, h).font = HEADER
    if sheet == "Event_Capex":
        ws[f"H{ROW_HELP_HDR}"] = "Equity before financing ($m)"
        ws[f"I{ROW_HELP_HDR}"] = "Funding need ($m, PV)"
        ws[f"G{ROW_HELP_HDR}"] = "FULL equity, existing holders ($m)"
    ws[f"A{ROW_HELP_MIRROR}"] = "Mirror"
    for i in range(N_REV_STEPS):
        ws[f"A{ROW_HELP_FIRST + i}"] = f"Grid step {i + 1}"


def _build_event_capex(wb: Workbook, cx: dict[str, Any]) -> None:
    """Customer / own capex-cycle event: own revenue, own capex, financing mix; FULL = all three."""
    ws = wb.create_sheet("Event_Capex")
    ws["A1"] = "EVENT: Capex cycle (customer capital expenditure shock) — ONE event (standalone / information / FULL)"
    ws["A1"].font = TITLE
    ws["A2"] = (
        "Capex = capital expenditure. FCFF = free cash flow to the firm, operating-only. A customer capex shock moves own "
        "revenue (via a yellow elasticity) and own capex (via a yellow response). Any funding need is financed ONCE: "
        "the equity-funded part lowers net debt and adds diluted shares in the same step; the debt-funded part adds debt "
        "and cash equally. Draft, not investment advice."
    )
    ws["A2"].font = MUTED
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:G3")
    ws["A4"] = (
        "Terms: NWC = net working capital; D&A = depreciation and amortisation; PV = present value; WACC = weighted "
        "average cost of capital; t0 = valuation date; %Δ = percentage change; $/sh = US dollars per diluted share."
    )
    ws["A4"].font = MUTED
    headers = [
        "Row", "Kind", "Channel",
        "Stressed equity, pre-event shares ($m, perp. method)", "%Δ vs own method base",
        "Blended-mid sensitivity proxy ($m)", "Proxy $/sh diluted",
    ]
    for c, h in enumerate(headers, 1):
        ws.cell(5, c, h)
        ws.cell(5, c).font = HEADER
        ws.cell(5, c).fill = LIGHT
        ws.cell(5, c).border = THIN

    ws["A20"] = "Capex-cycle event inputs"
    ws["A20"].font = HEADER
    ws["A20"].fill = LIGHT

    def put(r, label, val, fmt, judgment=False, note=""):
        ws[f"A{r}"] = label
        ws[f"B{r}"] = val
        ws[f"B{r}"].number_format = fmt
        (style_judgment if judgment else style_input)(ws[f"B{r}"])
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = MUTED

    put(21, "Customer capex shock % (enter the adverse direction)", cx.get("customer_capex_shock_pct", 0.30), "0.0%",
        note="ILLUSTRATIVE; one event only. Grid and mirror scale this cell.")
    put(22, "Own revenue elasticity to customer capex", cx.get("revenue_elasticity_to_customer_capex", 0.8), "0.00", True,
        "Judgment: own revenue %Δ per 1.0 customer capex %Δ")
    put(23, "Own capex response to customer capex", cx.get("own_capex_response", 1.0), "0.00", True,
        "Judgment: own capex %Δ per 1.0 customer capex %Δ")
    put(24, "Equity share of incremental funding (rest is debt)", cx.get("equity_funding_share", 0.5), "0.0%", True,
        "Judgment: financing mix")
    put(25, "Issue price per new share ($/sh)", cx.get("issue_price_per_share", 10.0), "0.00", True,
        "Judgment: below value per share ⇒ dilutive for existing holders")
    ws["A26"] = "Own revenue %Δ (= elasticity × shock)"
    ws["B26"] = "=B22*B21"
    ws["A27"] = "Own capex %Δ (= response × shock)"
    ws["B27"] = "=B23*B21"
    for r in (26, 27):
        ws[f"B{r}"].number_format = "0.0%"
        style_formula(ws[f"B{r}"])

    ws["A29"] = "Item"
    for i, yr in enumerate([1, 2, 3, 4, 5], start=2):
        ws.cell(29, i, f"Y{yr}").font = HEADER
    _helper_capex_fcff(ws, 30, "$B$21", rev_on=True, capex_on=False)
    ws["A30"] = "FCFF A — own revenue only (opex/royalty/NWC follow revenue; D&A and capex at base $)"
    _helper_capex_fcff(ws, 31, "$B$21", rev_on=False, capex_on=True)
    ws["A31"] = "FCFF B — own capex only (base revenue)"
    _helper_capex_fcff(ws, 32, "$B$21", rev_on=True, capex_on=True)
    ws["A32"] = "FCFF operating FULL (A+B, before financing)"
    ws["A33"] = "Incremental own capex outlay ($m)"
    for col in range(2, 7):
        cl = get_column_letter(col)
        ws.cell(33, col, f"=-DCF!{cl}12*$B$27")
        ws.cell(33, col).number_format = "#,##0.0"
        style_formula(ws.cell(33, col))

    rows = [
        (35, "PV of incremental own capex at selected WACC ($m)",
         "=" + "+".join(f"{get_column_letter(c)}33/(1+Inputs!$B$46)^{c - 1}" for c in range(2, 7)), "#,##0.0"),
        (36, "Funding need raised at t0 ($m) = MAX(0, PV incremental capex)", "=MAX(0,B35)", "#,##0.0"),
        (37, "Equity-funded part ($m)", "=B24*B36", "#,##0.0"),
        (38, "Debt-funded part ($m) — adds gross debt and cash equally (net debt unchanged)", "=(1-B24)*B36", "#,##0.0"),
        (39, "New diluted shares (million) = equity-funded / issue price", "=B37/B25", "#,##0.00"),
        (40, "Net-debt change from financing ($m, once)", "=-B37", "#,##0.0"),
        (41, "Post-event diluted shares (million)", "=Inputs!B39+B39", "#,##0.00"),
        (43, "Equity A standalone (base net debt, base shares)",
         "=" + _equity_formula([f"{get_column_letter(c)}30" for c in range(2, 7)]), "#,##0.0"),
        (44, "Equity B standalone (base net debt, base shares)",
         "=" + _equity_formula([f"{get_column_letter(c)}31" for c in range(2, 7)]), "#,##0.0"),
        (45, "Equity C standalone — financing mix on base operations, attributable to pre-event shares",
         "=(DCF!B28-B40)*Inputs!B39/B41", "#,##0.0"),
        (46, "Equity operating FULL before financing (information)",
         "=" + _equity_formula([f"{get_column_letter(c)}32" for c in range(2, 7)]), "#,##0.0"),
        (47, "Equity FULL A+B+C, attributable to pre-event shares", "=(B46-B40)*Inputs!B39/B41", "#,##0.0"),
        (48, "All-equity funding view (information): 100% of need issued at B25",
         "=(B46+B36)*Inputs!B39/(Inputs!B39+B36/B25)", "#,##0.0"),
        (49, "FULL $/sh on post-event diluted shares (own perp. method, display)", "=(B46-B40)/B41", "0.00"),
    ]
    for r, label, fx, fmt in rows:
        ws[f"A{r}"] = label
        ws[f"B{r}"] = fx
        ws[f"B{r}"].number_format = fmt
        style_formula(ws[f"B{r}"])

    channels = [
        (6, "A", "standalone", "Own revenue via customer capex elasticity", "=B43"),
        (7, "B", "standalone", "Own capex response (FCFF)", "=B44"),
        (8, "C", "standalone", "Financing mix: net debt and diluted shares change once, paired", "=B45"),
        (9, "D", "information", "FCFF carrier: operating A+B before financing (out of FULL)", "=B46"),
        (10, "E", "information", "All-equity funding view (out of FULL)", "=B48"),
        (11, "FULL", "cumulative", "A+B+C together (D/E information excluded)", "=B47"),
    ]
    for r, rowid, kind, channel, d_fx in channels:
        ws.cell(r, 1, rowid)
        ws.cell(r, 2, kind)
        ws.cell(r, 3, channel)
        ws.cell(r, 4, d_fx)
        _proxy_cols(ws, r)
        for c in range(1, 8):
            ws.cell(r, c).border = THIN
        if kind == "information":
            ws.cell(r, 2).fill = INFO_FILL
        if rowid == "FULL":
            for c in range(1, 8):
                ws.cell(r, c).font = BOLD
    ws["A13"] = "Own method base for %Δ"
    ws["B13"] = "=DCF!B28"
    ws["A14"] = "Blended mid (reference)"
    ws["B14"] = "=DCF!B30"
    for r in (13, 14):
        ws[f"B{r}"].number_format = "#,##0.0"
        style_formula(ws[f"B{r}"])
    ws["A16"] = (
        "FULL is sheet row 11. Column D is equity attributable to the pre-event diluted shares, so %Δ and $/sh "
        "already include dilution. Shock input for tests: B21."
    )
    ws["A17"] = "Never stack rates, commodity or tariff into this table. The shock is a permanent level shift (Y1–Y5 and terminal year)."
    set_widths(ws, [8, 16, 58, 20, 14, 18, 14, 18, 16])


def _build_tornado(wb: Workbook, event_list) -> None:
    tn = wb.create_sheet("Tornado")
    tn["A1"] = "Tornado — FULL-row blended-mid sensitivity proxies"
    tn["A1"].font = TITLE
    tn["A2"] = (
        "Each bar uses that event's FULL percentage change (%Δ) versus its own method base, applied to the blended mid. "
        "$/sh = US dollars per diluted share. Draft, not investment advice. Chart is formula-driven from this table."
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

    n = len(event_list)
    for i, (sheet, label, _cfg) in enumerate(event_list):
        r = 5 + i
        tn[f"A{r}"] = label
        tn[f"B{r}"] = f"={sheet}!E11"
        tn[f"C{r}"] = f"={sheet}!F11"
        tn[f"D{r}"] = f"={sheet}!G11"
        tn[f"E{r}"] = f"=C{r}-DCF!B30"
        tn[f"F{r}"] = f"=ABS(E{r})"
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

    tn["A11"] = "Chart source (signed Δ vs blended mid)"
    tn["A11"].font = HEADER
    tn["A12"] = "Shock"
    tn["B12"] = "Δ ($m)"
    for i in range(n):
        r = 13 + i
        tn[f"A{r}"] = f"=A{5 + i}"
        tn[f"B{r}"] = f"=E{5 + i}"
        style_formula(tn[f"A{r}"])
        style_formula(tn[f"B{r}"])
        tn[f"B{r}"].number_format = "#,##0.0"

    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = "Tornado — FULL blended-mid sensitivity proxy Δ"
    chart.y_axis.title = None
    chart.x_axis.title = "Δ vs blended mid ($m)"
    data = Reference(tn, min_col=2, min_row=12, max_row=12 + n)
    cats = Reference(tn, min_col=1, min_row=13, max_row=12 + n)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.legend = None
    tn.add_chart(chart, "D11")

    tn["A16"] = "Rule: never mix an exit-multiple FULL %Δ denominator with a perpetual-growth base in this tornado."
    set_widths(tn, [22, 14, 18, 12, 20, 14])


def _build_envelope(wb: Workbook, event_list) -> None:
    env = wb.create_sheet("Envelope")
    env["A1"] = "Envelope — MIN/MAX across FULL stressed equity proxies"
    env["A1"].font = TITLE
    env["A2"] = (
        "Live MIN/MAX on tornado proxy equities plus base blended mid; covers every enabled event. "
        "$/sh = US dollars per diluted share. Draft, not investment advice."
    )
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
    for i, (sheet, _label, _cfg) in enumerate(event_list):
        r = 6 + i
        env[f"A{r}"] = f"{EVENT_SPECS[sheet]['short']} FULL proxy"
        env[f"B{r}"] = f"=Tornado!C{5 + i}"
        env[f"C{r}"] = f"=Tornado!D{5 + i}"
    last = 5 + len(event_list)
    env["A9"] = "Envelope MIN"
    env["B9"] = f"=MIN(B5:B{last})"
    env["C9"] = f"=MIN(C5:C{last})"
    env["A10"] = "Envelope MAX"
    env["B10"] = f"=MAX(B5:B{last})"
    env["C10"] = f"=MAX(C5:C{last})"

    for r in list(range(5, last + 1)) + [9, 10]:
        for c in range(1, 4):
            env.cell(r, c).border = THIN
        style_formula(env.cell(r, 2))
        style_formula(env.cell(r, 3))
        env.cell(r, 2).number_format = "#,##0.0"
        env.cell(r, 3).number_format = "0.00"
    env["A9"].font = BOLD
    env["A10"].font = BOLD
    set_widths(env, [28, 14, 14])


def _football_last_row(event_list) -> int:
    return 11 if len(event_list) > 2 else 10


def _build_football(wb: Workbook, event_list) -> None:
    fb = wb.create_sheet("Football")
    fb["A1"] = "Football field — methods + FULL stress bands"
    fb["A1"].font = TITLE
    fb["A2"] = (
        "Draft, not investment advice. ILLUSTRATIVE. Stress bands are blended-mid sensitivity proxies from FULL rows. "
        "$/sh = US dollars per diluted share. Floating-bar chart uses Low and Width (=High−Low) formula columns."
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

    # Event bands: rows 8, 9 for the first two events; envelope stays on row 10; a third event uses row 11.
    band_rows = [8, 9, 11]
    for i, (sheet, _label, _cfg) in enumerate(event_list):
        r = band_rows[i]
        t = 5 + i
        fb[f"A{r}"] = f"{sheet.replace('Event_', '')} FULL stress band (proxy)"
        fb[f"B{r}"] = f"=MIN(DCF!B33,Tornado!D{t})"
        fb[f"C{r}"] = "=DCF!B33"
        fb[f"D{r}"] = f"=MAX(DCF!B33,Tornado!D{t})"
        fb[f"E{r}"] = f"Blended-mid sensitivity proxy from {sheet} FULL"

    fb["A10"] = "Envelope band"
    fb["B10"] = "=Envelope!C9"
    fb["C10"] = "=DCF!B33"
    fb["D10"] = "=Envelope!C10"
    fb["E10"] = "MIN/MAX of base and FULL proxies"

    last = _football_last_row(event_list)
    for r in range(5, last + 1):
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

    fb["A12"] = "Raw weighted average cost of capital (WACC) — reference, outside blend"
    fb["B12"] = "=Inputs!B45"
    fb["B12"].number_format = "0.00%"
    style_formula(fb["B12"])
    fb["C12"] = "Shown live when a yellow judgment WACC override is used — outside the blend"
    fb["C12"].fill = YELLOW
    fb["C12"].font = BLUE

    fb["A14"] = "Diluted shares (million)"
    fb["B14"] = "=Inputs!B39"
    style_formula(fb["B14"])

    chart = BarChart()
    chart.type = "bar"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "Football field ($/sh diluted)"
    chart.y_axis.title = None
    chart.x_axis.title = "$/sh"
    data = Reference(fb, min_col=6, min_row=4, max_col=7, max_row=last)
    cats = Reference(fb, min_col=1, min_row=5, max_row=last)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    from openpyxl.drawing.line import LineProperties
    s0 = chart.series[0]
    s0.graphicalProperties = GraphicalProperties(noFill=True)
    s0.graphicalProperties.line = LineProperties(noFill=True)
    fb.add_chart(chart, "A16")

    set_widths(fb, [40, 12, 12, 12, 55, 14, 12])


def _build_summary(wb: Workbook, event_list) -> None:
    """Rank every enabled event by absolute FULL change in value per share (live INDEX/MATCH/LARGE)."""
    sm = wb.create_sheet("Summary", 1)
    sm["A1"] = "Summary — events ranked by absolute FULL change in value per share"
    sm["A1"].font = TITLE
    sm["A2"] = (
        "Draft, not investment advice. ILLUSTRATIVE. FULL = all channels of one event applied together. "
        "$/sh = US dollars per diluted share, shown as a blended-mid sensitivity proxy (the FULL percentage change (%Δ) "
        "against its own perpetual-growth method base, applied to the blended midpoint). Every cell is a formula linked "
        "to the event sheets; nothing here feeds the blend. bp = basis points."
    )
    sm["A2"].font = MUTED
    sm["A2"].alignment = Alignment(wrap_text=True)
    sm.merge_cells("A2:K3")
    sm.row_dimensions[2].height = 30

    n = len(event_list)
    h0 = 13  # helper header row
    first, last = h0 + 1, h0 + n
    hdr = ["Rank", "Event", "Sheet", "Configured shock", "FULL $/sh down-case", "FULL $/sh up-case",
           "%Δ at configured shock", "Δ $/sh at configured shock", INTERACTION_LABEL + " ($/sh)",
           "Asymmetry ($/sh)", "Reverse-stress result (shock)"]
    for c, h in enumerate(hdr, 1):
        sm.cell(4, c, h).font = HEADER
        sm.cell(4, c).fill = LIGHT
        sm.cell(4, c).border = THIN
        sm.cell(4, c).alignment = Alignment(wrap_text=True)
    sm.row_dimensions[4].height = 30

    # Unranked source block (one row per event, same order as Tornado)
    sm[f"A{h0 - 1}"] = "Source block (unranked; formula links to each event sheet)"
    sm[f"A{h0 - 1}"].font = HEADER
    src_hdr = ["Event", "Sheet", "Configured shock", "FULL $/sh @ configured", "FULL $/sh @ mirrored",
               "%Δ @ configured", "Δ $/sh @ configured", INTERACTION_LABEL + " ($/sh)", "Asymmetry ($/sh)",
               "Reverse-stress result", "Rank key (|Δ $/sh|, tie-break by row)"]
    for c, h in enumerate(src_hdr, 1):
        sm.cell(h0, c, h).font = HEADER
        sm.cell(h0, c).border = THIN
    for i, (sheet, label, _cfg) in enumerate(event_list):
        r = first + i
        spec = EVENT_SPECS[sheet]
        sm[f"A{r}"] = label
        sm[f"B{r}"] = sheet
        if spec["unit"] == "bp":
            sm[f"C{r}"] = f'=TEXT({sheet}!C{ROW_MIRROR_CONF},"0")&" bp"'
        else:
            sm[f"C{r}"] = f'=TEXT({sheet}!C{ROW_MIRROR_CONF},"0.0%")'
        sm[f"D{r}"] = f"={sheet}!G11"
        sm[f"E{r}"] = f"={sheet}!G{ROW_MIRROR_MIRROR}"
        sm[f"F{r}"] = f"={sheet}!E11"
        sm[f"G{r}"] = f"=D{r}-DCF!$B$33"
        sm[f"H{r}"] = f"={sheet}!G{ROW_INTERACTION}"
        sm[f"I{r}"] = f"={sheet}!G{ROW_MIRROR_ASYM}"
        sm[f"J{r}"] = f"={sheet}!B{ROW_REV_RESULT}"
        sm[f"K{r}"] = f"=ABS(G{r})-ROW()/1000000000"
        for c, fmt in (("C", "@"), ("D", "0.00"), ("E", "0.00"), ("F", "0.0%"), ("G", "0.00"),
                       ("H", "0.00"), ("I", "0.00"), ("J", spec["fmt"]), ("K", "0.000000")):
            if c != "C":
                sm[f"{c}{r}"].number_format = fmt
            style_formula(sm[f"{c}{r}"])
        for c in range(1, 12):
            sm.cell(r, c).border = THIN

    keys = f"$K${first}:$K${last}"
    for k in range(1, n + 1):
        r = 4 + k
        sm[f"A{r}"] = k
        idx = f"MATCH(LARGE({keys},{k}),{keys},0)"
        def pick(col):
            return f"INDEX(${col}${first}:${col}${last},{idx})"
        sm[f"B{r}"] = f"={pick('A')}"
        sm[f"C{r}"] = f"={pick('B')}"
        sm[f"D{r}"] = f"={pick('C')}"
        sm[f"E{r}"] = f"=MIN({pick('D')},{pick('E')})"
        sm[f"F{r}"] = f"=MAX({pick('D')},{pick('E')})"
        sm[f"G{r}"] = f"={pick('F')}"
        sm[f"H{r}"] = f"={pick('G')}"
        sm[f"I{r}"] = f"={pick('H')}"
        sm[f"J{r}"] = f"={pick('I')}"
        sm[f"K{r}"] = f"={pick('J')}"
        for c, fmt in (("E", "0.00"), ("F", "0.00"), ("G", "0.0%"), ("H", "0.00"), ("I", "0.00"),
                       ("J", "0.00"), ("K", "General")):
            sm[f"{c}{r}"].number_format = fmt
        for c in range(2, 12):
            style_formula(sm.cell(r, c))
        for c in range(1, 12):
            sm.cell(r, c).border = THIN

    sm["A9"] = "Base blended mid ($/sh)"
    sm["B9"] = "=DCF!B33"
    sm["B9"].number_format = "0.00"
    style_formula(sm["B9"])
    sm["A10"] = (
        "Down-case / up-case = lower / higher of the configured and mirrored FULL proxies. Reverse-stress result = "
        "smallest tested shock that takes the FULL proxy $/sh below the yellow threshold on that event sheet "
        "(default threshold = base × (1 − 25%); n/a when the drop % or threshold is blank). Ranking uses the LARGE and INDEX/MATCH spreadsheet functions (Excel and LibreOffice)."
    )
    sm["A10"].font = MUTED
    sm["A10"].alignment = Alignment(wrap_text=True)
    sm.merge_cells("A10:K11")
    set_widths(sm, [26, 26, 16, 14, 14, 14, 14, 14, 14, 14, 18])


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


def _build_how_measured(wb: Workbook, event_list) -> None:
    hm = wb.create_sheet("How_Measured")
    hm["A1"] = "How_Measured — methods and stress channels"
    hm["A1"].font = TITLE
    hm["A2"] = "Draft, not investment advice. ILLUSTRATIVE. Acronyms spelled out on first use on this sheet."
    sheets = [e[0] for e in event_list]
    notes = [
        "",
        "DCF = discounted cash flow of FCFF = free cash flow to the firm. FCFF = NOPAT (net operating profit after tax) + D&A (depreciation and amortisation) − capex (capital expenditure) − ΔNWC (change in net working capital), with NOPAT from EBIT = earnings before interest and tax. Revenue and opex (operating expenses) are separate lines. Opex % of revenue INCLUDES D&A; FCFF adds D&A back, so D&A is never subtracted twice.",
        "WACC = weighted average cost of capital. Cost of equity from CAPM = capital asset pricing model: risk-free + beta × equity risk premium. Selected WACC = override if yellow cell filled, else raw CAPM.",
        "Net-debt bridge: equity = EV (enterprise value) − (gross debt − cash) ± mark-to-market (MTM). Financing is not inside FCFF. Per-share values use diluted shares; $/sh = US dollars per diluted share.",
        "Event_Rates: A standalone WACC (stressed WACC = selected WACC + Δ, in bp = basis points); B information floating interest; C standalone fixed-debt MTM; D information EPS (earnings per share); E information buyback; FULL cumulative = A+C only.",
    ]
    if "Event_Commodity" in sheets:
        notes.append("Event_Commodity: A revenue (price × base volume, opex fixed); B royalty at base calibration rate; C cost co-move via yellow cost elasticity; D capex pace via yellow capex–price elasticity; E diagnostic bridge view; FULL = A+B+C+D.")
    if "Event_Tariff" in sheets:
        notes.append("Event_Tariff: units × price × unit cost bridge, fixed cost separate. A unit cost (standalone); B, C, D cumulative steps adding pass-through, demand elasticity and markdowns; E diagnostic 100% pass-through posture; FULL adds sourcing mitigation. Pass-through counted once.")
    if "Event_Capex" in sheets:
        notes.append("Event_Capex: one customer capex shock. A own revenue via yellow elasticity; B own capex via yellow response; C financing mix — the funding need (PV of incremental own capex, floored at 0) is raised once at t0: the equity part lowers net debt and adds diluted shares at the yellow issue price in the same step, the debt part adds debt and cash equally. FULL = A+B+C, measured per pre-event share; D/E information rows are outside FULL.")
    notes += [
        "%Δ (percentage change): each FULL versus its own perpetual-growth method base, then × blended mid → blended-mid sensitivity proxy.",
        "Cross-channel interaction (not size non-linearity) — row 12 on each event sheet, labelled INT, diagnostic: FULL − sum of the standalone channel rows, as a change in $m and $/sh. It shows how the channels combine at one shock size; it is outside FULL and outside the blend. A zero interaction does NOT mean the value response is linear in shock size: read the mirror and asymmetry rows 61–63 for that and the Ladder sheet. On Event_Tariff the B–D rows are a sequential ladder, so INT isolates the cross-effect of sourcing mitigation with the other channels.",
        "Up / down mirror (rows 60–63): the same event at the configured shock and at the mirrored shock (opposite sign). Asymmetry = (configured − base) + (mirrored − base); 0 means symmetric at this size.",
        "Reverse stress test (rows 65–76): an 8-step formula grid of shock multiples (blue inputs, ascending) in the configured direction. The result cell reports the smallest tested shock whose FULL proxy $/sh falls below the yellow threshold. The threshold defaults to a live formula = base blended-mid $/sh × (1 − drop %), with the drop % a separate yellow input (default 25%); a blank drop % (or threshold) → n/a; typing a $/sh level over the threshold overrides the default. If no tested shock breaks the threshold, the cell reads 'Not reached: worst is -X% at …' with the worst fall and shock on the grid. No goal seek, no macros.",
        "Shock calibration (rows 54–57): yellow source and basis cells for each shock size (for example a historical 1-year 95th percentile move). Shipped as TO FILL; never invent a statistic.",
        "Ladder sheet: each event's FULL re-run at yellow multiples of its configured shock (default −1.5× to +1.5×), both directions, via helper rows 90–97 on each event sheet. Asymmetry per rung = (proxy at +k − base) + (proxy at −k − base); the ±1× rung must equal the mirror asymmetry (row 63) and the Summary, with a check cell. The reverse stress test there interpolates the shock that cuts FULL equity by each yellow % target; scripts/reverse_stress.py solves it exactly.",
        "Summary sheet: ranks every enabled event by absolute FULL change in $/sh with LARGE + INDEX/MATCH formulas; shows down-case / up-case, %Δ, interaction, asymmetry and reverse-stress result. Outside the blend.",
        "PassThrough: units × price × unit cost with fixed cost separate; invariant at 100% pass-through and zero elasticity/markdowns/mitigation.",
        "Review: run REVIEW_CHECKLIST.md and write CLEAR or BLOCK.",
    ]
    for i, line in enumerate(notes, start=3):
        hm[f"A{i}"] = line
        hm[f"A{i}"].alignment = Alignment(wrap_text=True)
    set_widths(hm, [130])
    for r in range(4, 4 + len(notes)):
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
    print(f"Wrote {args.out}")
    print("Note: formula caches are empty until you recalculate (LibreOffice or Excel).")
    return 0


# ---------------------------------------------------------------------------
# Shock ladder + reverse stress test: each event's FULL at several sizes, both directions
# ---------------------------------------------------------------------------
LADDER_MULTIPLES = [-1.5, -1.0, -0.5, 0.0, 0.5, 1.0, 1.5]
LADDER_TARGETS = [-0.10, -0.25]
LADDER_FIRST_BLOCK = 6
LADDER_BLOCK_ROWS = 12
ROW_LADDER_HELP_TITLE = 90
ROW_LADDER_HELP_FIRST = 91  # 7 ladder helpers 91..97 on each event sheet


def _build_ladder(wb: Workbook, event_list) -> None:
    from openpyxl.chart import ScatterChart, Series

    ld = wb.create_sheet("Ladder")
    ld["A1"] = "Shock ladder and reverse stress test: each event's FULL valuation at several sizes, both directions"
    ld["A1"].font = TITLE
    ld["A2"] = (
        "Each rung re-runs the SAME event's FULL (all valuation channels together) at a multiple of the configured shock, using the "
        "event sheet's own helper construction. Asymmetry ($/sh) = (proxy at +k − base) + (proxy at −k − base); the ±1× rung equals "
        "the event sheet's mirror asymmetry and the Summary. A straight-line response shows zero asymmetry on every rung. Multiples and "
        "drop targets are judgment inputs (yellow). Draft, not investment advice."
    )
    ld["A2"].font = MUTED
    ld["A2"].alignment = Alignment(wrap_text=True)
    ld.merge_cells("A2:J4")
    heads = ["Rung", "Multiple of configured shock", "Shock size", "FULL equity ($m, perp. method)",
             "Change vs base ($m)", "% change vs base", "Proxy $/sh diluted", "Asymmetry ($/sh)"]
    blocks = []
    for k, (sheet, label, _cfg) in enumerate(event_list):
        spec = EVENT_SPECS[sheet]
        shock = spec["shock_ref"]
        shock_q = f"{sheet}!{shock}" if shock.startswith("$") else shock
        ws = wb[sheet]
        writer = HELPER_WRITERS[sheet]
        ws[f"A{ROW_LADDER_HELP_TITLE}"] = "Helper rows for the Ladder sheet (live formulas; same construction as the FULL row)"
        ws[f"A{ROW_LADDER_HELP_TITLE}"].font = HEADER
        s = LADDER_FIRST_BLOCK + LADDER_BLOCK_ROWS * k
        ld.cell(s, 1, f"{label} ({spec['unit']}); event sheet {sheet}").font = HEADER
        for c, h in enumerate(heads, 1):
            cell = ld.cell(s + 1, c, h)
            cell.font = HEADER
            cell.fill = LIGHT
            cell.border = THIN
        rows = list(range(s + 2, s + 2 + len(LADDER_MULTIPLES)))
        n = len(rows)
        for i, (r, mult) in enumerate(zip(rows, LADDER_MULTIPLES)):
            ld.cell(r, 1, i + 1)
            ld.cell(r, 2, mult)
            style_judgment(ld.cell(r, 2))
            ld.cell(r, 2).number_format = '0.0"×"'
            ld.cell(r, 3, f"={shock_q}*$B{r}")
            ld.cell(r, 3).number_format = spec["fmt"]
            hrow = ROW_LADDER_HELP_FIRST + i
            ws[f"A{hrow}"] = f"Ladder rung {i + 1}"
            ref = writer(ws, hrow, f"Ladder!$C${r}")
            ld.cell(r, 4, f"={sheet}!{ref}")
            ld.cell(r, 5, f"=D{r}-DCF!$B$28")
            ld.cell(r, 6, f"=E{r}/DCF!$B$28")
            ld.cell(r, 7, f"=DCF!$B$33*(1+F{r})")
            mirror = rows[n - 1 - i]
            ld.cell(r, 8, f"=(G{r}-DCF!$B$33)+(G{mirror}-DCF!$B$33)")
            for c, fmt in ((4, "#,##0.0"), (5, "#,##0.0"), (6, "0.0%"), (7, "0.00"), (8, "0.00")):
                ld.cell(r, c).number_format = fmt
            for c in range(1, 9):
                ld.cell(r, c).border = THIN
            for c in range(3, 9):
                style_formula(ld.cell(r, c))
        r1 = rows[LADDER_MULTIPLES.index(1.0)]
        ld.cell(s + 1, 9, "Check: ±1× = mirror").font = HEADER
        ld.cell(s + 2, 9, f'=IF(ABS(H{r1}-{sheet}!$G${ROW_MIRROR_ASYM})<0.0001,"OK","CHECK")')
        style_formula(ld.cell(s + 2, 9))
        blocks.append((sheet, label, spec, rows))

    # ---- Reverse stress test by % fall in value ----
    rt = LADDER_FIRST_BLOCK + LADDER_BLOCK_ROWS * len(event_list) + 1
    ld.cell(rt, 1, "Reverse stress test: shock size that takes FULL equity down by the target (interpolated between rungs)").font = HEADER
    ld.cell(rt + 1, 1, "Target change in equity").font = HEADER
    ld.cell(rt + 1, 1).fill = LIGHT
    for k, (sheet, label, spec, rows) in enumerate(blocks):
        cell = ld.cell(rt + 1, 2 + k, f"{label}: shock needed ({spec['unit']})")
        cell.font = HEADER
        cell.fill = LIGHT
        cell.border = THIN
    helper_col0 = 12  # L onward: one helper column per (event, target)
    ld.cell(LADDER_FIRST_BLOCK - 1, helper_col0,
            "Helper: crossings between adjacent rungs for the reverse stress test (blank = target not in that interval)").font = MUTED
    for ti, tgt in enumerate(LADDER_TARGETS):
        rr = rt + 2 + ti
        ld.cell(rr, 1, tgt)
        style_judgment(ld.cell(rr, 1))
        ld.cell(rr, 1).number_format = "0%"
        for k, (sheet, label, spec, rows) in enumerate(blocks):
            hc = get_column_letter(helper_col0 + ti * len(blocks) + k)
            ld[f"{hc}{rows[0] - 1}"] = f"{sheet} target {ti + 1}"
            for a, b in zip(rows[:-1], rows[1:]):
                t = f"$A${rr}"
                ld[f"{hc}{a}"] = (
                    f'=IF(AND((F{a}-{t})*(F{b}-{t})<=0,F{a}<>F{b}),C{a}+({t}-F{a})*(C{b}-C{a})/(F{b}-F{a}),"")'
                )
                style_formula(ld[f"{hc}{a}"])
            rng = f"{hc}{rows[0]}:{hc}{rows[-2]}"
            first_r, last_r = rows[0], rows[-1]
            idx = f'INDEX(C{first_r}:C{last_r},MATCH(MIN(F{first_r}:F{last_r}),F{first_r}:F{last_r},0))'
            if spec["unit"] == "bp":
                shock_txt = f'IF({idx}>=0,"+","")&TEXT({idx},"0")&" basis points"'
            else:
                shock_txt = f'IF({idx}>=0,"+","")&TEXT({idx},"0.0%")'
            miss = f'"Not reached: worst is "&TEXT(MIN(F{first_r}:F{last_r}),"0.0%")&" at "&{shock_txt}'
            out = ld.cell(rr, 2 + k, f'=IF(COUNT({rng})=0,{miss},AVERAGE({rng}))')
            out.number_format = spec["fmt"]
            out.border = THIN
            style_formula(out)
    ld.cell(rt + 2 + len(LADDER_TARGETS), 1,
            "'Not reached: worst is -X% at …' means even the largest tested rung does not take value "
            "down by that target; the cell still shows the worst fall on the ladder. scripts/reverse_stress.py "
            "solves the same question exactly (no interpolation) from the Inputs sheet.").font = MUTED

    # ---- Charts: FULL equity vs shock size, one per event ----
    for k, (sheet, label, spec, rows) in enumerate(blocks):
        ch = ScatterChart()
        ch.title = f"{label}: FULL equity vs shock size"
        ch.style = 13
        ch.x_axis.title = f"Shock size ({spec['unit']})"
        ch.y_axis.title = "Equity ($m)"
        xs = Reference(ld, min_col=3, min_row=rows[0], max_row=rows[-1])
        ys = Reference(ld, min_col=4, min_row=rows[0], max_row=rows[-1])
        ser = Series(ys, xs, title="FULL equity")
        ser.marker.symbol = "circle"
        ch.series.append(ser)
        ch.height = 7
        ch.width = 13
        anchor_col = get_column_letter(helper_col0 + len(LADDER_TARGETS) * len(blocks) + 1)
        ld.add_chart(ch, f"{anchor_col}{LADDER_FIRST_BLOCK + 15 * k}")
    set_widths(ld, [30, 14, 12, 18, 14, 12, 12, 14, 16, 4, 4] + [10] * 8)


if __name__ == "__main__":
    sys.exit(main())
