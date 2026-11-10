"""Independent reference DCF + joint event scenarios for checker reconciliation.

Does not read channel rows or FULL formulas — only Inputs (plus the blue / yellow
event inputs on Event_Tariff B21:B29 and Event_Capex B21:B25 when those sheets are
present). Must reproduce base perpetual equity exactly before joint-scenario compares.
Every joint function takes an optional shock so the up/down mirror and the reverse
stress grid can be recomputed independently too.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from openpyxl import load_workbook


@dataclass
class ModelInputs:
    revenue_y0: float
    growth: float
    opex_pct: float
    tax: float
    da_pct: float
    capex_pct: float
    nwc_pct: float
    royalty_rate: float
    cost_elasticity: float
    capex_price_elasticity: float
    commodity_shock: float
    risk_free: float
    erp: float
    beta: float
    pre_tax_kd: float
    debt_weight: float
    equity_weight: float
    rate_shock_bp: float
    wacc_override: float | None
    terminal_g: float
    exit_ebitda_multiple: float
    gross_debt: float
    cash: float
    floating_share: float
    duration: float
    diluted_shares: float
    # tariff bridge (optional)
    tariff_units: float | None = None
    tariff_price: float | None = None
    tariff_unit_cost: float | None = None
    tariff_fixed_cost: float | None = None
    tariff_cost_shock: float | None = None
    tariff_pass_through: float | None = None
    tariff_elasticity: float | None = None
    tariff_markdown: float | None = None
    tariff_mitigation: float | None = None
    # capex-cycle event (optional)
    capex_shock: float | None = None
    capex_rev_elasticity: float | None = None
    capex_own_response: float | None = None
    capex_equity_share: float | None = None
    capex_issue_price: float | None = None


def _num(v: Any, default: float | None = None) -> float | None:
    if v is None or v == "":
        return default
    return float(v)


def load_inputs_from_workbook(path) -> ModelInputs:
    wb = load_workbook(path, data_only=True)
    inp = wb["Inputs"]

    def b(cell: str):
        return inp[cell].value

    override = b("B28")
    # blank override may be None
    wacc_override = None if override is None or override == "" else float(override)

    tu = tp = tuc = tfc = tcs = tpt = te = tmd = tmit = None
    if "Event_Tariff" in wb.sheetnames:
        et = wb["Event_Tariff"]
        tu = _num(et["B21"].value)
        tp = _num(et["B22"].value)
        tuc = _num(et["B23"].value)
        tfc = _num(et["B24"].value)
        tcs = _num(et["B25"].value)
        tpt = _num(et["B26"].value)
        te = _num(et["B27"].value)
        tmd = _num(et["B28"].value)
        tmit = _num(et["B29"].value)
    cxs = cxe = cxk = cxq = cxp = None
    if "Event_Capex" in wb.sheetnames:
        ex = wb["Event_Capex"]
        cxs = _num(ex["B21"].value)
        cxe = _num(ex["B22"].value)
        cxk = _num(ex["B23"].value)
        cxq = _num(ex["B24"].value)
        cxp = _num(ex["B25"].value)

    mi = ModelInputs(
        revenue_y0=float(b("B5")),
        growth=float(b("B6")),
        opex_pct=float(b("B7")),
        tax=float(b("B8")),
        da_pct=float(b("B9")),
        capex_pct=float(b("B10")),
        nwc_pct=float(b("B11")),
        royalty_rate=float(b("B12")),
        cost_elasticity=float(b("B13")),
        capex_price_elasticity=float(b("B14")),
        commodity_shock=float(b("B18")),
        risk_free=float(b("B21")),
        erp=float(b("B22")),
        beta=float(b("B23")),
        pre_tax_kd=float(b("B24")),
        debt_weight=float(b("B25")),
        equity_weight=float(b("B26")),
        rate_shock_bp=float(b("B27")),
        wacc_override=wacc_override,
        terminal_g=float(b("B31")),
        exit_ebitda_multiple=float(b("B32")),
        gross_debt=float(b("B35")),
        cash=float(b("B36")),
        floating_share=float(b("B37")),
        duration=float(b("B38")),
        diluted_shares=float(b("B39")),
        tariff_units=tu,
        tariff_price=tp,
        tariff_unit_cost=tuc,
        tariff_fixed_cost=tfc,
        tariff_cost_shock=tcs,
        tariff_pass_through=tpt,
        tariff_elasticity=te,
        tariff_markdown=tmd,
        tariff_mitigation=tmit,
        capex_shock=cxs,
        capex_rev_elasticity=cxe,
        capex_own_response=cxk,
        capex_equity_share=cxq,
        capex_issue_price=cxp,
    )
    wb.close()
    return mi


def selected_wacc(m: ModelInputs) -> float:
    ke = m.risk_free + m.beta * m.erp
    kd_at = m.pre_tax_kd * (1.0 - m.tax)
    raw = m.equity_weight * ke + m.debt_weight * kd_at
    if m.wacc_override is None:
        return raw
    return float(m.wacc_override)


def net_debt(m: ModelInputs) -> float:
    return m.gross_debt - m.cash


def _fcff_path(
    m: ModelInputs,
    *,
    price_factor: float = 1.0,
    opex_factor: float = 1.0,
    royalty_on_stressed_rev: bool = True,
    royalty_factor_on_base_rev: float | None = None,
    capex_pct_override: float | None = None,
    da_capex_on_base_volume: bool = False,
    nwc_on_stressed: bool = True,
) -> list[float]:
    """Build 5y FCFF list matching the workbook's operating construction."""
    capex_pct = m.capex_pct if capex_pct_override is None else capex_pct_override
    fcffs: list[float] = []
    prev_rev_for_nwc = m.revenue_y0  # Y0
    for t in range(1, 6):
        base_rev = m.revenue_y0 * ((1.0 + m.growth) ** t)
        stressed_rev = base_rev * price_factor
        base_opex = base_rev * m.opex_pct
        opex = base_opex * opex_factor
        if royalty_factor_on_base_rev is not None:
            royalty = base_rev * m.royalty_rate * royalty_factor_on_base_rev
        elif royalty_on_stressed_rev:
            royalty = stressed_rev * m.royalty_rate
        else:
            royalty = base_rev * m.royalty_rate  # base dollars
        ebit = stressed_rev - opex - royalty
        # For channels that hold revenue at base, stressed_rev should be base_rev
        # Caller sets price_factor=1 for those.
        nopat = ebit * (1.0 - m.tax)
        vol_rev = base_rev if da_capex_on_base_volume else stressed_rev
        da = vol_rev * m.da_pct
        capex = -(vol_rev * capex_pct)
        nwc_rev = stressed_rev if nwc_on_stressed else base_rev
        dnwc = -(nwc_rev - prev_rev_for_nwc) * m.nwc_pct
        prev_rev_for_nwc = nwc_rev
        fcffs.append(nopat + da + capex + dnwc)
    return fcffs


def equity_from_fcff(fcffs: list[float], wacc: float, terminal_g: float, nd: float) -> float:
    assert len(fcffs) == 5
    total = 0.0
    for t, f in enumerate(fcffs, start=1):
        total += f / ((1.0 + wacc) ** t)
    tv = fcffs[4] * (1.0 + terminal_g) / (wacc - terminal_g)
    total += tv / ((1.0 + wacc) ** 5)
    return total - nd


def base_fcff(m: ModelInputs) -> list[float]:
    return _fcff_path(
        m,
        price_factor=1.0,
        opex_factor=1.0,
        royalty_on_stressed_rev=True,
        da_capex_on_base_volume=False,
        nwc_on_stressed=True,
    )


def base_equity_perp(m: ModelInputs) -> float:
    w = selected_wacc(m)
    return equity_from_fcff(base_fcff(m), w, m.terminal_g, net_debt(m))


def base_equity_exit(m: ModelInputs) -> float:
    """Exit-multiple equity to support blend (matches DCF sheet)."""
    w = selected_wacc(m)
    fcffs = base_fcff(m)
    # Y5 EBITDA = EBIT + D&A
    base_rev5 = m.revenue_y0 * ((1.0 + m.growth) ** 5)
    opex5 = base_rev5 * m.opex_pct
    roy5 = base_rev5 * m.royalty_rate
    ebit5 = base_rev5 - opex5 - roy5
    da5 = base_rev5 * m.da_pct
    ebitda5 = ebit5 + da5
    # sum PV explicit
    pv_exp = sum(f / ((1.0 + w) ** t) for t, f in enumerate(fcffs, start=1))
    tv = ebitda5 * m.exit_ebitda_multiple
    pv_tv = tv / ((1.0 + w) ** 5)
    return pv_exp + pv_tv - net_debt(m)


def blended_mid(m: ModelInputs) -> float:
    return 0.5 * (base_equity_perp(m) + base_equity_exit(m))


def rates_joint_equity(m: ModelInputs, shock_bp: float | None = None) -> float:
    """FULL = equity @ selected_wacc+Δ with base FCFF, + fixed-debt MTM bridge relief."""
    w0 = selected_wacc(m)
    delta = (m.rate_shock_bp if shock_bp is None else shock_bp) / 10000.0
    w_s = w0 + delta
    eq_wacc = equity_from_fcff(base_fcff(m), w_s, m.terminal_g, net_debt(m))
    fixed_debt = m.gross_debt * (1.0 - m.floating_share)
    mtm = m.duration * fixed_debt * delta  # rates up → debt FV down → equity up
    return eq_wacc + mtm


def commodity_joint_equity(m: ModelInputs, shock: float | None = None) -> float:
    """FULL A+B+C+D: price→rev; opex via cost elasticity; royalty on stressed; adj capex on base vol."""
    shock = m.commodity_shock if shock is None else shock
    price_factor = 1.0 + shock
    opex_factor = 1.0 + m.cost_elasticity * shock
    capex_pct = m.capex_pct * (1.0 + m.capex_price_elasticity * shock)
    fcffs = _fcff_path(
        m,
        price_factor=price_factor,
        opex_factor=opex_factor,
        royalty_on_stressed_rev=True,
        capex_pct_override=capex_pct,
        da_capex_on_base_volume=True,
        nwc_on_stressed=True,
    )
    return equity_from_fcff(fcffs, selected_wacc(m), m.terminal_g, net_debt(m))


def _tariff_delta_nopat_years(
    m: ModelInputs,
    *,
    pass_through: float,
    elasticity: float,
    markdown: float,
    mitigation: float,
    cost_shock: float | None = None,
) -> list[float]:
    assert m.tariff_units is not None
    units0 = m.tariff_units
    price = m.tariff_price
    unit_cost = m.tariff_unit_cost
    cost_shock = m.tariff_cost_shock if cost_shock is None else cost_shock
    deltas = []
    for t in range(1, 6):
        units = units0 * ((1.0 + m.growth) ** t)
        brev = units * price
        bvar = units * unit_cost
        bgp = brev - bvar
        eff = cost_shock * (1.0 - mitigation)
        suc = unit_cost * (1.0 + eff)
        pup = unit_cost * eff * pass_through
        sp = price + pup
        ppct = 0.0 if price == 0 else (sp - price) / price
        su = units * (1.0 + elasticity * ppct)
        srev = su * sp * (1.0 - markdown)
        svar = su * suc
        sgp = srev - svar
        dgp = sgp - bgp
        deltas.append(dgp * (1.0 - m.tax))
    return deltas


def tariff_posture_equity(
    m: ModelInputs, *, pass_through: float, elasticity: float, markdown: float, mitigation: float
) -> float:
    """Equity at an arbitrary tariff posture (used for the mitigation-only interaction term)."""
    deltas = _tariff_delta_nopat_years(
        m, pass_through=pass_through, elasticity=elasticity, markdown=markdown, mitigation=mitigation
    )
    stressed = [b + d for b, d in zip(base_fcff(m), deltas)]
    return equity_from_fcff(stressed, selected_wacc(m), m.terminal_g, net_debt(m))


def tariff_joint_equity(m: ModelInputs, shock: float | None = None) -> float:
    """FULL = PV(base FCFF + joint ΔNOPAT) - net debt at selected WACC."""
    deltas = _tariff_delta_nopat_years(
        m,
        pass_through=float(m.tariff_pass_through),
        elasticity=float(m.tariff_elasticity),
        markdown=float(m.tariff_markdown),
        mitigation=float(m.tariff_mitigation),
        cost_shock=shock,
    )
    base = base_fcff(m)
    stressed = [b + d for b, d in zip(base, deltas)]
    return equity_from_fcff(stressed, selected_wacc(m), m.terminal_g, net_debt(m))


def _capex_fcff(m: ModelInputs, x: float, rev_on: bool, capex_on: bool) -> list[float]:
    """Capex-cycle FCFF: own revenue × (1 + elasticity·x); own capex × (1 + response·x).

    Opex, royalty and ΔNWC follow revenue; D&A and capex stay at base dollars unless capex_on.
    """
    e = float(m.capex_rev_elasticity)
    k = float(m.capex_own_response)
    out: list[float] = []
    prev = m.revenue_y0
    for t in range(1, 6):
        base_rev = m.revenue_y0 * ((1.0 + m.growth) ** t)
        rev = base_rev * (1.0 + e * x) if rev_on else base_rev
        nopat = rev * (1.0 - m.opex_pct - m.royalty_rate) * (1.0 - m.tax)
        da = base_rev * m.da_pct
        capex = base_rev * m.capex_pct * ((1.0 + k * x) if capex_on else 1.0)
        dnwc = -(rev - prev) * m.nwc_pct
        prev = rev
        out.append(nopat + da - capex + dnwc)
    return out


def capex_components(m: ModelInputs, shock: float | None = None) -> dict[str, float]:
    """Standalone A (revenue), B (own capex), C (financing mix) and FULL, per pre-event share basis.

    Funding need = MAX(0, PV of incremental own capex) raised once at t0: the equity part lowers net
    debt and adds new diluted shares at the issue price; the debt part adds debt and cash equally.
    Values are equity attributable to the pre-event diluted shares ($m).
    """
    assert m.capex_rev_elasticity is not None, "Event_Capex inputs missing"
    x = float(m.capex_shock) if shock is None else shock
    w = selected_wacc(m)
    nd = net_debt(m)
    g = m.terminal_g
    n0 = m.diluted_shares
    base = base_equity_perp(m)
    a = equity_from_fcff(_capex_fcff(m, x, True, False), w, g, nd)
    b = equity_from_fcff(_capex_fcff(m, x, False, True), w, g, nd)
    ops = equity_from_fcff(_capex_fcff(m, x, True, True), w, g, nd)
    k = float(m.capex_own_response)
    pv_inc = sum(
        m.revenue_y0 * ((1.0 + m.growth) ** t) * m.capex_pct * k * x / ((1.0 + w) ** t) for t in range(1, 6)
    )
    need = max(0.0, pv_inc)
    e_new = float(m.capex_equity_share) * need
    n_new = e_new / float(m.capex_issue_price)
    scale = n0 / (n0 + n_new)
    c = (base + e_new) * scale
    full = (ops + e_new) * scale
    return {"A": a, "B": b, "C": c, "ops": ops, "full": full, "need": need, "new_shares": n_new}


def capex_joint_equity(m: ModelInputs, shock: float | None = None) -> float:
    return capex_components(m, shock)["full"]


def full_at(m: ModelInputs, sheet: str, shock: float) -> float:
    """Independent FULL equity for `sheet` at an arbitrary shock in that sheet's units (bp for rates)."""
    if sheet == "Event_Rates":
        return rates_joint_equity(m, shock)
    if sheet == "Event_Commodity":
        return commodity_joint_equity(m, shock)
    if sheet == "Event_Tariff":
        return tariff_joint_equity(m, shock)
    if sheet == "Event_Capex":
        return capex_joint_equity(m, shock)
    raise ValueError(sheet)


def configured_shock(m: ModelInputs, sheet: str) -> float:
    return {
        "Event_Rates": m.rate_shock_bp,
        "Event_Commodity": m.commodity_shock,
        "Event_Tariff": m.tariff_cost_shock,
        "Event_Capex": m.capex_shock,
    }[sheet]


def blended_proxy(full_equity: float, m: ModelInputs) -> float:
    base = base_equity_perp(m)
    blend = blended_mid(m)
    return blend * (1.0 + (full_equity - base) / base)


def reconcile_events(path, second_sheet: str, tol_rel: float = 1e-4, tol_abs: float = 0.01):
    """Return (failures, details, independent_fulls dict)."""
    failures: list[str] = []
    details: list[str] = []
    m = load_inputs_from_workbook(path)

    # Base match first (perpetual equity)
    wb = load_workbook(path, data_only=True)
    wb_base = wb["DCF"]["B28"].value
    wb_blend = wb["DCF"]["B30"].value
    wb_rates = wb["Event_Rates"]["D11"].value
    wb_second = wb[second_sheet]["D11"].value
    has_capex = "Event_Capex" in wb.sheetnames
    wb_capex = wb["Event_Capex"]["D11"].value if has_capex else None
    wb_env_lo = wb["Envelope"]["B9"].value
    wb_env_hi = wb["Envelope"]["B10"].value
    wb_env_lo_sh = wb["Envelope"]["C9"].value
    wb_env_hi_sh = wb["Envelope"]["C10"].value
    wb_fb_lo = wb["Football"]["B10"].value
    wb_fb_hi = wb["Football"]["D10"].value
    wb.close()

    ref_base = base_equity_perp(m)
    ref_blend = blended_mid(m)
    details.append(f"Ref base equity={ref_base:.6f} workbook={wb_base}")
    if wb_base is None:
        failures.append("FULL reconciliation: workbook base equity cache missing")
        return failures, details, {}
    if abs(ref_base - float(wb_base)) > max(tol_abs, abs(float(wb_base)) * tol_rel):
        failures.append(
            f"FULL reconciliation: reference base {ref_base} != workbook DCF!B28 {wb_base}"
        )
        return failures, details, {}
    details.append("Reference base equity matches DCF!B28")

    if abs(ref_blend - float(wb_blend)) > max(tol_abs, abs(float(wb_blend)) * tol_rel):
        failures.append(
            f"FULL reconciliation: reference blend {ref_blend} != workbook DCF!B30 {wb_blend}"
        )
    else:
        details.append("Reference blend matches DCF!B30")

    ind: dict[str, float] = {}
    ind["rates"] = rates_joint_equity(m)
    details.append(f"Rates FULL ref={ind['rates']:.6f} workbook={wb_rates}")
    if wb_rates is None or abs(ind["rates"] - float(wb_rates)) > max(
        tol_abs, abs(float(wb_rates)) * tol_rel
    ):
        failures.append(
            f"Rates FULL reconciliation: ref {ind['rates']} != workbook {wb_rates}"
        )
    else:
        details.append("Rates FULL reconciliation OK")

    if second_sheet == "Event_Commodity":
        ind["commodity"] = commodity_joint_equity(m)
        key = "commodity"
    else:
        ind["tariff"] = tariff_joint_equity(m)
        key = "tariff"
    details.append(f"{key} FULL ref={ind[key]:.6f} workbook={wb_second}")
    if wb_second is None or abs(ind[key] - float(wb_second)) > max(
        tol_abs, abs(float(wb_second)) * tol_rel
    ):
        failures.append(
            f"{key} FULL reconciliation: ref {ind[key]} != workbook {wb_second}"
        )
    else:
        details.append(f"{key} FULL reconciliation OK")

    if has_capex:
        ind["capex"] = capex_joint_equity(m)
        details.append(f"capex FULL ref={ind['capex']:.6f} workbook={wb_capex}")
        if wb_capex is None or abs(ind["capex"] - float(wb_capex)) > max(
            tol_abs, abs(float(wb_capex)) * tol_rel
        ):
            failures.append(f"capex FULL reconciliation: ref {ind['capex']} != workbook {wb_capex}")
        else:
            details.append("capex FULL reconciliation OK")

    # Independent blended proxies for envelope (every enabled event)
    proxies = {
        "base": ref_blend,
        "rates": blended_proxy(ind["rates"], m),
        key: blended_proxy(ind[key], m),
    }
    if has_capex:
        proxies["capex"] = blended_proxy(ind["capex"], m)
    ind_env_lo = min(proxies.values())
    ind_env_hi = max(proxies.values())
    details.append(
        f"Envelope ref MIN/MAX={ind_env_lo:.6f}/{ind_env_hi:.6f} "
        f"workbook={wb_env_lo}/{wb_env_hi}"
    )
    if wb_env_lo is None or abs(ind_env_lo - float(wb_env_lo)) > max(
        tol_abs, abs(float(wb_env_lo)) * tol_rel
    ):
        failures.append(
            f"Envelope low reconciliation: ref {ind_env_lo} != workbook {wb_env_lo}"
        )
    if wb_env_hi is None or abs(ind_env_hi - float(wb_env_hi)) > max(
        tol_abs, abs(float(wb_env_hi)) * tol_rel
    ):
        failures.append(
            f"Envelope high reconciliation: ref {ind_env_hi} != workbook {wb_env_hi}"
        )
    if not any("Envelope" in f for f in failures):
        details.append("Envelope MIN/MAX reconciliation OK")

    # Football $/sh band = envelope $/sh
    shares = m.diluted_shares
    ind_lo_sh = ind_env_lo / shares
    ind_hi_sh = ind_env_hi / shares
    if wb_fb_lo is None or abs(ind_lo_sh - float(wb_fb_lo)) > max(1e-6, abs(ind_lo_sh) * tol_rel):
        failures.append(
            f"Football envelope low: ref {ind_lo_sh} != workbook {wb_fb_lo}"
        )
    if wb_fb_hi is None or abs(ind_hi_sh - float(wb_fb_hi)) > max(1e-6, abs(ind_hi_sh) * tol_rel):
        failures.append(
            f"Football envelope high: ref {ind_hi_sh} != workbook {wb_fb_hi}"
        )
    if wb_env_lo_sh is not None and abs(float(wb_env_lo_sh) - float(wb_fb_lo or 0)) > 1e-9:
        failures.append(
            f"Football!B10 {wb_fb_lo} != Envelope!C9 {wb_env_lo_sh}"
        )
    if wb_env_hi_sh is not None and abs(float(wb_env_hi_sh) - float(wb_fb_hi or 0)) > 1e-9:
        failures.append(
            f"Football!D10 {wb_fb_hi} != Envelope!C10 {wb_env_hi_sh}"
        )
    if not any("Football" in f for f in failures):
        details.append("Football envelope band matches envelope $/sh OK")

    ind["_proxies"] = proxies  # type: ignore
    return failures, details, ind


def format_not_reached(pct: float, shock: float, unit: str) -> str:
    """Label when no tested shock breaks the threshold. Matches the Excel TEXT() result.

    Example: 'Not reached: worst is -18.2% at +200 basis points'.
    """
    pct_txt = f"{pct:.1%}"
    if unit == "bp":
        shock_txt = f"{shock:+.0f} basis points"
    else:
        shock_txt = f"{shock:+.1%}"
    return f"Not reached: worst is {pct_txt} at {shock_txt}"


# ---------------------------------------------------------------------------
# Shock ladder + reverse stress test (independent of the Ladder sheet's formulas)
# Layout must match generate.py LADDER_* constants.
# ---------------------------------------------------------------------------
LADDER_FIRST_BLOCK = 6
LADDER_BLOCK_ROWS = 12
LADDER_N_RUNGS = 7
LADDER_N_TARGETS = 2


def ladder_layout(n_events: int) -> dict:
    blocks = []
    for k in range(n_events):
        s = LADDER_FIRST_BLOCK + LADDER_BLOCK_ROWS * k
        blocks.append({"title": s, "rows": list(range(s + 2, s + 2 + LADDER_N_RUNGS))})
    rev_title = LADDER_FIRST_BLOCK + LADDER_BLOCK_ROWS * n_events + 1
    return {"blocks": blocks, "rev_title": rev_title, "rev_hdr": rev_title + 1,
            "rev_rows": list(range(rev_title + 2, rev_title + 2 + LADDER_N_TARGETS))}


def event_sheets(sheetnames) -> list[str]:
    return [s for s in sheetnames if s in ("Event_Rates", "Event_Commodity", "Event_Tariff", "Event_Capex")]


def interpolate_reverse(grid: list[float], pct: list[float], target: float):
    """Same rule as the Ladder sheet: average of crossings over adjacent rungs (rungs sorted by multiple)."""
    hits = []
    for (x0, y0), (x1, y1) in zip(zip(grid, pct), zip(grid[1:], pct[1:])):
        if (y0 - target) * (y1 - target) <= 0 and y0 != y1:
            hits.append(x0 + (target - y0) * (x1 - x0) / (y1 - y0))
    return None if not hits else sum(hits) / len(hits)


def solve_shock_for_drop(f, base: float, target: float, lo: float, hi: float, tol: float = 1e-10):
    """Bisection for shock s in [lo, hi] with f(s)/base - 1 = target. None if not bracketed."""
    g = lambda s: f(s) / base - 1.0 - target  # noqa: E731
    glo, ghi = g(lo), g(hi)
    if glo * ghi > 0:
        return None
    for _ in range(300):
        mid = 0.5 * (lo + hi)
        gm = g(mid)
        if abs(gm) < tol:
            return mid
        if glo * gm <= 0:
            hi = mid
        else:
            lo, glo = mid, gm
    return 0.5 * (lo + hi)


def reconcile_ladder(path, tol_rel: float = 1e-4, tol_abs: float = 0.01):
    """Rebuild every ladder rung, asymmetry and reverse-stress interpolation from Inputs + grid inputs."""
    failures: list[str] = []
    details: list[str] = []
    wb = load_workbook(path, data_only=True)
    if "Ladder" not in wb.sheetnames:
        wb.close()
        return ["Ladder sheet missing"], details
    ld = wb["Ladder"]
    sheets = event_sheets(wb.sheetnames)
    mirror_asym = {sh: wb[sh]["G63"].value for sh in sheets}
    wb.close()
    m = load_inputs_from_workbook(path)
    base = base_equity_perp(m)
    blend = blended_mid(m)
    blend_sh = blend / m.diluted_shares
    lay = ladder_layout(len(sheets))
    pct_by_event = {}
    for k, (sh, blk) in enumerate(zip(sheets, lay["blocks"])):
        rows = blk["rows"]
        conf = configured_shock(m, sh)
        mults, shocks, pct, sh_vals = [], [], [], []
        for r in rows:
            mult = ld.cell(r, 2).value
            if mult is None:
                failures.append(f"Ladder {sh} row {r}: missing multiple")
                continue
            x = conf * float(mult)
            ref = full_at(m, sh, x)
            got_c, got_d, got_f, got_g = (ld.cell(r, c).value for c in (3, 4, 6, 7))
            if got_c is None or abs(float(got_c) - x) > 1e-9 * max(1.0, abs(x)):
                failures.append(f"Ladder {sh} row {r}: shock {got_c} != configured x multiple {x}")
            if got_d is None or abs(ref - float(got_d)) > max(tol_abs, abs(ref) * tol_rel):
                failures.append(f"Ladder {sh} row {r} (x{mult}): FULL ref {ref:.4f} != workbook {got_d}")
            p = ref / base - 1.0
            if got_f is None or abs(float(got_f) - p) > 1e-6:
                failures.append(f"Ladder {sh} row {r}: % change {got_f} != ref {p:.6f}")
            ps = blend_sh * (1.0 + p)
            if got_g is None or abs(float(got_g) - ps) > max(1e-4, abs(ps) * tol_rel):
                failures.append(f"Ladder {sh} row {r}: proxy $/sh {got_g} != ref {ps:.6f}")
            mults.append(float(mult)); shocks.append(x); pct.append(p); sh_vals.append(ps)
        n = len(rows)
        if len(mults) == n:
            if mults != sorted(mults) or any(abs(mults[i] + mults[n - 1 - i]) > 1e-12 for i in range(n)):
                failures.append(f"Ladder {sh}: multiples must be ascending and symmetric around 0, got {mults}")
            for i, r in enumerate(rows):
                ref_a = (sh_vals[i] - blend_sh) + (sh_vals[n - 1 - i] - blend_sh)
                got = ld.cell(r, 8).value
                if got is None or abs(float(got) - ref_a) > max(1e-4, abs(ref_a) * tol_rel):
                    failures.append(f"Ladder {sh} row {r}: asymmetry $/sh {got} != ref {ref_a:.6f}")
            # acceptance: the +/-1x rung reproduces the event sheet's mirror asymmetry (and Summary)
            if 1.0 in mults:
                i1 = mults.index(1.0)
                a1 = (sh_vals[i1] - blend_sh) + (sh_vals[n - 1 - i1] - blend_sh)
                ma = mirror_asym.get(sh)
                if ma is None or abs(float(ma) - a1) > max(1e-4, abs(a1) * tol_rel):
                    failures.append(f"Ladder {sh}: +/-1x asymmetry {a1:.6f} != {sh}!G63 mirror asymmetry {ma}")
            else:
                failures.append(f"Ladder {sh}: no +/-1x rung to tie to the mirror asymmetry")
            pct_by_event[sh] = (shocks, pct)
    for ti, rr in enumerate(lay["rev_rows"]):
        wb2 = load_workbook(path, data_only=True)
        ldv = wb2["Ladder"]
        tgt = ldv.cell(rr, 1).value
        for k, sh in enumerate(sheets):
            got = ldv.cell(rr, 2 + k).value
            if tgt is None or sh not in pct_by_event:
                continue
            shocks, pct = pct_by_event[sh]
            ref = interpolate_reverse(shocks if shocks[0] <= shocks[-1] else shocks, pct, float(tgt))
            if ref is None:
                unit = "bp" if sh == "Event_Rates" else "%"
                wi = min(range(len(pct)), key=lambda i: pct[i])
                exp = format_not_reached(pct[wi], shocks[wi], unit)
                if isinstance(got, (int, float)):
                    failures.append(f"Ladder reverse {sh} target {tgt}: workbook {got} but ref is outside the ladder")
                elif got != exp:
                    failures.append(f"Ladder reverse {sh} target {tgt}: expected {exp!r}, got {got!r}")
            elif not isinstance(got, (int, float)) or abs(ref - float(got)) > max(1e-6, abs(ref) * tol_rel):
                failures.append(f"Ladder reverse {sh} target {tgt}: ref {ref} != workbook {got}")
        wb2.close()
    if not failures:
        details.append("Ladder rungs, asymmetry (tied to each mirror at +/-1x) and reverse-stress interpolation reconcile to the independent reference")
    return failures, details
