"""Independent reference DCF + joint event scenarios for checker reconciliation.

Does not read channel rows or FULL formulas — only Inputs (and Event_Tariff bridge
inputs when present). Must reproduce base perpetual equity exactly before
joint-scenario compares.
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

    tariff = None
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


def rates_joint_equity(m: ModelInputs) -> float:
    """FULL = equity @ selected_wacc+Δ with base FCFF, + fixed-debt MTM bridge relief."""
    w0 = selected_wacc(m)
    delta = m.rate_shock_bp / 10000.0
    w_s = w0 + delta
    eq_wacc = equity_from_fcff(base_fcff(m), w_s, m.terminal_g, net_debt(m))
    fixed_debt = m.gross_debt * (1.0 - m.floating_share)
    mtm = m.duration * fixed_debt * delta  # rates up → debt FV down → equity up
    return eq_wacc + mtm


def commodity_joint_equity(m: ModelInputs) -> float:
    """FULL A+B+C+D: price→rev; opex via cost elasticity; royalty on stressed; adj capex on base vol."""
    shock = m.commodity_shock
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
) -> list[float]:
    assert m.tariff_units is not None
    units0 = m.tariff_units
    price = m.tariff_price
    unit_cost = m.tariff_unit_cost
    cost_shock = m.tariff_cost_shock
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


def tariff_joint_equity(m: ModelInputs) -> float:
    """FULL = PV(base FCFF + joint ΔNOPAT) - net debt at selected WACC."""
    deltas = _tariff_delta_nopat_years(
        m,
        pass_through=float(m.tariff_pass_through),
        elasticity=float(m.tariff_elasticity),
        markdown=float(m.tariff_markdown),
        mitigation=float(m.tariff_mitigation),
    )
    base = base_fcff(m)
    stressed = [b + d for b, d in zip(base, deltas)]
    return equity_from_fcff(stressed, selected_wacc(m), m.terminal_g, net_debt(m))


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

    # Independent blended proxies for envelope
    proxies = {
        "base": ref_blend,
        "rates": blended_proxy(ind["rates"], m),
        key: blended_proxy(ind[key], m),
    }
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
