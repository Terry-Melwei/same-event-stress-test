# Worked examples (illustrative only)

**Draft, not investment advice.** All three cases below use invented names and round ILLUSTRATIVE numbers. They are teaching walkthroughs of which shocks to run and which channels each shock moves. They are not valuations of any real issuer.

Terms: DCF = discounted cash flow. Capex = capital expenditure. PV = present value. INT = cross-channel interaction row (not size non-linearity). WACC = weighted average cost of capital. FCFF = free cash flow to the firm. EV = enterprise value. EPS = earnings per share. FX = foreign exchange. bp = basis points. FULL = all channels of one event applied together. EBIT = earnings before interest and tax. D&A = depreciation and amortisation. EBITDA = earnings before interest, tax, depreciation and amortisation (in this model, EBIT + D&A). opex = operating expenses. %Δ = percentage change. $/sh = US dollars per diluted share.

---

## Case 1 — Illustrative gold miner (“North Ridge Gold”, ILLUSTRATIVE)

**Setup (illustrative).** Long-term gold price assumption $1,800/oz; annual production held flat in the base; royalty 5% of revenue; sustaining capex tied to reserve replacement; net debt $250 million; diluted shares 120 million; WACC 9%; terminal growth 2%; exit multiple 7× earnings before interest, tax, depreciation and amortisation (EBITDA).

**Shocks run (each in its own table):**

1. **Interest rates ±100 bp**
   - **A — WACC:** risk-free +100 bp lifts cost of equity and after-tax cost of debt; illustrative base equity value $12.00/sh → FULL down-case about $10.40/sh before other channels.
   - **B — Refinancing at maturities:** floating and near-maturity debt (~30% of gross debt) re-prices; annual interest cash up ~$4 million in the adverse case.
   - **C — Fixed-debt mark-to-market:** fixed-rate book debt fair-value decline reduces economic net debt in the bridge (value *up* slightly via the bridge only — not via FCFF).
   - **D — Profit / EPS view:** higher interest expense cuts EPS; used for communication, not double-counted into FCFF.
   - **E — Buyback capacity:** higher cash interest trims buyback cash; diluted share path unchanged unless buybacks were in the base.
   - **FULL:** cumulative A + C only (WACC + mark-to-market bridge). B/D/E are information rows outside FULL so interest cash is not subtracted twice. Own-base %Δ → blended-mid sensitivity proxy on the tornado.

2. **Commodity price ±20%** (gold price to $1,440 / $2,160)
   - **A — Revenue:** price × flat ounces.
   - **B — Royalties / tax:** royalty dollars move with revenue; mineral tax bracket checked so the shocked rate is **not** reused to re-fit historical cost (avoids the royalty sign bug).
   - **C — Cost co-move:** ~20% of operating costs assumed to move 0.3× with gold (energy, contractors).
   - **D — Capex / reserve pace:** adverse price defers one development tranche; sustaining capex eases down slightly with weaker cash.
   - **E — Net debt / dividends:** weaker free cash raises net debt and cuts the dividend assumption in the bridge.
   - **FULL:** cumulative A+B+C+D (E diagnostic bridge excluded). Price moves revenue only; opex moves only via cost elasticity. Illustrative adverse FULL equity value below base on the blended-mid proxy (teaching scale only).

3. **Jurisdiction (optional third table):** one named event — “Permit pause at Red Gulch pit for 18 months” — driving outage volume, a temporary fiscal royalty uplift, and a +50 bp country risk add-on in WACC. Not stacked into the gold-price table.

**Channels deliberately not mixed.** FX and tariff tables were not run for this case. Financing never entered FCFF; interest and debt mark-to-market sat in the net-debt bridge / WACC only.

---

## Case 2 — Illustrative specialty apparel retailer (“Harbor Lane Apparel”, ILLUSTRATIVE)

**Setup (illustrative).** Revenue $800 million; gross margin 55%; store occupancy largely fixed; online mix 35%; net debt $80 million; diluted shares 40 million; WACC about 8% (no override); terminal growth 2.5%; exit multiple 9× EBITDA. Config: [configs/apparel_retailer.yaml](configs/apparel_retailer.yaml). The workbook also has a blue net debt / Year-1 EBITDA covenant limit of 0.35×, labelled ILLUSTRATIVE, replace with your own loan limit. A real covenant limit comes from the company's loan agreement; the yellow source cell ships as `TO FILL: cite the loan agreement clause`. Covenant flags show BREACH in red and PASS in light green (live conditional formatting).

**Shocks run (each in its own table):**

1. **Interest rates ±100 bp** (lighter retail debt book)
   - Channels A–E as in the miner case, but refinancing and mark-to-market channels are small.

2. **Tariffs / input costs (+15% unit product cost, with partial pass-through)**
   - Explicit **units × price × unit cost** bridge; fixed occupancy kept on a separate line.
   - **A — Unit cost:** landed cost +15%.
   - **B — Price pass-through:** management recovery 60% of the unit-cost hit in shelf price — counted **once** (no second add-back in margin).
   - **C — Demand elasticity / markdowns:** volume response and markdowns as in the config.
   - **D — Inventory / working capital** and **E — Sourcing mitigation** complete the bridge.
   - **FULL:** all of A–E. Pass-through invariant checked: at 100% recovery and zero elasticity/markdowns/mitigation, gross-profit dollars unchanged.

3. **Foreign exchange (FX), home-currency appreciation +10%** (enabled in this workbook)
   - **A — Translation of foreign revenue** (foreign-currency revenue share 40%).
   - **B — Foreign-currency cost offset (natural hedge, 15% of opex) + financial hedge** (yellow hedge ratio 25% of foreign revenue), counted once.
   - **C — Foreign-currency debt remeasurement** (30% of gross debt) once in the net-debt bridge, not inside FCFF.
   - **FULL:** A+B+C. This event is linear in the shock by construction, so the +10% and −10% cases mirror; do not claim FX non-linearity. Cash is not remeasured.
   - Illustrative order of magnitude in the shipped workbook (teaching scale only, not a forecast): FULL at +10% home-currency appreciation is about −7% on the blended-mid proxy; the covenant ratio stays below the 0.35× limit (PASS).

4. **Consumer demand / volume −15%** (enabled in this workbook)
   - **A — Volume** on revenue only (opex at base).
   - **B — Price / mix elasticity** (yellow, 0.20).
   - **C — Operating leverage** (yellow fixed-cost share of opex 40%).
   - **D — Year-1 inventory cash use** that applies only when volume falls (inside FCFF, not a t0 net-debt restatement).
   - **FULL:** revenue = base × (1 + volume) × (1 + elasticity × volume), leveraged opex, inventory overlay. Not stacked with the tariff or FX tables.
   - Honest limit: fixed-cost leverage alone does not make dollar FCFF non-linear in volume; the shipped workbook's mirror asymmetry comes from the revenue product and the downside-only inventory cash use.
   - Illustrative order of magnitude in the shipped workbook (teaching scale only): FULL at −15% volume is about −41% on the blended-mid proxy, and Year-1 net debt / EBITDA rises above the 0.35× limit (BREACH). The +15% mirror does not have the inventory cash use.

5. **Capex cycle (optional, not built in the apparel workbook):** mall remodel wave — own revenue lift from refreshed stores, own capex spike, FCFF trough, and a temporary shift toward debt funding in the bridge (equity vs debt mix). Own-base %Δ only. Case 3 shows the generated `Event_Capex` sheet.

**Checks emphasised.** No double-counted tariff recovery; FX financing counted once in the bridge; demand inventory overlay inside FCFF only; covenant flags recomputed independently; labels’ years matched the first affected forecast column; EV for any peer comps recomputed as market cap + debt − cash rather than a vendor EV field. Every shock-size calibration cell ships as `TO FILL`.

---

## Case 3 — Illustrative optical-components supplier (“Lumen Bay Photonics”, ILLUSTRATIVE)

**Setup (illustrative).** Revenue $1,500 million growing 8%; operating costs 82% of revenue; capex 10% of revenue; net debt $350 million; diluted shares 100 million; beta 1.3; terminal growth 3%; exit multiple 12× EBITDA. Config: [configs/optical_supplier.yaml](configs/optical_supplier.yaml).

**Shocks run (each in its own table):**

1. **Interest rates ±100 bp** — as in Case 1.
2. **Component tariff +15%** — units × price × unit cost bridge with 50% pass-through, −0.2 demand elasticity and 20% sourcing mitigation. The INT row is non-zero here: mitigation shrinks the shock that pass-through and elasticity respond to.
3. **Customer capex +30%** (for example a data-centre build-out), entered as the adverse case:
   - **A — Own revenue:** +0.6 × 30% = +18% (yellow elasticity); opex, royalty and net working capital follow revenue.
   - **B — Own capex:** +1.2 × 30% = +36% (yellow response) — the FCFF channel.
   - **C — Financing mix:** the PV of the extra capex is raised once: 50% equity at a yellow $18.00 issue price (below value, so dilutive) lowers net debt and adds diluted shares in the same step; 50% debt adds debt and cash equally.
   - **D / E — information:** operating-only FCFF view and an all-equity funding view, both outside FULL.
   - **FULL:** A+B+C per pre-event diluted share. The mirror (customer capex −30%) needs no financing, so the two directions are not symmetric.
   - **Reverse stress:** an ILLUSTRATIVE typed $25.00 threshold overrides the default (base × (1 − 25%)) on this sheet; the grid reports the smallest tested capex shock that takes the FULL proxy below it.

Every shock-size calibration cell ships as `TO FILL`. No number in this case is sourced from a real company or vendor.

---

## How to use these examples with the template

Generate with `python scripts/generate.py --config examples/configs/gold_miner.yaml` (or `apparel_retailer.yaml`, `optical_supplier.yaml`), recalculate, and run `python scripts/check_workbook.py`. Or copy [templates/stress_template.xlsx](../templates/stress_template.xlsx). Gate with [REVIEW_CHECKLIST.md](../REVIEW_CHECKLIST.md) and write a CLEAR or BLOCK verdict.


Shipped recalculated workbooks: [workbooks/gold_miner.xlsx](workbooks/gold_miner.xlsx), [workbooks/apparel_retailer.xlsx](workbooks/apparel_retailer.xlsx), [workbooks/optical_supplier.xlsx](workbooks/optical_supplier.xlsx).
