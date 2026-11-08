# Worked examples (illustrative only)

**Draft, not investment advice.** Both cases below use invented names and round ILLUSTRATIVE numbers. They are teaching walkthroughs of which shocks to run and which channels each shock moves. They are not valuations of any real issuer.

Terms: DCF = discounted cash flow. WACC = weighted average cost of capital. FCFF = free cash flow to the firm. EV = enterprise value. EPS = earnings per share. FX = foreign exchange. bp = basis points. FULL = all channels of one event applied together.

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

**Setup (illustrative).** Revenue $800 million; gross margin 55%; store occupancy largely fixed; online mix 35%; net debt $80 million; diluted shares 40 million; WACC 10%; terminal growth 2.5%; exit multiple 9× EBITDA.

**Shocks run (each in its own table):**

1. **Tariffs / input costs (+15% unit product cost, with partial pass-through)**
   - Explicit **units × price × unit cost** bridge; fixed occupancy kept on a separate line.
   - **A — Unit cost:** landed cost +15%.
   - **B — Price pass-through:** management recovery 60% of the unit-cost hit in shelf price — counted **once** (no second add-back in margin).
   - **C — Demand elasticity / markdowns:** volume −3% from higher shelf prices; markdown rate +1 percentage point.
   - **D — Inventory / working capital:** inventory days +5 on the transition.
   - **E — Sourcing mitigation:** 25% of volume shifted to an alternate country over two years (partial offset).
   - **FULL:** all of A–E. Pass-through invariant checked: at 100% recovery and zero elasticity/markdowns/mitigation, gross-profit dollars unchanged.

2. **Consumer demand / comps (−5% comparable sales)**
   - **A — Volume / comps:** −5% transactions.
   - **B — Price / mix:** mix shift toward promotional tiers (−50 bp margin).
   - **C — Markdowns:** +150 bp of sales.
   - **D — Inventory:** clearance raises short-term cash then normalises.
   - **E — Store opex:** fixed occupancy largely unchanged (highlighting operating leverage).
   - **FULL:** combined demand table; not stacked with the tariff table.

3. **Interest rates ±100 bp** (lighter retail debt book)
   - Channels A–E as in the miner case, but refinancing and mark-to-market channels are small; buyback capacity matters more because the base plan included modest buybacks.

4. **Capex cycle (optional):** mall remodel wave — own revenue lift from refreshed stores, own capex spike, FCFF trough, and a temporary shift toward debt funding in the bridge (equity vs debt mix). Own-base %Δ only.

**Checks emphasised.** No double-counted tariff recovery; labels’ years matched the first affected forecast column; EV for any peer comps recomputed as market cap + debt − cash rather than a vendor EV field.

---

## How to use these examples with the template

Generate with `python scripts/generate.py --config examples/configs/gold_miner.yaml` (or `apparel_retailer.yaml`), recalculate, and run `python scripts/check_workbook.py`. Or copy [templates/stress_template.xlsx](../templates/stress_template.xlsx). Gate with [REVIEW_CHECKLIST.md](../REVIEW_CHECKLIST.md) and write a CLEAR or BLOCK verdict.


Shipped recalculated workbooks: [workbooks/gold_miner.xlsx](workbooks/gold_miner.xlsx), [workbooks/apparel_retailer.xlsx](workbooks/apparel_retailer.xlsx).
