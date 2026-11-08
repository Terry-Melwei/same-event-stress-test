# Review checklist: same-event multi-channel stress tests

Use this checklist to review (gate) a valuation workbook before anyone relies on it.
The verdict is **CLEAR** only if every Blocking check passes. Otherwise it is **BLOCK**, with a numbered list of fixes that cites cells.
CLEAR means the workbook is built correctly. It does not mean the valuation assumptions are endorsed.

Terms: DCF = discounted cash flow. WACC = weighted average cost of capital. FCFF = free cash flow to firm. EV = enterprise value. GP = gross profit. FULL row = the result of one event with all of its channels applied together.

## 1. Live wiring (Blocking)
1. **Live-input test.**
   - Copy the workbook. Change one core input (for example WACC +100 basis points) and recalculate (LibreOffice headless works).
   - Every output that should depend on the changed input must move: FULL values, tornado endpoints, envelope bounds and football-field stress-band cells. Outputs of events that don't use that input may stay unchanged. Say which outputs you expected to move.
   - Repeat with one operating input (for example gross margin, or the long-term commodity price).
2. **No pasted outputs.** Scan the scenario, FULL, summary, sensitivity-grid and chart-source ranges for typed-in numbers.
   - A "cached engine" of values counts as a defect even if it is documented.
   - Static cross-checks are allowed only when they are labelled, flagged STALE when inputs change, and do not feed any valuation output.
3. **Blend lock.** The blended midpoint averages the valuation-method rows only.
   - Stress bands, combined rows and reference rows (stale quotes, raw-WACC views) stay outside the blend.
4. **Cached values present.** Save after recalculating, so that formula results exist in the file. Zero formula errors.

## 2. Event design (Blocking)
5. **One event per table.** Each table has single-channel rows (A, B, C…) and a FULL row that applies all of them at once. Never stack separate events.
6. **Every channel present.** Minimum channels by event type:
   - **Interest rates:** WACC (cost of debt and cost of equity), interest on floating debt and cash, refinancing / mark-to-market of fixed debt, buyback or dividend capacity.
   - **Commodity price:** revenue, royalties and taxes, cost co-movement, capex/growth pace, financing and dividend capacity.
   - **Tariffs / input costs:** unit cost, price pass-through, demand elasticity, markdowns/promotions, inventory and working capital, sourcing mitigation.
   - **Foreign exchange (FX):** revenue translation, cost translation, foreign-currency debt.
   - **Customer / industry capex:** own revenue, own capex, equity vs debt financing mix.
   - **Jurisdiction:** one named event (for example a specific mine or country) that drives the outage, fiscal and risk-premium channels together.
   - Diagnostic rows (for example a standalone bridge view of a channel already counted in the FULL row) are allowed, but they must be labelled as diagnostic and kept out of the FULL row, so nothing is counted twice.
7. **The label matches the code.** The years, segments and weights stated in each row label are the ones the engine actually applies. Check the first affected year column.

## 3. Economic invariants (Blocking)
8. **Pass-through invariant.** Set 100% cost recovery and zero elasticity / markdowns / mitigation. Gross-profit dollars must stay unchanged in every forecast year.
   - Build an explicit units × price × unit cost bridge, with fixed costs kept separate.
9. **Sign tests.** An isolated adverse input must lower value: a higher royalty, a higher tariff, a higher cost.
   - Historical calibration inputs must be separate from forward-looking shock inputs, so a shock never re-fits history.
10. **No financing double count.** EV uses operating cash flows only; no interest goes into FCFF.
    - Financing effects enter the net-debt bridge or the WACC weights once.
    - Share counts change once, alongside the matching cash.
11. **No period overlap.** Stub-period cash, the balance-sheet date and the first explicit forecast year don't overlap.
12. **Lease consistency.** If lease liabilities are in net debt, use EBITDAR-based cash flows (earnings before interest, tax, depreciation, amortisation and rent). Otherwise leave leases out of both. Never subtract them twice.

## 4. Consistent bases (Blocking)
13. Measure each FULL % change against its own method's base (an exit-multiple shock against the exit-method value, the rest against perpetual growth). Then apply that % change to the blended midpoint.
    - Label the result a "blended-mid sensitivity proxy".
    - Never divide one method's result by another method's base.

## 5. Market data and comps (Blocking if material)
14. Use a live price with its date and source. Confirm the listing actually trades; otherwise value the primary listing and show any stale secondary quote as a reference only.
15. Recompute peer EV as market cap + debt − cash from raw inputs. Spot-check 2–3 peer multiples against an outside source, and investigate any gap above about 20%.

## 6. Disclosure (Blocking for big judgment calls)
16. Any judgment that changes value by more than about 10% (for example a chosen WACC instead of the raw CAPM, capital asset pricing model, rate) sits in a yellow cell. Show the raw-input value live on the Cover and Football, outside the blend.
17. List data gaps as gaps. Never fill them with invented numbers.

## 7. Presentation (fix before CLEAR)
18. Colours: inputs in blue font, formulas in black, judgment cells with a yellow fill.
19. A How_Measured sheet explains every method and every stress channel. Spell out each acronym on first use on each sheet.
20. A real football-field chart and a tornado chart, both driven by formula ranges.
21. A "Draft, not investment advice" label on the Cover.

## Verdict template
```
VERDICT: CLEAR | BLOCK
Recomputed vs claimed: <method values, blended mid, 2+ FULL values>
Live-input test: <input changed> -> <outputs moved: n/n>
Fixes (BLOCK only): 1. <issue> - <cell refs> - <required change>
Non-blocking notes: ...
```
