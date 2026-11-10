# Review checklist: same-event multi-channel stress tests

Use this checklist to review (gate) a valuation workbook before anyone relies on it.
The verdict is **CLEAR** only if every Blocking check passes. Otherwise it is **BLOCK**, with a numbered list of fixes that cites cells.
CLEAR means the workbook is built correctly. It does not mean the valuation assumptions are endorsed.

Terms: DCF = discounted cash flow. WACC = weighted average cost of capital. FCFF = free cash flow to firm. EV = enterprise value. GP = gross profit. Capex = capital expenditure. PV = present value. $/sh = US dollars per diluted share. %Δ = percentage change. FULL row = the result of one event with all of its channels applied together. INT row = the interaction diagnostic (FULL minus the standalone channel rows).

## 1. Live wiring (Blocking)
1. **Live-input test.**
   - Copy the workbook. Change one core input (for example WACC +100 basis points) and recalculate (LibreOffice headless works).
   - Every output that should depend on the changed input must move: FULL values, tornado endpoints, envelope bounds and football-field stress-band cells. Outputs of events that don't use that input may stay unchanged. Say which outputs you expected to move.
   - Repeat with one operating input (for example gross margin, or the long-term commodity price).
2. **No pasted outputs.** Scan the scenario, FULL, interaction, mirror, reverse-stress grid and result, Summary, sensitivity-grid and chart-source ranges for typed-in numbers.
   - A "cached engine" of values counts as a defect even if it is documented.
   - Static cross-checks are allowed only when they are labelled, flagged STALE when inputs change, and do not feed any valuation output.
3. **Blend lock.** The blended midpoint averages the valuation-method rows only.
   - Stress bands, combined rows, interaction rows, mirror and reverse-stress results, the Summary sheet and reference rows (stale quotes, raw-WACC views) stay outside the blend.
4. **Cached values present.** Save after recalculating, so that formula results exist in the file. Zero formula errors.

## 2. Event design (Blocking)
5. **One event per table.** Each table has single-channel rows (A, B, C…) and a FULL row that applies all of them at once. Never stack separate events.
6. **Every channel present.** Minimum channels by event type:
   - **Interest rates:** WACC (cost of debt and cost of equity), interest on floating debt and cash, refinancing / mark-to-market of fixed debt, buyback or dividend capacity.
   - **Commodity price:** revenue, royalties and taxes, cost co-movement, capex/growth pace, financing and dividend capacity.
   - **Tariffs / input costs:** unit cost, price pass-through, demand elasticity, markdowns/promotions, inventory and working capital, sourcing mitigation.
   - **Foreign exchange (FX):** revenue translation, cost translation, foreign-currency debt.
   - **Capex cycle (customer or own):** own revenue, own capex (the FCFF channel), and the equity vs debt financing mix in the net-debt / diluted-share bridge. The funding need is financed once: the equity part lowers net debt and adds diluted shares in the same step; the debt part adds debt and cash equally. FULL is measured per pre-event diluted share.
   - **Jurisdiction:** one named event (for example a specific mine or country) that drives the outage, fiscal and risk-premium channels together.
   - Diagnostic rows (for example a standalone bridge view of a channel already counted in the FULL row) are allowed, but they must be labelled as diagnostic and kept out of the FULL row, so nothing is counted twice.
7. **The label matches the code.** The years, segments and weights stated in each row label are the ones the engine actually applies. Check the first affected year column.
8. **Interaction-row integrity.** Each event table has an INT row, labelled diagnostic, that is a live formula equal to (FULL − base) − Σ(standalone row − base).
   - It sums standalone rows only: never information or diagnostic rows. On a sequential ladder (each row adds one response to the previous row, as on the tariff sheet) it isolates the remaining separable channel; say so on the sheet.
   - It is outside FULL and outside the blend, and is not pasted. Recompute it from the channel rows and compare.
9. **Both directions and reverse stress.** Each event shows FULL at the configured shock and at the mirrored shock (opposite sign) with an asymmetry line. Each event has a formula reverse-stress grid (ascending multiples) whose result cell reports the smallest tested shock that takes the FULL $/sh below the yellow threshold, or `n/a` when the threshold is blank.
   - Test it: set a threshold between two grid steps, recalculate, and confirm the result is the first step below the threshold while the previous step is at or above it.
10. **Calibration source.** Every shock size has a yellow calibration **source** (data set or document, with date) and **basis** (for example a historical 1-year 95th percentile move).
    - `TO FILL` is acceptable only in an ILLUSTRATIVE template. A real-issuer workbook with `TO FILL`, or with a statistic that has no cited source, is BLOCK.

## 3. Economic invariants (Blocking)
11. **Pass-through invariant.** Set 100% cost recovery and zero elasticity / markdowns / mitigation. Gross-profit dollars must stay unchanged in every forecast year.
   - Build an explicit units × price × unit cost bridge, with fixed costs kept separate.
12. **Sign tests.** An isolated adverse input must lower value: a higher royalty, a higher tariff, a higher cost.
   - Historical calibration inputs must be separate from forward-looking shock inputs, so a shock never re-fits history.
13. **No financing double count.** EV uses operating cash flows only; no interest goes into FCFF.
    - Financing effects enter the net-debt bridge or the WACC weights once.
    - Share counts change once, alongside the matching cash.
    - Capex-cycle funding: the capex outflow is already in FCFF, so only the equity-funded part moves net debt (paired with the new diluted shares); the debt-funded part adds debt and cash equally. Never add the full funding need to net debt as well.
14. **No period overlap.** Stub-period cash, the balance-sheet date and the first explicit forecast year don't overlap.
15. **Lease consistency.** If lease liabilities are in net debt, use EBITDAR-based cash flows (earnings before interest, tax, depreciation, amortisation and rent). Otherwise leave leases out of both. Never subtract them twice.

## 4. Consistent bases (Blocking)
16. Measure each FULL % change against its own method's base (an exit-multiple shock against the exit-method value, the rest against perpetual growth). Then apply that % change to the blended midpoint.
    - Label the result a "blended-mid sensitivity proxy".
    - Never divide one method's result by another method's base.

## 5. Market data and comps (Blocking if material)
17. Use a live price with its date and source. Confirm the listing actually trades; otherwise value the primary listing and show any stale secondary quote as a reference only.
18. Recompute peer EV as market cap + debt − cash from raw inputs. Spot-check 2–3 peer multiples against an outside source, and investigate any gap above about 20%.

## 6. Disclosure (Blocking for big judgment calls)
19. Any judgment that changes value by more than about 10% (for example a chosen WACC instead of the raw CAPM, capital asset pricing model, rate) sits in a yellow cell. Show the raw-input value live on the Cover and Football, outside the blend.
20. List data gaps as gaps. Never fill them with invented numbers.

## 7. Presentation (fix before CLEAR)
21. Colours: inputs in blue font, formulas in black, judgment cells with a yellow fill.
22. A How_Measured sheet explains every method and every stress channel, plus the interaction row, mirror, reverse stress test, calibration cells and Summary sheet. Spell out each acronym on first use on each sheet.
23. A real football-field chart and a tornado chart, both driven by formula ranges, covering every enabled event. A Summary sheet right after the Cover ranks the events by absolute FULL change in $/sh with live formulas.
24. A "Draft, not investment advice" label on the Cover.

## Verdict template
```
VERDICT: CLEAR | BLOCK
Recomputed vs claimed: <method values, blended mid, 2+ FULL values>
Live-input test: <input changed> -> <outputs moved: n/n>
Interaction / asymmetry: <INT $/sh per event; mirrored FULL $/sh per event>
Reverse stress: <threshold set -> result; previous step at/above threshold: yes/no>
Calibration: <source + basis per shock, or TO FILL (BLOCK for a real issuer)>
Fixes (BLOCK only): 1. <issue> - <cell refs> - <required change>
Non-blocking notes: ...
```
