# Review checklist: same-event multi-channel stress tests

Use this checklist to review (gate) a valuation workbook before anyone relies on it.
The verdict is **CLEAR** only if every Blocking check passes. Otherwise it is **BLOCK**, with a numbered list of fixes that cites cells.
CLEAR means the workbook is built correctly. It does not mean the valuation assumptions are endorsed.

Terms: DCF = discounted cash flow. WACC = weighted average cost of capital. FCFF = free cash flow to firm. EV = enterprise value. GP = gross profit. Capex = capital expenditure. PV = present value. $/sh = US dollars per diluted share. %Δ = percentage change. FULL row = the result of one event with all of its channels applied together. INT row = the cross-channel interaction (not size non-linearity) diagnostic (FULL minus the standalone channel rows at one shock size). D&A = depreciation and amortisation. EBIT = earnings before interest and tax. EBITDA = earnings before interest, tax, depreciation and amortisation (in this model, EBIT + D&A). FX = foreign exchange.

## 1. Live wiring (Blocking)
1. **Live-input test.**
   - Copy the workbook. Change one core input (for example WACC +100 basis points) and recalculate (LibreOffice headless works).
   - Every output that should depend on the changed input must move: FULL values, tornado endpoints, envelope bounds and football-field stress-band cells. Outputs of events that don't use that input may stay unchanged. Say which outputs you expected to move.
   - Repeat with one operating input (for example gross margin, or the long-term commodity price).
2. **No pasted outputs.** Scan the scenario, FULL, interaction, mirror, covenant, reverse-stress grid and result, Summary, sensitivity-grid and chart-source ranges for typed-in numbers.
   - A "cached engine" of values counts as a defect even if it is documented.
   - Static cross-checks are allowed only when they are labelled, flagged STALE when inputs change, and do not feed any valuation output.
3. **Blend lock.** The blended midpoint averages the valuation-method rows only.
   - Stress bands, combined rows, interaction rows, mirror and reverse-stress results, the covenant row, the Summary sheet and reference rows (stale quotes, raw-WACC views) stay outside the blend.
4. **Cached values present.** Save after recalculating, so that formula results exist in the file. Zero formula errors.

## 2. Event design (Blocking)
5. **One event per table.** Each table has single-channel rows (A, B, C…) and a FULL row that applies all of them at once. Never stack separate events.
6. **Every channel present.** Minimum channels by event type:
   - **Interest rates:** WACC (cost of debt and cost of equity), interest on floating debt and cash, refinancing / mark-to-market of fixed debt, buyback or dividend capacity.
   - **Commodity price:** revenue, royalties and taxes, cost co-movement, capex/growth pace, financing and dividend capacity.
   - **Tariffs / input costs:** unit cost, price pass-through, demand elasticity, markdowns/promotions, inventory and working capital, sourcing mitigation.
   - **Foreign exchange (FX):** translation of foreign revenue, foreign-currency cost offset (natural hedge) plus a financial hedge counted once, and foreign-currency debt remeasured once in the net-debt bridge. This event is linear in the shock by construction; do not claim FX non-linearity.
   - **Consumer demand / comps:** volume, price/mix elasticity, operating leverage on a fixed-cost share of opex, and a Year-1 inventory / working-capital cash use that applies only when volume falls (inside FCFF, not a t0 net-debt restatement). Fixed-cost leverage alone does not make dollar FCFF non-linear in volume; any asymmetry must come from the revenue product or the inventory overlay.
   - **Capex cycle (customer or own):** own revenue, own capex (the FCFF channel), and the equity vs debt financing mix in the net-debt / diluted-share bridge. The funding need is financed once: the equity part lowers net debt and adds diluted shares in the same step; the debt part adds debt and cash equally. FULL is measured per pre-event diluted share.
   - **Jurisdiction:** one named event (for example a specific mine or country) that drives the outage, fiscal and risk-premium channels together.
   - Diagnostic rows (for example a standalone bridge view of a channel already counted in the FULL row) are allowed, but they must be labelled as diagnostic and kept out of the FULL row, so nothing is counted twice.
7. **The label matches the code.** The years, segments and weights stated in each row label are the ones the engine actually applies. Check the first affected year column.
8. **Interaction-row integrity.** Each event table has an INT row, labelled "Cross-channel interaction (not size non-linearity)" and kind diagnostic, that is a live formula equal to (FULL − base) − Σ(standalone row − base).
   - It sums standalone rows only: never information or diagnostic rows. On a sequential ladder (each row adds one response to the previous row, as on the tariff sheet) it isolates the remaining separable channel; say so on the sheet.
   - It is outside FULL and outside the blend, and is not pasted. Recompute it from the channel rows and compare.
9. **Both directions, reverse stress and the covenant row.** Each event shows FULL at the configured shock and at the mirrored shock (opposite sign) with an asymmetry line. Each event has a formula reverse-stress grid (ascending multiples) whose result cell reports the smallest tested shock that takes the FULL $/sh below the yellow threshold, or `n/a` when the threshold is blank, or `Not reached: worst is −X% at …` (the worst fall and shock on the grid) when nothing tested breaks the threshold. The default threshold is a live formula = base $/sh × (1 − drop %), drop % a yellow input (default 25%); a blank drop % gives `n/a`. A typed threshold is an override and needs a stated reason.
   - Test it: set a threshold between two grid steps, recalculate, and confirm the result is the first step below the threshold while the previous step is at or above it.
   - **Covenant row** (columns I–M of the mirror rows, and on the Summary): live net debt / Year-1 EBITDA under the FULL shock and its mirror, versus the blue Inputs limit (blank = n/a), with a PASS / BREACH / n/a flag. The Inputs limit is labelled ILLUSTRATIVE, replace with your own loan limit; a real covenant limit comes from the company's loan agreement, and its source cell ships as `TO FILL: cite the loan agreement clause`. It is a diagnostic, outside FULL and the blend. EBITDA = EBIT + D&A, matching the model's existing definition (opex already includes D&A, which is added back). A non-positive EBITDA, or a ratio above the limit, is BREACH. A BREACH flag is shown in red; PASS is shown in light green (live conditional formatting). Recompute the ratio independently and compare.
10. **Shock ladder.** A Ladder sheet re-runs every event's FULL at yellow multiples of the configured shock (default −1.5× to +1.5×) with live formulas. The asymmetry at the ±1× rung equals the event sheet's mirror asymmetry (row 63) and the Summary. If a yellow % fall target is outside the ladder, the reverse-stress cell reads `Not reached: worst is −X% at …` rather than a bare "outside range".
11. **Calibration source.** Every shock size has a yellow calibration **source** (data set or document, with date), **basis** (for example a historical 1-year 95th percentile move) and **severity** (plausible, severe or extreme).
    - `TO FILL` is acceptable only in an ILLUSTRATIVE template. A real-issuer workbook with `TO FILL`, or with a statistic that has no cited source, is BLOCK. Run the checker with `--strict-calibration`; it must PASS (without the flag, TO FILL is only a WARNING).

## 3. Economic invariants (Blocking)
12. **Pass-through invariant.** Set 100% cost recovery and zero elasticity / markdowns / mitigation. Gross-profit dollars must stay unchanged in every forecast year.
   - Build an explicit units × price × unit cost bridge, with fixed costs kept separate.
13. **Sign tests.** An isolated adverse input must lower value: a higher royalty, a higher tariff, a higher cost.
   - Historical calibration inputs must be separate from forward-looking shock inputs, so a shock never re-fits history.
14. **No financing double count.** EV uses operating cash flows only; no interest goes into FCFF.
    - Financing effects enter the net-debt bridge or the WACC weights once.
    - Share counts change once, alongside the matching cash.
    - Capex-cycle funding: the capex outflow is already in FCFF, so only the equity-funded part moves net debt (paired with the new diluted shares); the debt-funded part adds debt and cash equally. Never add the full funding need to net debt as well.
15. **No period overlap.** Stub-period cash, the balance-sheet date and the first explicit forecast year don't overlap.
16. **Lease consistency.** If lease liabilities are in net debt, use EBITDAR-based cash flows (earnings before interest, tax, depreciation, amortisation and rent). Otherwise leave leases out of both. Never subtract them twice.

## 4. Consistent bases (Blocking)
17. Measure each FULL % change against its own method's base (an exit-multiple shock against the exit-method value, the rest against perpetual growth). Then apply that % change to the blended midpoint.
    - Label the result a "blended-mid sensitivity proxy".
    - Never divide one method's result by another method's base.

## 5. Market data and comps (Blocking if material)
18. Use a live price with its date and source. Confirm the listing actually trades; otherwise value the primary listing and show any stale secondary quote as a reference only.
19. Recompute peer EV as market cap + debt − cash from raw inputs. Spot-check 2–3 peer multiples against an outside source, and investigate any gap above about 20%.

## 6. Disclosure (Blocking for big judgment calls)
20. Any judgment that changes value by more than about 10% (for example a chosen WACC instead of the raw CAPM, capital asset pricing model, rate) sits in a yellow cell. Show the raw-input value live on the Cover and Football, outside the blend.
21. List data gaps as gaps. Never fill them with invented numbers.

## 7. Presentation (fix before CLEAR)
22. Colours: inputs in blue font, formulas in black, judgment cells with a yellow fill.
23. A How_Measured sheet explains every method and every stress channel, plus the interaction row, mirror, covenant row (and that EBITDA = EBIT + D&A), reverse stress test, Ladder sheet, calibration cells and Summary sheet. Spell out each acronym on first use on each sheet.
24. A real football-field chart and a tornado chart, both driven by formula ranges, covering every enabled event. A Summary sheet right after the Cover ranks the events by absolute FULL change in $/sh with live formulas, including the covenant ratio and flag.
25. A "Draft, not investment advice" label on the Cover.

## Verdict template
```
VERDICT: CLEAR | BLOCK
Recomputed vs claimed: <method values, blended mid, 2+ FULL values>
Live-input test: <input changed> -> <outputs moved: n/n>
Interaction / asymmetry: <INT $/sh per event; mirrored FULL $/sh per event>
Reverse stress: <threshold set -> result; previous step at/above threshold: yes/no; Not reached wording if none>
Covenant: <limit; configured and mirror ratios and flags; independent recompute matches: yes/no>
Ladder: <±1× asymmetry matches mirror / Summary: yes/no>
Calibration: <source + basis + severity per shock, or TO FILL (BLOCK for a real issuer)>
Fixes (BLOCK only): 1. <issue> - <cell refs> - <required change>
Non-blocking notes: ...
```
