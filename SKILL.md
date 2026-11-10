---
name: same-event-stress-test
description: Use when stress-testing a stock's valuation against one market shock at a time (interest rates, commodity prices, tariffs, foreign exchange, demand, regulation), or when building or reviewing a same-event multi-channel stress test (every transmission channel, FULL row, tornado, envelope and football-field charts) for a discounted cash flow (DCF) or similar model. Also use for "what happens to the share price if rates rise 100bp", "how exposed is this stock to copper", "is the downside bigger than the upside", or "how big a shock would cut value by a quarter".
---

# Same-event multi-channel stress test

**Draft, not investment advice.** This method is a documentation-and-template package for analysts and for people who run AI agents. It does not recommend buying, selling, or holding any security. Illustrative numbers are labelled ILLUSTRATIVE and are not real company data.

## When to use

Load this skill when you need to:

- Stress a discounted cash flow (DCF) or related valuation under **one named event at a time**, showing every channel that event moves.
- Produce a tornado, envelope, and football-field view that are **formula-linked** to those FULL rows, plus a Summary sheet that ranks the events.
- See how non-linear an event is (interaction row, up/down mirror) and how big a shock has to be to break a value threshold (reverse stress test).
- Gate a finished workbook with [REVIEW_CHECKLIST.md](REVIEW_CHECKLIST.md) and issue a written CLEAR / BLOCK verdict.

Do **not** use this skill to stack multiple unrelated events in one table, to paste static scenario outputs, or to treat a CLEAR review as an endorsement of assumptions.

## Method in one paragraph

A **same-event multi-channel stress test** takes a single shock (for example rates +100 basis points, or commodity price −20%) and builds **one table for that shock**. Rows A, B, C… each isolate one transmission channel. The **FULL** row applies every channel of that same event together. Never put two different events in the same table. Measure each FULL percentage change (%Δ) against **that method’s own base**, apply the %Δ to the blended midpoint, and label the result a **blended-mid sensitivity proxy**. Wire tornado, envelope, and football-field ranges to the FULL rows with live Excel formulas. Under every table, show what the channel rows cannot: an **interaction** row (FULL minus the sum of the standalone rows), the same event **mirrored** (opposite sign) so the asymmetry is visible, a **reverse stress** grid, and the **calibration source** of the shock size.

## Choosing the shock size

A stress test is only as useful as its shock is plausible. Before building a table, write down **why** this size:

- **Anchor to history.** Prefer a move the market has actually made, such as the largest 12-month change in the last 10 to 20 years, a named past episode, or a one-in-ten-year move from the series' own volatility. Put the source and window in the yellow calibration source cell.
- **Say how severe it is.** Label each shock *plausible* (seen several times in the window), *severe* (seen once) or *extreme* (beyond the window). Don't present an extreme shock as a base case.
- **Match the horizon.** The valuation assumes the shock persists through the forecast. If history says the move usually reverses within a year, say so next to the result, because a persistent-shock valuation will overstate the damage.

## Reading the result: asymmetry and non-linearity

Each event sheet already shows the shock and its mirror (the same size, opposite sign). In any write-up:

- **Lead with the downside / upside ratio** = |change in value per share from the adverse case| ÷ |change from the favourable case|. A ratio well above 1 means the stock loses more than it gains from the same-sized move.
- **Name the channel behind it:** operating leverage on fixed costs, a royalty or tax step, debt that amplifies equity moves, or a dominant terminal value.
- **Don't read a zero interaction row as "linear".** The interaction row measures how channels *combine* at one shock size. A shock can still hurt more than proportionally as it grows, for example through the convexity of value in the weighted average cost of capital (WACC). Use the mirror and, where built, a multi-size shock ladder to see that curve.
- **Commodity events are linear by construction in this generator.** Every channel scales in proportion to price, so the +x% and −x% cases mirror exactly. Floors on equity, untaxed losses and shutting high-cost output are not modelled yet, so don't claim commodity non-linearity from this workbook.

## Reverse stress test

The reverse-stress grid (wiring rule 11) answers: how big a shock pushes value per share below the threshold in the yellow cell? If nobody has a specific level in mind, use **25% below the base value per share** as the default and say so. Then compare the break-point shock with the calibration source: a break point inside the historically plausible range is a finding to lead with; one far beyond the extreme range is reassurance. For a debt covenant, read net debt ÷ earnings before interest, tax, depreciation and amortisation (EBITDA) at the break-point shock by hand until a covenant row exists.

## Per-shock channel checklists

For each event type below, include at least the listed channels (add more if material). Every acronym is spelled out on first use in this section.

### Interest rates

Shock example: ±100 basis points (bp). Stress the **selected** weighted average cost of capital (WACC): stressed WACC = selected WACC + Δ (the selected WACC already includes any yellow judgment override).

| Row | Kind | Channel | What moves |
|-----|------|---------|------------|
| A | standalone | WACC | Re-discount operating free cash flow to the firm (FCFF) at selected WACC + Δ |
| B | information | Refinancing / floating interest | Shown for communication; **outside FULL** |
| C | standalone | Fixed-debt mark-to-market | Fair-value adjustment inside the **net-debt bridge** only |
| D | information | Profit / earnings per share (EPS) | Communication only; **outside FULL** |
| E | information | Buyback capacity | Communication only; **outside FULL** so the same interest cash is not subtracted twice |
| FULL | cumulative | A + C | WACC channel + MTM bridge — financing once; B/D/E excluded |

### Commodity price

Shock example: ±20% on the long-term commodity price. Model **revenue and operating expenses (opex) separately**. Price moves revenue only; costs move only via the documented cost elasticity (yellow judgment input).

| Row | Kind | Channel | What moves |
|-----|------|---------|------------|
| A | standalone | Revenue | Price × volume (volume held); opex fixed at base |
| B | standalone | Royalties / tax | Base calibration rate on Inputs — never re-fit history with the shock |
| C | standalone | Cost co-move | Opex × (1 + cost elasticity × price shock) only |
| D | standalone | Capital expenditure (capex) / reserve pace | Capex % × (1 + capex–price elasticity × shock) |
| E | diagnostic | Net debt / dividends bridge view | Labelled diagnostic; **outside FULL** |
| FULL | cumulative | A+B+C+D | Diagnostic E excluded |

### Cost inflation

Channels: unit operating cost; price pass-through; volume / mix response; working capital; net debt. Keep fixed occupancy or fixed overhead separate from unit costs.

### Production / volume

Channels: revenue units; variable cost; royalties/tax on volume; working capital; capex/reserve depletion pace; net debt.

### Foreign exchange (FX)

Channels: sales translation; local-currency cost translation; foreign-currency debt translation in the net-debt bridge; any FX hedges treated once.

### Jurisdiction / regulatory

**One named event** (for example a specific mine permit pause or a named country fiscal change) driving together: outage / volume; fiscal (royalty, tax, export levy); equity risk-premium or country WACC add-on. Do not mix unrelated jurisdictions in one table.

### Tariffs and input costs

Build an explicit **units × price × unit cost** bridge. Keep fixed occupancy separate. Count **pass-through once**. Channels: unit cost; price recovery; demand elasticity / markdowns; inventory and working capital; sourcing mitigation.

### Consumer demand / comps

Channels: comparable-store or volume; price/mix; markdowns; inventory; store / capacity opex; net debt.

### Capex cycle (customer or own)

Shock example: customer capital expenditure (capex) +30% (for example a data-centre build-out), entered in its adverse direction. Own revenue and own capex respond through yellow judgment inputs; the funding need is financed **once**.

| Row | Kind | Channel | What moves |
|-----|------|---------|------------|
| A | standalone | Own revenue | Revenue × (1 + revenue elasticity × shock); opex, royalty and net working capital (NWC) follow revenue; depreciation and amortisation (D&A) and capex stay at base dollars |
| B | standalone | Own capex | Capex × (1 + own-capex response × shock), base revenue; this is the free cash flow to the firm (FCFF) channel |
| C | standalone | Financing mix | Funding need = MAX(0, present value (PV) of incremental own capex) raised at the valuation date: the equity share lowers net debt **and** adds diluted shares at the yellow issue price in the same step; the debt share adds debt and cash equally (net debt unchanged) |
| D | information | FCFF carrier | Operating A+B before financing; **outside FULL** |
| E | information | All-equity funding view | Communication only; **outside FULL** |
| FULL | cumulative | A+B+C | Equity attributable to the **pre-event** diluted shares, so %Δ and value per share (US dollars per diluted share, $/sh) already include dilution |

A capex cut frees cash inside FCFF and needs no financing, so the financing channel switches off on the downside: the mirror row shows that asymmetry.

### Terminal growth

Single channel on the perpetual-growth terminal value. Measure %Δ against the perpetual-growth method base.

### Exit multiple

Single channel on the exit-multiple terminal value. Measure %Δ against the **exit-method** base (not the perpetual-growth base).

### Net debt

Channels that restate cash, debt, leases, or other bridge items without changing operating FCFF. Lease rule: if lease liabilities sit in net debt, use earnings before interest, tax, depreciation, amortisation and rent (EBITDAR)-consistent cash flows; otherwise leave leases out of both.

### Dilution / buyback

Share count changes **once**, alongside the matching cash in the bridge. Always use **diluted** shares for per-share value.

## Wiring rules

1. **FCFF is operating-only.** Free cash flow to the firm excludes interest and financing cash. Financing enters **once** in the net-debt bridge (and WACC weights as designed). On the rates sheet, interest and buyback capacity are **information** rows outside FULL so the same cash interest is never subtracted twice. Never push interest into FCFF or double-count the tax shield. Stress **selected WACC + Δ**, not a rebuilt CAPM path that ignores a yellow WACC override.
2. **One event per table.** Channel rows A/B/C… plus FULL. No stacking of separate events.
3. **Own-base %Δ → blended mid.** For each FULL row, %Δ = (stressed method value − that method’s base) / that method’s base. Apply %Δ to the blended midpoint. Label outputs **blended-mid sensitivity proxy**. Never divide one method’s result by another method’s base.
4. **Charts follow FULL rows.** Tornado, envelope (MIN/MAX over **every** enabled event), football-field stress bands and the Summary ranking are formula-linked to FULL outputs.
5. **Live formulas only.** Scenario outputs are Excel formulas, never pasted values. After changing an input and recalculating, every output that should depend on the changed input must move (FULL values, tornado endpoints, envelope bounds, and football-field stress-band cells). Outputs of events that do not use that input may stay unchanged.
6. **Colour conventions.** Blue font = inputs. Black font = formulas. Yellow fill = judgment (and disclose large judgments on the cover / football outside the blend).
7. **Diluted shares.** Equity value per share uses diluted share count; share-count changes pair with cash once.
8. **Blend lock.** The blended midpoint averages valuation-method rows only. Stress bands, the interaction row, mirror and reverse-stress results, the Summary sheet and reference rows stay outside the blend.
9. **Interaction row.** Row 12 on each event sheet (labelled INT, kind diagnostic) = (FULL − base) − Σ(standalone row − base), in $m and $/sh. Never sum information or diagnostic rows into it, never feed it into FULL or the blend, and never paste it. Where the channel rows are a sequential ladder (the tariff sheet), the increments already sum to the last step, so INT isolates the cross-effect of the remaining separable channel (sourcing mitigation). INT = 0 means the channels add linearly; in the template the commodity and rates channels are additive by construction, so their INT is 0 and that is correct.
10. **Both directions.** Show FULL at the configured shock and at the mirrored shock (same size, opposite sign) with an asymmetry line = (configured − base) + (mirrored − base).
11. **Reverse stress test.** An 8-step formula grid of ascending shock multiples (blue inputs) in the configured direction, and a result cell that reports the smallest tested shock whose FULL proxy $/sh falls below a yellow threshold. Blank threshold → `n/a`. No goal seek, no macros, no pasted answers.
12. **Calibration rule.** Every shock size has a yellow **source** cell (data set or document, with date) and a yellow **basis** cell (for example a historical 1-year 95th percentile move). Ship them as `TO FILL: source` / `TO FILL: basis`; never invent a statistic. A real-issuer workbook cannot be CLEAR while either says TO FILL.
13. **Enter shocks in their adverse direction.** The grid and the default checker tests treat the configured sign as the adverse case; flip the input to test the other side.
14. **Summary sheet.** Right after the Cover, rank every enabled event by absolute FULL change in $/sh with LARGE + INDEX/MATCH formulas (no SORT, no macros), showing down-case / up-case $/sh, %Δ, interaction, asymmetry and the reverse-stress result.

## Colour and labelling conventions

| Signal | Meaning |
|--------|---------|
| Blue font | Editable input |
| Black font | Formula |
| Yellow fill + blue font | Judgment / discretionary assumption |
| ILLUSTRATIVE | Fake numbers for teaching; not a real issuer |
| blended-mid sensitivity proxy | FULL %Δ applied to blended midpoint |
| INT (diagnostic) | Interaction row: FULL minus the standalone channel rows; outside FULL and the blend |
| TO FILL | Calibration source / basis not yet cited; blocks a real-issuer CLEAR |
| Draft, not investment advice | Required cover label |

## Workflow

1. **Build** the base DCF (or related model) with live formulas, diluted shares, operating-only FCFF, separate revenue/opex lines, and a clean net-debt bridge. Or generate from YAML: `python scripts/generate.py --config examples/configs/<case>.yaml --out <out.xlsx>`.
2. **Add one event table per shock**, with channel rows labelled standalone / cumulative / information / diagnostic, and a FULL row that excludes information/diagnostic rows. Add the interaction row, the mirror, the reverse-stress grid and the calibration cells under it. Fill the calibration source/basis or leave them `TO FILL`.
3. **Wire** tornado, envelope, football field (tables + charts) and the Summary sheet to the FULL rows; label blended-mid sensitivity proxies.
4. **Recalculate** (LibreOffice or Excel). Regeneration leaves formula caches empty until recalc.
5. **Run** `python scripts/check_workbook.py <out.xlsx> --config <case.yaml>` (live tests, pasted-output scan, pass-through invariant, independent FULL / mirror / grid recompute, interaction, reverse-stress and Summary checks), then every Blocking item in [REVIEW_CHECKLIST.md](REVIEW_CHECKLIST.md).
6. **Report both directions** and the downside / upside ratio, name the channel behind any asymmetry, and state the shock's historical anchor and severity label.
7. **Write a pass/fail verdict** using the checklist’s verdict template (`CLEAR` or `BLOCK` with numbered cell-cited fixes). CLEAR means the workbook is built correctly; it does not endorse the assumptions.

## Pitfalls (from real reviews) and fixes

| Pitfall | Fix |
|---------|-----|
| Orphan pasted values that do not move when inputs change | Replace with formulas; run the live-input test in REVIEW_CHECKLIST.md §1 |
| Royalty calibration that reused the shocked rate for historical cost (sign bug) | Keep historical calibration inputs separate from forward shock inputs so a shock never re-fits history |
| Financing counted in FCFF **and** the bridge **and** the tax shield | FCFF operating-only; financing once in the net-debt bridge / WACC weights |
| Mixed valuation bases in one tornado | Each FULL %Δ vs its own method base; then apply to blended mid; label proxy |
| Tariff price recovery counted twice | Units × price × unit cost bridge; pass-through once; fixed occupancy separate |
| Labels whose timing did not match the formulas | Row label years/segments/weights must match the first affected year column |
| Stale or wrong data-vendor enterprise value (EV) fields | Recompute EV = market capitalisation + debt − cash from raw inputs |
| Capex funding counted in FCFF **and** again as new debt in the bridge | Debt-funded part adds debt and cash equally; only the equity part moves net debt, paired with new diluted shares |
| A "shock size" with no source | Yellow calibration source + basis cells; `TO FILL` until sourced |
| Interaction row typed in, or summing information rows | Live formula over standalone rows only; checker recomputes it |
| One-directional shock reported as "the" sensitivity | Run adverse and favourable sizes; report both and the downside / upside ratio |
| Shock size picked without justification | Anchor to a cited historical move and label it plausible, severe or extreme |

## Package pointers

- Template workbook: [templates/stress_template.xlsx](templates/stress_template.xlsx)
- Generator / checker: [scripts/generate.py](scripts/generate.py), [scripts/check_workbook.py](scripts/check_workbook.py)
- Example configs: [examples/configs/](examples/configs/) (rates + commodity, rates + tariff, and rates + tariff + capex cycle in `optical_supplier.yaml`)
- Worked prose examples: [examples/README.md](examples/README.md)
- Review gate: [REVIEW_CHECKLIST.md](REVIEW_CHECKLIST.md)
- Package overview and how to load this skill into agents: [README.md](README.md)
