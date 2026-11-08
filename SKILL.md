---
name: same-event-stress-test
description: Use when building or reviewing a same-event multi-channel valuation stress test (one shock, every transmission channel, FULL row, tornado/envelope/football field) for a DCF or similar model.
---

# Same-event multi-channel stress test

**Draft, not investment advice.** This method is a documentation-and-template package for analysts and for people who run AI agents. It does not recommend buying, selling, or holding any security. Illustrative numbers are labelled ILLUSTRATIVE and are not real company data.

## When to use

Load this skill when you need to:

- Stress a discounted cash flow (DCF) or related valuation under **one named event at a time**, showing every channel that event moves.
- Produce a tornado, envelope, and football-field view that are **formula-linked** to those FULL rows.
- Gate a finished workbook with [REVIEW_CHECKLIST.md](REVIEW_CHECKLIST.md) and issue a written CLEAR / BLOCK verdict.

Do **not** use this skill to stack multiple unrelated events in one table, to paste static scenario outputs, or to treat a CLEAR review as an endorsement of assumptions.

## Method in one paragraph

A **same-event multi-channel stress test** takes a single shock (for example rates +100 basis points, or commodity price −20%) and builds **one table for that shock**. Rows A, B, C… each isolate one transmission channel. The **FULL** row applies every channel of that same event together. Never put two different events in the same table. Measure each FULL percentage change (%Δ) against **that method’s own base**, apply the %Δ to the blended midpoint, and label the result a **blended-mid sensitivity proxy**. Wire tornado, envelope, and football-field ranges to the FULL rows with live Excel formulas.

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

Channels: own revenue; own capex; free cash flow to the firm (FCFF); equity versus debt financing mix in the net-debt / share-count bridge.

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
4. **Charts follow FULL rows.** Tornado, envelope (MIN/MAX), and football-field stress bands are formula-linked to FULL outputs.
5. **Live formulas only.** Scenario outputs are Excel formulas, never pasted values. After changing an input and recalculating, every output that should depend on the changed input must move (FULL values, tornado endpoints, envelope bounds, and football-field stress-band cells). Outputs of events that do not use that input may stay unchanged.
6. **Colour conventions.** Blue font = inputs. Black font = formulas. Yellow fill = judgment (and disclose large judgments on the cover / football outside the blend).
7. **Diluted shares.** Equity value per share uses diluted share count; share-count changes pair with cash once.
8. **Blend lock.** The blended midpoint averages valuation-method rows only. Stress bands and reference rows stay outside the blend.

## Colour and labelling conventions

| Signal | Meaning |
|--------|---------|
| Blue font | Editable input |
| Black font | Formula |
| Yellow fill + blue font | Judgment / discretionary assumption |
| ILLUSTRATIVE | Fake numbers for teaching; not a real issuer |
| blended-mid sensitivity proxy | FULL %Δ applied to blended midpoint |
| Draft, not investment advice | Required cover label |

## Workflow

1. **Build** the base DCF (or related model) with live formulas, diluted shares, operating-only FCFF, separate revenue/opex lines, and a clean net-debt bridge. Or generate from YAML: `python scripts/generate.py --config examples/configs/<case>.yaml --out <out.xlsx>`.
2. **Add one event table per shock**, with channel rows labelled standalone / cumulative / information / diagnostic, and a FULL row that excludes information/diagnostic rows.
3. **Wire** tornado, envelope, and football field (tables + charts) to the FULL rows; label blended-mid sensitivity proxies.
4. **Recalculate** (LibreOffice or Excel). Regeneration leaves formula caches empty until recalc.
5. **Run** `python scripts/check_workbook.py <out.xlsx> --config <case.yaml>` (live tests, pasted-output scan, pass-through invariant), then every Blocking item in [REVIEW_CHECKLIST.md](REVIEW_CHECKLIST.md).
6. **Write a pass/fail verdict** using the checklist’s verdict template (`CLEAR` or `BLOCK` with numbered cell-cited fixes). CLEAR means the workbook is built correctly; it does not endorse the assumptions.

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

## Package pointers

- Template workbook: [templates/stress_template.xlsx](templates/stress_template.xlsx)
- Generator / checker: [scripts/generate.py](scripts/generate.py), [scripts/check_workbook.py](scripts/check_workbook.py)
- Example configs: [examples/configs/](examples/configs/)
- Worked prose examples: [examples/README.md](examples/README.md)
- Review gate: [REVIEW_CHECKLIST.md](REVIEW_CHECKLIST.md)
- Package overview and how to load this skill into agents: [README.md](README.md)
