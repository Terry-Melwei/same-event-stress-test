# Contributing

**Draft, not investment advice.** Contributions improve the method documentation, generator, checker, and illustrative template only.

## Scope

- Keep the method **same-event**: one shock per table, channel rows plus FULL.
- Preserve wiring rules: free cash flow to the firm (FCFF) operating-only; financing once in the net-debt bridge / weighted average cost of capital (WACC); rates interest and buyback as information rows outside FULL; commodity revenue/opex separated; live formulas; own-base %Δ applied to the blended mid and labelled blended-mid sensitivity proxy.
- Label rows standalone / cumulative / information / diagnostic; keep information and diagnostic rows out of FULL.
- Keep the per-event blocks at their fixed rows (interaction row 12, shock calibration 54–58, mirror 60–63, covenant note 64, reverse stress 65–76, helpers 80–88, ladder helpers 90–97); `check_workbook.py`, `reference_dcf.py` and the tests read those cells. If you move them, update all three.
- Shock calibration cells (rows 54–58) are: 54 title, 55 the shock size used by FULL, 56 yellow **source**, 57 yellow **basis**, 58 yellow **severity** (plausible, severe or extreme). Ship the yellow cells as `TO FILL: source` / `TO FILL: basis` / `TO FILL: severity`. Never replace them with an uncited statistic.
- Run `python -m pytest -v` (needs LibreOffice for the recalculation tests) and `check_workbook.py --regression` on every shipped workbook before opening a pull request. Code must stay valid on Python 3.11.
- Discretionary assumptions use **yellow fill + blue font**.
- Do not add real issuer data, vendor prices, or pasted scenario outputs.
- Leave [REVIEW_CHECKLIST.md](REVIEW_CHECKLIST.md) Blocking intent intact unless you are deliberately tightening a gate; document why.

## Events, foreign exchange, demand and the covenant row

Enabled event sheets are `Event_Rates` plus one of `Event_Commodity` or `Event_Tariff`, and optionally `Event_Capex` (capital expenditure), `Event_FX` (foreign exchange) and `Event_Demand` (consumer demand). Switch the optional events on in the YAML (YAML Ain't Markup Language) config: `events.fx.enabled` / `events.demand.enabled` (and the matching capex flag) with their inputs, as documented in [README.md](README.md) and [schemas/config.schema.json](schemas/config.schema.json). They are disabled by default; the apparel retailer example turns FX and demand on. `scripts/generate.py` builds one sheet per enabled event and wires it into the Summary, Ladder, Tornado, Envelope and Football views.

The debt-covenant row is a diagnostic, outside FULL and the blend. The limit is the blue input `inputs.covenant_nd_ebitda_limit` on Inputs!B52 (net debt / Year-1 earnings before interest, tax, depreciation and amortisation (EBITDA); blank or omitted = n/a). The Inputs label reads "ILLUSTRATIVE, replace with your own loan limit". A real covenant limit comes from the company's loan agreement; the yellow source cell Inputs!B55 ships as `TO FILL: cite the loan agreement clause`. Each event sheet shows the ratio and a PASS / BREACH / n/a flag in column M of the mirror rows (M61 configured, M62 mirrored); the Summary repeats those flags. A BREACH flag is shown in red; PASS is shown in light green, via live conditional formatting. EBITDA = EBIT (earnings before interest and tax) + D&A (depreciation and amortisation). To change the limit or its source, edit the config / Inputs cells, not the flag formulas. The checker independently recomputes the ratio and flag, warns (or fails under `--strict-calibration`) while Inputs!B55 still says TO FILL, and the tests include a deliberately broken flag the checker must catch.

`scripts/reverse_stress.py` solves the exact shock that cuts perpetual-method equity by a target. It derives column labels from each event sheet's A1 title (falling back to the sheet name), so an unknown `Event_` sheet must not raise `KeyError`. Tests run the script on the apparel example and check that fallback.

## Regenerating and checking

```bash
source .venv/bin/activate
mkdir -p build/generated build/recalculated
python scripts/generate.py --config examples/configs/default_illustrative.yaml --out build/generated/stress_template.xlsx
# caches empty until you recalculate into a SEPARATE folder (LibreOffice will not overwrite its source):
soffice --headless --calc --convert-to xlsx --outdir build/recalculated build/generated/stress_template.xlsx
cp build/recalculated/stress_template.xlsx templates/stress_template.xlsx
python scripts/check_workbook.py templates/stress_template.xlsx --config examples/configs/default_illustrative.yaml --regression
python -m pytest -v
```

Configs must validate against [schemas/config.schema.json](schemas/config.schema.json).

## Style

Precise, plain English. Spell out acronyms on first use. Keep paths relative inside the package; no machine-absolute paths in published files.
