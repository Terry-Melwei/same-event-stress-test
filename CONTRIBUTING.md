# Contributing

**Draft, not investment advice.** Contributions improve the method documentation, generator, checker, and illustrative template only.

## Scope

- Keep the method **same-event**: one shock per table, channel rows plus FULL.
- Preserve wiring rules: free cash flow to the firm (FCFF) operating-only; financing once in the net-debt bridge / weighted average cost of capital (WACC); rates interest and buyback as information rows outside FULL; commodity revenue/opex separated; live formulas; own-base %Δ applied to the blended mid and labelled blended-mid sensitivity proxy.
- Label rows standalone / cumulative / information / diagnostic; keep information and diagnostic rows out of FULL.
- Keep the per-event blocks at their fixed rows (interaction row 12, calibration 54–57, mirror 60–63, reverse stress 65–76, helpers 80–88); `check_workbook.py`, `reference_dcf.py` and the tests read those cells. If you move them, update all three.
- Shock calibration cells ship as `TO FILL: source` / `TO FILL: basis`. Never replace them with an uncited statistic.
- Run `python -m pytest -v` (needs LibreOffice for the recalculation tests) and `check_workbook.py --regression` on every shipped workbook before opening a pull request. Code must stay valid on Python 3.11.
- Discretionary assumptions use **yellow fill + blue font**.
- Do not add real issuer data, vendor prices, or pasted scenario outputs.
- Leave [REVIEW_CHECKLIST.md](REVIEW_CHECKLIST.md) Blocking intent intact unless you are deliberately tightening a gate; document why.

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
