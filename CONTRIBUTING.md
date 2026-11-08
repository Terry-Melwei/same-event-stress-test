# Contributing

**Draft, not investment advice.** Contributions improve the method documentation, generator, checker, and illustrative template only.

## Scope

- Keep the method **same-event**: one shock per table, channel rows plus FULL.
- Preserve wiring rules: free cash flow to the firm (FCFF) operating-only; financing once in the net-debt bridge / weighted average cost of capital (WACC); rates interest and buyback as information rows outside FULL; commodity revenue/opex separated; live formulas; own-base %Δ applied to the blended mid and labelled blended-mid sensitivity proxy.
- Label rows standalone / cumulative / information / diagnostic; keep information and diagnostic rows out of FULL.
- Discretionary assumptions use **yellow fill + blue font**.
- Do not add real issuer data, vendor prices, or pasted scenario outputs.
- Leave [REVIEW_CHECKLIST.md](REVIEW_CHECKLIST.md) Blocking intent intact unless you are deliberately tightening a gate; document why.

## Regenerating and checking

```bash
source .venv/bin/activate
python scripts/generate.py --config examples/configs/default_illustrative.yaml --out templates/stress_template.xlsx
# caches empty until:
soffice --headless --calc --convert-to xlsx --outdir templates templates/stress_template.xlsx
python scripts/check_workbook.py templates/stress_template.xlsx --config examples/configs/default_illustrative.yaml
```

Configs must validate against [schemas/config.schema.json](schemas/config.schema.json).

## Style

Precise, plain English. Spell out acronyms on first use. Keep paths relative inside the package; no machine-absolute paths in published files.
