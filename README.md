# Same-event multi-channel stress test

**Draft, not investment advice.** This package documents a valuation stress-test method and ships an illustrative Excel template plus a YAML-driven generator. It does not recommend any security, issuer, or trade. Numbers marked ILLUSTRATIVE are teaching placeholders, not real company data or vendor prices.

## What it is

**Markets don't send warnings.** A central bank surprises by 100 basis points (bp, hundredths of a percentage point). Copper drops 20% in a month. A tariff lands overnight. The question isn't whether a shock will come; it's what that shock does to the value of the stock you own, and whether you will see it coming in your own model before the market shows you.

This skill answers that question one shock at a time. Pick a single event, such as interest rates ±100 bp, a commodity price ±20% or a customer capital-expenditure (capex) boom, and the workbook traces **every channel that same event pushes through**: revenue, margins, the discount rate, the debt load, the share count. Each one is a row (A, B, C…), and a **FULL** row adds them up into one number for equity value.

What you get from it:

- **The hidden asymmetry.** Shocks rarely cut both ways equally. Every event is run in both directions, side by side, with an asymmetry line, and a cross-channel interaction row shows how much the channels amplify (or cancel) each other beyond their simple sum. In this generator the asymmetry comes from discounting on the rates sheet, the funding floor on the capex sheet, and mitigation and the demand response on the tariff sheet. Commodity shocks are linear by construction here; floors such as equity at zero, tax losses and shut-ins are future work.
- **The channel that really matters.** Instead of one blended sensitivity number, you see which channel drives the move, so you know what to watch and what to hedge.
- **The shock that breaks you.** Name the value per share you cannot live with, and a reverse stress grid tells you how big the shock has to be to get there.
- **A plan before the headline.** Run the shocks you fear in advance, see them ranked on one Summary page and the downside range on a football-field chart, and decide your response while it is still a calm decision.

Under the hood, the method is built to be checked, not trusted. Free cash flow to the firm (FCFF) stays operating-only, and financing enters once, in the net-debt bridge or through the weighted average cost of capital (WACC), so nothing is double-counted. Interest and buyback capacity on the rates sheet are shown as information rows outside FULL; new shares issued to fund a capex boom change the share count and net debt together, once. Every shock size carries a calibration source cell that stays `TO FILL` until someone cites real data. The tornado, envelope, football-field and Summary views are all formula-linked to the FULL rows, and a bundled checker independently rebuilds every stress total, every mirrored shock and every reverse-stress grid step. Reviewers gate finished workbooks with [REVIEW_CHECKLIST.md](REVIEW_CHECKLIST.md).

Audience: equity and credit analysts, and people who run AI agents and want to hand the agent this method.

## Features

| Feature | Where | Notes |
|---------|-------|-------|
| One event per table, channel rows + FULL | `Event_Rates`, `Event_Commodity` or `Event_Tariff`, optional `Event_Capex` (rows 6–11) | Information / diagnostic rows stay outside FULL |
| Capex-cycle event | `Event_Capex` | Own revenue, own capex, equity vs debt financing mix (net debt and diluted shares change once, paired) |
| Cross-channel interaction (not size non-linearity) | Row 12 on every event sheet | FULL − Σ standalone rows at one shock size, in $m and US dollars per share ($/sh); diagnostic, outside FULL and the blend. Zero does not mean linear in shock size; read the mirror / asymmetry |
| Up / down mirror | Rows 60–63 | Same event at the opposite sign, plus an asymmetry line |
| Shock ladder | Ladder sheet (helpers rows 90–97 on each event sheet) | FULL at yellow multiples −1.5×…+1.5× of the configured shock; asymmetry per rung; ±1× rung checked equal to the mirror asymmetry (row 63) and the Summary; chart of FULL equity vs shock size; reverse stress by yellow % fall targets (−10%, −25%), interpolated between rungs |
| Reverse stress test | Rows 65–76 (helpers 80–88) | 8-step formula grid; yellow threshold defaults to base × (1 − 25%) via a yellow drop-% input; result = smallest tested shock below it, `n/a` when the drop % or threshold is blank |
| Shock calibration disclosure | Rows 54–57 | Yellow source + basis cells, shipped as `TO FILL`; checker WARNING, or FAIL with `--strict-calibration` |
| Summary ranking | `Summary` (right after Cover) | Events ranked by absolute FULL change in $/sh with LARGE + INDEX/MATCH |
| Tornado / envelope / football field | `Tornado`, `Envelope`, `Football` | Formula-linked to FULL rows; envelope MIN/MAX covers every enabled event |
| Checker | `scripts/check_workbook.py` | Live-input, shock and channel tests; independent recompute; pasted-output scan; regression mutations |

## Supported environment

| Tool | Supported | Tested |
|------|-----------|--------|
| Python | >=3.11,<3.14 | 3.11.18, 3.13.5 |
| LibreOffice | >=24 (headless `soffice`) | 25.2.3.2 |
| Packages | see [requirements.txt](requirements.txt) (pinned) | openpyxl 3.1.5, PyYAML 6.0.3, jsonschema 4.26.0 |

## Folder layout

```
stress-test-skill/
├── SKILL.md
├── REVIEW_CHECKLIST.md
├── README.md
├── LICENSE
├── CONTRIBUTING.md
├── requirements.txt
├── .gitignore
├── schemas/
│   └── config.schema.json
├── tests/
│   └── test_workbooks.py     # pytest: formula match, checker PASS, interaction, reverse stress, Summary
├── scripts/
│   ├── generate.py           # YAML → workbook
│   ├── check_workbook.py     # recalc + live tests + pass-through + FULL reconciliation
│   ├── reverse_stress.py     # exact solve: shock size that cuts equity by a target %
│   ├── reference_dcf.py      # independent DCF / joint-event reference for checker
│   └── build_stress_template.py  # thin shim → generate.py
├── templates/
│   └── stress_template.xlsx      # shipped WITH formula caches
├── examples/
│   ├── README.md
│   ├── configs/
│   │   ├── default_illustrative.yaml
│   │   ├── gold_miner.yaml
│   │   ├── apparel_retailer.yaml
│   │   └── optical_supplier.yaml    # rates + tariff + capex cycle
│   └── workbooks/               # shipped WITH formula caches
│       ├── gold_miner.xlsx
│       ├── apparel_retailer.xlsx
│       └── optical_supplier.xlsx
└── build/                       # local generate/recalc outputs (see commands)
    └── recalculated/
```

## Exact commands (venv → install → generate → recalc → check)

```bash
cd stress-test-skill
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -U pip
python -m pip install -r requirements.txt

# Generate into a staging folder (caches empty until recalc)
mkdir -p build/generated build/lo_in build/recalculated examples/workbooks
python scripts/generate.py \
  --config examples/configs/default_illustrative.yaml \
  --out build/generated/stress_template.xlsx

# Recalculate into a SEPARATE output folder (never convert onto the source path)
cp build/generated/stress_template.xlsx build/lo_in/stress_template.xlsx
soffice --headless --calc --convert-to xlsx \
  --outdir build/recalculated build/lo_in/stress_template.xlsx

# Ship the recalculated artifact (has formula caches)
cp build/recalculated/stress_template.xlsx templates/stress_template.xlsx

# Checker requires caches in the file under test (empty-cache files FAIL)
python scripts/check_workbook.py templates/stress_template.xlsx \
  --config examples/configs/default_illustrative.yaml --regression

# Gold miner example
python scripts/generate.py --config examples/configs/gold_miner.yaml --out build/generated/gold_miner.xlsx
cp build/generated/gold_miner.xlsx build/lo_in/gold_miner.xlsx
soffice --headless --calc --convert-to xlsx --outdir build/recalculated build/lo_in/gold_miner.xlsx
cp build/recalculated/gold_miner.xlsx examples/workbooks/gold_miner.xlsx
python scripts/check_workbook.py examples/workbooks/gold_miner.xlsx \
  --config examples/configs/gold_miner.yaml --regression

# Apparel example (tariff / input-cost event)
python scripts/generate.py --config examples/configs/apparel_retailer.yaml --out build/generated/apparel_retailer.xlsx
cp build/generated/apparel_retailer.xlsx build/lo_in/apparel_retailer.xlsx
soffice --headless --calc --convert-to xlsx --outdir build/recalculated build/lo_in/apparel_retailer.xlsx
cp build/recalculated/apparel_retailer.xlsx examples/workbooks/apparel_retailer.xlsx
python scripts/check_workbook.py examples/workbooks/apparel_retailer.xlsx \
  --config examples/configs/apparel_retailer.yaml

# Optical supplier example (rates + component tariff + customer capex cycle)
python scripts/generate.py --config examples/configs/optical_supplier.yaml --out build/generated/optical_supplier.xlsx
cp build/generated/optical_supplier.xlsx build/lo_in/optical_supplier.xlsx
soffice --headless --calc --convert-to xlsx --outdir build/recalculated build/lo_in/optical_supplier.xlsx
cp build/recalculated/optical_supplier.xlsx examples/workbooks/optical_supplier.xlsx
python scripts/check_workbook.py examples/workbooks/optical_supplier.xlsx \
  --config examples/configs/optical_supplier.yaml --regression

# Exact reverse stress test (bisection, no interpolation) for any shipped workbook
python scripts/reverse_stress.py examples/workbooks/gold_miner.xlsx --targets -0.10 -0.25

# Tests (LibreOffice-dependent tests are skipped when soffice is missing)
python -m pip install pytest
python -m pytest -v
```

**Important:** `generate.py` writes live formulas only. Formula caches stay empty until you recalculate into a **separate** outdir (LibreOffice errors if asked to overwrite the source path). `check_workbook.py` **FAILS** when the input file has empty caches (for example `DCF!B30 is None`). Always check the recalculated artifact.

## How to load this method into an agent

There are two different mechanisms. Do not confuse them.

### A) Native Agent Skills install (discovered automatically)

Per [Cursor Agent Skills docs](https://cursor.com/docs/skills), Cursor auto-discovers skills only from specific directories (for example `.cursor/skills/`, `.agents/skills/`, and compatibility trees such as `.claude/skills/`). Each skill is a **folder containing `SKILL.md`**. A plain `skills/` folder at the repo root is **not** Cursor auto-discovery.

**Cursor**

1. Copy this package to `.cursor/skills/same-event-stress-test/` (project) or `~/.cursor/skills/same-event-stress-test/` (user).
2. Keep the YAML frontmatter (`name`, `description`) on `SKILL.md`. The `name` should match the folder name.
3. Invoke with `/same-event-stress-test` or let the agent pick it from context.

**Claude Code / Agent Skills folder**

1. Copy to `.claude/skills/same-event-stress-test/` (or the user-level Claude skills directory your build uses).
2. Same `SKILL.md` frontmatter contract.

### B) Upload as knowledge or instructions (not auto-discovery)

**Custom GPTs / Claude Projects**

1. Upload `SKILL.md` and `REVIEW_CHECKLIST.md` as knowledge files.
2. Add instructions: follow SKILL.md; gate with REVIEW_CHECKLIST.md; never stack multiple events in one table.

**LangChain-style / other frameworks**

1. Load `SKILL.md` into the system prompt or a tool description.
2. Expose `REVIEW_CHECKLIST.md` as a required review step.
3. Point the agent at `scripts/generate.py` / `scripts/check_workbook.py` when it can run local tools.

## Template contents

[templates/stress_template.xlsx](templates/stress_template.xlsx) (shipped with caches) includes: Cover, Summary (events ranked by absolute FULL change in $/sh), Inputs, a five-year discounted cash flow (DCF), Event_Rates, Event_Commodity (or Event_Tariff in the apparel example; plus Event_Capex in the optical example), Tornado (table + bar chart), Envelope, Football (table + floating-bar chart), PassThrough (units × price × unit cost invariant), Ladder (each event's FULL at seven sizes in both directions, with a chart per event and a reverse stress test by % fall in value), and How_Measured. Every event sheet also carries the interaction row, the up/down mirror, the reverse stress grid and the calibration cells. Blue font = inputs; black = formulas; yellow fill + blue font = judgment.

To switch on the capex-cycle event in your own config, add `events.capex` with `enabled: true` and its inputs (`customer_capex_shock_pct`, `revenue_elasticity_to_customer_capex`, `own_capex_response`, `equity_funding_share`, `issue_price_per_share`). Any event can also take `threshold_drop_pct` (reverse-stress default threshold = base × (1 − drop); default 0.25, `null` for `n/a`), `threshold_value_per_share` (a typed $/sh threshold that overrides the default), `calibration_source`, `calibration_basis` and `reverse_stress_multiples` (8 ascending numbers). See [schemas/config.schema.json](schemas/config.schema.json).

## Checker limits (read before adapting)

- **Independent reference inputs.** `scripts/reference_dcf.py` rebuilds every stress total from the `Inputs` sheet only, never from channel rows or FULL formulas. Two exceptions: it also reads the raw tariff-bridge inputs at `Event_Tariff!B21:B29` and the capex-cycle inputs at `Event_Capex!B21:B25` (blue / yellow input cells). If you adapt those events, keep the inputs in those ranges or move them to `Inputs` and update `reference_dcf.py` to match.
- **Linear channels give a zero interaction.** In the template, the commodity channels and the rates WACC + mark-to-market channels add up linearly, so their interaction row is 0 by construction (the checker confirms it). Cross-effects show up where channels multiply: sourcing mitigation on the tariff sheet, and the capex funding floor plus dilution. A zero interaction does not mean the value response is linear in shock size: the rates mirror is asymmetric through discounting even though its INT is 0. Read the mirror / asymmetry rows and the Ladder sheet, which re-runs each FULL at −1.5× to +1.5× of the configured shock. In the examples, commodity events are exactly linear by construction (zero asymmetry on every rung), while rates are convex through discounting and the tariff bridge curves mildly.
- **Opex includes D&A.** Opex % of revenue includes depreciation and amortisation (D&A); FCFF adds D&A back, so it is never subtracted twice.
- **Calibration TO FILL.** The checker prints a WARNING naming every event whose calibration source or basis is still `TO FILL`; pass `--strict-calibration` for a real-issuer CLEAR to turn that into FAIL.
- **Capex-cycle simplifications.** The funding need is the PV of the incremental own capex, raised once at the valuation date and floored at zero (a capex cut needs no financing). The debt-funded part is value-neutral at fair value; the WACC weights are not re-levered. The shock is a permanent level shift through the terminal year.
- **Direction checks come from the config, not the labels.** The expected sign of each channel (for example, higher rates lower equity value) is taken from the YAML (YAML Ain't Markup Language) config's test expectations. Prose labels in the workbook are not validated. If you build a deliberately unusual scenario, such as a negative rate shock, set the matching expected directions in the config, or the checker will FAIL it (or, worse, a wrong label will go unnoticed).

## Review gate

Before anyone relies on a workbook built with this method, run [REVIEW_CHECKLIST.md](REVIEW_CHECKLIST.md) end-to-end and record the CLEAR / BLOCK verdict. `check_workbook.py` PASS is necessary but not sufficient for CLEAR.

## License

MIT License — Copyright (c) 2026 一然 王. See [LICENSE](LICENSE).
