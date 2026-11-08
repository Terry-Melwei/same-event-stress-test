# Same-event multi-channel stress test

**Draft, not investment advice.** This package documents a valuation stress-test method and ships an illustrative Excel template plus a YAML-driven generator. It does not recommend any security, issuer, or trade. Numbers marked ILLUSTRATIVE are teaching placeholders, not real company data or vendor prices.

## What it is

**Markets don't send warnings.** A central bank surprises by 100 basis points (bp, hundredths of a percentage point). Copper drops 20% in a month. A tariff lands overnight. The question isn't whether a shock will come; it's what that shock does to the value of the stock you own, and whether you will see it coming in your own model before the market shows you.

This skill answers that question one shock at a time. Pick a single event, such as interest rates ±100 bp or a commodity price ±20%, and the workbook traces **every channel that same event pushes through**: revenue, margins, the discount rate, the debt load. Each one is a row (A, B, C…), and a **FULL** row adds them up into one number for equity value.

What you get from it:

- **The hidden asymmetry.** Shocks rarely cut both ways equally. A +20% move and a −20% move can change value by very different amounts, and the tornado and envelope views show where the pain is non-linear, so a "small" shock that does outsized damage is easy to spot.
- **The channel that really matters.** Instead of one blended sensitivity number, you see which channel drives the move, so you know what to watch and what to hedge.
- **A plan before the headline.** Run the shocks you fear in advance, see the downside range on a football-field chart, and decide your response while it is still a calm decision.

Under the hood, the method is built to be checked, not trusted. Free cash flow to the firm (FCFF) stays operating-only, and financing enters once, in the net-debt bridge or through the weighted average cost of capital (WACC), so nothing is double-counted. Interest and buyback capacity on the rates sheet are shown as information rows outside FULL. The tornado, envelope and football-field charts are all formula-linked to the FULL rows, and a bundled checker independently rebuilds every stress total. Reviewers gate finished workbooks with [REVIEW_CHECKLIST.md](REVIEW_CHECKLIST.md).

Audience: equity and credit analysts, and people who run AI agents and want to hand the agent this method.

## Supported environment

| Tool | Supported | Tested |
|------|-----------|--------|
| Python | >=3.11,<3.14 | 3.13.5 |
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
├── scripts/
│   ├── generate.py           # YAML → workbook
│   ├── check_workbook.py     # recalc + live tests + pass-through + FULL reconciliation
│   ├── reference_dcf.py      # independent DCF / joint-event reference for checker
│   └── build_stress_template.py  # thin shim → generate.py
├── templates/
│   └── stress_template.xlsx      # shipped WITH formula caches
├── examples/
│   ├── README.md
│   ├── configs/
│   │   ├── default_illustrative.yaml
│   │   ├── gold_miner.yaml
│   │   └── apparel_retailer.yaml
│   └── workbooks/               # shipped WITH formula caches
│       ├── gold_miner.xlsx
│       └── apparel_retailer.xlsx
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

[templates/stress_template.xlsx](templates/stress_template.xlsx) (shipped with caches) includes: Cover, Inputs, a five-year discounted cash flow (DCF), Event_Rates, Event_Commodity (or Event_Tariff in the apparel example), Tornado (table + bar chart), Envelope, Football (table + floating-bar chart), PassThrough (units × price × unit cost invariant), and How_Measured. Blue font = inputs; black = formulas; yellow fill + blue font = judgment.

## Checker limits (read before adapting)

- **Independent reference inputs.** `scripts/reference_dcf.py` rebuilds every stress total from the `Inputs` sheet only, never from channel rows or FULL formulas. One exception: in the apparel example it also reads the raw tariff-bridge inputs (blue-font input cells) at `Event_Tariff!B21:B29`. If you adapt the apparel event, keep those inputs in that range or move them to `Inputs` and update `reference_dcf.py` to match.
- **Direction checks come from the config, not the labels.** The expected sign of each channel (for example, higher rates lower equity value) is taken from the YAML (YAML Ain't Markup Language) config's test expectations. Prose labels in the workbook are not validated. If you build a deliberately unusual scenario, such as a negative rate shock, set the matching expected directions in the config, or the checker will FAIL it (or, worse, a wrong label will go unnoticed).

## Review gate

Before anyone relies on a workbook built with this method, run [REVIEW_CHECKLIST.md](REVIEW_CHECKLIST.md) end-to-end and record the CLEAR / BLOCK verdict. `check_workbook.py` PASS is necessary but not sufficient for CLEAR.

## License

MIT License — Copyright (c) 2026 一然 王. See [LICENSE](LICENSE).
