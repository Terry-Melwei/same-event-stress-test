#!/usr/bin/env python3
"""Reverse stress test: how big a single shock takes equity value down by a target?

Reads only the input cells of a recalculated workbook (Inputs plus each event sheet's yellow
and blue assumptions) and solves exactly by bisection with scripts/reference_dcf.py, so it does
not depend on the Ladder sheet's grid or interpolation. Draft, not investment advice.

Example:
  python scripts/reverse_stress.py examples/workbooks/gold_miner.xlsx --targets -0.1 -0.25
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from openpyxl import load_workbook  # noqa: E402

import reference_dcf as ref  # noqa: E402


LABELS = {"Event_Rates": "Rates (bp)", "Event_Commodity": "Commodity price", "Event_Tariff": "Input cost",
          "Event_Capex": "Customer capex"}


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("workbook", type=Path)
    p.add_argument("--targets", type=float, nargs="+", default=[-0.10, -0.25],
                   help="Target change in perpetual-method equity, e.g. -0.25 for a 25%% fall")
    p.add_argument("--max-bp", type=float, default=1000.0, help="Search range for rates, +/- basis points")
    p.add_argument("--max-shock", type=float, default=0.9, help="Search range for % events, +/- fraction")
    p.add_argument("--json", action="store_true", help="Print JSON instead of a Markdown table")
    a = p.parse_args(argv)

    wb = load_workbook(a.workbook, read_only=True)
    sheets = ref.event_sheets(wb.sheetnames)
    wb.close()
    m = ref.load_inputs_from_workbook(a.workbook)
    base = ref.base_equity_perp(m)

    rows = []
    for t in a.targets:
        res = {"target": t}
        for sh in sheets:
            lim = a.max_bp if sh == "Event_Rates" else a.max_shock
            sign = 1.0 if ref.configured_shock(m, sh) >= 0 else -1.0
            f = lambda x, sh=sh: ref.full_at(m, sh, x)  # noqa: E731
            # configured direction first, then the opposite direction
            hit = ref.solve_shock_for_drop(f, base, t, *sorted((0.0, sign * lim)))
            if hit is None:
                hit = ref.solve_shock_for_drop(f, base, t, *sorted((0.0, -sign * lim)))
            res[sh] = hit
        rows.append(res)

    if a.json:
        print(json.dumps({"base_equity_perp": base, "events": sheets, "results": rows}, indent=2))
        return 0
    print(f"Base equity (perpetual method): {base:,.2f}")
    print("| Target change | " + " | ".join(f"{LABELS[s]} shock needed" for s in sheets) + " |")
    print("|---" * (len(sheets) + 1) + "|")
    for r in rows:
        cells = []
        for sh in sheets:
            v = r[sh]
            cells.append("beyond search range" if v is None else (f"{v:+.1f} bp" if sh == "Event_Rates" else f"{v:+.2%}"))
        print(f"| {r['target']:+.0%} | " + " | ".join(cells) + " |")
    print("Draft, not investment advice. Exact solve on the same FULL definitions as the event sheets.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
