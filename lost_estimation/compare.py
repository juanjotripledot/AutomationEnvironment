"""AI vs human accuracy comparison.

For LOST Story/Task/Bug tickets that reached "Verified for Production" within the
trailing 14 days, compute:
  - actual cycle (business days) from First Time In Progress → First Time In Verified For Production
  - implied SP from the threshold table in baseline_patterns.json
  - human SP error  = |human_sp - implied_sp_from_cycle|   (in Fibonacci-rank distance)
  - AI    SP error  = |ai_sp    - implied_sp_from_cycle|

Outputs:
  reports/runs/<YYYY-MM-DD>/accuracy.csv  — one row per verified ticket
  reports/runs/<YYYY-MM-DD>/accuracy_summary.json — aggregate MAE for human and AI
"""

from __future__ import annotations

import csv
import datetime as dt
import json
import os
import sys
from pathlib import Path

from dateutil import parser as dtparse

from jira_client import JiraClient

ROOT = Path(__file__).resolve().parent
RUN_DATE = dt.date.today().isoformat()
RUN_DIR = ROOT / "reports" / "runs" / RUN_DATE
RUN_DIR.mkdir(parents=True, exist_ok=True)

PROJECT = os.environ.get("JIRA_PROJECT", "LOST")
AI_FIELD = os.environ.get("AI_SP_FIELD", "customfield_13643")
HUMAN_FIELD = os.environ.get("HUMAN_SP_FIELD", "customfield_10023")
IN_PROGRESS_FIELD = os.environ.get("IN_PROGRESS_FIELD", "customfield_10362")
VERIFIED_FIELD = os.environ.get("VERIFIED_FIELD", "customfield_10388")
WINDOW_DAYS = int(os.environ.get("ACCURACY_WINDOW_DAYS", "14"))

ALLOWED_SP = (1, 3, 5, 8, 13, 21)
SP_RANK = {sp: i for i, sp in enumerate(ALLOWED_SP)}


def business_days(a_iso: str, b_iso: str) -> int | None:
    try:
        a = dtparse.parse(a_iso).date()
        b = dtparse.parse(b_iso).date()
    except Exception:
        return None
    if b < a:
        return None
    d, n = a, 0
    while d <= b:
        if d.weekday() < 5:
            n += 1
        d += dt.timedelta(days=1)
    return n


def implied_sp(cycle_days: int, thresholds: list[dict]) -> int:
    for t in thresholds:
        upper = t["if_cycle_days_le"]
        try:
            upper_n = float(upper)
        except (TypeError, ValueError):
            upper_n = float("inf")
        if cycle_days <= upper_n:
            return int(t["then_sp"])
    return int(thresholds[-1]["then_sp"])


def rank_distance(a: int | None, b: int | None) -> int | None:
    if a is None or b is None:
        return None
    return abs(SP_RANK[a] - SP_RANK[b])


def main() -> int:
    baseline = json.loads((ROOT / "baseline_patterns.json").read_text())
    thresholds = baseline["cycle_to_sp_thresholds"]
    jira = JiraClient()

    since = (dt.date.today() - dt.timedelta(days=WINDOW_DAYS)).isoformat()
    jql = (
        f'project = {PROJECT} '
        f'AND issuetype in (Story, Task, Bug) '
        f'AND "First Time In Verified For Production" >= "{since}" '
        f'ORDER BY "First Time In Verified For Production" DESC'
    )
    fields = ["issuetype", "components", IN_PROGRESS_FIELD, VERIFIED_FIELD, HUMAN_FIELD, AI_FIELD]
    print(f"Comparing accuracy on tickets verified since {since}")

    rows: list[dict] = []
    for issue in jira.search(jql, fields):
        f = issue["fields"]
        in_prog = f.get(IN_PROGRESS_FIELD)
        verified = f.get(VERIFIED_FIELD)
        if not in_prog or not verified:
            continue
        cycle = business_days(in_prog, verified)
        if cycle is None or cycle <= 0:
            continue

        human_sp = f.get(HUMAN_FIELD)
        ai_sp = f.get(AI_FIELD)
        try:
            human_sp = int(human_sp) if human_sp is not None else None
        except (TypeError, ValueError):
            human_sp = None
        try:
            ai_sp = int(ai_sp) if ai_sp is not None else None
        except (TypeError, ValueError):
            ai_sp = None

        if human_sp not in ALLOWED_SP:
            human_sp = None
        if ai_sp not in ALLOWED_SP:
            ai_sp = None

        implied = implied_sp(cycle, thresholds)

        rows.append({
            "key": issue["key"],
            "issuetype": (f.get("issuetype") or {}).get("name", ""),
            "components": "|".join(c.get("name", "") for c in (f.get("components") or [])),
            "cycle_business_days": cycle,
            "implied_sp_from_cycle": implied,
            "human_sp": human_sp,
            "ai_sp": ai_sp,
            "human_rank_error": rank_distance(human_sp, implied),
            "ai_rank_error": rank_distance(ai_sp, implied),
        })

    out_csv = RUN_DIR / "accuracy.csv"
    if rows:
        with out_csv.open("w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)

    def mae(values: list[int | None]) -> float | None:
        clean = [v for v in values if v is not None]
        return round(sum(clean) / len(clean), 3) if clean else None

    human_errors = [r["human_rank_error"] for r in rows]
    ai_errors = [r["ai_rank_error"] for r in rows]
    summary = {
        "run_date": RUN_DATE,
        "window_days": WINDOW_DAYS,
        "tickets_verified": len(rows),
        "tickets_with_human_sp": sum(1 for r in rows if r["human_sp"] is not None),
        "tickets_with_ai_sp": sum(1 for r in rows if r["ai_sp"] is not None),
        "human_mae_rank": mae(human_errors),
        "ai_mae_rank": mae(ai_errors),
        "human_exact_pct": (sum(1 for v in human_errors if v == 0) * 100 // max(sum(1 for v in human_errors if v is not None), 1)) if any(v is not None for v in human_errors) else None,
        "ai_exact_pct": (sum(1 for v in ai_errors if v == 0) * 100 // max(sum(1 for v in ai_errors if v is not None), 1)) if any(v is not None for v in ai_errors) else None,
    }
    (RUN_DIR / "accuracy_summary.json").write_text(json.dumps(summary, indent=2))
    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
