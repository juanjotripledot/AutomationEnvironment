"""LOST AI estimation step.

Pulls every open LOST Story/Task/Bug with empty Story Points, asks Claude for a
Fibonacci value using the 12-month baseline patterns, and writes the result to
customfield_13643 ("Story Points IA estimated").

Outputs:
  reports/runs/<YYYY-MM-DD>/estimates.csv  — one row per ticket estimated
  reports/runs/<YYYY-MM-DD>/summary.json   — counts by SP and confidence
"""

from __future__ import annotations

import csv
import datetime as dt
import json
import os
import re
import sys
from pathlib import Path

from anthropic import Anthropic

from jira_client import JiraClient

ROOT = Path(__file__).resolve().parent
RUN_DATE = dt.date.today().isoformat()
RUN_DIR = ROOT / "reports" / "runs" / RUN_DATE
RUN_DIR.mkdir(parents=True, exist_ok=True)

PROJECT = os.environ.get("JIRA_PROJECT", "LOST")
AI_FIELD = os.environ.get("AI_SP_FIELD", "customfield_13643")
HUMAN_FIELD = os.environ.get("HUMAN_SP_FIELD", "customfield_10023")
DRY_RUN = os.environ.get("DRY_RUN", "false").lower() == "true"
MAX_TICKETS = int(os.environ.get("MAX_TICKETS_PER_RUN", "200"))

ALLOWED_SP = (1, 3, 5, 8, 13, 21)


def load_baseline() -> dict:
    return json.loads((ROOT / "baseline_patterns.json").read_text())


def description_shape(text: str) -> dict[str, bool | str | int]:
    """Lightweight features extracted from the ticket description."""
    if not text:
        return {"has_ac": False, "has_steps": False, "has_code_blocks": False, "length_band": "empty", "length": 0}
    has_ac = bool(re.search(r"(?i)(AC0?\d|acceptance criteria|given.{0,30}when.{0,30}then)", text))
    has_steps = bool(re.search(r"(?i)(steps to reproduce|preconditions:|actual result:|expected result:)", text))
    has_code = text.count("```") >= 2 or text.count("{{") >= 2
    n = len(text.strip())
    if n < 200:
        band = "short"
    elif n < 1500:
        band = "medium"
    else:
        band = "long"
    return {"has_ac": has_ac, "has_steps": has_steps, "has_code_blocks": has_code, "length_band": band, "length": n}


def nearest_fibonacci(x: float | int) -> int:
    return min(ALLOWED_SP, key=lambda f: abs(f - x))


def make_prompt(baseline: dict, ticket: dict) -> str:
    """Construct the user prompt for Claude. The system prompt is short and fixed."""
    per_sp = baseline["per_sp"]
    per_type = baseline["per_sp_per_issuetype"]
    thresholds = baseline["cycle_to_sp_thresholds"]
    adjustments = baseline["issue_type_adjustments"]

    rows = []
    for sp in ("1", "3", "5", "8", "13", "21"):
        s = per_sp[sp]
        flag = " (extrapolated)" if s.get("extrapolated") else ""
        rows.append(f"  SP {sp}: median={s['median_business_days']}d, p90-mean={s['mean_p90_capped']}d, p90={s['p90']}d, n={s['sample_size']}{flag}")

    by_type_lines = []
    for it in ("Story", "Task", "Bug"):
        cells = []
        for sp in ("1", "3", "5", "8", "13", "21"):
            s = per_type[it][sp]
            mark = "*" if s.get("extrapolated") else ""
            cells.append(f"SP{sp}={s['median_business_days']}d{mark}")
        by_type_lines.append(f"  {it} ({adjustments[it]}x): " + ", ".join(cells))

    thr_lines = [f"  ≤ {t['if_cycle_days_le']} bus.days → SP {t['then_sp']} ({t['anchor']})" for t in thresholds]

    shape = description_shape(ticket["description"])

    return f"""You are estimating Story Points for one Jira ticket in the LOST (Hawk LiveOps) project.

The team uses a Fibonacci scale: 1, 3, 5, 8, 13, 21. Pick exactly ONE of these.

## What the team's data says

Per-SP median business days from In Progress → Verified for Production (last 12 months, n={baseline['n_usable']} usable tickets):
{chr(10).join(rows)}

Median cycle by issue type and SP (asterisk = extrapolated, no observed data):
{chr(10).join(by_type_lines)}

Issue-type adjustment multipliers (relative to overall median): {adjustments}.
A Bug at "5 SP complexity" tends to actually finish in ~40% the time a Story at the same complexity does.

Cycle-time → SP mapping derived from team data:
{chr(10).join(thr_lines)}

## The ticket

Key: {ticket['key']}
Issue type: {ticket['issuetype']}
Components: {', '.join(ticket['components']) or '(none)'}
Labels: {', '.join(ticket['labels']) or '(none)'}
Status: {ticket['status']}

Description shape: length={shape['length']} chars ({shape['length_band']}), has_acceptance_criteria={shape['has_ac']}, has_steps_to_reproduce={shape['has_steps']}, has_code_blocks={shape['has_code_blocks']}.

Summary:
{ticket['summary']}

Description:
{ticket['description'] or '(no description)'}

## Your job

Estimate the SP for this ticket. Reason briefly about:
1. The likely cycle time in business days (use the patterns above as priors).
2. The issue type adjustment.
3. Whether the description is detailed enough to estimate confidently.

Then output STRICTLY this JSON, nothing else:

{{
  "predicted_cycle_business_days": <number>,
  "proposed_sp": <one of 1, 3, 5, 8, 13, 21>,
  "confidence": "<low | medium | high>",
  "reasoning": "<2-3 sentence justification>"
}}

Confidence rules:
- high: description has clear AC or repro steps, scope is unambiguous.
- medium: description is present but partial.
- low: no description, or one-line summary only.

If unsure between two SP values, pick the smaller one and lower confidence to low.
"""


def call_claude(client: Anthropic, baseline: dict, ticket: dict) -> dict:
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=600,
        system="You are a precise software project estimator. Always reply with valid JSON only, no surrounding prose.",
        messages=[{"role": "user", "content": make_prompt(baseline, ticket)}],
    )
    raw = msg.content[0].text.strip()
    # Best-effort JSON extraction in case the model wraps in code fences.
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.DOTALL)
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        parsed = json.loads(m.group(0)) if m else {}
    sp = parsed.get("proposed_sp")
    if sp not in ALLOWED_SP:
        # Snap to nearest legal value
        try:
            sp = nearest_fibonacci(float(sp))
        except (TypeError, ValueError):
            sp = 3
    return {
        "predicted_cycle_business_days": parsed.get("predicted_cycle_business_days"),
        "proposed_sp": sp,
        "confidence": parsed.get("confidence", "low"),
        "reasoning": (parsed.get("reasoning") or "")[:500],
    }


def main() -> int:
    baseline = load_baseline()
    jira = JiraClient()
    claude = Anthropic()

    jql = (
        f'project = {PROJECT} '
        f'AND issuetype in (Story, Task, Bug) '
        f'AND "Story Points" is EMPTY '
        f'AND statusCategory != Done '
        f'ORDER BY updated DESC'
    )
    fields = ["summary", "description", "labels", "components", "issuetype", "status", AI_FIELD, HUMAN_FIELD]
    print(f"Querying: {jql}")

    rows: list[dict] = []
    pulled = 0
    for issue in jira.search(jql, fields):
        if pulled >= MAX_TICKETS:
            break
        pulled += 1
        f = issue["fields"]
        ticket = {
            "key": issue["key"],
            "summary": f.get("summary", ""),
            "description": JiraClient.adf_to_text(f.get("description")),
            "issuetype": (f.get("issuetype") or {}).get("name", ""),
            "components": [c.get("name", "") for c in (f.get("components") or [])],
            "labels": f.get("labels") or [],
            "status": (f.get("status") or {}).get("name", ""),
            "current_ai_sp": f.get(AI_FIELD),
        }

        # Skip tickets we've already estimated this cycle (idempotency).
        if ticket["current_ai_sp"] in ALLOWED_SP:
            print(f"  - {ticket['key']} already has AI estimate {ticket['current_ai_sp']}; skipping")
            continue

        print(f"  estimating {ticket['key']} ({ticket['issuetype']}) — {ticket['summary'][:60]}")
        try:
            estimate = call_claude(claude, baseline, ticket)
        except Exception as e:
            print(f"    Claude error: {e}")
            estimate = {"proposed_sp": None, "confidence": "error", "reasoning": str(e)[:300], "predicted_cycle_business_days": None}

        wrote = False
        if estimate["proposed_sp"] in ALLOWED_SP and not DRY_RUN:
            wrote = jira.set_field(ticket["key"], AI_FIELD, estimate["proposed_sp"])
        elif estimate["proposed_sp"] in ALLOWED_SP and DRY_RUN:
            print(f"    [DRY RUN] would write {AI_FIELD}={estimate['proposed_sp']} to {ticket['key']}")

        rows.append({
            "key": ticket["key"],
            "issuetype": ticket["issuetype"],
            "summary": ticket["summary"][:200],
            "components": "|".join(ticket["components"]),
            "labels": "|".join(ticket["labels"]),
            "status": ticket["status"],
            "predicted_cycle_business_days": estimate["predicted_cycle_business_days"],
            "proposed_sp": estimate["proposed_sp"],
            "confidence": estimate["confidence"],
            "reasoning": estimate["reasoning"],
            "written_to_jira": wrote,
            "dry_run": DRY_RUN,
        })

    # Save CSV
    out_csv = RUN_DIR / "estimates.csv"
    if rows:
        with out_csv.open("w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)

    # Save summary JSON
    summary = {
        "run_date": RUN_DATE,
        "n_evaluated": len(rows),
        "n_written": sum(1 for r in rows if r["written_to_jira"]),
        "dry_run": DRY_RUN,
        "by_sp": {sp: sum(1 for r in rows if r["proposed_sp"] == sp) for sp in ALLOWED_SP},
        "by_confidence": {c: sum(1 for r in rows if r["confidence"] == c) for c in ("high", "medium", "low", "error")},
    }
    (RUN_DIR / "summary.json").write_text(json.dumps(summary, indent=2))

    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
