"""Post a short Slack digest with the latest run's results.

Reads the two summary JSONs written by estimate.py and compare.py from today's
run folder, then posts via incoming webhook.

Channel: set in the webhook itself when you create it in the Slack admin.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import sys
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent
RUN_DATE = dt.date.today().isoformat()
RUN_DIR = ROOT / "reports" / "runs" / RUN_DATE


def fmt_mae(v: float | None) -> str:
    return f"{v:.2f}" if isinstance(v, (int, float)) else "—"


def main() -> int:
    webhook = os.environ.get("SLACK_WEBHOOK_URL")
    if not webhook:
        print("SLACK_WEBHOOK_URL not set — skipping Slack digest.")
        return 0

    est_path = RUN_DIR / "summary.json"
    acc_path = RUN_DIR / "accuracy_summary.json"
    est = json.loads(est_path.read_text()) if est_path.exists() else {}
    acc = json.loads(acc_path.read_text()) if acc_path.exists() else {}

    by_sp = est.get("by_sp", {})
    by_conf = est.get("by_confidence", {})

    n_est = est.get("n_evaluated", 0)
    n_wrote = est.get("n_written", 0)
    n_verified = acc.get("tickets_verified", 0)
    human_mae = acc.get("human_mae_rank")
    ai_mae = acc.get("ai_mae_rank")
    human_pct = acc.get("human_exact_pct")
    ai_pct = acc.get("ai_exact_pct")

    verdict = "—"
    if isinstance(human_mae, (int, float)) and isinstance(ai_mae, (int, float)):
        if ai_mae < human_mae:
            verdict = f":white_check_mark: AI was *more* accurate (Δ = {round(human_mae - ai_mae, 2)} ranks)"
        elif ai_mae > human_mae:
            verdict = f":warning: AI was *less* accurate (Δ = {round(ai_mae - human_mae, 2)} ranks)"
        else:
            verdict = ":scales: AI tied with humans"

    blocks = [
        {
            "type": "header",
            "text": {"type": "plain_text", "text": f"LOST AI Estimation — {RUN_DATE}"},
        },
        {
            "type": "section",
            "fields": [
                {"type": "mrkdwn", "text": f"*New estimates*\n{n_est} tickets · {n_wrote} written to Jira"},
                {"type": "mrkdwn", "text": f"*Verified since last window*\n{n_verified} tickets"},
            ],
        },
        {
            "type": "section",
            "fields": [
                {"type": "mrkdwn", "text": f"*Confidence*\nhigh: {by_conf.get('high', 0)}  ·  med: {by_conf.get('medium', 0)}  ·  low: {by_conf.get('low', 0)}"},
                {"type": "mrkdwn", "text": f"*By SP*\n" + " · ".join(f"{sp}={by_sp.get(str(sp), 0) if isinstance(list(by_sp.keys())[0] if by_sp else '', str) else by_sp.get(sp, 0)}" for sp in (1, 3, 5, 8, 13, 21))},
            ],
        },
        {
            "type": "section",
            "fields": [
                {"type": "mrkdwn", "text": f"*Human MAE (rank)*\n{fmt_mae(human_mae)}  ·  exact: {human_pct if human_pct is not None else '—'}%"},
                {"type": "mrkdwn", "text": f"*AI MAE (rank)*\n{fmt_mae(ai_mae)}  ·  exact: {ai_pct if ai_pct is not None else '—'}%"},
            ],
        },
        {
            "type": "context",
            "elements": [{"type": "mrkdwn", "text": verdict}],
        },
    ]

    payload = {"text": f"LOST AI Estimation — {RUN_DATE}", "blocks": blocks}
    r = requests.post(webhook, json=payload, timeout=10)
    if r.status_code >= 300:
        print(f"Slack post failed: {r.status_code} {r.text[:200]}")
        return 1
    print("Slack digest posted.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
