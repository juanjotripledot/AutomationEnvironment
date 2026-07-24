# LOST AI Estimation

A GitHub Actions workflow that auto-estimates Story Points for unestimated tickets in the LOST (Hawk LiveOps) Jira project, using Anthropic's Claude calibrated against the team's actual cycle-time history.

## What it does, each run

The workflow fires every Sunday and Tuesday at 06:00 UTC (= 08:00 Madrid CEST / 07:00 CET):

1. **Estimates** every open LOST `Story`/`Task`/`Bug` ticket where the human Story Points field is empty. For each, it sends Claude the ticket's summary + description + labels + components plus the 12-month cycle-time baseline (`baseline_patterns.json`). The model returns one Fibonacci value (1, 3, 5, 8, 13, 21), a confidence, and a one-sentence rationale.
2. **Writes** the value to `customfield_13643` ("Story Points IA estimated"). The human Story Points field (`customfield_10023`) is never touched.
3. **Compares** AI vs human accuracy for tickets that reached Verified for Production in the trailing 14 days. The "ground truth" is the SP implied by each ticket's actual cycle time (via the threshold table in the baseline).
4. **Posts** a digest to Slack and commits the run's CSVs to `reports/runs/<YYYY-MM-DD>/` for history.

## Quick setup (10–15 minutes)

### 1. Drop these files into the repo

Place the bundle at the root of `tripledotstudios/<your-hawk-liveops-repo>`:

```
.github/workflows/lost-ai-estimation.yml
scripts/lost_estimation/
  ├── estimate.py
  ├── compare.py
  ├── slack_digest.py
  ├── jira_client.py
  ├── baseline_patterns.json
  ├── requirements.txt
  └── README.md   (this file)
```

### 2. Create a Jira service account + API token

1. Atlassian admin → invite a new user, e.g. `ai-estimation@tripledotstudios.com`.
2. Grant it `Browse Projects` and `Edit Issues` on the LOST project.
3. Log in as that user at <https://id.atlassian.com/manage-profile/security/api-tokens> and create an API token. Copy it once — Atlassian won't show it again.

### 3. Get an Anthropic API key

1. Check first if Tripledot already has an organisation-level Anthropic account; if so, ask for a workspace + key with `claude-sonnet-4-6` access.
2. Otherwise create one at <https://console.anthropic.com/>. The job uses ~3-5K input tokens per ticket × 50–200 tickets/run × 2 runs/week. Budget ≤ $5/month at typical volumes.

### 4. Create the Slack webhook

In Slack: app directory → "Incoming Webhooks" → add to channel → copy the webhook URL. The webhook is bound to one channel; pick whichever channel should receive the digest (the workflow does not need to know the channel name — the webhook carries it).

### 5. Add GitHub secrets

Repo → Settings → Secrets and variables → Actions → New repository secret. Add four:

| Secret name | Value |
|---|---|
| `ANTHROPIC_API_KEY` | from step 3 |
| `JIRA_EMAIL` | the service account email |
| `JIRA_API_TOKEN` | from step 2 |
| `SLACK_WEBHOOK_URL` | from step 4 |

### 6. First run — dry run

Actions tab → `LOST AI Estimation` → `Run workflow` → tick **Dry run** → Run. The job will pull tickets, call Claude, log what *would* be written, and post nothing to Slack or Jira. Inspect the run artifact (`scripts/lost_estimation/reports/runs/<date>/estimates.csv`) before enabling live writes.

### 7. Enable live runs

Once dry-run output looks sensible, do nothing — the cron schedule is already active. The next Sunday or Tuesday at 06:00 UTC the workflow will fire for real.

## How estimation works

The Python step builds a prompt for Claude that includes:

- Per-SP-bucket cycle-time stats (median, p90-capped mean, sample size).
- Per-SP-bucket cycle-time stats *split by issue type*. This is critical: in the team's data, Bugs cycle at ~0.4× the speed of Stories at the same SP, so a "5-SP bug" tends to actually finish in 4 days, not 9.
- The cycle-time → SP threshold table derived from the team's medians.
- The current ticket's summary, description, labels, components, and a small set of description-shape features (length band, presence of acceptance criteria / steps to reproduce / code blocks).

Claude is instructed to return JSON only, with `proposed_sp`, `predicted_cycle_business_days`, `confidence`, and `reasoning`. If it returns anything not in {1, 3, 5, 8, 13, 21}, the script snaps to the nearest Fibonacci value.

## Tuning knobs

All set via environment variables in the workflow YAML:

| Variable | Default | Meaning |
|---|---|---|
| `MAX_TICKETS_PER_RUN` | `200` | Hard cap so a runaway backlog doesn't burn budget. |
| `ACCURACY_WINDOW_DAYS` | `14` | How far back `compare.py` looks for newly-verified tickets. |
| `DRY_RUN` | `false` | When `true`, no Jira writes and no Slack post. |

## Updating the baseline

`baseline_patterns.json` is a frozen snapshot from 2026-05-19. As more tickets accumulate, you'll want to refresh it — probably quarterly. To regenerate:

1. Pull the latest 12 months from Jira (the original analysis lived in Cowork; the procedure is in the project Confluence page).
2. Re-run the cycle-time aggregation script to produce a new `baseline_patterns.json`.
3. Commit the new JSON. Next workflow run picks it up automatically.

A quarterly refresh job could be added as a second workflow if you want it automated.

## Schedule and DST

Cron is UTC. `0 6 * * 0,2` = Sun + Tue 06:00 UTC, which is 08:00 Madrid local during CEST (late March–late October). During CET (winter) it fires at 07:00 Madrid. If you need the run to always land at exactly 08:00 Madrid, either use a self-hosted runner with `TZ=Europe/Madrid`, or accept the one-hour shift and adjust the cron expression twice a year.

## Idempotency and safety

- `estimate.py` skips tickets whose `customfield_13643` already holds a Fibonacci value, so re-running the same day is safe.
- The human Story Points field is never touched.
- All writes use the LOST project; the workflow has no permission elsewhere because the service account scope is project-bounded.
- Each Slack post links to the run's commit, so you can audit any value the bot wrote.

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `401 Unauthorized` from Jira | API token expired or scoped to the wrong account. Regenerate. |
| `429 Too Many Requests` from Anthropic | Lower `MAX_TICKETS_PER_RUN` or upgrade the plan. |
| Slack post returns 200 but no message | Webhook deactivated or channel renamed. Recreate. |
| Workflow runs but `estimates.csv` is empty | No open tickets matched the JQL — check the LOST backlog has empty-SP Story/Task/Bug tickets. |
