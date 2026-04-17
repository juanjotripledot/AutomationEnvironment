"""
GS CTO Monthly Performance Report Generator
============================================
Reads GS_jira_year_report.xlsx (extracted from GS_jira_year_report.zip)
and produces a PDF report replicating the monthly CTO deck:

  Page 1 – Cover
  Page 2 – Time in Development  (histogram: tickets count per rounded days, 3/5/8/13 SP × 2-week + monthly)
  Page 3 – Time in QA           (histogram: tickets count per rounded days, 3/5/8/13 SP × 2-week + monthly)
  Page 4 – Velocity             (sum SP verified for production × 2-week + monthly)
  Page 5 – Effort Distribution  (story points by studio & game × monthly + quarterly)

Usage
-----
  python gs_cto_report.py \
      --xlsx  GS_jira_year_report.xlsx \
      --mapping GamesStudios.xlsx \
      --output  202602_Team_Velocity.pdf \
      [--month 2026-02]

Environment variables (alternative to --mapping):
  MAPPING_SHEET_ID   Google Sheets ID for the GamesStudios mapping sheet
                     (requires GOOGLE_API_KEY or service-account credentials)

Dependencies
------------
  pip install openpyxl matplotlib reportlab requests
"""

import argparse
import json
import math
import os
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import openpyxl
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Paragraph,
    Spacer, Image, Table, TableStyle, NextPageTemplate, PageBreak
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ─── Constants ────────────────────────────────────────────────────────────────

SP_VALID = {3, 5, 8, 13}

DEV_STATUSES = {
    "In Progress",
    "Re-Opened",
    "Code Review",
    "Deployed on Feature branch",
    "In QA feature branch",
    "Verified on feature branch",
}

QA_STATUSES = {"Staging", "In QA"}

PERIODS_BACK = 10        # how many biweekly / monthly periods to show
OUTLIER_IQR_MULT = 2.5   # IQR multiplier for outlier detection

PALETTE = {
    "blue":   "#1e6fc9",
    "green":  "#2da06e",
    "amber":  "#e8a020",
    "red":    "#e84c30",
    "purple": "#8b4fc9",
    "teal":   "#26b3b3",
    "pink":   "#c94f8b",
    "olive":  "#6b8e23",
    "brown":  "#a07050",
    "gray":   "#888780",
}
SP_COLORS = {3: PALETTE["blue"], 5: PALETTE["green"],
             8: PALETTE["amber"], 13: PALETTE["red"]}
STUDIO_COLORS = list(PALETTE.values())
GAME_COLORS   = list(PALETTE.values())

# ─── Date helpers ─────────────────────────────────────────────────────────────

def is_weekend(d: date) -> bool:
    return d.weekday() >= 5


def working_days_between(start: date, end: date) -> float:
    """Count working (Mon–Fri) days from start up to but NOT including end."""
    if end <= start:
        return 0.0
    count = 0
    cur = start
    while cur < end:
        if not is_weekend(cur):
            count += 1
        cur += timedelta(days=1)
    return float(count)


def iso_week_biweekly_label(d: date) -> str:
    """
    Return a biweekly period label based on ISO week number.
    Weeks are paired as (W01,W02)→'26W02', (W03,W04)→'26W04', etc.
    Odd week n is paired with week n+1.
    """
    iso_year, iso_week, _ = d.isocalendar()
    paired = iso_week if iso_week % 2 == 0 else iso_week + 1
    yy = str(iso_year)[-2:]
    return f"{yy}W{paired:02d}"


def monthly_label(d: date) -> str:
    yy = str(d.year)[-2:]
    return f"{yy}M{d.month:02d}"


def quarterly_label(d: date) -> str:
    yy = str(d.year)[-2:]
    q = (d.month - 1) // 3 + 1
    return f"{yy}Q{q}"


def last_completed_month_end(ref: date) -> date:
    """Return the last day of the month before ref's month.
    e.g. ref=2026-02-01 → 2026-01-31
         ref=2026-04-17 → 2026-03-31
    """
    return date(ref.year, ref.month, 1) - timedelta(days=1)


def last_n_biweekly_labels(n: int, ref: date = None) -> list:
    """Return last n distinct biweekly ISO-week labels ending at last completed month.
    Walks back week by week from the last day of the previous month.
    """
    if ref is None:
        ref = date.today()
    anchor = last_completed_month_end(ref)
    labels = []
    seen = set()
    d = anchor
    while len(labels) < n:
        lbl = iso_week_biweekly_label(d)
        if lbl not in seen:
            seen.add(lbl)
            labels.insert(0, lbl)
        d -= timedelta(weeks=1)
    return labels


def last_n_monthly_labels(n: int, ref: date = None) -> list:
    """Return last n monthly labels ending at last completed month.
    e.g. ref=2026-04-17, n=10 → ['25M07','25M08',...,'26M03']
    """
    if ref is None:
        ref = date.today()
    # anchor = last day of previous month
    anchor = last_completed_month_end(ref)
    labels = []
    d = date(anchor.year, anchor.month, 1)
    for _ in range(n):
        labels.insert(0, monthly_label(d))
        d = date(d.year, d.month, 1) - timedelta(days=1)  # go to prev month
        d = date(d.year, d.month, 1)
    return labels


def last_n_quarterly_labels(n: int, ref: date = None) -> list:
    """Return last n quarterly labels ending at last completed quarter."""
    if ref is None:
        ref = date.today()
    anchor = last_completed_month_end(ref)
    labels = []
    seen = set()
    d = anchor
    while len(labels) < n:
        lbl = quarterly_label(d)
        if lbl not in seen:
            seen.add(lbl)
            labels.insert(0, lbl)
        d -= timedelta(weeks=13)
    return labels


# ─── Changelog parser ─────────────────────────────────────────────────────────

def parse_ts(ts_str: str) -> datetime:
    if not ts_str:
        return None
    s = ts_str.strip()
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        return None


def parse_changelog(changelog_json: str) -> dict:
    """
    Parse the Changelog (JSON) column and return:
      devDays            – total working days spent in DEV_STATUSES
      qaDays             – total working days spent in QA_STATUSES
      verifiedForProdDate – date (or None) first entered 'Verified for production'
    """
    result = {"devDays": 0.0, "qaDays": 0.0, "verifiedForProdDate": None}
    if not changelog_json:
        return result

    try:
        histories = json.loads(changelog_json)
    except (json.JSONDecodeError, TypeError):
        return result

    transitions = []
    for h in histories:
        ts = parse_ts(h.get("created", ""))
        if ts is None:
            continue
        for item in h.get("items", []):
            if item.get("field") == "status":
                transitions.append({"ts": ts, "to": item.get("toString", "")})

    transitions.sort(key=lambda x: x["ts"])

    dev_hours = 0.0
    qa_hours  = 0.0
    verified_date = None

    for i, tr in enumerate(transitions):
        status = tr["to"]
        enter  = tr["ts"]
        exit_  = transitions[i + 1]["ts"] if i + 1 < len(transitions) else datetime.now(tz=timezone.utc)

        # working hours between enter and exit
        start_d = enter.date()
        end_d   = exit_.date()
        wdays   = working_days_between(start_d, end_d)

        if status in DEV_STATUSES:
            dev_hours += wdays
        if status in QA_STATUSES:
            qa_hours  += wdays

        if status == "Verified for production" and verified_date is None:
            verified_date = enter.date()

    result["devDays"]             = dev_hours   # already in days (not hours)
    result["qaDays"]              = qa_hours
    result["verifiedForProdDate"] = verified_date
    return result


# ─── Outlier detection ────────────────────────────────────────────────────────

def has_outlier_iqr(values: list) -> bool:
    if len(values) < 4:
        return False
    s = sorted(values)
    q1 = s[len(s) // 4]
    q3 = s[(len(s) * 3) // 4]
    iqr = q3 - q1
    if iqr == 0:
        return False
    return s[-1] > q3 + OUTLIER_IQR_MULT * iqr or s[0] < q1 - OUTLIER_IQR_MULT * iqr


# ─── Load component → studio/game mapping ────────────────────────────────────

def load_mapping(mapping_path: str) -> dict:
    """
    Returns {component_lower: {studio, game}}
    Accepts .xlsx or .csv.
    """
    comp_map = {}
    if not mapping_path or not os.path.exists(mapping_path):
        return comp_map

    if mapping_path.endswith(".csv"):
        import csv
        with open(mapping_path, encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                comp = str(row.get("Component") or row.get("component") or "").strip().lower()
                if comp:
                    comp_map[comp] = {
                        "studio": str(row.get("Studio") or "Other").strip(),
                        "game":   str(row.get("Game")   or "Other").strip(),
                    }
    else:
        wb = openpyxl.load_workbook(mapping_path, read_only=True, data_only=True)
        ws = wb.active
        headers = None
        for row in ws.iter_rows(values_only=True):
            if headers is None:
                headers = [str(c).strip() if c else "" for c in row]
                continue
            if not any(row):
                continue
            rdict = dict(zip(headers, row))
            comp = str(rdict.get("Component") or rdict.get("component") or "").strip().lower()
            if comp:
                comp_map[comp] = {
                    "studio": str(rdict.get("Studio") or "Other").strip(),
                    "game":   str(rdict.get("Game")   or "Other").strip(),
                }
        wb.close()
    return comp_map


def resolve_studio_game(components_str: str, comp_map: dict) -> tuple:
    parts = [p.strip().lower() for p in str(components_str or "").split(",")]
    for p in parts:
        if p in comp_map:
            return comp_map[p]["studio"], comp_map[p]["game"]
        # partial match
        for k, v in comp_map.items():
            if p and (p in k or k in p):
                return v["studio"], v["game"]
    return "Other", "Other"


# ─── Load XLSX ────────────────────────────────────────────────────────────────

def load_tickets(xlsx_path: str, comp_map: dict) -> list:
    """
    Returns list of dicts:
      sp, devDays, qaDays, verifiedForProdDate, studio, game
    Only SP in {3,5,8,13} are included.
    """
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active

    headers = None
    tickets = []
    row_count = 0

    for row in ws.iter_rows(values_only=True):
        if headers is None:
            headers = [str(c).strip() if c is not None else "" for c in row]
            continue
        if not any(row):
            continue

        rdict = dict(zip(headers, row))
        row_count += 1

        # Story points
        sp_raw = rdict.get("Story points") or rdict.get("story points") or rdict.get("SP")
        try:
            sp = int(float(sp_raw))
        except (TypeError, ValueError):
            continue
        if sp not in SP_VALID:
            continue

        # Changelog
        changelog_raw = (
            rdict.get("Changelog (JSON)") or
            rdict.get("Changelog")        or
            rdict.get("changelog")        or ""
        )
        parsed = parse_changelog(str(changelog_raw) if changelog_raw else "")

        # Skip tickets that have 0 dev AND 0 qa AND no verified date
        if (parsed["devDays"] == 0 and parsed["qaDays"] == 0
                and parsed["verifiedForProdDate"] is None):
            continue

        # Studio / game
        components = rdict.get("Components") or rdict.get("components") or ""
        studio, game = resolve_studio_game(str(components), comp_map)

        tickets.append({
            "sp":                   sp,
            "devDays":              parsed["devDays"],
            "qaDays":               parsed["qaDays"],
            "verifiedForProdDate":  parsed["verifiedForProdDate"],
            "studio":               studio,
            "game":                 game,
        })

    wb.close()
    print(f"  Loaded {row_count} rows → {len(tickets)} valid tickets (3/5/8/13 SP)")
    return tickets


# ─── Aggregation ─────────────────────────────────────────────────────────────

def aggregate(tickets: list, ref_date: date = None) -> dict:
    if ref_date is None:
        ref_date = date.today()

    bw_labels = last_n_biweekly_labels(PERIODS_BACK, ref_date)
    m_labels  = last_n_monthly_labels(PERIODS_BACK, ref_date)
    q_labels  = last_n_quarterly_labels(8, ref_date)

    bw_dev  = defaultdict(lambda: defaultdict(list))   # [sp][label] = [days]
    m_dev   = defaultdict(lambda: defaultdict(list))
    bw_qa   = defaultdict(lambda: defaultdict(list))
    m_qa    = defaultdict(lambda: defaultdict(list))
    all_dev = defaultdict(list)    # [sp] = all dev day values (for histogram)
    all_qa  = defaultdict(list)    # [sp] = all qa day values (for histogram)
    bw_vel  = defaultdict(float)   # [label] = sp sum
    m_vel   = defaultdict(float)
    st_mon  = defaultdict(lambda: defaultdict(float))  # [month][studio] = sp
    st_qtr  = defaultdict(lambda: defaultdict(float))
    gm_mon  = defaultdict(lambda: defaultdict(float))
    gm_qtr  = defaultdict(lambda: defaultdict(float))
    st_tot  = defaultdict(float)
    gm_tot  = defaultdict(float)

    for t in tickets:
        sp  = t["sp"]
        vd  = t["verifiedForProdDate"]
        if vd is None:
            continue

        bw_lbl = iso_week_biweekly_label(vd)
        mo_lbl = monthly_label(vd)
        qt_lbl = quarterly_label(vd)

        # Velocity (all valid SP tickets)
        bw_vel[bw_lbl] += sp
        m_vel[mo_lbl]  += sp

        # Dev days (exclude 0)
        if t["devDays"] > 0:
            bw_dev[sp][bw_lbl].append(t["devDays"])
            m_dev[sp][mo_lbl].append(t["devDays"])
            all_dev[sp].append(t["devDays"])

        # QA days (exclude 0)
        if t["qaDays"] > 0:
            bw_qa[sp][bw_lbl].append(t["qaDays"])
            m_qa[sp][mo_lbl].append(t["qaDays"])
            all_qa[sp].append(t["qaDays"])

        # Effort
        st_mon[mo_lbl][t["studio"]] += sp
        st_qtr[qt_lbl][t["studio"]] += sp
        gm_mon[mo_lbl][t["game"]]   += sp
        gm_qtr[qt_lbl][t["game"]]   += sp
        st_tot[t["studio"]]          += sp
        gm_tot[t["game"]]            += sp

    return {
        "bw_labels": bw_labels, "m_labels": m_labels, "q_labels": q_labels,
        "bw_dev": bw_dev, "m_dev": m_dev,
        "bw_qa":  bw_qa,  "m_qa":  m_qa,
        "all_dev": all_dev, "all_qa": all_qa,
        "bw_vel": bw_vel, "m_vel": m_vel,
        "st_mon": st_mon, "st_qtr": st_qtr,
        "gm_mon": gm_mon, "gm_qtr": gm_qtr,
        "st_tot": st_tot, "gm_tot": gm_tot,
    }


# ─── Chart builders ───────────────────────────────────────────────────────────

def _avg(lst):
    return sum(lst) / len(lst) if lst else None


def make_histogram_bar(days_values: list, color: str, title: str,
                       figsize=(5, 2.6), outlier_days=None) -> str:
    """
    Bar chart: X = number of days (rounded integers), Y = number of tickets.
    outlier_days: set of day values to highlight in red.
    """
    if not days_values:
        fig, ax = plt.subplots(figsize=figsize)
        fig.patch.set_facecolor("#EBF3FD")
        ax.set_facecolor("#EBF3FD")
        ax.set_title(title, fontsize=8, pad=4, color="#1a1917")
        ax.text(0.5, 0.5, "No data", transform=ax.transAxes,
                ha="center", va="center", fontsize=8, color="#a09e98")
        path = f"/tmp/chart_{title.replace(' ','_').replace('/','_')[:40]}.png"
        fig.savefig(path, dpi=120, bbox_inches="tight")
        plt.close(fig)
        return path

    # Round all values to nearest integer
    rounded = [int(round(v)) for v in days_values]
    counts = defaultdict(int)
    for d in rounded:
        counts[d] += 1

    all_days = sorted(counts.keys())
    # Fill gaps so bars are contiguous
    min_d, max_d = all_days[0], all_days[-1]
    x_vals = list(range(min_d, max_d + 1))
    y_vals = [counts.get(x, 0) for x in x_vals]

    bar_colors = []
    for x in x_vals:
        if outlier_days and x in outlier_days:
            bar_colors.append("#e84c30")
        else:
            bar_colors.append("#1e6fc9")

    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor("#EBF3FD")
    ax.set_facecolor("#EBF3FD")

    ax.bar(x_vals, y_vals, color=bar_colors, width=0.7, zorder=3,
           edgecolor="white", linewidth=0.5)

    ax.set_xticks(x_vals)
    ax.set_xticklabels([str(x) for x in x_vals], fontsize=7)
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax.tick_params(axis="y", labelsize=7)
    ax.set_title(title, fontsize=8, pad=4, color="#1a1917")
    ax.set_xlabel("Days", fontsize=7, color="#6b6860")
    ax.set_ylabel("Number of tickets", fontsize=7, color="#6b6860")
    ax.grid(axis="y", color="#c8d8ea", linewidth=0.5, linestyle="--", zorder=0)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for sp_ in ["left", "bottom"]:
        ax.spines[sp_].set_color("#b0c4d8")

    plt.tight_layout(pad=0.4)
    path = f"/tmp/chart_{title.replace(' ','_').replace('/','_')[:40]}.png"
    fig.savefig(path, dpi=120, bbox_inches="tight")
    plt.close(fig)
    return path


def make_velocity_line(labels, values, color, title, ylabel,
                       figsize=(5, 2.6)) -> str:
    """Line chart for velocity trend over time periods."""
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor("#EBF3FD")
    ax.set_facecolor("#EBF3FD")

    xs = range(len(labels))
    ys = [v if v is not None else float("nan") for v in values]

    ax.plot(xs, ys, color=color, linewidth=1.8, marker="o",
            markersize=4, zorder=3)

    ax.set_xticks(list(xs))
    ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=7)
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.0f"))
    ax.tick_params(axis="y", labelsize=7)
    ax.set_title(title, fontsize=8, pad=4, color="#1a1917")
    ax.set_ylabel(ylabel, fontsize=7, color="#6b6860")
    ax.grid(axis="y", color="#c8d8ea", linewidth=0.5, linestyle="--")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for sp_ in ["left", "bottom"]:
        ax.spines[sp_].set_color("#b0c4d8")

    plt.tight_layout(pad=0.4)
    path = f"/tmp/chart_{title.replace(' ','_').replace('/','_')[:40]}.png"
    fig.savefig(path, dpi=120, bbox_inches="tight")
    plt.close(fig)
    return path


def make_stacked_bar(labels, series_data: dict, colors_map: list,
                     title, ylabel, figsize=(6, 2.6)) -> str:
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor("#EBF3FD")
    ax.set_facecolor("#EBF3FD")

    xs = list(range(len(labels)))
    bottoms = [0.0] * len(labels)
    for i, (name, vals) in enumerate(series_data.items()):
        color = colors_map[i % len(colors_map)]
        ax.bar(xs, vals, bottom=bottoms, color=color, label=name, width=0.6)
        bottoms = [b + v for b, v in zip(bottoms, vals)]

    ax.set_xticks(xs)
    ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=7)
    ax.tick_params(axis="y", labelsize=7)
    ax.set_title(title, fontsize=8, pad=4, color="#1a1917")
    ax.set_ylabel(ylabel, fontsize=7, color="#6b6860")
    ax.grid(axis="y", color="#c8d8ea", linewidth=0.5, linestyle="--")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for sp_ in ["left", "bottom"]:
        ax.spines[sp_].set_color("#b0c4d8")

    plt.tight_layout(pad=0.4)
    path = f"/tmp/chart_{title.replace(' ','_').replace('/','_')[:40]}.png"
    fig.savefig(path, dpi=120, bbox_inches="tight")
    plt.close(fig)
    return path


def make_donut(labels, values, colors_list, title, figsize=(3.5, 3.0)) -> str:
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor("#EBF3FD")

    total = sum(values)
    pct_labels = [f"{l}\n{v/total*100:.1f}%" if total else l
                  for l, v in zip(labels, values)]

    wedges, texts = ax.pie(
        values, labels=None, colors=colors_list[:len(values)],
        startangle=90, wedgeprops={"width": 0.55, "edgecolor": "white", "linewidth": 0.8}
    )
    ax.legend(wedges, pct_labels, loc="center left", bbox_to_anchor=(1, 0.5),
              fontsize=6.5, frameon=False)
    ax.set_title(title, fontsize=8, pad=6, color="#1a1917")
    plt.tight_layout(pad=0.3)
    path = f"/tmp/chart_{title.replace(' ','_').replace('/','_')[:40]}.png"
    fig.savefig(path, dpi=120, bbox_inches="tight")
    plt.close(fig)
    return path


# ─── PDF builder ──────────────────────────────────────────────────────────────

LIGHT_BLUE  = colors.HexColor("#EBF3FD")
DARK_BLUE   = colors.HexColor("#1F4E79")
MID_BLUE    = colors.HexColor("#185FA5")
TEXT_DARK   = colors.HexColor("#1a1917")
TEXT_MID    = colors.HexColor("#6b6860")
RED_OUTLIER = colors.HexColor("#e84c30")


def build_pdf(agg: dict, output_path: str, report_month: str):
    styles  = getSampleStyleSheet()
    W, H    = A4   # 595.27 x 841.89 pt

    title_style = ParagraphStyle(
        "ReportTitle", fontName="Helvetica-Bold", fontSize=22,
        textColor=DARK_BLUE, alignment=TA_CENTER, spaceAfter=6
    )
    sub_style = ParagraphStyle(
        "ReportSub", fontName="Helvetica", fontSize=11,
        textColor=TEXT_MID, alignment=TA_CENTER, spaceAfter=4
    )
    section_style = ParagraphStyle(
        "SectionHead", fontName="Helvetica-Bold", fontSize=10,
        textColor=DARK_BLUE, spaceBefore=8, spaceAfter=4
    )
    note_style = ParagraphStyle(
        "Note", fontName="Helvetica-Oblique", fontSize=7,
        textColor=TEXT_MID, spaceAfter=2
    )
    outlier_style = ParagraphStyle(
        "Outlier", fontName="Helvetica-Oblique", fontSize=7,
        textColor=RED_OUTLIER, spaceAfter=2
    )

    doc = BaseDocTemplate(
        output_path, pagesize=A4,
        leftMargin=1.8*cm, rightMargin=1.8*cm,
        topMargin=2*cm, bottomMargin=2*cm
    )

    main_frame = Frame(
        doc.leftMargin, doc.bottomMargin,
        W - doc.leftMargin - doc.rightMargin,
        H - doc.topMargin - doc.bottomMargin,
        id="main"
    )

    def header_footer(canvas, doc):
        canvas.saveState()
        # Header bar
        canvas.setFillColor(DARK_BLUE)
        canvas.rect(0, H - 1.2*cm, W, 1.2*cm, fill=1, stroke=0)
        canvas.setFont("Helvetica-Bold", 9)
        canvas.setFillColor(colors.white)
        canvas.drawString(1.8*cm, H - 0.8*cm, "Game Server Report")
        canvas.drawRightString(W - 1.8*cm, H - 0.8*cm, report_month)
        # Footer
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(TEXT_MID)
        canvas.drawCentredString(W/2, 0.8*cm, f"Page {doc.page}")
        canvas.restoreState()

    doc.addPageTemplates([
        PageTemplate(id="main", frames=[main_frame], onPage=header_footer)
    ])

    story = []
    IMG_W_HALF = (W - doc.leftMargin - doc.rightMargin - 0.4*cm) / 2.0
    IMG_W_FULL = W - doc.leftMargin - doc.rightMargin

    def img(path, width=None):
        width = width or IMG_W_HALF
        from PIL import Image as PILImage
        with PILImage.open(path) as im:
            iw, ih = im.size
        height = width * ih / iw
        return Image(path, width=width, height=height)

    def two_charts(left_path, right_path):
        t = Table(
            [[img(left_path, IMG_W_HALF), img(right_path, IMG_W_HALF)]],
            colWidths=[IMG_W_HALF + 0.2*cm, IMG_W_HALF + 0.2*cm]
        )
        t.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP"),
                                ("LEFTPADDING",  (0,0), (-1,-1), 0),
                                ("RIGHTPADDING", (0,0), (-1,-1), 0)]))
        story.append(t)

    # ── Cover ────────────────────────────────────────────────────────────────
    story.append(Spacer(1, 3*cm))
    story.append(Paragraph("Game Server Report", title_style))
    story.append(Paragraph(f"{report_month} · Metrics", sub_style))
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph(
        "This report provides a comprehensive overview of key performance metrics "
        "across the GameServer engineering teams. Pages cover: Time in Development, "
        "Time in QA, Team Velocity, and Effort Distribution by studio and game.",
        ParagraphStyle("Intro", fontName="Helvetica", fontSize=9,
                       textColor=TEXT_DARK, leading=14, alignment=TA_LEFT)
    ))
    story.append(PageBreak())

    # ── Page 2: Time in Development ──────────────────────────────────────────
    story.append(Paragraph("Time in Development", section_style))
    story.append(Paragraph(
        "Distribution of tickets by days spent in development statuses "
        "(In Progress, Re-Opened, Code Review, Deployed on Feature branch, "
        "In QA feature branch, Verified on feature branch). "
        "X axis = rounded working days, Y axis = number of tickets. "
        "Red bars indicate outlier day values (IQR × 2.5). "
        "Left column: tickets with Verified for Production date in last 10 biweekly periods. "
        "Right column: last 10 monthly periods.",
        note_style
    ))

    bw_labels = agg["bw_labels"]
    m_labels  = agg["m_labels"]

    def outlier_day_values(days_list):
        """Return set of rounded day values that are outliers."""
        if len(days_list) < 4:
            return set()
        rounded = [int(round(v)) for v in days_list]
        s = sorted(rounded)
        q1 = s[len(s) // 4]
        q3 = s[(len(s) * 3) // 4]
        iqr = q3 - q1
        if iqr == 0:
            return set()
        return {x for x in set(rounded) if x > q3 + OUTLIER_IQR_MULT * iqr
                or x < q1 - OUTLIER_IQR_MULT * iqr}

    for sp in [3, 5, 8, 13]:
        color = SP_COLORS[sp]
        # Collect all dev days for tickets in biweekly window
        bw_days = []
        for lbl in bw_labels:
            bw_days.extend(agg["bw_dev"][sp].get(lbl, []))
        # Collect all dev days for tickets in monthly window
        mo_days = []
        for lbl in m_labels:
            mo_days.extend(agg["m_dev"][sp].get(lbl, []))

        bw_out = outlier_day_values(bw_days)
        mo_out = outlier_day_values(mo_days)

        p_bw = make_histogram_bar(
            bw_days, color,
            f"Time in Dev – {sp} SP (last {PERIODS_BACK} biweekly periods)",
            outlier_days=bw_out
        )
        p_mo = make_histogram_bar(
            mo_days, color,
            f"Time in Dev – {sp} SP (last {PERIODS_BACK} months)",
            outlier_days=mo_out
        )
        story.append(Spacer(1, 0.15*cm))
        story.append(Paragraph(f"{sp} SP", ParagraphStyle(
            f"SPLabel_{sp}", fontName="Helvetica-Bold", fontSize=8,
            textColor=colors.HexColor(color))))
        if bw_out or mo_out:
            all_out = bw_out | mo_out
            story.append(Paragraph(
                f"⚠ Outlier day values detected: {', '.join(str(d) for d in sorted(all_out))} days",
                outlier_style))
        two_charts(p_bw, p_mo)

    story.append(PageBreak())

    # ── Page 3: Time in QA ───────────────────────────────────────────────────
    story.append(Paragraph("Time in QA", section_style))
    story.append(Paragraph(
        "Distribution of tickets by days spent in QA statuses (Staging, In QA). "
        "X axis = rounded working days, Y axis = number of tickets. "
        "Red bars indicate outlier day values (IQR × 2.5).",
        note_style
    ))

    for sp in [3, 5, 8, 13]:
        color = SP_COLORS[sp]
        bw_days = []
        for lbl in bw_labels:
            bw_days.extend(agg["bw_qa"][sp].get(lbl, []))
        mo_days = []
        for lbl in m_labels:
            mo_days.extend(agg["m_qa"][sp].get(lbl, []))

        bw_out = outlier_day_values(bw_days)
        mo_out = outlier_day_values(mo_days)

        p_bw = make_histogram_bar(
            bw_days, color,
            f"Time in QA – {sp} SP (last {PERIODS_BACK} biweekly periods)",
            outlier_days=bw_out
        )
        p_mo = make_histogram_bar(
            mo_days, color,
            f"Time in QA – {sp} SP (last {PERIODS_BACK} months)",
            outlier_days=mo_out
        )
        story.append(Spacer(1, 0.15*cm))
        story.append(Paragraph(f"{sp} SP", ParagraphStyle(
            f"SPLabelQA_{sp}", fontName="Helvetica-Bold", fontSize=8,
            textColor=colors.HexColor(color))))
        if bw_out or mo_out:
            all_out = bw_out | mo_out
            story.append(Paragraph(
                f"⚠ Outlier day values detected: {', '.join(str(d) for d in sorted(all_out))} days",
                outlier_style))
        two_charts(p_bw, p_mo)

    story.append(PageBreak())

    # ── Page 4: Velocity ─────────────────────────────────────────────────────
    story.append(Paragraph("Team Velocity", section_style))
    story.append(Paragraph(
        "Sum of story points from tickets reaching 'Verified for production'. "
        "Left: biweekly (ISO week pairs). Right: monthly.",
        note_style
    ))

    bw_vel_vals = [agg["bw_vel"].get(l, 0) for l in bw_labels]
    mo_vel_vals = [agg["m_vel"].get(l, 0)  for l in m_labels]

    p_bw_vel = make_velocity_line(
        bw_labels, bw_vel_vals, PALETTE["blue"],
        "Velocity per 2 Weeks (To Verified for Production)",
        "Num Story Points", figsize=(5.5, 3.0)
    )
    p_mo_vel = make_velocity_line(
        m_labels, mo_vel_vals, PALETTE["blue"],
        "Velocity per Month (To Verified for Production)",
        "Num Story Points", figsize=(5.5, 3.0)
    )
    two_charts(p_bw_vel, p_mo_vel)
    story.append(PageBreak())

    # ── Page 5: Effort Distribution ──────────────────────────────────────────
    story.append(Paragraph("Effort Distribution", section_style))
    story.append(Paragraph(
        "Story points released per studio and per game, monthly and quarterly.",
        note_style
    ))

    studios = sorted(agg["st_tot"], key=lambda s: -agg["st_tot"][s])
    games   = sorted(agg["gm_tot"], key=lambda g: -agg["gm_tot"][g])[:10]
    st_colors_list = [STUDIO_COLORS[i % len(STUDIO_COLORS)] for i in range(len(studios))]
    gm_colors_list = [GAME_COLORS[i % len(GAME_COLORS)]    for i in range(len(games))]

    # Studio donut (global)
    if studios:
        st_vals = [agg["st_tot"][s] for s in studios]
        p_donut = make_donut(studios, st_vals, st_colors_list,
                             "Effort per Studio (Global)", figsize=(4.5, 3.2))
        story.append(img(p_donut, IMG_W_HALF))

    # Studio monthly bar
    if studios and m_labels:
        st_mon_series = {
            s: [agg["st_mon"].get(l, {}).get(s, 0) for l in m_labels]
            for s in studios
        }
        p_st_m = make_stacked_bar(
            m_labels, st_mon_series, st_colors_list,
            "Effort per Studio – Monthly", "Story Points"
        )
        # Studio quarterly bar
        q_labels = agg["q_labels"]
        st_qtr_series = {
            s: [agg["st_qtr"].get(l, {}).get(s, 0) for l in q_labels]
            for s in studios
        }
        p_st_q = make_stacked_bar(
            q_labels, st_qtr_series, st_colors_list,
            "Effort per Studio – Quarterly", "Story Points"
        )
        two_charts(p_st_m, p_st_q)

    # Game monthly bar
    if games and m_labels:
        gm_mon_series = {
            g: [agg["gm_mon"].get(l, {}).get(g, 0) for l in m_labels]
            for g in games
        }
        p_gm_m = make_stacked_bar(
            m_labels, gm_mon_series, gm_colors_list,
            "Effort per Game – Monthly", "Story Points"
        )
        gm_qtr_series = {
            g: [agg["gm_qtr"].get(l, {}).get(g, 0) for l in q_labels]
            for g in games
        }
        p_gm_q = make_stacked_bar(
            q_labels, gm_qtr_series, gm_colors_list,
            "Effort per Game – Quarterly", "Story Points"
        )
        two_charts(p_gm_m, p_gm_q)

    doc.build(story)
    print(f"  PDF written → {output_path}")


# ─── CLI ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Generate CTO monthly performance report PDF from Jira XLSX export."
    )
    parser.add_argument(
        "--xlsx", required=True,
        help="Path to GS_jira_year_report.xlsx (extracted from zip)"
    )
    parser.add_argument(
        "--mapping", default=None,
        help="Path to GamesStudios mapping file (.xlsx or .csv). "
             "Columns: Component, Studio, Game"
    )
    parser.add_argument(
        "--output", default="GS_CTO_Report.pdf",
        help="Output PDF filename (default: GS_CTO_Report.pdf)"
    )
    parser.add_argument(
        "--month", default=None,
        help="Report month label e.g. '2026-02' (default: current month)"
    )
    args = parser.parse_args()

    # Determine report month label and ref_date
    # ref_date must be the first day of the month AFTER the report month,
    # so that last_completed_month_end(ref_date) == last day of report month.
    if args.month:
        try:
            dt = datetime.strptime(args.month, "%Y-%m")
            report_month = dt.strftime("%B %Y")
            # ref = first day of the month after the report month
            next_month = dt.month % 12 + 1
            next_year  = dt.year + (1 if dt.month == 12 else 0)
            ref_date   = date(next_year, next_month, 1)
        except ValueError:
            report_month = args.month
            ref_date = date.today()
    else:
        ref_date     = date.today()
        report_month = last_completed_month_end(ref_date).strftime("%B %Y")

    print(f"\nGS CTO Report Generator")
    print(f"  Report month : {report_month}")
    print(f"  XLSX         : {args.xlsx}")
    print(f"  Mapping      : {args.mapping or '(none — studios/games = Other)'}")
    print(f"  Output       : {args.output}\n")

    print("Loading component → studio/game mapping...")
    comp_map = load_mapping(args.mapping)
    print(f"  {len(comp_map)} component mappings loaded")

    print("Loading Jira tickets...")
    tickets = load_tickets(args.xlsx, comp_map)

    print("Aggregating metrics...")
    agg = aggregate(tickets, ref_date)

    # Quick summary
    total_sp = sum(agg["m_vel"].values())
    print(f"  Total SP verified for production (all time in dataset): {total_sp:.0f}")
    print(f"  Biweekly periods: {agg['bw_labels']}")
    print(f"  Monthly periods:  {agg['m_labels']}")

    print("Building PDF...")
    build_pdf(agg, args.output, report_month)

    print("\nDone.")


if __name__ == "__main__":
    main()
