#!/usr/bin/env python3
"""
GS CTO Monthly Performance Report
===================================
Generates a 5-page PDF from a Jira XLSX export.

Pages:
  1. Cover
  2. Time in Development  — histograms (days → ticket count) per SP × period
  3. Time in QA           — histograms per SP × period
  4. Team Velocity        — SP delivered per biweekly + monthly period
  5. Effort Distribution  — donut + stacked bars by Studio and Game

Usage:
  python gs_cto_report.py --xlsx GS_jira_year_report.xlsx \
                          --mapping GamesStudios.csv \
                          --output GS_CTO_Report.pdf \
                          [--month 2026-05]
"""

import argparse
import csv
import json
import math
import numpy as np
import os
import sys
import tempfile
from collections import defaultdict
from datetime import date, datetime, timedelta

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
    BaseDocTemplate, Frame, Image, PageBreak, PageTemplate,
    Paragraph, Spacer, Table, TableStyle,
)

# ── Constants ─────────────────────────────────────────────────────────────────
PERIODS_BACK_BIWEEKLY  = 10
PERIODS_BACK_MONTHLY   = 10
PERIODS_BACK_QUARTERLY = 8

SP_SIZES    = [3, 5, 8, 13]
BRAND_BLUE  = "#1E6FA5"
BRAND_LIGHT = "#EBF3FD"
BRAND_DARK  = "#0D3B5E"
ACCENT      = "#E8A020"

def _wd(start, end):
    """Working days between two datetime/date objects."""
    if start is None or end is None:
        return None
    if isinstance(start, datetime):
        start = start.date()
    if isinstance(end, datetime):
        end = end.date()
    if end < start:
        return None
    return float(np.busday_count(start, end))

STUDIO_COLORS = {
    "Puzzle":    "#378ADD",
    "Word":      "#5CB85C",
    "Platform":  "#E8A020",
    "Multiple":  "#9B59B6",
    "Other":     "#95A5A6",
}

GAME_COLORS_LIST = [
    "#378ADD","#5CB85C","#E8A020","#9B59B6","#E74C3C",
    "#1ABC9C","#F39C12","#2ECC71","#3498DB","#E91E63",
]

IMG_W_FULL  = 17 * cm
IMG_W_HALF  = 8.2 * cm
IMG_H_CHART = 5.5 * cm

# ── Date helpers ──────────────────────────────────────────────────────────────
def iso_week_biweekly_label(d: date) -> str:
    """Return YYWnn label where nn is the odd week of the biweekly pair."""
    iso_year, iso_week, _ = d.isocalendar()
    pair_week = iso_week if iso_week % 2 == 1 else iso_week - 1
    return f"{str(iso_year)[-2:]}W{pair_week:02d}"

def monthly_label(d: date) -> str:
    return f"{str(d.year)[-2:]}M{d.month:02d}"

def quarterly_label(d: date) -> str:
    q = (d.month - 1) // 3 + 1
    return f"{str(d.year)[-2:]}Q{q}"

def last_n_biweekly_labels(n: int, anchor: date) -> list:
    labels = []
    seen = set()
    d = anchor
    while len(labels) < n:
        lbl = iso_week_biweekly_label(d)
        if lbl not in seen:
            seen.add(lbl)
            labels.append(lbl)
        d -= timedelta(days=7)
    return list(reversed(labels))

def last_n_monthly_labels(n: int, anchor: date) -> list:
    labels = []
    yr, mo = anchor.year, anchor.month
    for _ in range(n):
        labels.append(f"{str(yr)[-2:]}M{mo:02d}")
        mo -= 1
        if mo == 0:
            mo = 12; yr -= 1
    return list(reversed(labels))

def last_n_quarterly_labels(n: int, anchor: date) -> list:
    labels = []
    yr, mo = anchor.year, anchor.month
    for _ in range(n):
        q = (mo - 1) // 3 + 1
        labels.append(f"{str(yr)[-2:]}Q{q}")
        mo -= 3
        if mo <= 0:
            mo += 12; yr -= 1
    return list(reversed(labels))

# ── Data loading ──────────────────────────────────────────────────────────────
def load_mapping(path: str) -> dict:
    """Return {component_name: {'studio': str, 'game': str}}."""
    mapping = {}
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                comp = (row.get("Component") or row.get("GAME") or "").strip()
                studio = (row.get("Studio") or row.get("STUDIO") or row.get("TEAM") or "").strip()
                game = (row.get("Game") or row.get("GAME_NAME") or comp).strip()
                if comp:
                    mapping[comp] = {"studio": studio or "Other", "game": game or comp}
    else:
        wb = openpyxl.load_workbook(path, read_only=True)
        ws = wb.active
        hdrs = [c.value for c in next(ws.iter_rows(max_row=1))]
        for row in ws.iter_rows(min_row=2, values_only=True):
            r = dict(zip(hdrs, row))
            comp   = str(r.get("Component") or r.get("GAME") or "").strip()
            studio = str(r.get("Studio") or r.get("STUDIO") or r.get("TEAM") or "").strip()
            game   = str(r.get("Game") or r.get("GAME_NAME") or comp).strip()
            if comp:
                mapping[comp] = {"studio": studio or "Other", "game": game or comp}
        wb.close()
    return mapping

def parse_vfp_date(changelog_json: str):
    """Return the first date the ticket entered Verified for Production (or Production/Done)."""
    try:
        cl = json.loads(changelog_json or "[]")
    except Exception:
        return None
    target = {"VERIFIED FOR PRODUCTION", "VERIFIED FOR PROD", "PRODUCTION", "DONE"}
    for entry in cl:
        for item in entry.get("items", []):
            if item.get("field", "").lower() == "status":
                if str(item.get("toString", "")).upper().strip() in target:
                    created = entry.get("created", "")
                    try:
                        return datetime.fromisoformat(
                            created.replace("Z", "+00:00")
                        ).date()
                    except Exception:
                        try:
                            return datetime.strptime(created[:10], "%Y-%m-%d").date()
                        except Exception:
                            return None
    return None

def load_tickets(xlsx_path: str, comp_map: dict) -> list:
    """Return list of dicts with sp, dev_days, qa_days, studio, game, verified_date.

    Dev Time  = In Progress → Last Time in Staging  (from changelog timestamps, working days)
    QA Time   = In QA       → Verified for Production (from changelog timestamps, working days)
    """
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(max_row=1))]

    tickets = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        r = dict(zip(headers, row))

        sp = r.get("Story points")
        if sp not in SP_SIZES:
            continue

        # ── Parse key timestamps from Changelog ──────────────────────────────
        in_progress   = None
        first_in_qa   = None
        vfp_ts        = None
        staging_times = []

        try:
            cl = json.loads(r.get("Changelog (JSON)") or "[]")
        except Exception:
            cl = []

        for entry in cl:
            try:
                ts = datetime.fromisoformat(
                    entry["created"].replace("Z", "+00:00")
                )
            except Exception:
                continue
            for item in entry.get("items", []):
                if item.get("field", "").lower() != "status":
                    continue
                to_s = str(item.get("toString", "")).upper().strip()
                if to_s == "IN PROGRESS" and in_progress is None:
                    in_progress = ts
                if to_s == "STAGING":
                    staging_times.append(ts)
                if to_s == "IN QA" and first_in_qa is None:
                    first_in_qa = ts
                if to_s in {"VERIFIED FOR PRODUCTION", "VERIFIED FOR PROD"} and vfp_ts is None:
                    vfp_ts = ts

        last_staging = max(staging_times) if staging_times else None

        # ── Compute working-day metrics ───────────────────────────────────────
        dev_days = _wd(in_progress, last_staging)   # In Progress → Last Staging
        qa_days  = _wd(first_in_qa,  vfp_ts)        # In QA → Verified for Production

        # ── VfP date (already parsed from changelog above) ───────────────────
        vfp_date = vfp_ts.date() if vfp_ts else None
        if not vfp_date:
            continue

        # Map component → studio/game
        comp_str = str(r.get("Components") or "")
        studio, game = "Other", "Other"
        for comp in comp_str.split(","):
            comp = comp.strip()
            if comp in comp_map:
                studio = comp_map[comp]["studio"]
                game   = comp_map[comp]["game"]
                break
            # Fuzzy: partial match on [Game] prefix
            for k, v in comp_map.items():
                if comp and (comp in k or k in comp):
                    studio = v["studio"]
                    game   = v["game"]
                    break

        tickets.append({
            "key":           r.get("Key"),
            "sp":            sp,
            "dev_days":      float(dev_days) if dev_days is not None else None,
            "qa_days":       float(qa_days)  if qa_days  is not None else None,
            "studio":        studio,
            "game":          game,
            "verified_date": vfp_date,
        })

    wb.close()
    return tickets

# ── Aggregation ───────────────────────────────────────────────────────────────
def aggregate_periods(tickets: list, anchor: date) -> dict:
    bw_labels  = last_n_biweekly_labels(PERIODS_BACK_BIWEEKLY,  anchor)
    mon_labels = last_n_monthly_labels(PERIODS_BACK_MONTHLY,    anchor)
    qtr_labels = last_n_quarterly_labels(PERIODS_BACK_QUARTERLY, anchor)

    # dev/qa by SP and period
    dev_bw  = {sp: defaultdict(list) for sp in SP_SIZES}
    dev_mon = {sp: defaultdict(list) for sp in SP_SIZES}
    qa_bw   = {sp: defaultdict(list) for sp in SP_SIZES}
    qa_mon  = {sp: defaultdict(list) for sp in SP_SIZES}

    # velocity
    vel_bw  = defaultdict(float)
    vel_mon = defaultdict(float)

    # effort
    st_tot  = defaultdict(float)
    gm_tot  = defaultdict(float)
    st_mon  = defaultdict(lambda: defaultdict(float))
    gm_mon  = defaultdict(lambda: defaultdict(float))
    st_qtr  = defaultdict(lambda: defaultdict(float))
    gm_qtr  = defaultdict(lambda: defaultdict(float))

    for t in tickets:
        vd = t["verified_date"]
        if vd > anchor:
            continue

        sp  = t["sp"]
        bwl = iso_week_biweekly_label(vd)
        ml  = monthly_label(vd)
        ql  = quarterly_label(vd)

        if bwl in bw_labels:
            vel_bw[bwl] += sp
            if t["dev_days"] is not None and t["dev_days"] >= 0:
                dev_bw[sp][bwl].append(t["dev_days"])
            if t["qa_days"] is not None and t["qa_days"] >= 0:
                qa_bw[sp][bwl].append(t["qa_days"])

        if ml in mon_labels:
            vel_mon[ml] += sp
            if t["dev_days"] is not None and t["dev_days"] >= 0:
                dev_mon[sp][ml].append(t["dev_days"])
            if t["qa_days"] is not None and t["qa_days"] >= 0:
                qa_mon[sp][ml].append(t["qa_days"])
            # effort
            st_tot[t["studio"]] += sp
            gm_tot[t["game"]]   += sp
            st_mon[ml][t["studio"]] += sp
            gm_mon[ml][t["game"]]   += sp

        if ql in qtr_labels:
            st_qtr[ql][t["studio"]] += sp
            gm_qtr[ql][t["game"]]   += sp

    return {
        "bw_labels":  bw_labels,
        "m_labels":   mon_labels,
        "q_labels":   qtr_labels,
        "dev_bw":     dev_bw,
        "dev_mon":    dev_mon,
        "qa_bw":      qa_bw,
        "qa_mon":     qa_mon,
        "vel_bw":     vel_bw,
        "vel_mon":    vel_mon,
        "st_tot":     dict(st_tot),
        "gm_tot":     dict(gm_tot),
        "st_mon":     {k: dict(v) for k, v in st_mon.items()},
        "gm_mon":     {k: dict(v) for k, v in gm_mon.items()},
        "st_qtr":     {k: dict(v) for k, v in st_qtr.items()},
        "gm_qtr":     {k: dict(v) for k, v in gm_qtr.items()},
    }

# ── Chart helpers ─────────────────────────────────────────────────────────────
_tmp_files = []

def _tmp(suffix=".png"):
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    _tmp_files.append(path)
    return path

def make_histogram_bar(days_values: list, color: str, title: str,
                       figsize=(5, 2.8)) -> str:
    """X = rounded days, Y = number of tickets."""
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor(BRAND_LIGHT)
    ax.set_facecolor(BRAND_LIGHT)

    if days_values:
        rounded = [int(round(v)) for v in days_values if v is not None]
        if rounded:
            max_d = min(max(rounded), 20)
            bins = range(0, max_d + 2)
            counts = [rounded.count(d) for d in bins]
            ax.bar(list(bins), counts, color=color, width=0.7, zorder=3)
            ax.set_xlabel("Days", fontsize=7, color="#6b6860")
            ax.set_ylabel("Tickets", fontsize=7, color="#6b6860")
            ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
            ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
        else:
            ax.text(0.5, 0.5, "No data", ha="center", va="center",
                    transform=ax.transAxes, color="#aaa", fontsize=9)
    else:
        ax.text(0.5, 0.5, "No data", ha="center", va="center",
                transform=ax.transAxes, color="#aaa", fontsize=9)

    ax.set_title(title, fontsize=8, pad=4, color="#1a1917")
    ax.tick_params(labelsize=7)
    ax.grid(axis="y", color="#c8d8ea", linewidth=0.5, linestyle="--")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for sp_ in ["left", "bottom"]:
        ax.spines[sp_].set_color("#b0c4d8")

    plt.tight_layout(pad=0.4)
    path = _tmp()
    fig.savefig(path, dpi=120, bbox_inches="tight")
    plt.close(fig)
    return path

def make_line_chart(labels, values, color, title, ylabel, figsize=(5, 2.8)) -> str:
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor(BRAND_LIGHT)
    ax.set_facecolor(BRAND_LIGHT)

    xs = range(len(labels))
    ys = [v if v is not None else float("nan") for v in values]
    ax.plot(xs, ys, color=color, linewidth=2, marker="o", markersize=4, zorder=3)
    ax.fill_between(xs, ys, alpha=0.12, color=color)

    ax.set_xticks(list(xs))
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
    path = _tmp()
    fig.savefig(path, dpi=120, bbox_inches="tight")
    plt.close(fig)
    return path

def make_stacked_bar(labels, series_dict, color_map, title, ylabel, figsize=(5, 2.8)) -> str:
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor(BRAND_LIGHT)
    ax.set_facecolor(BRAND_LIGHT)

    xs = range(len(labels))
    bottom = [0.0] * len(labels)
    colors_list = list(color_map.values()) + GAME_COLORS_LIST
    for i, (name, vals) in enumerate(series_dict.items()):
        ys = [vals.get(l, 0) for l in labels]
        color = colors_list[i % len(colors_list)]
        ax.bar(xs, ys, bottom=bottom, label=name, color=color, width=0.7, zorder=3)
        bottom = [b + y for b, y in zip(bottom, ys)]

    ax.set_xticks(list(xs))
    ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=7)
    ax.tick_params(axis="y", labelsize=7)
    ax.set_title(title, fontsize=8, pad=4, color="#1a1917")
    ax.set_ylabel(ylabel, fontsize=7, color="#6b6860")
    ax.legend(fontsize=6, loc="upper left", framealpha=0.8)
    ax.grid(axis="y", color="#c8d8ea", linewidth=0.5, linestyle="--")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for sp_ in ["left", "bottom"]:
        ax.spines[sp_].set_color("#b0c4d8")

    plt.tight_layout(pad=0.4)
    path = _tmp()
    fig.savefig(path, dpi=120, bbox_inches="tight")
    plt.close(fig)
    return path

def make_donut(labels, values, color_map, title, figsize=(5, 3.2)) -> str:
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor(BRAND_LIGHT)
    ax.set_facecolor(BRAND_LIGHT)

    colors_list = [color_map.get(l, GAME_COLORS_LIST[i % len(GAME_COLORS_LIST)])
                   for i, l in enumerate(labels)]
    wedges, texts, autotexts = ax.pie(
        values, labels=labels, colors=colors_list,
        autopct="%1.0f%%", startangle=90,
        wedgeprops={"width": 0.55, "edgecolor": "white", "linewidth": 1.5},
        textprops={"fontsize": 7},
    )
    for at in autotexts:
        at.set_fontsize(7)
    ax.set_title(title, fontsize=8, pad=6, color="#1a1917")

    plt.tight_layout(pad=0.4)
    path = _tmp()
    fig.savefig(path, dpi=120, bbox_inches="tight")
    plt.close(fig)
    return path

# ── PDF builder ───────────────────────────────────────────────────────────────
def img(path: str, width: float, height: float = None) -> Image:
    i = Image(path, width=width)
    if height:
        i.drawHeight = height
    return i

def build_pdf(tickets: list, anchor: date, month_label: str, output_path: str):
    agg = aggregate_periods(tickets, anchor)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle("title", parent=styles["Title"],
                                 fontSize=28, textColor=colors.HexColor(BRAND_DARK),
                                 spaceAfter=0.4 * cm)
    h2_style = ParagraphStyle("h2", parent=styles["Heading2"],
                               fontSize=16, textColor=colors.HexColor(BRAND_BLUE),
                               spaceBefore=0.3 * cm, spaceAfter=0.2 * cm)
    h3_style = ParagraphStyle("h3", parent=styles["Heading3"],
                               fontSize=11, textColor=colors.HexColor(BRAND_DARK),
                               spaceBefore=0.2 * cm, spaceAfter=0.1 * cm)
    body_style = ParagraphStyle("body", parent=styles["Normal"],
                                fontSize=9, textColor=colors.HexColor("#3a3a3a"),
                                spaceAfter=0.15 * cm)
    note_style = ParagraphStyle("note", parent=styles["Normal"],
                                fontSize=8, textColor=colors.HexColor("#888888"),
                                spaceAfter=0.1 * cm, italics=True)

    doc = BaseDocTemplate(
        output_path, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin,
                  doc.width, doc.height, id="main")
    doc.addPageTemplates([PageTemplate(id="main", frames=[frame])])

    story = []

    def two_charts(p1, p2):
        t = Table([[img(p1, IMG_W_HALF), img(p2, IMG_W_HALF)]],
                  colWidths=[IMG_W_HALF + 0.3 * cm, IMG_W_HALF + 0.3 * cm])
        t.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"),
                                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                                ("RIGHTPADDING", (0, 0), (-1, -1), 4)]))
        story.append(t)

    # ── Page 1: Cover ──────────────────────────────────────────────────────────
    story.append(Spacer(1, 3 * cm))
    story.append(Paragraph("Game Server", title_style))
    story.append(Paragraph("CTO Monthly Performance Report", h2_style))
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph(month_label, ParagraphStyle(
        "month", parent=styles["Normal"], fontSize=18,
        textColor=colors.HexColor(BRAND_BLUE), spaceAfter=0.5 * cm
    )))
    story.append(Spacer(1, 1 * cm))
    story.append(Paragraph(
        "This report summarises engineering delivery performance for the Game Server team, "
        "covering development throughput, quality, velocity, and effort distribution across "
        f"studios and games. Data anchor: {anchor.strftime('%B %d, %Y')}.",
        body_style
    ))
    total_sp = sum(t["sp"] for t in tickets if t["verified_date"] <= anchor)
    total_tickets = sum(1 for t in tickets if t["verified_date"] <= anchor)
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph(
        f"Total tickets in dataset: <b>{total_tickets}</b> &nbsp;&nbsp; "
        f"Total story points: <b>{int(total_sp)}</b>",
        body_style
    ))
    story.append(PageBreak())

    # ── Page 2: Time in Development ──────────────────────────────────────────
    story.append(Paragraph("Time in Development", h2_style))
    story.append(Paragraph(
        "Distribution of working days spent in development (In Progress → last Staging) "
        f"per story-point size. Last {PERIODS_BACK_BIWEEKLY} biweekly and "
        f"{PERIODS_BACK_MONTHLY} monthly periods.",
        note_style
    ))
    story.append(Spacer(1, 0.2 * cm))

    for sp in SP_SIZES:
        story.append(Paragraph(f"{sp} Story Points", h3_style))
        bw_vals = [v for lbl in agg["bw_labels"] for v in agg["dev_bw"][sp].get(lbl, [])]
        mon_vals = [v for lbl in agg["m_labels"]  for v in agg["dev_mon"][sp].get(lbl, [])]
        p1 = make_histogram_bar(bw_vals,  BRAND_BLUE,  f"{sp} SP — Biweekly")
        p2 = make_histogram_bar(mon_vals, BRAND_DARK,  f"{sp} SP — Monthly")
        two_charts(p1, p2)
        story.append(Spacer(1, 0.1 * cm))

    story.append(PageBreak())

    # ── Page 3: Time in QA ───────────────────────────────────────────────────
    story.append(Paragraph("Time in QA", h2_style))
    story.append(Paragraph(
        "Distribution of working days spent in QA (In QA status → Verified for Production) "
        f"per story-point size. Last {PERIODS_BACK_BIWEEKLY} biweekly and "
        f"{PERIODS_BACK_MONTHLY} monthly periods.",
        note_style
    ))
    story.append(Spacer(1, 0.2 * cm))

    for sp in SP_SIZES:
        story.append(Paragraph(f"{sp} Story Points", h3_style))
        bw_vals  = [v for lbl in agg["bw_labels"] for v in agg["qa_bw"][sp].get(lbl, [])]
        mon_vals = [v for lbl in agg["m_labels"]  for v in agg["qa_mon"][sp].get(lbl, [])]
        p1 = make_histogram_bar(bw_vals,  "#E8A020", f"{sp} SP — Biweekly")
        p2 = make_histogram_bar(mon_vals, "#B87318", f"{sp} SP — Monthly")
        two_charts(p1, p2)
        story.append(Spacer(1, 0.1 * cm))

    story.append(PageBreak())

    # ── Page 4: Team Velocity ────────────────────────────────────────────────
    story.append(Paragraph("Team Velocity", h2_style))
    story.append(Paragraph(
        f"Story points delivered per biweekly and monthly period. "
        f"Last {PERIODS_BACK_BIWEEKLY} biweekly and {PERIODS_BACK_MONTHLY} monthly periods.",
        note_style
    ))
    story.append(Spacer(1, 0.3 * cm))

    bw_vel  = [agg["vel_bw"].get(l, 0)  for l in agg["bw_labels"]]
    mon_vel = [agg["vel_mon"].get(l, 0) for l in agg["m_labels"]]
    p1 = make_line_chart(agg["bw_labels"],  bw_vel,  "#5CB85C", "Velocity — Biweekly",  "Story Points")
    p2 = make_line_chart(agg["m_labels"],   mon_vel, "#3A7D3A", "Velocity — Monthly",   "Story Points")
    two_charts(p1, p2)
    story.append(PageBreak())

    # ── Page 5: Effort Distribution ──────────────────────────────────────────
    story.append(Paragraph("Effort Distribution", h2_style))
    story.append(Paragraph(
        f"Story points delivered per studio and game. "
        f"Last {PERIODS_BACK_MONTHLY} months and {PERIODS_BACK_QUARTERLY} quarters.",
        note_style
    ))
    story.append(Spacer(1, 0.2 * cm))

    studios = sorted(agg["st_tot"], key=lambda s: -agg["st_tot"][s])
    games   = sorted(agg["gm_tot"], key=lambda g: -agg["gm_tot"][g])[:10]

    if studios:
        vals = [agg["st_tot"][s] for s in studios]
        donut_path = make_donut(studios, vals, STUDIO_COLORS,
                                "Effort per Studio (All Time)", figsize=(5.5, 3.5))
        story.append(img(donut_path, IMG_W_FULL * 0.65))
        story.append(Spacer(1, 0.25 * cm))

    if studios:
        st_mon_series = {s: {l: agg["st_mon"].get(l, {}).get(s, 0) for l in agg["m_labels"]}
                         for s in studios}
        st_qtr_series = {s: {l: agg["st_qtr"].get(l, {}).get(s, 0) for l in agg["q_labels"]}
                         for s in studios}
        p1 = make_stacked_bar(agg["m_labels"], st_mon_series, STUDIO_COLORS,
                              "Effort per Studio – Monthly", "Story Points")
        p2 = make_stacked_bar(agg["q_labels"], st_qtr_series, STUDIO_COLORS,
                              "Effort per Studio – Quarterly", "Story Points")
        two_charts(p1, p2)
        story.append(Spacer(1, 0.2 * cm))

    if games:
        game_colors = {g: GAME_COLORS_LIST[i % len(GAME_COLORS_LIST)]
                       for i, g in enumerate(games)}
        gm_mon_series = {g: {l: agg["gm_mon"].get(l, {}).get(g, 0) for l in agg["m_labels"]}
                         for g in games}
        gm_qtr_series = {g: {l: agg["gm_qtr"].get(l, {}).get(g, 0) for l in agg["q_labels"]}
                         for g in games}
        p1 = make_stacked_bar(agg["m_labels"], gm_mon_series, game_colors,
                              "Effort per Game – Monthly (Top 10)", "Story Points")
        p2 = make_stacked_bar(agg["q_labels"], gm_qtr_series, game_colors,
                              "Effort per Game – Quarterly (Top 10)", "Story Points")
        two_charts(p1, p2)

    doc.build(story)
    print(f"  PDF saved → {output_path}")

# ── CLI ───────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Generate GS CTO Monthly PDF report.")
    parser.add_argument("--xlsx",    required=True, help="Path to GS_jira_year_report.xlsx")
    parser.add_argument("--mapping", required=True, help="Path to GamesStudios.csv")
    parser.add_argument("--output",  required=True, help="Output PDF path")
    parser.add_argument("--month",   default="",    help="Report month YYYY-MM (default: previous month)")
    args = parser.parse_args()

    # Determine anchor = last day of target month
    if args.month:
        yr, mo = map(int, args.month.split("-"))
    else:
        today = date.today()
        mo = today.month - 1 or 12
        yr = today.year if today.month > 1 else today.year - 1

    # Last day of the month
    if mo == 12:
        anchor = date(yr + 1, 1, 1) - timedelta(days=1)
    else:
        anchor = date(yr, mo + 1, 1) - timedelta(days=1)

    month_str = anchor.strftime("%B %Y")
    print(f"Generating report for {month_str} (anchor: {anchor})")

    print("Loading mapping...")
    comp_map = load_mapping(args.mapping)
    print(f"  {len(comp_map)} components mapped")

    print("Loading tickets...")
    tickets = load_tickets(args.xlsx, comp_map)
    print(f"  {len(tickets)} tickets loaded")

    print("Building PDF...")
    build_pdf(tickets, anchor, month_str, args.output)

    # Cleanup temp files
    for f in _tmp_files:
        try:
            os.remove(f)
        except Exception:
            pass

if __name__ == "__main__":
    main()
