"""
Utilization Credit Report — Web App
Run with: streamlit run utilization_app_v2.py
"""

import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Utilization Credit Report",
    page_icon="📈",
    layout="wide"
)

# ── Constants ─────────────────────────────────────────────────────────────────
NAVY     = "1e2c63"
TEAL     = "4472C4"
WHITE    = "FFFFFF"
LTGRAY   = "F2F2F2"
MID_GRAY = "BDC3C7"

TAG_COLORS = {
    "CREDITED":     "EAF9F1",
    "OVERRUN":      "FDECED",
    "PARTIAL":      "FEF9E7",
    "NON-BILLABLE": "EBEDEE",
    "UNCONFIGURED": "F2F2F2",
}
TAG_BADGE = {
    "CREDITED":     "🟢",
    "OVERRUN":      "🔴",
    "PARTIAL":      "🟡",
    "NON-BILLABLE": "⚫",
    "UNCONFIGURED": "⚪",
}

# ── Stored scope map ──────────────────────────────────────────────────────────
DEFAULT_SCOPE = {
    "Capture":                 20,
    "Approvals":               17,
    "Reconcile":               17,
    "PSP":                     18,
    "Payments":                30,
    "Reconcile 2.0":           20,
    "CC":                       6,
    "SFTP":                    12,
    "Premium - 10":            10,
    "Premium - 20":            20,
    "E-Invoicing":             15,
    "Capture and E-Invoicing": 30,
    "Additional Subsidiary":    2,
}

# ── Available hours by region/month (2026) ────────────────────────────────────
AVAIL_HOURS = {
    "Spain":            {"2026-01":155.00,"2026-02":155.00,"2026-03":170.50,"2026-04":162.75,"2026-05":155.00,"2026-06":170.50,"2026-07":178.25,"2026-08":162.75,"2026-09":170.50,"2026-10":162.75,"2026-11":155.00,"2026-12":155.00},
    "UK":               {"2026-01":157.50,"2026-02":150.00,"2026-03":165.00,"2026-04":150.00,"2026-05":142.50,"2026-06":165.00,"2026-07":172.50,"2026-08":150.00,"2026-09":165.00,"2026-10":165.00,"2026-11":157.50,"2026-12":157.50},
    "Northern Ireland": {"2026-01":157.50,"2026-02":150.00,"2026-03":157.50,"2026-04":150.00,"2026-05":142.50,"2026-06":165.00,"2026-07":165.00,"2026-08":150.00,"2026-09":165.00,"2026-10":165.00,"2026-11":157.50,"2026-12":157.50},
    "Netherlands":      {"2026-01":168.00,"2026-02":160.00,"2026-03":176.00,"2026-04":152.00,"2026-05":144.00,"2026-06":176.00,"2026-07":184.00,"2026-08":168.00,"2026-09":176.00,"2026-10":176.00,"2026-11":168.00,"2026-12":176.00},
    "Faroe Islands":    {"2026-01":168.00,"2026-02":160.00,"2026-03":176.00,"2026-04":144.00,"2026-05":144.00,"2026-06":176.00,"2026-07":168.00,"2026-08":168.00,"2026-09":176.00,"2026-10":176.00,"2026-11":168.00,"2026-12":168.00},
    "North Macedonia":  {"2026-01":160.00,"2026-02":160.00,"2026-03":168.00,"2026-04":168.00,"2026-05":160.00,"2026-06":176.00,"2026-07":184.00,"2026-08":168.00,"2026-09":168.00,"2026-10":168.00,"2026-11":168.00,"2026-12":176.00},
    "Czech Republic":   {"2026-01":168.00,"2026-02":160.00,"2026-03":176.00,"2026-04":160.00,"2026-05":152.00,"2026-06":176.00,"2026-07":176.00,"2026-08":168.00,"2026-09":168.00,"2026-10":168.00,"2026-11":160.00,"2026-12":168.00},
    "Serbia":           {"2026-01":152.00,"2026-02":152.00,"2026-03":176.00,"2026-04":160.00,"2026-05":160.00,"2026-06":176.00,"2026-07":184.00,"2026-08":168.00,"2026-09":176.00,"2026-10":176.00,"2026-11":160.00,"2026-12":184.00},
    "Canada":           {"2026-01":168.00,"2026-02":160.00,"2026-03":176.00,"2026-04":168.00,"2026-05":160.00,"2026-06":176.00,"2026-07":176.00,"2026-08":160.00,"2026-09":160.00,"2026-10":168.00,"2026-11":160.00,"2026-12":168.00},
    "USA":              {"2026-01":160.00,"2026-02":152.00,"2026-03":176.00,"2026-04":176.00,"2026-05":160.00,"2026-06":168.00,"2026-07":176.00,"2026-08":168.00,"2026-09":168.00,"2026-10":168.00,"2026-11":152.00,"2026-12":176.00},
    "Sydney (NSW)":     {"2026-01":152.00,"2026-02":152.00,"2026-03":167.20,"2026-04":144.40,"2026-05":159.60,"2026-06":159.60,"2026-07":174.80,"2026-08":152.00,"2026-09":167.20,"2026-10":159.60,"2026-11":159.60,"2026-12":159.60},
    "Manila (PH)":      {"2026-01":168.00,"2026-02":152.00,"2026-03":176.00,"2026-04":152.00,"2026-05":160.00,"2026-06":168.00,"2026-07":184.00,"2026-08":152.00,"2026-09":176.00,"2026-10":176.00,"2026-11":152.00,"2026-12":144.00},
}

# Fixed fee task keywords (Case/Task/Event column)
FF_TASKS = ["Configuration", "Enablement", "Training", "Post Go-live", "Project Management"]

def get_avail_hours(region, period):
    """Look up available hours for a region/period. Returns None if not found."""
    region_clean = str(region).strip()
    # Try exact match first, then case-insensitive
    for r, months in AVAIL_HOURS.items():
        if r.lower() == region_clean.lower():
            return months.get(str(period), None)
    return None

def match_ff_task(task_val):
    """Match a task value to one of the 4 standard FF task categories."""
    t = str(task_val).strip().lower()
    if "config" in t:             return "Configuration"
    if "enabl" in t or "train" in t: return "Enablement/Training"
    if "post" in t or "go-live" in t or "golive" in t: return "Post Go-live Support"
    if "project mgmt" in t or "project management" in t or "pm" == t: return "Project Management"
    return None

# ── Excel helpers ─────────────────────────────────────────────────────────────
def thin_border():
    s = Side(style="thin", color=MID_GRAY)
    return Border(left=s, right=s, top=s, bottom=s)

def hdr_fill(hex_color):  return PatternFill("solid", fgColor=hex_color)
def row_fill(hex_color):  return PatternFill("solid", fgColor=hex_color)

GROUP_COLORS = ["EEF2FB", "FFFFFF"]  # alternating soft blue/white for grouped rows

def group_bg(value, prev_value, group_idx):
    """Return bg color and updated group index for grouped first-column display."""
    if value != prev_value:
        group_idx = 1 - group_idx  # toggle group
    return GROUP_COLORS[group_idx], group_idx

def style_header(ws, row, headers, fill_color=NAVY):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font      = Font(name="Manrope", bold=True, color=WHITE, size=10)
        c.fill      = hdr_fill(fill_color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = thin_border()
    ws.row_dimensions[row].height = 22

def style_cell(cell, bg, fmt=None, bold=False, align="left"):
    cell.fill      = row_fill(bg)
    cell.font      = Font(name="Manrope", size=10, bold=bold)
    cell.border    = thin_border()
    cell.alignment = Alignment(horizontal=align, vertical="center")
    if fmt:
        cell.number_format = fmt

def write_title(ws, title, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font      = Font(name="Manrope", bold=True, size=14, color=WHITE)
    c.fill      = hdr_fill(NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

# ── Column detection ──────────────────────────────────────────────────────────
def auto_detect_columns(df):
    cols_lower = {c.lower().strip(): c for c in df.columns}
    checks = {
        "employee":      ["employee", "employee name", "name", "resource"],
        "project":       ["project", "project name", "job"],
        "project_type":  ["project type", "type", "project_type"],
        "date":          ["date", "time entry date", "entry date", "work date"],
        "hours":         ["hours", "duration", "hours logged", "time", "qty"],
        "approval":      ["approval status", "approval", "status"],
        "task":          ["case/task/event", "task", "case", "event", "memo"],
        "non_billable":  ["non-billable", "non billable", "nonbillable",
                          "non_billable", "is non billable"],
        "billing_type":  ["billing type", "billing_type", "bill type", "billtype"],
        "hours_to_date": ["hours to date", "hours_to_date", "htd", "prior hours",
                          "cumulative hours", "hours booked to date"],
        "region":        ["region", "location", "country", "office"],
    }
    mapping = {}
    unmatched = []
    for standard, candidates in checks.items():
        for candidate in candidates:
            if candidate in cols_lower:
                mapping[cols_lower[candidate]] = standard
                break
        else:
            if standard in ["employee", "project", "project_type", "hours", "non_billable"]:
                unmatched.append(standard)
    return mapping, unmatched

# ── Core credit logic ─────────────────────────────────────────────────────────
def assign_credits(df, scope_map):
    col_map, unmatched = auto_detect_columns(df)
    if unmatched:
        st.warning(f"⚠️ Could not auto-detect columns: {unmatched}. Check your file headers.")

    df = df.rename(columns=col_map)
    df["non_billable"] = df["non_billable"].astype(str).str.strip().str.upper()
    df["hours"]  = pd.to_numeric(df["hours"], errors="coerce").fillna(0)
    df["date"]   = pd.to_datetime(df["date"], errors="coerce")
    df["period"] = df["date"].dt.strftime("%Y-%m").fillna("Unknown")

    consumed = {}
    credit_hrs_list    = []
    variance_hrs_list  = []
    credit_tag_list    = []
    notes_list         = []
    htd_start_list     = []  # track starting HTD per row for output

    for _, row in df.iterrows():
        proj      = str(row.get("project", "")).strip()
        ptype     = str(row.get("project_type", "")).strip()
        hrs       = float(row.get("hours", 0))
        nb        = str(row.get("non_billable", "NO")).strip().upper()
        bill_type = str(row.get("billing_type", "")).strip().lower()
        is_tm     = bill_type == "t&m"

        if hrs <= 0:
            credit_hrs_list.append(0); variance_hrs_list.append(0)
            credit_tag_list.append("SKIPPED"); notes_list.append("Zero or missing hours")
            htd_start_list.append(0)
            continue

        # Rule 1: Internal = always 0 credit
        if bill_type == "internal":
            credit_hrs_list.append(0); variance_hrs_list.append(0)
            credit_tag_list.append("NON-BILLABLE"); notes_list.append("Internal: excluded from utilization")
            htd_start_list.append(0)
            continue

        # Rule 2: T&M = always full credit, no cap
        if is_tm:
            credit_hrs_list.append(hrs); variance_hrs_list.append(0)
            credit_tag_list.append("CREDITED"); notes_list.append("T&M: full credit")
            htd_start_list.append(0)
            continue

        # Rule 3: Fixed Fee = capped at scope (longest match wins)
        _ptype_lower = ptype.strip().lower()
        _matches = [(k, float(v)) for k, v in scope_map.items()
                    if k.strip().lower() in _ptype_lower]
        scope_hrs = max(_matches, key=lambda x: len(x[0]))[1] if _matches else None

        if scope_hrs is None:
            credit_hrs_list.append(0); variance_hrs_list.append(hrs)
            credit_tag_list.append("UNCONFIGURED"); notes_list.append(f"Fixed Fee but no scope defined for: {ptype}")
            htd_start_list.append(0)
            continue

        # Seed starting balance from hours_to_date if first time seeing this project
        if proj not in consumed:
            htd = row.get("hours_to_date", None)
            try:
                consumed[proj] = float(htd) if htd is not None and str(htd).strip() not in ("", "nan") else 0
            except (ValueError, TypeError):
                consumed[proj] = 0

        already    = consumed[proj]
        remaining  = scope_hrs - already
        htd_start_list.append(already)

        if remaining <= 0:
            credit_hrs_list.append(0); variance_hrs_list.append(hrs)
            credit_tag_list.append("OVERRUN"); notes_list.append(f"Scope exhausted (cap: {scope_hrs:.0f}h)")
        elif hrs <= remaining:
            consumed[proj] = already + hrs
            credit_hrs_list.append(hrs); variance_hrs_list.append(0)
            credit_tag_list.append("CREDITED"); notes_list.append(f"NB within scope ({already:.1f}/{scope_hrs:.0f}h used)")
        else:
            consumed[proj] = already + remaining
            credit_hrs_list.append(remaining); variance_hrs_list.append(hrs - remaining)
            credit_tag_list.append("PARTIAL")
            notes_list.append(f"Split: {remaining:.2f}h credited / {hrs - remaining:.2f}h overrun")

    df["credit_hrs"]    = credit_hrs_list
    df["variance_hrs"]  = variance_hrs_list
    df["credit_tag"]    = credit_tag_list
    df["notes"]         = notes_list
    df["htd_start"]     = htd_start_list

    # Updated hours to date = htd_start + total hours booked this period (credited + overrun)
    df["updated_htd"] = df["htd_start"] + df["hours"]

    # Tag FF tasks
    df["ff_task"] = df["task"].apply(match_ff_task) if "task" in df.columns else None

    # Collect skipped rows for reporting
    skipped_df = df[df["credit_tag"] == "SKIPPED"][
        [c for c in ["employee","project","project_type","billing_type","date","hours","notes"] if c in df.columns]
    ].copy()

    return df, consumed, skipped_df


# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(df, scope_map, consumed):
    wb  = Workbook()
    wb.remove(wb.active)
    bgs = [WHITE, LTGRAY]

    # ── 1. PROCESSED DATA ─────────────────────────────────────
    ws = wb.create_sheet("PROCESSED_DATA")
    ws.sheet_properties.tabColor = TEAL
    ws.freeze_panes = "A3"

    headers = ["Employee","Region","Project","Project Type","Billing Type","Hrs to Date",
               "Date","Hours Logged","Approval","Task/Case","Non-Billable",
               "Credit Hrs","Variance Hrs","Updated Hrs to Date","Credit Tag","Period","Notes"]
    widths  = [20,16,35,20,14,13,14,13,14,25,13,12,12,18,16,12,45]
    cols    = ["employee","region","project","project_type","billing_type","hours_to_date",
               "date","hours","approval","task","non_billable","credit_hrs",
               "variance_hrs","updated_htd","credit_tag","period","notes"]

    write_title(ws, "PROCESSED DATA — Utilization Credit Detail", len(headers))
    style_header(ws, 2, headers, TEAL)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for r_idx, (_, row) in enumerate(df.iterrows(), 3):
        tag = str(row.get("credit_tag","")).strip()
        bg  = TAG_COLORS.get(tag, "F2F2F2")
        for c_idx, col in enumerate(cols, 1):
            val  = row.get(col, "")
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            fmt, bold, align = None, False, "left"
            if col == "date" and pd.notna(val):
                fmt = "YYYY-MM-DD"; align = "center"
            elif col in ("hours","credit_hrs","variance_hrs","hours_to_date","updated_htd"):
                fmt = "#,##0.00"; align = "right"
            elif col == "credit_tag":
                bold = True; align = "center"
            elif col in ("period","billing_type","region"):
                align = "center"
            style_cell(cell, bg, fmt=fmt, bold=bold, align=align)

    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}2"

    # ── 2. EMPLOYEE SUMMARY ───────────────────────────────────
    ws2 = wb.create_sheet("SUMMARY - By Employee")
    ws2.sheet_properties.tabColor = NAVY
    ws2.freeze_panes = "A3"

    eh = ["Employee","Region","Period","Avail Hrs","Hours This Period",
          "Utilization Credits","Project Overrun Hrs","Admin Hrs","Util %"]
    ew = [22,18,12,12,14,18,18,14,10]
    write_title(ws2, "SUMMARY — Utilization by Employee", len(eh))
    style_header(ws2, 2, eh, TEAL)
    for i, w in enumerate(ew, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # Get region per employee (first occurrence)
    emp_region = {}
    if "region" in df.columns:
        emp_region = df.dropna(subset=["region"]).groupby("employee")["region"].first().to_dict()

    # Admin hours = all Internal billing type rows
    admin_hrs = df[df["billing_type"].str.lower() == "internal"].groupby(
        ["employee","period"])["hours"].sum().reset_index().rename(columns={"hours":"admin_hrs"})

    emp_sum = df[df["credit_tag"] != "SKIPPED"].groupby(
        ["employee","period"], as_index=False
    ).agg(
        hours_this_period=("hours","sum"),
        credit_hrs=("credit_hrs","sum"),
        overrun_hrs=("variance_hrs","sum"),
    ).sort_values(["employee","period"])

    emp_sum = emp_sum.merge(admin_hrs, on=["employee","period"], how="left")
    emp_sum["admin_hrs"] = emp_sum["admin_hrs"].fillna(0)

    _prev_emp = None; _grp_idx = 0
    for r_idx, (_, row) in enumerate(emp_sum.iterrows(), 3):
        emp     = row["employee"]
        period  = row["period"]
        region  = emp_region.get(emp, "")
        avail   = get_avail_hours(region, period) if region else None
        util    = row["credit_hrs"] / avail if avail else 0
        bg, _grp_idx = group_bg(emp, _prev_emp, _grp_idx)
        _prev_emp = emp
        util_bg = ("EAF9F1" if util >= 0.8 else "FEF9E7" if util >= 0.6 else "FDECED")

        vals = [emp, region, period, avail or "—",
                row["hours_this_period"], row["credit_hrs"],
                row["overrun_hrs"], row.get("admin_hrs", 0),
                util if avail else "—"]
        fmts = [None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","0.0%"]

        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, util_bg if c_idx == 9 else bg, fmt=fmt,
                       align="right" if c_idx > 3 else "center" if c_idx == 3 else "left")

    # ── 3. PROJECT SUMMARY ────────────────────────────────────
    ws3 = wb.create_sheet("SUMMARY - By Project")
    ws3.sheet_properties.tabColor = "E67E22"
    ws3.freeze_panes = "A3"

    ph = ["Project","Project Type","Scoped Hrs","Hours to Date",
          "Hours This Period","Credit Hrs","Variance Hrs","Updated Hrs to Date","Burn %","Status"]
    pw = [35,20,12,15,15,12,12,18,10,12]
    write_title(ws3, "SUMMARY — Utilization by Project", len(ph))
    style_header(ws3, 2, ph, TEAL)
    for i, w in enumerate(pw, 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    proj_sum = df[df["credit_tag"] != "SKIPPED"].groupby(
        ["project","project_type"], as_index=False
    ).agg(
        hours_this_period=("hours","sum"),
        credit_hrs=("credit_hrs","sum"),
        variance_hrs=("variance_hrs","sum"),
        htd_start=("htd_start","first"),
    ).sort_values("project")

    # HTD seed now comes directly from aggregation above
    htd_seeds = dict(zip(proj_sum["project"], proj_sum["htd_start"]))

    _prev_ptype = None; _grp_idx_p = 0
    for r_idx, (_, row) in enumerate(proj_sum.iterrows(), 3):
        ptype   = str(row["project_type"]).strip()
        _pm     = [(k, float(v)) for k, v in scope_map.items() if k.strip().lower() in ptype.lower()]
        scope_h = max(_pm, key=lambda x: len(x[0]))[1] if _pm else 0

        seed      = float(row["htd_start"]) if row["htd_start"] else 0
        # Updated HTD = prior HTD seed + all hours booked this period
        updated_h = seed + row["hours_this_period"]

        # Burn % = updated HTD / scoped hrs (includes prior periods)
        burn = updated_h / scope_h if scope_h > 0 else 0

        vari_h  = row["variance_hrs"]
        if vari_h > 0 or burn > 1:  status = "OVERRUN"
        elif burn >= 0.9:            status = "AT RISK"
        elif burn > 0:               status = "ON TRACK"
        else:                        status = "—"

        status_bg = {"OVERRUN":"FDECED","AT RISK":"FEF9E7","ON TRACK":"EAF9F1"}.get(status, LTGRAY)
        bg, _grp_idx_p = group_bg(ptype, _prev_ptype, _grp_idx_p)
        _prev_ptype = ptype

        vals = [row["project"], ptype, scope_h or "—", row["htd_start"],
                row["hours_this_period"], row["credit_hrs"], vari_h, updated_h,
                burn if scope_h > 0 else "—", status]
        fmts = [None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","0.0%",None]

        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, status_bg if c_idx == 10 else bg,
                       fmt=fmt, bold=(c_idx == 10),
                       align="right" if c_idx in (3,4,5,6,7,8,9) else "center" if c_idx == 10 else "left")

    # ── 4. ZCO NON-BILLABLE BREAKDOWN ─────────────────────────
    ws5 = wb.create_sheet("ZCO Non-Billable")
    ws5.sheet_properties.tabColor = "95A5A6"
    ws5.freeze_panes = "A3"

    znh = ["Task / Activity","Employee","Period","Hours"]
    znw = [35,22,12,12]
    write_title(ws5, "ZCO NON-BILLABLE — Hours by Employee by Activity", len(znh))
    style_header(ws5, 2, znh, TEAL)
    for i, w in enumerate(znw, 1):
        ws5.column_dimensions[get_column_letter(i)].width = w

    zco_df = df[df["credit_tag"] == "NON-BILLABLE"].copy()
    if "task" in zco_df.columns and len(zco_df) > 0:
        zco_sum = zco_df.groupby(
            ["task","employee","period"], as_index=False
        ).agg(hours=("hours","sum")).sort_values(["task","employee","period"])

        # Total hours per employee per period (all rows incl billable)
        emp_period_totals = df[df["credit_tag"] != "SKIPPED"].groupby(
            ["employee","period"])["hours"].sum().to_dict()

        znh[3] = "Hours"
        # Update headers to include % col
        ws5.cell(row=2, column=5, value="% of Total Hrs").font = Font(name="Manrope", bold=True, color=WHITE, size=10)
        ws5.cell(row=2, column=5).fill = hdr_fill(TEAL)
        ws5.cell(row=2, column=5).alignment = Alignment(horizontal="center", vertical="center")
        ws5.cell(row=2, column=5).border = thin_border()
        ws5.column_dimensions["E"].width = 16

        _prev_task_z = None; _grp_idx_z = 0
        for r_idx, (_, row) in enumerate(zco_sum.iterrows(), 3):
            bg, _grp_idx_z = group_bg(row.get("task",""), _prev_task_z, _grp_idx_z)
            _prev_task_z = row.get("task","")
            total_hrs  = emp_period_totals.get((row["employee"], row["period"]), 0)
            pct        = row["hours"] / total_hrs if total_hrs > 0 else 0
            vals = [row.get("task",""), row["employee"], row["period"], row["hours"], pct]
            fmts = [None, None, None, "#,##0.00", "0.0%"]
            for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
                cell = ws5.cell(row=r_idx, column=c_idx, value=val)
                style_cell(cell, bg, fmt=fmt,
                           align="right" if c_idx in (4,5) else "center" if c_idx == 3 else "left")
    else:
        ws5.cell(row=3, column=1, value="No Non-Billable (Internal) entries in this period.")

    # ── 5. TASK ANALYSIS ───────────────────────
    ws6 = wb.create_sheet("Task Analysis")
    ws6.sheet_properties.tabColor = "27AE60"
    ws6.freeze_panes = "A3"

    tah = ["Task Category","Project Type","Hours This Period","Avg Hrs / Project","% of Type Hrs"]
    taw = [28,25,16,18,16]
    write_title(ws6, "TASK ANALYSIS — Hours by Task › Project Type", len(tah))
    style_header(ws6, 2, tah, TEAL)
    for i, w in enumerate(taw, 1):
        ws6.column_dimensions[get_column_letter(i)].width = w

    # Only Fixed Fee rows with a matched FF task
    ff_df = df[(df["billing_type"].str.lower() == "fixed fee") & (df["ff_task"].notna())].copy() \
        if "billing_type" in df.columns else df[df["ff_task"].notna()].copy()

    if len(ff_df) > 0:
        task_sum = ff_df.groupby(
            ["ff_task","project_type"], as_index=False
        ).agg(
            hours=("hours","sum"),
            project_count=("project","nunique"),
        ).sort_values(["ff_task","project_type"])

        # Total hours per type for % calc
        type_totals = ff_df.groupby("project_type")["hours"].sum().to_dict()

        _prev_task_t = None; _grp_idx_t = 0
        for r_idx, (_, row) in enumerate(task_sum.iterrows(), 3):
            type_total = type_totals.get(row["project_type"], 0)
            pct        = row["hours"] / type_total if type_total > 0 else 0
            proj_cnt   = row["project_count"] if row["project_count"] > 0 else 1
            raw_avg    = row["hours"] / proj_cnt
            # Round to nearest .25
            avg_hrs    = round(raw_avg * 4) / 4
            bg         = bgs[r_idx % 2]

            task_colors = {
                "Configuration":        "EBF5FB",
                "Enablement/Training":  "EAF9F1",
                "Post Go-live Support": "FEF9E7",
                "Project Management":   "F4ECF7",
            }
            bg, _grp_idx_t = group_bg(row["ff_task"], _prev_task_t, _grp_idx_t)
            _prev_task_t = row["ff_task"]
            task_bg = task_colors.get(row["ff_task"], bg)

            vals = [row["ff_task"], row["project_type"], row["hours"], avg_hrs, pct]
            fmts = [None, None, "#,##0.00", "#,##0.00", "0.0%"]
            for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
                cell = ws6.cell(row=r_idx, column=c_idx, value=val)
                style_cell(cell, task_bg if c_idx == 1 else bg, fmt=fmt,
                           align="right" if c_idx in (3,4,5) else "left")
    else:
        ws6.cell(row=3, column=1, value="No Fixed Fee task data found. Check Billing Type and Task/Case columns.")

    # ── 6. PROJECT COUNT BY TYPE ─────────────────────────────
    ws_pc = wb.create_sheet("Project Count")
    ws_pc.sheet_properties.tabColor = "2980B9"
    ws_pc.freeze_panes = "A3"

    pch = ["Project Type","Billing Type","Project Count","Projects"]
    pcw = [25,14,14,60]
    write_title(ws_pc, "PROJECT COUNT — Distinct Projects by Type (excl. Internal)", len(pch))
    style_header(ws_pc, 2, pch, TEAL)
    for i, w in enumerate(pcw, 1):
        ws_pc.column_dimensions[get_column_letter(i)].width = w

    # Exclude internal, count distinct projects per type + billing type
    pc_df = df[df["billing_type"].str.lower() != "internal"].copy()         if "billing_type" in df.columns else df.copy()

    pc_sum = pc_df.groupby(["project_type","billing_type"], as_index=False).agg(
        project_count=("project","nunique"),
        projects=("project", lambda x: ", ".join(sorted(x.unique())))
    ).sort_values(["project_type","billing_type"])

    _prev_ptype_pc = None; _grp_idx_pc = 0
    for r_idx, (_, row) in enumerate(pc_sum.iterrows(), 3):
        bg, _grp_idx_pc = group_bg(row["project_type"], _prev_ptype_pc, _grp_idx_pc)
        _prev_ptype_pc = row["project_type"]
        vals = [row["project_type"], row["billing_type"],
                row["project_count"], row["projects"]]
        fmts = [None, None, "#,##0", None]
        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_pc.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, bg, fmt=fmt,
                       align="center" if c_idx in (2,3) else "left")

    # ── 7. SKIPPED ROWS ──────────────────────────────────────
    ws7 = wb.create_sheet("Skipped Rows")
    ws7.sheet_properties.tabColor = "E74C3C"
    ws7.freeze_panes = "A3"

    skh = ["Employee","Project","Project Type","Billing Type","Date","Hours","Reason"]
    skw = [20,35,20,14,14,10,45]
    write_title(ws7, "SKIPPED ROWS — Not Included in Utilization Calculations", len(skh))
    style_header(ws7, 2, skh, "C0392B")
    for i, w in enumerate(skw, 1):
        ws7.column_dimensions[get_column_letter(i)].width = w

    skip_cols = ["employee","project","project_type","billing_type","date","hours","notes"]
    skipped = df[df["credit_tag"] == "SKIPPED"]
    if len(skipped) > 0:
        for r_idx, (_, row) in enumerate(skipped.iterrows(), 3):
            for c_idx, col in enumerate(skip_cols, 1):
                val  = row.get(col, "")
                cell = ws7.cell(row=r_idx, column=c_idx, value=val)
                fmt  = "YYYY-MM-DD" if col == "date" else "#,##0.00" if col == "hours" else None
                style_cell(cell, "FDECED", fmt=fmt,
                           align="right" if col == "hours" else "center" if col == "date" else "left")
    else:
        ws7.cell(row=3, column=1, value="No skipped rows — all entries were processed.")

    # ── Reorder sheets: Project Count first, Processed Data last ────────────
    sheet_order = [
        "Project Count",
        "SUMMARY - By Employee",
        "SUMMARY - By Project",
        "ZCO Non-Billable",
        "Task Analysis",
        "Skipped Rows",
        "PROCESSED_DATA",
    ]
    # Rebuild workbook sheet order directly
    existing = [s for s in sheet_order if s in wb.sheetnames]
    # Any sheets not in our list go at the end
    remaining = [s for s in wb.sheetnames if s not in existing]
    wb._sheets = [wb[s] for s in existing + remaining]

    # Save
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Streamlit UI ──────────────────────────────────────────────────────────────
def main():
    st.markdown("""
        <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
        <style>
            html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
            h1, h2, h3, .stMarkdown, .stDataFrame, label, button { font-family: 'Manrope', sans-serif !important; }
        </style>
        <div style='background-color:#1e2c63;padding:24px 32px;border-radius:8px;margin-bottom:24px;font-family:Manrope,sans-serif'>
            <h1 style='color:white;margin:0;font-size:28px;font-family:Manrope,sans-serif'>Utilization Credit Report</h1>
            <p style='color:#aac4d0;margin:6px 0 0 0;font-size:14px;font-family:Manrope,sans-serif'>
                Upload your NetSuite time detail export to generate a utilization credit report.
            </p>
        </div>
    """, unsafe_allow_html=True)

    # ── Upload ────────────────────────────────────────────────
    st.subheader("Step 1 — Upload NetSuite Time Detail Export")
    st.caption("Supported columns: Employee, Region, Project, Project Type, Billing Type, "
               "Hours to Date, Date, Hours, Approval Status, Case/Task/Event, Non-Billable")

    uploaded = st.file_uploader(
        "Drop your file here or click to browse",
        type=["csv", "xlsx", "xls"],
        help="Supports CSV and Excel files exported from NetSuite"
    )

    if not uploaded:
        # Show stored config as reference
        with st.expander("📋 View stored scope map & available hours"):
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Fixed Fee Scope Hours**")
                scope_df = pd.DataFrame(list(DEFAULT_SCOPE.items()), columns=["Project Type","Scoped Hrs"])
                st.dataframe(scope_df, hide_index=True, use_container_width=True)
            with c2:
                st.markdown("**Available Hours by Region (2026)**")
                avail_df = pd.DataFrame([
                    {"Region": r, **months} for r, months in AVAIL_HOURS.items()
                ])
                st.dataframe(avail_df, hide_index=True, use_container_width=True)
        st.info("👆 Upload your NetSuite export to continue.")
        return

    # Load file
    try:
        ext = os.path.splitext(uploaded.name)[1].lower()
        df_raw = pd.read_excel(uploaded) if ext in (".xlsx",".xls") else \
                 pd.read_csv(uploaded, encoding="utf-8") if True else \
                 pd.read_csv(uploaded, encoding="latin-1")
    except Exception:
        try:
            df_raw = pd.read_csv(uploaded, encoding="latin-1")
        except Exception as e:
            st.error(f"Could not read file: {e}"); return

    st.success(f"✅ Loaded **{len(df_raw):,} rows** from `{uploaded.name}`")
    with st.expander("Preview raw data (first 5 rows)"):
        st.dataframe(df_raw.head(), use_container_width=True)

    st.divider()

    # ── Process ───────────────────────────────────────────────
    st.subheader("Step 2 — Generate Report")

    if st.button("▶️ Run Utilization Engine", type="primary"):
        with st.spinner("Processing..."):
            try:
                df, consumed, skipped_df = assign_credits(df_raw.copy(), DEFAULT_SCOPE)
            except Exception as e:
                st.error(f"Processing error: {e}"); return

        st.success("✅ Processing complete!")

        # Metrics
        total_rows     = len(df[df["credit_tag"] != "SKIPPED"])
        total_credit   = df["credit_hrs"].sum()
        total_variance = df["variance_hrs"].sum()
        total_nb       = len(df[df["credit_tag"] == "NON-BILLABLE"])
        total_overrun  = len(df[df["credit_tag"] == "OVERRUN"])

        m1,m2,m3,m4,m5 = st.columns(5)
        m1.metric("Rows Processed",   f"{total_rows:,}")
        m2.metric("Total Credit Hrs", f"{total_credit:,.1f}")
        m3.metric("Variance Hrs",     f"{total_variance:,.1f}")
        m4.metric("Non-Billable Rows",f"{total_nb:,}")
        m5.metric("Overrun Rows",     f"{total_overrun:,}")

        st.markdown("**Credit Tag Breakdown**")
        tag_counts = df[df["credit_tag"] != "SKIPPED"]["credit_tag"].value_counts()
        tcols = st.columns(len(tag_counts))
        for i, (tag, count) in enumerate(tag_counts.items()):
            tcols[i].markdown(f"{TAG_BADGE.get(tag,'⚪')} **{tag}**  \n{count:,} rows")

        st.divider()

        # Previews
        tab1, tab2, tab3, tab4, tab5 = st.tabs(
            ["By Employee", "By Project", "ZCO Non-Billable", "Task Analysis", "Detail"]
        )

        emp_region = {}
        if "region" in df.columns:
            emp_region = df.dropna(subset=["region"]).groupby("employee")["region"].first().to_dict()

        with tab1:
            emp_sum = df[df["credit_tag"] != "SKIPPED"].groupby(
                ["employee","period"], as_index=False
            ).agg(hours_this_period=("hours","sum"), credit_hrs=("credit_hrs","sum"),
                  variance_hrs=("variance_hrs","sum")).sort_values(["employee","period"])
            emp_sum["region"]    = emp_sum["employee"].map(emp_region)
            emp_sum["avail_hrs"] = emp_sum.apply(
                lambda r: get_avail_hours(r["region"], r["period"]), axis=1)
            emp_sum["util_pct"]  = emp_sum.apply(
                lambda r: f"{r['credit_hrs']/r['avail_hrs']*100:.1f}%" if r["avail_hrs"] else "—", axis=1)
            st.dataframe(emp_sum, use_container_width=True, hide_index=True)

        with tab2:
            proj_sum = df[df["credit_tag"] != "SKIPPED"].groupby(
                ["project","project_type"], as_index=False
            ).agg(hours_this_period=("hours","sum"), credit_hrs=("credit_hrs","sum"),
                  variance_hrs=("variance_hrs","sum")).sort_values("project")
            proj_sum["scope_hrs"]  = proj_sum["project_type"].apply(
                lambda pt: (lambda m: max(m, key=lambda x: len(x[0]))[1] if m else "—")(
                    [(k,v) for k,v in DEFAULT_SCOPE.items() if k.strip().lower() in str(pt).strip().lower()]))
            htd_seeds_ui = df[df["credit_tag"] != "SKIPPED"].groupby("project")["htd_start"].first()
            proj_sum["htd_seed"]    = proj_sum["project"].map(htd_seeds_ui).fillna(0)
            proj_sum["updated_htd"] = proj_sum["htd_seed"] + proj_sum["hours_this_period"]
            proj_sum["burn_pct"]    = proj_sum.apply(
                lambda r: f"{r['updated_htd']/r['scope_hrs']*100:.1f}%"
                if isinstance(r["scope_hrs"], (int,float)) and r["scope_hrs"] > 0 else "—", axis=1)
            st.dataframe(proj_sum, use_container_width=True, hide_index=True)

        with tab3:
            zco_df = df[df["credit_tag"] == "NON-BILLABLE"]
            if "task" in zco_df.columns and len(zco_df) > 0:
                zco_sum = zco_df.groupby(["task","employee","period"], as_index=False
                ).agg(hours=("hours","sum")).sort_values(["task","employee","period"])
                st.dataframe(zco_sum, use_container_width=True, hide_index=True)
            else:
                st.info("No ZCO Non-Billable entries in this dataset.")

        with tab4:
            ff_df = df[df["ff_task"].notna()] if "ff_task" in df.columns else pd.DataFrame()
            if len(ff_df) > 0:
                task_sum = ff_df.groupby(["ff_task","project_type"], as_index=False
                ).agg(hours=("hours","sum")).sort_values(["ff_task","project_type"])
                type_totals = ff_df.groupby("project_type")["hours"].sum()
                task_sum["pct_of_type"] = task_sum.apply(
                    lambda r: f"{r['hours']/type_totals.get(r['project_type'],1)*100:.1f}%", axis=1)
                st.dataframe(task_sum, use_container_width=True, hide_index=True)
            else:
                st.info("No Fixed Fee task data found. Check Billing Type and Task/Case columns.")

        with tab5:
            display_cols = ["employee","region","project","project_type","billing_type",
                            "hours_to_date","date","hours","credit_hrs","variance_hrs",
                            "updated_htd","credit_tag","notes"]
            existing = [c for c in display_cols if c in df.columns]
            st.dataframe(df[existing].head(100), use_container_width=True, hide_index=True)
            if len(df) > 100:
                st.caption(f"Showing first 100 of {len(df):,} rows. Full data in Excel download.")

        st.divider()

        # Download
        st.subheader("Download Report")
        with st.spinner("Building Excel file..."):
            excel_buf = build_excel(df, DEFAULT_SCOPE, consumed)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        filename  = f"utilization_report_{timestamp}.xlsx"

        st.download_button(
            label="⬇️ Download Excel Report",
            data=excel_buf,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption(f"`{filename}` — 6 tabs: Processed Data · By Employee · By Project · "
                   f"ZCO Non-Billable · Task Analysis · Skipped Rows")


if __name__ == "__main__":
    main()
