"""
Utilization Credit Report — Web App
Run with: streamlit run utilization_app.py
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
    page_icon="📊",
    layout="wide"
)

# ── Styles ────────────────────────────────────────────────────────────────────
NAVY     = "1F2D4E"
TEAL     = "2A7F8A"
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

def thin_border():
    s = Side(style="thin", color=MID_GRAY)
    return Border(left=s, right=s, top=s, bottom=s)

def hdr_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def row_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


# ── Core logic (same as CLI script) ──────────────────────────────────────────

def auto_detect_columns(df):
    cols_lower = {c.lower().strip(): c for c in df.columns}
    checks = {
        "employee":     ["employee", "employee name", "name", "resource"],
        "project":      ["project", "project name", "job"],
        "project_type": ["project type", "type", "project_type"],
        "date":         ["date", "time entry date", "entry date", "work date"],
        "hours":        ["hours", "duration", "hours logged", "time", "qty"],
        "approval":     ["approval status", "approval", "status"],
        "task":         ["case/task/event", "task", "case", "event", "memo"],
        "non_billable": ["non-billable", "non billable", "nonbillable",
                         "non_billable", "is non billable"],
        "billing_type": ["billing type", "billing_type", "bill type", "billtype"],
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


def assign_credits(df, scope_map):
    col_map, unmatched = auto_detect_columns(df)
    if unmatched:
        st.warning(f"⚠️ Could not auto-detect columns: {unmatched}. Check your file headers.")

    df = df.rename(columns=col_map)
    df["non_billable"] = df["non_billable"].astype(str).str.strip().str.upper()
    df["hours"] = pd.to_numeric(df["hours"], errors="coerce").fillna(0)
    df["date"]  = pd.to_datetime(df["date"], errors="coerce")
    df["period"] = df["date"].dt.strftime("%Y-%m").fillna("Unknown")

    consumed = {}
    credit_hrs_list   = []
    variance_hrs_list = []
    credit_tag_list   = []
    notes_list        = []

    for _, row in df.iterrows():
        proj  = str(row.get("project", "")).strip()
        ptype = str(row.get("project_type", "")).strip()
        hrs   = float(row.get("hours", 0))
        nb    = str(row.get("non_billable", "NO")).strip().upper()
        is_zco    = "ZCO" in ptype.upper()
        bill_type = str(row.get("billing_type", "")).strip().lower()
        is_tm     = bill_type == "t&m"

        if hrs <= 0:
            credit_hrs_list.append(0)
            variance_hrs_list.append(0)
            credit_tag_list.append("SKIPPED")
            notes_list.append("Zero or missing hours")
            continue

        # Rule 1: NB + ZCO = excluded
        if nb == "YES" and is_zco:
            credit_hrs_list.append(0)
            variance_hrs_list.append(0)
            credit_tag_list.append("NON-BILLABLE")
            notes_list.append("Excluded: ZCO Internal Project")
            continue

        # Rule 2: T&M billing type = always full credit, no cap
        if is_tm:
            credit_hrs_list.append(hrs)
            variance_hrs_list.append(0)
            credit_tag_list.append("CREDITED")
            notes_list.append("T&M: full credit")
            continue

        # Rule 3: Billable (non-billable = No, no billing type) = full credit
        if nb == "NO" and not is_tm:
            credit_hrs_list.append(hrs)
            variance_hrs_list.append(0)
            credit_tag_list.append("CREDITED")
            notes_list.append("Billable: full credit")
            continue

        # Rule 4: NB non-ZCO = capped at scope
        # Match by specificity — longest matching key wins
        # e.g. "Capture and E-Invoicing" beats "Capture" for that type
        _ptype_lower = ptype.strip().lower()
        _matches = [
            (k, float(v)) for k, v in scope_map.items()
            if k.strip().lower() in _ptype_lower
        ]
        if _matches:
            # Pick the match with the longest key (most specific)
            scope_hrs = max(_matches, key=lambda x: len(x[0]))[1]
        else:
            scope_hrs = None

        if scope_hrs is None:
            credit_hrs_list.append(0)
            variance_hrs_list.append(hrs)
            credit_tag_list.append("UNCONFIGURED")
            notes_list.append(f"No scope defined for: {ptype}")
            continue

        already = consumed.get(proj, 0)
        remaining = scope_hrs - already

        if remaining <= 0:
            credit_hrs_list.append(0)
            variance_hrs_list.append(hrs)
            credit_tag_list.append("OVERRUN")
            notes_list.append(f"Scope exhausted (cap: {scope_hrs:.0f}h)")
        elif hrs <= remaining:
            consumed[proj] = already + hrs
            credit_hrs_list.append(hrs)
            variance_hrs_list.append(0)
            credit_tag_list.append("CREDITED")
            notes_list.append(f"NB within scope ({already:.1f}/{scope_hrs:.0f}h used)")
        else:
            consumed[proj] = already + remaining
            credit_hrs_list.append(remaining)
            variance_hrs_list.append(hrs - remaining)
            credit_tag_list.append("PARTIAL")
            notes_list.append(
                f"Split: {remaining:.2f}h credited / {hrs - remaining:.2f}h overrun"
            )

    df["credit_hrs"]   = credit_hrs_list
    df["variance_hrs"] = variance_hrs_list
    df["credit_tag"]   = credit_tag_list
    df["notes"]        = notes_list

    return df


# ── Excel builder ─────────────────────────────────────────────────────────────

def style_header(ws, row, headers, fill_color=NAVY):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font      = Font(name="Arial", bold=True, color=WHITE, size=10)
        c.fill      = hdr_fill(fill_color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = thin_border()
    ws.row_dimensions[row].height = 22

def style_cell(cell, bg, fmt=None, bold=False, align="left"):
    cell.fill      = row_fill(bg)
    cell.font      = Font(name="Arial", size=10, bold=bold)
    cell.border    = thin_border()
    cell.alignment = Alignment(horizontal=align, vertical="center")
    if fmt:
        cell.number_format = fmt

def write_title(ws, title, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font      = Font(name="Arial", bold=True, size=14, color=WHITE)
    c.fill      = hdr_fill(NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30


def build_excel(df, scope_map):
    wb = Workbook()
    wb.remove(wb.active)

    # ── PROCESSED DATA ────────────────────────────────────────
    ws = wb.create_sheet("PROCESSED_DATA")
    ws.sheet_properties.tabColor = TEAL
    ws.freeze_panes = "A3"

    headers = ["Employee","Project","Project Type","Billing Type","Date","Hours Logged",
               "Approval","Task/Case","Non-Billable","Credit Hrs",
               "Variance Hrs","Credit Tag","Period","Notes"]
    widths  = [20,35,20,14,14,13,14,25,13,12,12,16,12,45]
    cols    = ["employee","project","project_type","billing_type","date","hours",
               "approval","task","non_billable","credit_hrs",
               "variance_hrs","credit_tag","period","notes"]

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
            fmt  = None
            bold = False
            align = "left"
            if col == "date" and pd.notna(val):
                fmt = "YYYY-MM-DD"; align = "center"
            elif col in ("hours","credit_hrs","variance_hrs"):
                fmt = "#,##0.00"; align = "right"
            elif col == "credit_tag":
                bold = True; align = "center"
            elif col in ("period","billing_type"):
                align = "center"
            style_cell(cell, bg, fmt=fmt, bold=bold, align=align)

    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}2"

    # ── EMPLOYEE SUMMARY ──────────────────────────────────────
    ws2 = wb.create_sheet("SUMMARY - By Employee")
    ws2.sheet_properties.tabColor = NAVY
    ws2.freeze_panes = "A3"

    eh = ["Employee","Period","Hours Booked","Credit Hrs","Variance Hrs","Util %"]
    ew = [22,12,14,14,14,10]
    write_title(ws2, "SUMMARY — Utilization by Employee", len(eh))
    style_header(ws2, 2, eh, TEAL)
    for i, w in enumerate(ew, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    emp_sum = df[df["credit_tag"] != "SKIPPED"].groupby(
        ["employee","period"], as_index=False
    ).agg(
        hours_booked=("hours","sum"),
        credit_hrs=("credit_hrs","sum"),
        variance_hrs=("variance_hrs","sum"),
    ).sort_values(["employee","period"])

    AVAIL = 160
    bgs = [WHITE, LTGRAY]
    for r_idx, (_, row) in enumerate(emp_sum.iterrows(), 3):
        util   = row["credit_hrs"] / AVAIL if AVAIL > 0 else 0
        bg     = bgs[r_idx % 2]
        util_bg = ("EAF9F1" if util >= 0.8 else "FEF9E7" if util >= 0.6 else "FDECED")
        vals = [row["employee"], row["period"], row["hours_booked"],
                row["credit_hrs"], row["variance_hrs"], util]
        fmts = [None, None, "#,##0.00", "#,##0.00", "#,##0.00", "0.0%"]
        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, util_bg if c_idx == 6 else bg, fmt=fmt,
                       align="right" if c_idx > 2 else "left")

    # ── PROJECT SUMMARY ───────────────────────────────────────
    ws3 = wb.create_sheet("SUMMARY - By Project")
    ws3.sheet_properties.tabColor = "E67E22"
    ws3.freeze_panes = "A3"

    ph = ["Project","Project Type","Scoped Hrs","Hours Booked",
          "Credit Hrs","Variance Hrs","Burn %","Status"]
    pw = [35,20,12,12,12,12,10,12]
    write_title(ws3, "SUMMARY — Utilization by Project", len(ph))
    style_header(ws3, 2, ph, TEAL)
    for i, w in enumerate(pw, 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    proj_sum = df[df["credit_tag"] != "SKIPPED"].groupby(
        ["project","project_type"], as_index=False
    ).agg(
        hours_booked=("hours","sum"),
        credit_hrs=("credit_hrs","sum"),
        variance_hrs=("variance_hrs","sum"),
    ).sort_values("project")

    for r_idx, (_, row) in enumerate(proj_sum.iterrows(), 3):
        ptype   = str(row["project_type"]).strip()
        _pt = ptype.lower()
        _pm = [(k, float(v)) for k, v in scope_map.items()
               if k.strip().lower() in _pt]
        scope_h = max(_pm, key=lambda x: len(x[0]))[1] if _pm else 0
        cred_h  = row["credit_hrs"]
        vari_h  = row["variance_hrs"]
        burn    = cred_h / scope_h if scope_h > 0 else 0
        if vari_h > 0:      status = "OVERRUN"
        elif burn >= 0.9:   status = "AT RISK"
        elif burn > 0:      status = "ON TRACK"
        else:               status = "—"

        status_bg = {"OVERRUN":"FDECED","AT RISK":"FEF9E7","ON TRACK":"EAF9F1"}.get(status, LTGRAY)
        bg = bgs[r_idx % 2]

        vals = [row["project"], ptype, scope_h or "—", row["hours_booked"],
                cred_h, vari_h, burn if scope_h > 0 else "—", status]
        fmts = [None, None, "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00", "0.0%", None]

        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, status_bg if c_idx == 8 else bg,
                       fmt=fmt, bold=(c_idx == 8),
                       align="right" if c_idx in (3,4,5,6,7) else "center" if c_idx == 8 else "left")

    # ── PERIOD SUMMARY ────────────────────────────────────────
    ws4 = wb.create_sheet("SUMMARY - Period vs Period")
    ws4.sheet_properties.tabColor = "8E44AD"
    ws4.freeze_panes = "A3"

    penh = ["Employee","Period","Credit Hrs","Util %"]
    penw = [22,12,14,10]
    write_title(ws4, "SUMMARY — Credit Hours by Employee by Period", len(penh))
    style_header(ws4, 2, penh, TEAL)
    for i, w in enumerate(penw, 1):
        ws4.column_dimensions[get_column_letter(i)].width = w

    per_sum = df[df["credit_tag"].isin(["CREDITED","PARTIAL"])].groupby(
        ["employee","period"], as_index=False
    ).agg(credit_hrs=("credit_hrs","sum")).sort_values(["employee","period"])

    for r_idx, (_, row) in enumerate(per_sum.iterrows(), 3):
        util   = row["credit_hrs"] / AVAIL
        bg     = bgs[r_idx % 2]
        util_bg = ("EAF9F1" if util >= 0.8 else "FEF9E7" if util >= 0.6 else "FDECED")
        vals = [row["employee"], row["period"], row["credit_hrs"], util]
        fmts = [None, None, "#,##0.00", "0.0%"]
        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws4.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, util_bg if c_idx == 4 else bg, fmt=fmt,
                       align="right" if c_idx > 2 else "left")

    # Save to bytes buffer
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Streamlit UI ──────────────────────────────────────────────────────────────

def main():
    # Header
    st.markdown("""
        <div style='background-color:#1F2D4E;padding:24px 32px;border-radius:8px;margin-bottom:24px'>
            <h1 style='color:white;margin:0;font-size:28px'>📊 Utilization Credit Report</h1>
            <p style='color:#aac4d0;margin:6px 0 0 0;font-size:14px'>
                Upload your NetSuite time detail export to generate a utilization credit report.
            </p>
        </div>
    """, unsafe_allow_html=True)

    # ── Step 1: Scope configuration ──────────────────────────
    st.subheader("Step 1 — Configure Scope Hours by Project Type")
    st.caption("Define the budget cap (hours) for each project type that has a scope limit. "
               "T&M / Billable projects don't need an entry — they're always credited 1:1.")

    col1, col2 = st.columns([2, 1])
    with col1:
        scope_input = st.text_area(
            "Project Type → Scoped Hours (one per line, format: Type Name = Hours)",
            value="Capture = 20\nApprovals = 17\nReconcile = 17\nPSP = 18\nPayments = 30\nReconcile 2.0 = 20\nCC = 6\nSFTP = 12\nPremium - 10 = 10\nPremium - 20 = 20\nE-Invoicing = 15\nCapture and E-Invoicing = 30\nAdditional Subsidiary = 2",
            height=120,
        )

    scope_map = {}
    for line in scope_input.strip().split("\n"):
        if "=" in line:
            k, v = line.split("=", 1)
            try:
                scope_map[k.strip()] = float(v.strip())
            except ValueError:
                st.warning(f"Could not parse: {line}")

    with col2:
        st.markdown("**Current scope map:**")
        if scope_map:
            for k, v in scope_map.items():
                st.markdown(f"- **{k}**: {v:.0f} hrs")
        else:
            st.caption("No scope entries defined.")

    st.divider()

    # ── Step 2: File upload ───────────────────────────────────
    st.subheader("Step 2 — Upload NetSuite Time Detail Export")
    uploaded = st.file_uploader(
        "Drop your file here or click to browse",
        type=["csv", "xlsx", "xls"],
        help="Supports CSV and Excel files exported from NetSuite"
    )

    if not uploaded:
        st.info("👆 Upload your NetSuite export to continue.")
        return

    # Load file
    try:
        ext = os.path.splitext(uploaded.name)[1].lower()
        if ext in (".xlsx", ".xls"):
            df_raw = pd.read_excel(uploaded)
        else:
            try:
                df_raw = pd.read_csv(uploaded, encoding="utf-8")
            except UnicodeDecodeError:
                df_raw = pd.read_csv(uploaded, encoding="latin-1")
    except Exception as e:
        st.error(f"Could not read file: {e}")
        return

    st.success(f"✅ Loaded **{len(df_raw):,} rows** from `{uploaded.name}`")

    with st.expander("Preview raw data (first 5 rows)"):
        st.dataframe(df_raw.head(), use_container_width=True)

    st.divider()

    # ── Step 3: Process ───────────────────────────────────────
    st.subheader("Step 3 — Generate Report")

    if st.button("▶️ Run Utilization Engine", type="primary", use_container_width=False):
        with st.spinner("Processing..."):
            try:
                df = assign_credits(df_raw.copy(), scope_map)
            except Exception as e:
                st.error(f"Processing error: {e}")
                return

        st.success("✅ Processing complete!")

        # Summary metrics
        total_rows    = len(df[df["credit_tag"] != "SKIPPED"])
        total_credit  = df["credit_hrs"].sum()
        total_variance = df["variance_hrs"].sum()
        total_nb      = len(df[df["credit_tag"] == "NON-BILLABLE"])
        total_overrun = len(df[df["credit_tag"] == "OVERRUN"])

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Rows Processed", f"{total_rows:,}")
        m2.metric("Total Credit Hrs", f"{total_credit:,.1f}")
        m3.metric("Total Variance Hrs", f"{total_variance:,.1f}")
        m4.metric("Non-Billable Rows", f"{total_nb:,}")
        m5.metric("Overrun Rows", f"{total_overrun:,}")

        # Tag breakdown
        st.markdown("**Credit Tag Breakdown**")
        tag_counts = df[df["credit_tag"] != "SKIPPED"]["credit_tag"].value_counts()
        cols = st.columns(len(tag_counts))
        for i, (tag, count) in enumerate(tag_counts.items()):
            badge = TAG_BADGE.get(tag, "⚪")
            cols[i].markdown(f"{badge} **{tag}**  \n{count:,} rows")

        st.divider()

        # Preview tabs
        tab1, tab2, tab3 = st.tabs(["By Employee", "By Project", "Detail"])

        with tab1:
            emp_sum = df[df["credit_tag"] != "SKIPPED"].groupby(
                ["employee","period"], as_index=False
            ).agg(
                hours_booked=("hours","sum"),
                credit_hrs=("credit_hrs","sum"),
                variance_hrs=("variance_hrs","sum"),
            ).sort_values(["employee","period"])
            emp_sum["util_pct"] = (emp_sum["credit_hrs"] / 160 * 100).round(1).astype(str) + "%"
            st.dataframe(emp_sum, use_container_width=True, hide_index=True)

        with tab2:
            proj_sum = df[df["credit_tag"] != "SKIPPED"].groupby(
                ["project","project_type"], as_index=False
            ).agg(
                hours_booked=("hours","sum"),
                credit_hrs=("credit_hrs","sum"),
                variance_hrs=("variance_hrs","sum"),
            ).sort_values("project")
            proj_sum["scope_hrs"] = proj_sum["project_type"].apply(
                lambda pt: (
                    lambda matches: max(matches, key=lambda x: len(x[0]))[1]
                    if matches else "—"
                )([(k, v) for k, v in scope_map.items()
                   if k.strip().lower() in str(pt).strip().lower()])
            )
            proj_sum["burn_pct"] = proj_sum.apply(
                lambda r: f"{r['credit_hrs']/r['scope_hrs']*100:.1f}%"
                if isinstance(r["scope_hrs"], float) and r["scope_hrs"] > 0 else "—", axis=1
            )
            st.dataframe(proj_sum, use_container_width=True, hide_index=True)

        with tab3:
            display_cols = ["employee","project","project_type","billing_type","date",
                            "hours","credit_hrs","variance_hrs","credit_tag","notes"]
            existing = [c for c in display_cols if c in df.columns]
            st.dataframe(df[existing].head(100), use_container_width=True, hide_index=True)
            if len(df) > 100:
                st.caption(f"Showing first 100 of {len(df):,} rows. Full data in Excel download.")

        st.divider()

        # ── Download ──────────────────────────────────────────
        st.subheader("Download Report")
        with st.spinner("Building Excel file..."):
            excel_buf = build_excel(df, scope_map)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        filename  = f"utilization_report_{timestamp}.xlsx"

        st.download_button(
            label="⬇️ Download Excel Report",
            data=excel_buf,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=False,
        )
        st.caption(f"File: `{filename}` — 4 tabs: Processed Data, By Employee, By Project, Period Summary")


if __name__ == "__main__":
    main()
