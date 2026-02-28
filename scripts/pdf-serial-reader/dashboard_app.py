#!/usr/bin/env python3
"""
Streamlit dashboard for tool inventory tracking.
Design inspired by seed.com - dark olive, clean typography, minimal.

Usage:
    streamlit run scripts/pdf-serial-reader/dashboard_app.py
"""

import os
import sys
from collections import Counter
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit_authenticator as stauth
import yaml
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from googleapiclient.discovery import build

SHEET_ID = "1wK92FpXq-07LdYYPCwZi7-C2vruLPs59JM14w4nAggs"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
PROJECT_DIR = Path(__file__).resolve().parent.parent.parent


# ── Google Sheets helpers ──────────────────────────────────────────


def get_credentials():
    """Get Google credentials - cloud (service account) or local (OAuth token)."""
    # Cloud: use service account from Streamlit secrets
    if "gcp_service_account" in st.secrets:
        return service_account.Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=SCOPES
        )

    # Local: use OAuth token file
    for name in ["token_gmail.json", "token.json"]:
        path = PROJECT_DIR / name
        if path.exists():
            creds = Credentials.from_authorized_user_file(str(path), SCOPES)
            if creds and creds.valid:
                return creds
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
                with open(path, "w") as f:
                    f.write(creds.to_json())
                return creds

    st.error("No valid Google credentials found.")
    st.stop()


def get_sheets_service():
    creds = get_credentials()
    return build("sheets", "v4", credentials=creds)


@st.cache_data(ttl=60)
def load_data():
    """Load Out, Inventory, and Modem data from Google Sheet."""
    service = get_sheets_service()
    sheets = service.spreadsheets()

    out_result = sheets.values().get(spreadsheetId=SHEET_ID, range="Out!A:I").execute()
    out_rows = out_result.get("values", [])

    inv_result = sheets.values().get(spreadsheetId=SHEET_ID, range="Inventory!A:C").execute()
    inv_rows = inv_result.get("values", [])

    try:
        modem_result = sheets.values().get(spreadsheetId=SHEET_ID, range="Modem!A:F").execute()
        modem_rows = modem_result.get("values", [])
    except Exception:
        modem_rows = []

    return out_rows, inv_rows, modem_rows


def build_dashboard_df(out_rows, inv_rows):
    """Build dashboard DataFrame from raw data."""
    counts = Counter()
    for row in out_rows[1:]:
        if len(row) >= 6 and row[5]:
            counts[row[5]] += 1

    records = []
    for row in inv_rows[1:]:
        if not row:
            continue
        tool = row[0]
        # Skip crossovers/subs (items with "Box" in the name)
        if "box" in tool.lower():
            continue
        total = int(row[1]) if len(row) > 1 and row[1] else 0
        redress = int(row[2]) if len(row) > 2 and row[2] else 0
        load_out = counts.get(tool, 0)
        ready = total - redress - load_out
        records.append({
            "Tool": tool,
            "Redress": redress,
            "Ready": ready,
            "Load Out": load_out,
            "Total Stock": total,
        })

    return pd.DataFrame(records)


def build_loadout_df(out_rows):
    """Build load out detail DataFrame."""
    if len(out_rows) <= 1:
        return pd.DataFrame()
    headers = out_rows[0]
    data = []
    for row in out_rows[1:]:
        padded = row + [""] * (len(headers) - len(row))
        data.append(padded[:len(headers)])
    return pd.DataFrame(data, columns=headers)


def build_active_jobs_df(out_rows):
    """Build active jobs overview grouped by Customer + Well Name."""
    if len(out_rows) <= 1:
        return pd.DataFrame()
    jobs = {}
    for row in out_rows[1:]:
        if len(row) < 6:
            continue
        customer = row[1] if len(row) > 1 else ""
        well = row[2] if len(row) > 2 else ""
        date = row[3] if len(row) > 3 else ""
        desc = row[5] if len(row) > 5 else ""
        key = (customer, well)
        if key not in jobs:
            jobs[key] = {"Customer": customer, "Well Name": well, "Load Out Date": date, "Tools": [], "Count": 0}
        jobs[key]["Tools"].append(desc)
        jobs[key]["Count"] += 1
    records = []
    for job in jobs.values():
        unique_tools = sorted(set(job["Tools"]))
        records.append({
            "Customer": job["Customer"],
            "Well Name": job["Well Name"],
            "Date": job["Load Out Date"],
            "Tools": len(unique_tools),
            "Items": job["Count"],
        })
    return pd.DataFrame(records)


def build_modem_df(modem_rows):
    """Build upcoming jobs DataFrame from Modem sheet."""
    if len(modem_rows) <= 1:
        return pd.DataFrame(columns=["Customer", "Well Name", "Date", "Coordinator", "Tools Needed", "Status"])
    headers = modem_rows[0]
    data = []
    for row in modem_rows[1:]:
        padded = row + [""] * (len(headers) - len(row))
        data.append(padded[:len(headers)])
    return pd.DataFrame(data, columns=headers)


def save_modem(modem_data):
    """Save modem (upcoming jobs) back to Google Sheet."""
    service = get_sheets_service()
    rows = [["Customer", "Well Name", "Date", "Coordinator", "Tools Needed", "Status"]]
    for _, row in modem_data.iterrows():
        rows.append([str(row[c]) for c in modem_data.columns])

    service.spreadsheets().values().clear(
        spreadsheetId=SHEET_ID, range="Modem!A:F"
    ).execute()
    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range="Modem!A1",
        valueInputOption="RAW",
        body={"values": rows},
    ).execute()


def ensure_modem_sheet():
    """Create the Modem sheet if it doesn't exist."""
    service = get_sheets_service()
    try:
        meta = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
        sheet_names = [s["properties"]["title"] for s in meta["sheets"]]
        if "Modem" not in sheet_names:
            service.spreadsheets().batchUpdate(
                spreadsheetId=SHEET_ID,
                body={"requests": [{"addSheet": {"properties": {"title": "Modem"}}}]},
            ).execute()
            service.spreadsheets().values().update(
                spreadsheetId=SHEET_ID,
                range="Modem!A1",
                valueInputOption="RAW",
                body={"values": [["Customer", "Well Name", "Date", "Coordinator", "Tools Needed", "Status"]]},
            ).execute()
    except Exception:
        pass


def save_inventory(inv_data):
    """Save updated inventory back to Google Sheet."""
    service = get_sheets_service()
    rows = [["Tool", "Total Stock", "Redress"]]
    for _, row in inv_data.iterrows():
        rows.append([row["Tool"], int(row["Total Stock"]), int(row["Redress"])])

    service.spreadsheets().values().clear(
        spreadsheetId=SHEET_ID, range="Inventory!A:C"
    ).execute()
    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range="Inventory!A1",
        valueInputOption="RAW",
        body={"values": rows},
    ).execute()


# ── Seed-inspired HTML table builder ──────────────────────────────


def render_dashboard_table(df):
    """Render a custom HTML table with Seed-inspired styling."""
    html = '<div class="seed-table-wrap"><table class="seed-table">'

    # Header
    html += "<thead><tr>"
    for col in df.columns:
        align = "left" if col == "Tool" else "center"
        html += f'<th style="text-align:{align}">{col}</th>'
    html += "</tr></thead>"

    # Body
    html += "<tbody>"
    for idx, row in df.iterrows():
        html += "<tr>"
        for col in df.columns:
            val = row[col]
            align = "left" if col == "Tool" else "center"
            cell_class = ""

            if col == "Ready":
                if val <= 0:
                    cell_class = "cell-danger"
                elif val <= 2:
                    cell_class = "cell-warn"
                else:
                    cell_class = "cell-ok"
            elif col == "Load Out" and val > 0:
                cell_class = "cell-active"

            html += f'<td class="{cell_class}" style="text-align:{align}">{val}</td>'
        html += "</tr>"
    html += "</tbody></table></div>"

    return html


def render_metric_card(label, value, icon=""):
    """Render a single metric card."""
    return f"""
    <div class="metric-card">
        <div class="metric-value">{value}</div>
        <div class="metric-label">{label}</div>
    </div>
    """


# ── Page config ────────────────────────────────────────────────────


st.set_page_config(
    page_title="Tool Inventory",
    page_icon="wrench",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Seed-inspired CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

    /* ── Global ── */
    .stApp {
        background-color: #2d3a2d;
        font-family: 'Inter', -apple-system, sans-serif;
    }

    .block-container {
        padding: 2rem 3rem !important;
        max-width: 1200px;
    }

    /* ── Hide Streamlit chrome ── */
    #MainMenu, footer, header { visibility: hidden; }
    .stDeployButton { display: none; }

    /* ── Typography ── */
    h1, h2, h3, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
        color: #f5f0e8 !important;
        font-weight: 300 !important;
        letter-spacing: -0.02em;
    }

    p, span, label, .stMarkdown p {
        color: #c8c0b4 !important;
    }

    /* ── Navigation pills ── */
    .stTabs [data-baseweb="tab-list"] {
        background: rgba(255,255,255,0.05);
        border-radius: 100px;
        padding: 4px;
        gap: 4px;
        border: none;
        width: fit-content;
    }

    .stTabs [data-baseweb="tab"] {
        background: transparent;
        border: none;
        border-radius: 100px;
        color: #a09888 !important;
        font-size: 14px;
        font-weight: 500;
        padding: 8px 24px;
        letter-spacing: 0.01em;
    }

    .stTabs [data-baseweb="tab"]:hover {
        color: #f5f0e8 !important;
        background: rgba(255,255,255,0.05);
    }

    .stTabs [aria-selected="true"] {
        background: rgba(255,255,255,0.12) !important;
        color: #f5f0e8 !important;
    }

    .stTabs [data-baseweb="tab-highlight"] {
        display: none;
    }

    .stTabs [data-baseweb="tab-border"] {
        display: none;
    }

    /* ── Metric cards ── */
    .metrics-row {
        display: flex;
        gap: 16px;
        margin-bottom: 32px;
    }

    .metric-card {
        flex: 1;
        background: rgba(255,255,255,0.06);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 16px;
        padding: 24px;
        text-align: center;
        transition: all 0.2s ease;
    }

    .metric-card:hover {
        background: rgba(255,255,255,0.09);
        border-color: rgba(255,255,255,0.12);
    }

    .metric-value {
        font-size: 36px;
        font-weight: 300;
        color: #f5f0e8;
        letter-spacing: -0.03em;
        line-height: 1;
        margin-bottom: 8px;
    }

    .metric-label {
        font-size: 12px;
        font-weight: 500;
        color: #8a8070;
        text-transform: uppercase;
        letter-spacing: 0.1em;
    }

    /* ── Dashboard table ── */
    .seed-table-wrap {
        border-radius: 16px;
        overflow: hidden;
        border: 1px solid rgba(255,255,255,0.08);
    }

    .seed-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
    }

    .seed-table thead th {
        background: rgba(255,255,255,0.08);
        color: #8a8070;
        font-weight: 500;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        padding: 14px 20px;
        border: none;
    }

    .seed-table tbody td {
        padding: 16px 20px;
        color: #f5f0e8;
        border-bottom: 1px solid rgba(255,255,255,0.05);
        font-weight: 400;
        transition: all 0.15s ease;
    }

    .seed-table tbody tr:last-child td {
        border-bottom: none;
    }

    .seed-table tbody tr:hover td {
        background: rgba(255,255,255,0.04);
    }

    /* Status cells */
    .cell-ok {
        color: #7dbe6a !important;
        font-weight: 500 !important;
    }

    .cell-warn {
        color: #e0a84e !important;
        font-weight: 600 !important;
    }

    .cell-danger {
        color: #e05a5a !important;
        font-weight: 600 !important;
    }

    .cell-active {
        color: #6aadbe !important;
        font-weight: 500 !important;
    }

    /* ── Streamlit data editor overrides ── */
    .stDataFrame, [data-testid="stDataFrame"] {
        border-radius: 16px;
        overflow: hidden;
        border: 1px solid rgba(255,255,255,0.08) !important;
    }

    [data-testid="stDataFrame"] > div {
        background: rgba(255,255,255,0.03);
    }

    /* ── Button ── */
    .stButton > button {
        background: rgba(255,255,255,0.1) !important;
        color: #f5f0e8 !important;
        border: 1px solid rgba(255,255,255,0.15) !important;
        border-radius: 100px !important;
        font-weight: 500 !important;
        font-size: 14px !important;
        padding: 10px 32px !important;
        letter-spacing: 0.01em;
        transition: all 0.2s ease !important;
    }

    .stButton > button:hover {
        background: rgba(255,255,255,0.18) !important;
        border-color: rgba(255,255,255,0.25) !important;
    }

    .stButton > button[kind="primary"] {
        background: #f5f0e8 !important;
        color: #2d3a2d !important;
        border-color: #f5f0e8 !important;
    }

    .stButton > button[kind="primary"]:hover {
        background: #ffffff !important;
    }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background: #242e24;
        border-right: 1px solid rgba(255,255,255,0.06);
    }

    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2 {
        color: #f5f0e8 !important;
    }

    /* ── Divider ── */
    hr {
        border-color: rgba(255,255,255,0.08) !important;
    }

    /* ── Success/info messages ── */
    .stSuccess, .stInfo {
        background: rgba(255,255,255,0.06) !important;
        border: 1px solid rgba(255,255,255,0.1) !important;
        color: #c8c0b4 !important;
        border-radius: 12px !important;
    }

    /* ── Caption ── */
    .header-subtitle {
        color: #8a8070;
        font-size: 14px;
        font-weight: 400;
        margin-top: -8px;
        margin-bottom: 32px;
    }
</style>
""", unsafe_allow_html=True)


# ── Authentication ─────────────────────────────────────────────────


def check_login():
    """Show login page and return True if authenticated."""
    # Skip auth if no credentials configured (local dev)
    if "credentials" not in st.secrets:
        return True

    def _deep_dict(obj):
        if hasattr(obj, "items"):
            return {k: _deep_dict(v) for k, v in obj.items()}
        return obj

    credentials = {"usernames": _deep_dict(st.secrets["credentials"]["usernames"])}
    authenticator = stauth.Authenticate(
        credentials,
        "tool_inventory",
        st.secrets.get("cookie_key", "some_secret_key"),
        cookie_expiry_days=30,
    )

    authenticator.login()

    if st.session_state.get("authentication_status"):
        with st.sidebar:
            st.markdown(f'<p style="color:#8a8070; font-size:13px;">Logged in as <strong style="color:#f5f0e8">{st.session_state["name"]}</strong></p>', unsafe_allow_html=True)
            authenticator.logout("Logout", "sidebar")
        return True
    elif st.session_state.get("authentication_status") is False:
        st.error("Feil brukernavn eller passord.")
        st.stop()
    else:
        st.stop()


# ── Main app ───────────────────────────────────────────────────────


check_login()

# Header
st.markdown('<h1 style="font-size:32px; margin-bottom:0;">Tool Inventory</h1>', unsafe_allow_html=True)
st.markdown('<p class="header-subtitle">Equipment tracking and load out management</p>', unsafe_allow_html=True)

# Ensure Modem sheet exists
ensure_modem_sheet()

# Load data
out_rows, inv_rows, modem_rows = load_data()
dashboard_df = build_dashboard_df(out_rows, inv_rows)

# Tabs
tab_dash, tab_modem, tab_inventory, tab_raw = st.tabs(["Dashboard", "Modem", "Inventory", "Load Out Data"])

# ── Tab 1: Dashboard ──
with tab_dash:
    if dashboard_df.empty:
        st.info("No inventory data yet. Add tools in the Inventory tab.")
    else:
        # Metric cards
        total_tools = len(dashboard_df)
        total_ready = int(dashboard_df["Ready"].sum())
        total_loadout = int(dashboard_df["Load Out"].sum())
        total_redress = int(dashboard_df["Redress"].sum())

        st.markdown(f"""
        <div class="metrics-row">
            {render_metric_card("Tools", total_tools)}
            {render_metric_card("Ready", total_ready)}
            {render_metric_card("Load Out", total_loadout)}
            {render_metric_card("Redress", total_redress)}
        </div>
        """, unsafe_allow_html=True)

        # Dashboard table
        st.markdown(render_dashboard_table(dashboard_df), unsafe_allow_html=True)

    # Refresh button at bottom
    st.markdown("<br>", unsafe_allow_html=True)
    col_left, col_mid, col_right = st.columns([1, 1, 1])
    with col_mid:
        if st.button("Refresh Data", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

# ── Tab 2: Modem (Job Overview) ──
with tab_modem:
    st.markdown('<p class="header-subtitle">Active jobs and upcoming orders from coordinators.</p>', unsafe_allow_html=True)

    # Active jobs
    st.markdown('<h3 style="color:#f5f0e8; font-size:20px; margin-top:16px;">Active Jobs</h3>', unsafe_allow_html=True)
    active_df = build_active_jobs_df(out_rows)
    if active_df.empty:
        st.info("No active load outs.")
    else:
        st.dataframe(active_df, use_container_width=True, hide_index=True, column_config={
            "Customer": st.column_config.TextColumn("Customer", width="medium"),
            "Well Name": st.column_config.TextColumn("Well Name", width="medium"),
            "Date": st.column_config.TextColumn("Date", width="small"),
            "Tools": st.column_config.NumberColumn("Tool Types", width="small"),
            "Items": st.column_config.NumberColumn("Total Items", width="small"),
        })

    # Upcoming jobs
    st.markdown('<h3 style="color:#f5f0e8; font-size:20px; margin-top:32px;">Upcoming Jobs</h3>', unsafe_allow_html=True)
    modem_df = build_modem_df(modem_rows)

    edited_modem = st.data_editor(
        modem_df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Customer": st.column_config.TextColumn("Customer", width="medium"),
            "Well Name": st.column_config.TextColumn("Well Name", width="medium"),
            "Date": st.column_config.TextColumn("Date", width="small"),
            "Coordinator": st.column_config.TextColumn("Coordinator", width="small"),
            "Tools Needed": st.column_config.TextColumn("Tools Needed", width="large"),
            "Status": st.column_config.SelectboxColumn("Status", options=["Planned", "Confirmed", "Ready", "Cancelled"], width="small"),
        },
        height=min(500, 40 + 35 * max(len(modem_df), 3)),
    )

    st.markdown("<br>", unsafe_allow_html=True)
    col_l, col_m, col_r = st.columns([1, 1, 1])
    with col_m:
        if st.button("Save Upcoming Jobs", type="primary", use_container_width=True):
            save_modem(edited_modem)
            st.cache_data.clear()
            st.success("Upcoming jobs saved.")
            st.rerun()

# ── Tab 3: Inventory Editor ──
with tab_inventory:
    st.markdown('<p class="header-subtitle">Edit Total Stock and Redress, then save changes back to Google Sheet.</p>', unsafe_allow_html=True)

    if len(inv_rows) > 1:
        inv_records = []
        for row in inv_rows[1:]:
            if not row:
                continue
            inv_records.append({
                "Tool": row[0],
                "Total Stock": int(row[1]) if len(row) > 1 and row[1] else 0,
                "Redress": int(row[2]) if len(row) > 2 and row[2] else 0,
            })

        inv_df = pd.DataFrame(inv_records)
        edited_df = st.data_editor(
            inv_df,
            use_container_width=True,
            hide_index=True,
            disabled=["Tool"],
            column_config={
                "Tool": st.column_config.TextColumn("Tool", width="large"),
                "Total Stock": st.column_config.NumberColumn("Total Stock", min_value=0, step=1),
                "Redress": st.column_config.NumberColumn("Redress", min_value=0, step=1),
            },
            height=min(800, 40 + 35 * len(inv_df)),
        )

        st.markdown("<br>", unsafe_allow_html=True)
        col_left, col_mid, col_right = st.columns([1, 1, 1])
        with col_mid:
            if st.button("Save to Google Sheet", type="primary", use_container_width=True):
                save_inventory(edited_df)
                st.cache_data.clear()
                st.success("Inventory saved.")
                st.rerun()
    else:
        st.info("No inventory data. Run setup_dashboard.py first.")

# ── Tab 3: Raw Load Out Data ──
with tab_raw:
    loadout_df = build_loadout_df(out_rows)
    if loadout_df.empty:
        st.info("No load out data yet.")
    else:
        st.markdown(f'<p class="header-subtitle">{len(loadout_df)} records from {loadout_df["Filename"].nunique()} PDFs</p>', unsafe_allow_html=True)
        st.dataframe(loadout_df, use_container_width=True, hide_index=True)
