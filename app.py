# /// script
# dependencies = [
#     "streamlit>=1.30.0",
#     "playwright>=1.40.0",
#     "pandas>=2.0.0",
#     "openpyxl>=3.1.0",
#     "python-dateutil>=2.8.2",
# ]
# ///

import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
import pandas as pd
import streamlit as st

try:
    if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

# Set Streamlit Page Configuration
st.set_page_config(
    page_title="Logical Data",
    page_icon=":material/table_chart:",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Pure Google Sheets Office Aesthetic
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap');
    
    html, body, [class*="css"], .stApp {
        font-family: 'Roboto', -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif !important;
        background-color: #FFFFFF !important;
        color: #202124 !important;
    }
    
    /* Enforce Dark Readable Text for all elements */
    p, span, label, h1, h2, h3, h4, h5, h6, div, li, small {
        color: #202124 !important;
    }
    
    /* Google Sheets Top Header */
    .sheets-header {
        background-color: #FFFFFF;
        border: 1px solid #DADCE0;
        padding: 10px 18px;
        margin-bottom: 1.25rem;
        display: flex;
        align-items: center;
        gap: 14px;
        border-radius: 4px;
        box-shadow: 0 1px 2px rgba(60,64,67,0.06);
    }
    .sheets-icon {
        width: 36px;
        height: 36px;
        background-color: #0F9D58;
        border-radius: 3px;
        display: flex;
        align-items: center;
        justify-content: center;
        color: #FFFFFF !important;
        font-weight: 700;
        font-size: 18px;
    }
    .sheets-title-box h1 {
        color: #202124 !important;
        font-size: 1.35rem !important;
        font-weight: 600 !important;
        margin: 0 !important;
        line-height: 1.2 !important;
    }
    .sheets-title-box p {
        color: #5F6368 !important;
        font-size: 0.8rem !important;
        margin: 2px 0 0 0 !important;
    }
    
    /* Spreadsheet KPI Summary Cells */
    .kpi-card {
        background-color: #FFFFFF;
        border: 1px solid #DADCE0;
        border-radius: 4px;
        padding: 12px 10px;
        text-align: center;
        box-shadow: 0 1px 2px rgba(60, 64, 67, 0.04);
    }
    .kpi-title {
        color: #5F6368 !important;
        font-size: 0.75rem !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 2px;
    }
    .kpi-value {
        color: #137333 !important;
        font-size: 1.55rem !important;
        font-weight: 700 !important;
    }
    .kpi-subtext {
        color: #70757A !important;
        font-size: 0.72rem !important;
        margin-top: 1px;
    }
    
    /* Clean Office Buttons */
    .stDownloadButton > button {
        background-color: #137333 !important;
        color: #FFFFFF !important;
        font-weight: 600 !important;
        border: 1px solid #0F9D58 !important;
        border-radius: 4px !important;
        padding: 0.55rem 1.4rem !important;
        box-shadow: 0 1px 2px rgba(60,64,67,0.1) !important;
        width: 100%;
        transition: background-color 0.2s ease;
    }
    .stDownloadButton > button:hover {
        background-color: #0d5a27 !important;
        color: #FFFFFF !important;
    }
    
    button[kind="primary"] {
        background-color: #1A73E8 !important;
        color: #FFFFFF !important;
        font-weight: 600 !important;
        border: none !important;
        border-radius: 4px !important;
        box-shadow: 0 1px 2px rgba(60,64,67,0.15) !important;
    }
    button[kind="primary"]:hover {
        background-color: #1557B0 !important;
    }

    /* Google Sheets Sidebar - Light & Crisp */
    section[data-testid="stSidebar"] {
        background-color: #F8F9FA !important;
        border-right: 1px solid #DADCE0 !important;
    }
    section[data-testid="stSidebar"] * {
        color: #202124 !important;
    }
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        color: #202124 !important;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
    }
    
    /* Dropdown and Select inputs */
    div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        border: 1px solid #DADCE0 !important;
        border-radius: 4px !important;
    }
    div[data-baseweb="select"] * {
        color: #202124 !important;
    }
    
    /* Radio options */
    div[role="radiogroup"] label {
        color: #202124 !important;
    }
    div[role="radiogroup"] label span {
        color: #202124 !important;
    }
    
    /* Table styling like Google Sheets */
    div[data-testid="stDataFrame"] {
        border: 1px solid #DADCE0;
        border-radius: 4px;
        background-color: #FFFFFF;
    }
    
    /* Status Box */
    .status-box {
        background: #F8F9FA;
        border-left: 3px solid #1A73E8;
        color: #202124;
        padding: 0.8rem;
        border-radius: 4px;
        font-family: monospace;
        max-height: 200px;
        overflow-y: auto;
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)

# Ensure Playwright chromium binary is installed
import subprocess
@st.cache_resource
def setup_playwright():
    # If system chromium exists (e.g. from packages.txt on Debian / Streamlit Cloud), skip download
    for p in ["/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome"]:
        if os.path.exists(p):
            return True
    try:
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False, timeout=60)
    except Exception:
        pass
    return True

setup_playwright()

# Title Banner - Google Sheets Style
st.markdown("""
<div class="sheets-header">
    <div class="sheets-icon">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect><line x1="3" y1="9" x2="21" y2="9"></line><line x1="3" y1="15" x2="21" y2="15"></line><line x1="9" y1="3" x2="9" y2="21"></line><line x1="15" y1="3" x2="15" y2="21"></line></svg>
    </div>
    <div class="sheets-title-box">
        <h1>Logical Data</h1>
        <p>Operational Performance Audit & Resolution Tracking</p>
    </div>
</div>
""", unsafe_allow_html=True)

# Sidebar Inputs
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Sidebar Inputs
with st.sidebar:
    st.subheader("Audit Configuration")
    
    # Employee Selection
    emp_options = [
        "ALL TEAM (Ajit, Shashikant, Om, Sabin, Sunil, Sanjeev, Chandramani, Rajesh)",
        "Ajit Shrestha",
        "Shashikant Chaudhary",
        "Om Neupane",
        "Sabin Giri",
        "Sunil Chaudhary",
        "Sanjeev Giri",
        "Chandramani Tharu",
        "Rajesh Maharjan",
        "Custom Employee Name / List..."
    ]
    selected_emp_type = st.selectbox("Employee / Team Selection", emp_options, index=0)
    
    if selected_emp_type == "Custom Employee Name / List...":
        employee_name = st.text_input("Enter Employee Name(s) (comma separated for multiple)", value="Om Neupane")
    elif selected_emp_type.startswith("ALL TEAM"):
        employee_name = "Ajit Shrestha, Shashikant Chaudhary, Om Neupane, Sabin Giri, Sunil Chaudhary, Sanjeev Giri, Chandramani Tharu, Rajesh Maharjan"
    else:
        employee_name = selected_emp_type
        
    st.markdown("---")
    st.subheader("Date Range Selector")
    
    today = datetime.now().date()
    yesterday = today - timedelta(days=1)
    
    date_preset = st.radio(
        "Quick Range", 
        ["Today", "Yesterday & Today (2 Days)", "Last 7 Days", "Last 30 Days", "Full Current Month", "Previous Month", "Custom Range"], 
        index=0
    )
    
    if date_preset == "Today":
        from_date_obj = today
        to_date_obj = today
    elif date_preset == "Yesterday & Today (2 Days)":
        from_date_obj = yesterday
        to_date_obj = today
    elif date_preset == "Last 7 Days":
        from_date_obj = today - timedelta(days=6)
        to_date_obj = today
    elif date_preset == "Last 30 Days":
        from_date_obj = today - timedelta(days=29)
        to_date_obj = today
    elif date_preset == "Full Current Month":
        from_date_obj = today.replace(day=1)
        to_date_obj = today
    elif date_preset == "Previous Month":
        first_day_curr = today.replace(day=1)
        last_day_prev = first_day_curr - timedelta(days=1)
        from_date_obj = last_day_prev.replace(day=1)
        to_date_obj = last_day_prev
    else:
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            from_date_obj = st.date_input("From Date", value=today)
        with col_d2:
            to_date_obj = st.date_input("To Date", value=today)

    from_date_str = from_date_obj.strftime("%d %b %Y")  # e.g., "01 Aug 2026"
    to_date_str = to_date_obj.strftime("%d %b %Y")      # e.g., "31 Aug 2026"

    if from_date_str == to_date_str:
        st.info(f"Target Date: **{from_date_str}** (Today's updates only)")
    else:
        st.info(f"Target Period: **{from_date_str}** to **{to_date_str}**")
    st.markdown("---")
    
    run_btn = st.button("Run Audit Scraper", type="primary", use_container_width=True)

# Helper function to generate Master Executive Excel Report from multiple DataFrames
def build_executive_team_excel(master_df, team_summary_df, cat_summary_df, matrix_summary_df=None):
    output_buf = io.BytesIO()
    with pd.ExcelWriter(output_buf, engine='openpyxl') as writer:
        team_summary_df.to_excel(writer, sheet_name="Team Executive Summary", index=False)
        cat_summary_df.to_excel(writer, sheet_name="Category Summary", index=False)
        if matrix_summary_df is not None:
            matrix_summary_df.to_excel(writer, sheet_name="Technician Matrix Pivot")
        master_df.to_excel(writer, sheet_name="Master Audit Details", index=False)
        
    output_buf.seek(0)
    wb = openpyxl.load_workbook(output_buf)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    cell_font = Font(name="Segoe UI", size=10)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    
    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                cell.font = cell_font
                cell.border = thin_border
                if (sheetname == "Team Executive Summary" or sheetname == "Technician Matrix Pivot") and row == ws.max_row:
                    cell.font = Font(name="Segoe UI", size=11, bold=True, color="1F4E78")
                    cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                elif sheetname == "Technician Matrix Pivot" and col == ws.max_column:
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="1F4E78")
                    cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                    
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)
            
    final_buf = io.BytesIO()
    wb.save(final_buf)
    final_buf.seek(0)
    return final_buf.getvalue()

def classify_work_type(row):
    remark = str(row.get('Grid Remark', '')).strip()
    solution = str(row.get('Solution Given', '')).strip()
    employee_note = str(row.get('Employee Remark / Solution Note', '')).strip()
    title = str(row.get('Title', '')).strip()
    category = str(row.get('Category', '')).strip()
    
    full_text = f"{title} {remark} {solution} {employee_note}".upper()
    
    # Robust WiFi 6 detection (Title, Category, Sub Category, Remark, Solution Note, Serial Number)
    wifi6_regex = r'wifi\s*6|wifi-6|wifi6|wi-fi\s*6|router\s*6|alcl|nokia.*wifi|upgrade.*wifi|wifi.*upgrade|dual\s*band'
    is_wifi6_upgrade = (
        bool(re.search(wifi6_regex, full_text, re.I)) or
        'WIFI 6' in full_text or 'WIFI-6' in full_text or 'WIFI6' in full_text or 'ALCL' in full_text
    )
    
    if is_wifi6_upgrade:
        return 'WiFi 6 Upgrade'
    elif 'IPTV' in full_text or 'IPTV' in category.upper():
        return 'IPTV Issue'
    else:
        return 'General / Other Issues'

# Mode Selection Tabs in Main Area
main_mode_tab1, main_mode_tab_repeats, main_mode_tab_scorecard, main_mode_tab2 = st.tabs([
    "Performance Analytics & Scraper",
    "Repeat Issue / Customer Risk (Idea 3)",
    "Technician Leaderboard & Scorecard (Idea 5)",
    "Combine Uploaded Reports (Manager Tool)"
])

with main_mode_tab1:
    if run_btn:
        st.info(f"Starting automated scraper for **{employee_name}** from `{from_date_str}` to `{to_date_str}`...")
        
        log_container = st.empty()
        logs_list = []
        
        def ui_log(msg, level="INFO"):
            timestamp = datetime.now().strftime("%H:%M:%S")
            log_line = f"[{timestamp}] [{level}] {msg}"
            logs_list.append(log_line)
            log_html = "<br>".join(logs_list[-8:])
            log_container.markdown(f'<div class="status-box">{log_html}</div>', unsafe_allow_html=True)

        progress_bar = st.progress(0.1, text="Starting Scraper process...")

        try:
            script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "audit_automation.py")
            cmd = [
                sys.executable,
                script_path,
                "--employee", employee_name,
                "--from-date", from_date_str,
                "--to-date", to_date_str,
                "--non-interactive"
            ]
            
            env = os.environ.copy()
            env["PYTHONUNBUFFERED"] = "1"
            env["HEADLESS"] = "true"

            ui_log(f"Spawning scraper process for {employee_name} ({from_date_str} to {to_date_str})...")
            
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                env=env
            )
            
            while True:
                line = process.stdout.readline()
                if not line and process.poll() is not None:
                    break
                if line:
                    clean_line = line.strip()
                    if clean_line:
                        ui_log(clean_line)
                        if "Navigating to" in clean_line:
                            progress_bar.progress(0.25, text="Connecting to CGNET portal...")
                        elif "Logged in successfully" in clean_line:
                            progress_bar.progress(0.45, text="Logged in, loading Audit page...")
                        elif "Scraping page" in clean_line or "Found" in clean_line:
                            progress_bar.progress(0.70, text="Scraping records...")
                        elif "compiled report saved" in clean_line:
                            progress_bar.progress(0.90, text="Generating Excel report...")

            return_code = process.poll()
            
            if return_code != 0:
                st.error(f"Scraper process finished with exit code {return_code}. Review the logs above.")
            else:
                progress_bar.progress(1.0, text="Scraping completed!")
                st.success("Audit Scraper completed successfully.")
                
                if ',' in employee_name or 'ALL TEAM' in employee_name.upper():
                    safe_emp = "ALL_TEAM"
                else:
                    safe_emp = re.sub(r'[^a-zA-Z0-9]', '_', employee_name)
                safe_from = re.sub(r'[^a-zA-Z0-9]', '_', from_date_str)
                output_file = f"audit_report_{safe_emp}_{safe_from}.xlsx"
                
                if os.path.exists(output_file):
                    st.session_state["last_output_file"] = output_file
                    st.session_state["last_run_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                else:
                    import glob
                    matching_files = glob.glob(f"audit_report_{safe_emp}_*.xlsx")
                    if matching_files:
                        st.session_state["last_output_file"] = matching_files[0]
                        st.session_state["last_run_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
        except Exception as e:
            st.error(f"Scraper error encountered: {e}")
            st.exception(e)

    # Automatically discover master historical database and available daily reports
    @st.cache_data(ttl=60, show_spinner=False)
    def get_cached_master_csv(file_path):
        if os.path.exists(file_path):
            try:
                return pd.read_csv(file_path, dtype=str)
            except Exception:
                return None
        return None

    @st.cache_data(ttl=60, show_spinner=False)
    def get_cached_excel_report(file_path):
        if file_path and os.path.exists(file_path):
            try:
                xls = pd.ExcelFile(file_path)
                df_det = pd.read_excel(xls, sheet_name="Audit Details") if "Audit Details" in xls.sheet_names else None
                df_sum = pd.read_excel(xls, sheet_name="Summary Report", nrows=5) if "Summary Report" in xls.sheet_names else None
                return df_det, df_sum
            except Exception:
                return None, None
        return None, None

    import glob
    master_csv_file = "master_audit_history.csv"
    df_master_all = get_cached_master_csv(master_csv_file)

    existing_reports = glob.glob("audit_report_*.xlsx")
    nepal_tz = timezone(timedelta(hours=5, minutes=45))
    today_npt_dt = datetime.now(nepal_tz)
    today_npt_str = today_npt_dt.strftime("%d_%b_%Y")
    today_date_str = today_npt_dt.strftime("%d %b %Y")
    yest_npt_dt = today_npt_dt - timedelta(days=1)
    yest_npt_str = yest_npt_dt.strftime("%d_%b_%Y")
    yest_date_str = yest_npt_dt.strftime("%d %b %Y")

    # Discover latest report file on disk
    latest_on_disk = None
    if existing_reports:
        today_reports = [f for f in existing_reports if today_npt_str in f]
        if today_reports:
            today_reports.sort(key=os.path.getmtime, reverse=True)
            latest_on_disk = today_reports[0]
        else:
            existing_reports.sort(key=os.path.getmtime, reverse=True)
            latest_on_disk = existing_reports[0]
            
    output_file = st.session_state.get("last_output_file", latest_on_disk)
    if output_file and os.path.exists(output_file):
        st.session_state["last_output_file"] = output_file
        mod_time = datetime.fromtimestamp(os.path.getmtime(output_file)).strftime("%Y-%m-%d %H:%M:%S")
        st.session_state["last_run_time"] = mod_time

    # Build View Period Options for Google Sheets dropdown
    period_options = [f"Today ({today_date_str})"]
    past_dates = []
    if df_master_all is not None and "Report Date" in df_master_all.columns:
        for d in df_master_all["Report Date"].dropna().unique():
            d_clean = str(d).strip()
            if d_clean and d_clean != today_date_str and d_clean not in past_dates:
                past_dates.append(d_clean)
    try:
        past_dates = sorted(past_dates, key=lambda d: datetime.strptime(d, "%d %b %Y"), reverse=True)
    except Exception:
        past_dates = sorted(past_dates, reverse=True)

    for pd_date in past_dates:
        if pd_date == yest_date_str:
            period_options.append(f"Yesterday ({yest_date_str})")
        else:
            period_options.append(pd_date)

    current_month_str = today_npt_dt.strftime("%B %Y")
    period_options.append("Weekly (Last 7 Days Combined)")
    period_options.append(f"Month to Date ({current_month_str})")
    period_options.append("All-Time Master History")

    has_data = (output_file and os.path.exists(output_file)) or (df_master_all is not None and not df_master_all.empty)

    if has_data:
        dash_header_col, dash_period_col, dash_refresh_col = st.columns([3, 2, 1])
        with dash_header_col:
            st.markdown("### Performance Analytics Dashboard")
            st.caption(f"**Live View** | Last updated: {st.session_state.get('last_run_time', 'Recently')}")
        with dash_period_col:
            selected_period = st.selectbox(
                "Select Report View Period",
                options=period_options,
                index=0,
                help="Switch between Today, Yesterday, Weekly, or Month-to-Date data without re-scraping!"
            )
        with dash_refresh_col:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            if st.button("Refresh", use_container_width=True):
                st.cache_data.clear()
                st.session_state.clear()
                st.rerun()
        
        try:
            df_details = None
            df_summary = None
            if selected_period.startswith("Today"):
                today_file = [f for f in existing_reports if today_npt_str in f]
                if today_file:
                    df_details, df_summary = get_cached_excel_report(today_file[0])
                elif df_master_all is not None and "Report Date" in df_master_all.columns:
                    df_details = df_master_all[df_master_all["Report Date"].str.strip() == today_date_str].copy()
            elif selected_period.startswith("Yesterday"):
                yest_file = [f for f in existing_reports if yest_npt_str in f]
                if yest_file:
                    df_details, df_summary = get_cached_excel_report(yest_file[0])
                elif df_master_all is not None and "Report Date" in df_master_all.columns:
                    df_details = df_master_all[df_master_all["Report Date"].str.strip() == yest_date_str].copy()
            elif "Weekly" in selected_period:
                if df_master_all is not None and "Report Date" in df_master_all.columns:
                    cutoff = today_npt_dt - timedelta(days=7)
                    def in_week(d):
                        try:
                            return datetime.strptime(str(d).strip(), "%d %b %Y").date() >= cutoff.date()
                        except Exception:
                            return False
                    df_details = df_master_all[df_master_all["Report Date"].apply(in_week)].copy()
            elif "Month to Date" in selected_period:
                curr_my = today_npt_dt.strftime("%b %Y")
                if df_master_all is not None and "Report Date" in df_master_all.columns:
                    df_details = df_master_all[df_master_all["Report Date"].str.contains(curr_my, case=False, na=False)].copy()
            elif "All-Time" in selected_period:
                if df_master_all is not None:
                    df_details = df_master_all.copy()
            else:
                clean_target = selected_period.strip()
                past_file = [f for f in existing_reports if clean_target.replace(" ", "_") in f]
                if past_file:
                    df_details, df_summary = get_cached_excel_report(past_file[0])
                elif df_master_all is not None and "Report Date" in df_master_all.columns:
                    df_details = df_master_all[df_master_all["Report Date"].str.strip() == clean_target].copy()

            if df_details is None or df_details.empty:
                if output_file and os.path.exists(output_file):
                    df_details, df_summary = get_cached_excel_report(output_file)
                elif df_master_all is not None and not df_master_all.empty:
                    df_details = df_master_all.copy()
                else:
                    df_details = pd.DataFrame()

            df_details['Task / Issue Type'] = df_details.apply(classify_work_type, axis=1)
            
            def is_solved_row(row):
                st_val = str(row.get("Status", "")).strip().lower()
                rm_val = str(row.get("Grid Remark", "")).strip().lower()
                if st_val in ["completed", "closed"]:
                    return True
                if "ms" in rm_val or "assign" in rm_val or "transfer" in rm_val or "forward" in rm_val:
                    return True
                return False

            df_details["Is_Solved_Val"] = df_details.apply(is_solved_row, axis=1)

            col1, col2, col3, col4, col5 = st.columns(5)
            
            tot_kpi_records = len(df_details)
            tot_kpi_solved = int(df_details["Is_Solved_Val"].sum()) if not df_details.empty else 0
            kpi_rate_val = f"{(tot_kpi_solved / tot_kpi_records * 100):.1f}%" if tot_kpi_records > 0 else "0.0%"

            def get_val(metric_name):
                if df_summary is not None and not df_summary.empty:
                    row = df_summary[df_summary["Metric"].str.contains(metric_name, case=False, na=False)]
                    if not row.empty:
                        return str(row.iloc[0]["Value"])
                return "N/A"

            with col1:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Total Records</div>
                    <div class="kpi-value">{tot_kpi_records}</div>
                    <div class="kpi-subtext">Scraped logs</div>
                </div>
                """, unsafe_allow_html=True)
                
            with col2:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Solved / Handled</div>
                    <div class="kpi-value">{tot_kpi_solved}</div>
                    <div class="kpi-subtext">Completed or MS Assigned</div>
                </div>
                """, unsafe_allow_html=True)
                
            with col3:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Solution Rate</div>
                    <div class="kpi-value">{kpi_rate_val}</div>
                    <div class="kpi-subtext">Completion %</div>
                </div>
                """, unsafe_allow_html=True)
                
            with col4:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Avg Time (Assigned)</div>
                    <div class="kpi-value" style="font-size:1.4rem;">{get_val("Average Completion Time (From Assigned Date)")}</div>
                    <div class="kpi-subtext">From assigned timestamp</div>
                </div>
                """, unsafe_allow_html=True)
                
            with col5:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Avg Time (Created)</div>
                    <div class="kpi-value" style="font-size:1.4rem;">{get_val("Average Completion Time (From Created Date)")}</div>
                    <div class="kpi-subtext">From initial creation</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(f"### Download Executive Report ({selected_period})")

            emp_name_col = "Grid Employee Name" if "Grid Employee Name" in df_details.columns else "Employee Name"
            has_multiple_emps = (emp_name_col in df_details.columns and df_details[emp_name_col].dropna().nunique() > 1) or ("ALL_TEAM" in str(output_file)) or ("," in str(employee_name)) or (tot_kpi_records > 0)

            if has_multiple_emps:
                period_slug = re.sub(r'[^a-zA-Z0-9]', '_', selected_period.split(' ')[0])
                exec_file_name = f"EXECUTIVE_TEAM_AUDIT_REPORT_{period_slug}_{today_npt_str}.xlsx"
                
                # Always build dynamically from the current active df_details to guarantee 100% match with the screen
                if "Employee Name" not in df_details.columns and emp_name_col in df_details.columns:
                    df_details["Employee Name"] = df_details[emp_name_col]
                elif "Employee Name" not in df_details.columns:
                    df_details["Employee Name"] = df_details.get("Target Employee", "Team Member")
                
                DEFAULT_TEAM_MEMBERS = [
                    "Ajit Shrestha", "Chandramani Tharu", "Om Neupane", "Rajesh Maharjan",
                    "Sabin Giri", "Sanjeev Giri", "Shashikant Chaudhary", "Sunil Chaudhary"
                ]
                
                summary_rows = []
                all_display_emps = sorted(list(set([e for e in df_details["Employee Name"].dropna().unique().tolist() if e])))
                for emp in all_display_emps:
                    grp = df_details[df_details["Employee Name"] == emp]
                    tot = len(grp)
                    if tot == 0:
                        continue  # Exclude technicians on day off (0 tickets)
                    solved = grp["Is_Solved_Val"].sum()
                    rate = f"{(solved / tot * 100):.1f}%" if tot > 0 else "0.0%"
                    summary_rows.append({
                        "Employee Name": emp,
                        "Total Scraped Tickets": tot,
                        "Solved / Handled Count": solved,
                        "Solution Rate %": rate
                    })
                team_summary_df = pd.DataFrame(summary_rows)
                tot_tickets_team = len(df_details)
                tot_solved_team = df_details["Is_Solved_Val"].sum()
                team_rate_val = f"{(tot_solved_team / tot_tickets_team * 100):.1f}%" if tot_tickets_team > 0 else "0.0%"
                total_team_row = pd.DataFrame([{
                    "Employee Name": "GRAND TOTAL (ALL TEAM)",
                    "Total Scraped Tickets": tot_tickets_team,
                    "Solved / Handled Count": tot_solved_team,
                    "Solution Rate %": team_rate_val
                }])
                team_summary_df = pd.concat([team_summary_df, total_team_row], ignore_index=True)
                
                work_summary_df = df_details.groupby("Task / Issue Type", dropna=False).agg(
                    Total_Tickets=("Ticket Number", "count"),
                    Solved_Count=("Is_Solved_Val", "sum")
                ).reset_index()
                work_summary_df["Solution Rate %"] = (work_summary_df["Solved_Count"] / work_summary_df["Total_Tickets"] * 100).round(1).astype(str) + '%'
                work_summary_df["% Share of Total"] = (work_summary_df["Total_Tickets"] / len(df_details) * 100).round(1).astype(str) + '%'
                
                matrix_cat_col = 'Sub Category' if 'Sub Category' in df_details.columns else 'Task / Issue Type'
                df_details[matrix_cat_col] = df_details[matrix_cat_col].fillna('Other / Uncategorized')
                df_excel_dedup = df_details.drop_duplicates(subset=['Ticket Number'], keep='first') if 'Ticket Number' in df_details.columns else df_details
                matrix_pivot_excel = pd.crosstab(
                    df_excel_dedup[matrix_cat_col],
                    df_excel_dedup['Employee Name'],
                    margins=True,
                    margins_name="Grand Total"
                )
                cols_ex = sorted([c for c in matrix_pivot_excel.columns if c != "Grand Total"]) + (["Grand Total"] if "Grand Total" in matrix_pivot_excel.columns else [])
                matrix_pivot_excel = matrix_pivot_excel[cols_ex]
                if "Grand Total" in matrix_pivot_excel.index:
                    d_rows = matrix_pivot_excel.drop(index="Grand Total").sort_values(by="Grand Total", ascending=False)
                    t_row = matrix_pivot_excel.loc[["Grand Total"]]
                    matrix_pivot_excel = pd.concat([d_rows, t_row])
                
                export_master = df_details.drop(columns=["Is_Solved_Val"], errors="ignore")
                exec_bytes = build_executive_team_excel(export_master, team_summary_df, work_summary_df, matrix_pivot_excel)

                st.download_button(
                    label=f"Download Combined Executive Team Report ({exec_file_name})",
                    data=exec_bytes,
                    file_name=exec_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.caption("Combined Executive Team Report including Technician Pivot Matrix, Category Summary & Master Audit Details")
            else:
                with open(output_file, "rb") as f:
                    bytes_data = f.read()
                st.download_button(
                    label=f"Download Excel Audit Report ({os.path.basename(output_file)})",
                    data=bytes_data,
                    file_name=os.path.basename(output_file),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.markdown("<br>", unsafe_allow_html=True)
            
            tab0, tab_team, tab_cat, tab1, tab3 = st.tabs([
                "Technician Category Matrix (Pivot)",
                "Team Executive Summary (Technicians)",
                "Category Summary",
                "Scraped Audit Details",
                "Raw Summary Sheet"
            ])
            
            with tab0:
                st.subheader("Technician Category Breakdown Matrix (Pivot Table)")
                st.caption("Exact category and sub-category ticket count per technician matching Excel Pivot Table layout.")
                
                col_p1, col_p2, col_p3 = st.columns([3, 2, 2])
                with col_p1:
                    status_opts = sorted(df_details["Status"].dropna().unique().tolist()) if "Status" in df_details.columns else []
                    chosen_statuses = st.multiselect(
                        "Status Filter (Multiple Items)",
                        options=status_opts,
                        default=status_opts,
                        help="Filter tickets by status (Completed, Closed, In Progress, etc.). Defaults to all to show full picture."
                    )
                with col_p2:
                    breakdown_col_choice = st.selectbox(
                        "Row Category Level",
                        options=["Category & Sub Category", "Sub Category", "Category", "Task / Issue Type"],
                        index=0
                    )
                with col_p3:
                    dedup_choice = st.checkbox(
                        "Remove Duplicate Tickets (Unique Only)",
                        value=True,
                        help="Prevents double-counting tickets that have multiple updates on the same day"
                    )
                
                # Filter dataset for matrix
                if chosen_statuses and "Status" in df_details.columns:
                    matrix_filtered = df_details[df_details["Status"].isin(chosen_statuses)].copy()
                else:
                    matrix_filtered = df_details.copy()

                if dedup_choice and "Ticket Number" in matrix_filtered.columns:
                    matrix_filtered = matrix_filtered.drop_duplicates(subset=["Ticket Number"], keep="first")
                    
                if breakdown_col_choice == "Category & Sub Category":
                    cat_col = 'Category' if 'Category' in matrix_filtered.columns else 'Task / Issue Type'
                    sub_col = 'Sub Category' if 'Sub Category' in matrix_filtered.columns else None
                    if sub_col:
                        matrix_filtered["Cat_SubCat"] = matrix_filtered.apply(
                            lambda r: f"{str(r.get(cat_col, 'Other')).strip()} ➔ {str(r.get(sub_col, '')).strip()}" if pd.notna(r.get(sub_col)) and str(r.get(sub_col)).strip() != '' else str(r.get(cat_col, 'Other')).strip(),
                            axis=1
                        )
                    else:
                        matrix_filtered["Cat_SubCat"] = matrix_filtered[cat_col].fillna("Other")
                    actual_breakdown_col = "Cat_SubCat"
                else:
                    actual_breakdown_col = breakdown_col_choice if breakdown_col_choice in matrix_filtered.columns else "Task / Issue Type"
                
                actual_emp_col = "Employee Name" if "Employee Name" in matrix_filtered.columns else ("Target Employee" if "Target Employee" in matrix_filtered.columns else "Grid Employee Name")
                matrix_filtered[actual_breakdown_col] = matrix_filtered[actual_breakdown_col].fillna("Other / Uncategorized")
                
                if not matrix_filtered.empty and actual_emp_col in matrix_filtered.columns:
                    pivot_table = pd.crosstab(
                        matrix_filtered[actual_breakdown_col],
                        matrix_filtered[actual_emp_col],
                        margins=True,
                        margins_name="Grand Total"
                    )
                    
                    # Sort active employee columns alphabetically with Grand Total strictly at the end
                    emp_cols = sorted([c for c in pivot_table.columns if c != "Grand Total"])
                    pivot_table = pivot_table[emp_cols + (["Grand Total"] if "Grand Total" in pivot_table.columns else [])]
                    
                    # Sort rows by Grand Total descending (keep Grand Total strictly at bottom)
                    if "Grand Total" in pivot_table.index:
                        data_rows = pivot_table.drop(index="Grand Total").sort_values(by="Grand Total", ascending=False)
                        total_row = pivot_table.loc[["Grand Total"]]
                        sorted_pivot = pd.concat([data_rows, total_row])
                    else:
                        sorted_pivot = pivot_table
                        
                    st.markdown(f"**Showing `{len(matrix_filtered)}` {'unique tickets' if dedup_choice else 'scraped log events'} across `{len(sorted_pivot.index) - 1}` categories:**")
                    st.dataframe(sorted_pivot, use_container_width=True)
                    
                    pivot_csv = sorted_pivot.to_csv().encode('utf-8')
                    st.download_button(
                        label="Download Pivot Matrix Table (CSV)",
                        data=pivot_csv,
                        file_name=f"technician_pivot_matrix_{datetime.now().strftime('%d_%b_%Y')}.csv",
                        mime="text/csv"
                    )
                else:
                    st.warning("No records found for the selected filter.")
            
            with tab_team:
                st.subheader("Team Executive Summary (Technician Performance)")
                st.caption("Overall ticket volume, solved counts, and solution percentage per technician (matching Excel Sheet 1).")
                st.dataframe(team_summary_df, use_container_width=True)
                
                team_csv = team_summary_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download Team Executive Summary (CSV)",
                    data=team_csv,
                    file_name=f"team_executive_summary_{datetime.now().strftime('%d_%b_%Y')}.csv",
                    mime="text/csv"
                )
                
            with tab_cat:
                st.subheader("Category & Issue Type Summary")
                st.caption("Distribution of tickets and solution rate by category (matching Excel Sheet 2).")
                st.dataframe(work_summary_df, use_container_width=True)
                
                cat_csv = work_summary_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download Category Summary (CSV)",
                    data=cat_csv,
                    file_name=f"category_summary_{datetime.now().strftime('%d_%b_%Y')}.csv",
                    mime="text/csv"
                )
                
                st.markdown("---")
                st.subheader("Category Volume Chart")
                st.bar_chart(work_summary_df.set_index('Task / Issue Type')['Total_Tickets'])
                
                st.markdown("---")
                st.subheader("Portal Category & Sub-Category Detailed Breakdown")
                if 'Category' in df_details.columns:
                    sub_col = 'Sub Category' if 'Sub Category' in df_details.columns else ('Sub Sub Category' if 'Sub Sub Category' in df_details.columns else None)
                    group_cols = ['Category'] + ([sub_col] if sub_col else [])
                    cat_counts = df_details.groupby(group_cols, dropna=False).size().reset_index(name='Total Tickets Count')
                    cat_counts = cat_counts.sort_values(by='Total Tickets Count', ascending=False)
                    st.dataframe(cat_counts, use_container_width=True)
            
            with tab1:
                st.subheader("Filterable Audit Details Table")
                search_query = st.text_input("Search records by ticket #, remark, or account name...", "")
                
                if search_query:
                    filtered_df = df_details[df_details.astype(str).apply(lambda row: row.str.contains(search_query, case=False).any(), axis=1)]
                else:
                    filtered_df = df_details
                    
                st.dataframe(filtered_df, use_container_width=True, height=400)
                
            with tab3:
                st.subheader("Summary Report Sheet View")
                if 'df_summary' in locals() and df_summary is not None and not df_summary.empty:
                    st.dataframe(df_summary.astype(str), use_container_width=True)
                elif 'xls' in locals() and hasattr(xls, 'sheet_names') and "Summary Report" in xls.sheet_names:
                    df_full_summary = pd.read_excel(xls, sheet_name="Summary Report")
                    st.dataframe(df_full_summary.astype(str), use_container_width=True)
                else:
                    st.dataframe(team_summary_df, use_container_width=True)

        except Exception as read_err:
            st.error(f"Could not load output preview: {read_err}")
    else:
        st.info("Select an Employee or **ALL TEAM**, set Date Range, then click **Run Audit Scraper** to generate your report.")

with main_mode_tab_repeats:
    st.markdown("""
    <div class="sheets-header">
        <div class="sheets-icon" style="background-color: #D93025;">🔁</div>
        <div class="sheets-title-box">
            <h1>Repeat Issue & Chronic Complaint Analyzer</h1>
            <p>Automatically flags accounts, phone numbers, or customers with multiple field visits and tickets.</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    rep_csv_file = "master_audit_history.csv"
    df_rep_raw = get_cached_master_csv(rep_csv_file)
    
    if df_rep_raw is None or df_rep_raw.empty:
        st.info("No master historical records found in `master_audit_history.csv` yet.")
    else:
        df_rep = df_rep_raw.copy()
        
        # Filter controls for Repeat analysis
        r_col1, r_col2, r_col3 = st.columns([2, 2, 2])
        with r_col1:
            id_mode = st.selectbox(
                "Group / Identify By",
                options=["User Id", "Mobile No", "Customer Name", "Account Name"],
                index=0,
                help="Select which identifier to use for spotting repeat customer complaints."
            )
        with r_col2:
            min_tickets = st.slider("Minimum Ticket Count (Repeat Threshold)", min_value=2, max_value=10, value=2)
        with r_col3:
            search_cust = st.text_input("Search Customer / ID / Mobile", placeholder="e.g. 9841...")
            
        # Group and calculate repeats
        valid_rows = df_rep[df_rep[id_mode].notna() & (df_rep[id_mode].astype(str).str.strip() != "") & (df_rep[id_mode].astype(str).str.strip().str.lower() != "nan")].copy()
        
        if valid_rows.empty:
            st.warning(f"No valid records found for identifier: {id_mode}")
        else:
            cust_group = valid_rows.groupby(id_mode).agg(
                Total_Tickets=('Ticket Number', 'count'),
                Unique_Tickets=('Ticket Number', 'nunique'),
                Technicians_Involved=('Target Employee', lambda s: ", ".join(sorted(set([str(x) for x in s.dropna() if str(x).strip()])))),
                Categories_Reported=('Sub Category', lambda s: ", ".join(sorted(set([str(x) for x in s.dropna() if str(x).strip()]))[:3])),
                First_Reported_Date=('Grid Date', 'min'),
                Last_Reported_Date=('Grid Date', 'max')
            ).reset_index()
            
            repeat_df = cust_group[cust_group['Total_Tickets'] >= min_tickets].sort_values(by='Total_Tickets', ascending=False)
            
            if search_cust:
                repeat_df = repeat_df[repeat_df[id_mode].astype(str).str.contains(search_cust.strip(), case=False, na=False)]
                
            # Summary Metrics for Repeats
            tot_repeat_entities = len(repeat_df)
            tot_repeat_tickets = repeat_df['Total_Tickets'].sum() if not repeat_df.empty else 0
            max_repeat_count = int(repeat_df['Total_Tickets'].max()) if not repeat_df.empty else 0
            
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Chronic Accounts</div>
                    <div class="kpi-value" style="color: #D93025;">{tot_repeat_entities}</div>
                    <div class="kpi-subtext">Accounts with &ge;{min_tickets} tickets</div>
                </div>
                """, unsafe_allow_html=True)
            with m2:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Repeat Work Volume</div>
                    <div class="kpi-value">{tot_repeat_tickets}</div>
                    <div class="kpi-subtext">Total tickets involved</div>
                </div>
                """, unsafe_allow_html=True)
            with m3:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Max Visits on 1 Client</div>
                    <div class="kpi-value">{max_repeat_count}</div>
                    <div class="kpi-subtext">Highest repetition</div>
                </div>
                """, unsafe_allow_html=True)
            with m4:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Repeat Ticket %</div>
                    <div class="kpi-value">{((tot_repeat_tickets / len(df_rep)) * 100):.1f}%</div>
                    <div class="kpi-subtext">Of all master history ({len(df_rep)})</div>
                </div>
                """, unsafe_allow_html=True)
                
            st.markdown("<br>", unsafe_allow_html=True)
            st.subheader(f"Top Chronic Accounts (Sorted by Ticket Frequency)")
            st.dataframe(repeat_df, use_container_width=True, height=350)
            
            # Drill-down view into a selected customer
            st.markdown("---")
            st.subheader("Customer Complaint History & Audit Trail Drill-Down")
            
            top_options = repeat_df[id_mode].astype(str).tolist()[:50]
            if top_options:
                selected_entity = st.selectbox(
                    f"Select {id_mode} to view full timeline and technician logs:",
                    options=top_options,
                    index=0
                )
                
                drilldown_rows = valid_rows[valid_rows[id_mode].astype(str) == selected_entity].copy()
                cols_to_show = [c for c in ['Grid Date', 'Ticket Number', 'Target Employee', 'Status', 'Category', 'Sub Category', 'Grid Remark', 'Resolution Time (From Assigned)'] if c in drilldown_rows.columns]
                st.dataframe(drilldown_rows[cols_to_show], use_container_width=True)
            else:
                st.info("No repeat accounts meet the current filter criteria.")

with main_mode_tab_scorecard:
    st.markdown("""
    <div class="sheets-header">
        <div class="sheets-icon" style="background-color: #F9AB00;">🏆</div>
        <div class="sheets-title-box">
            <h1>Technician Leaderboard & Gamification Scorecard</h1>
            <p>Objective monthly and weekly performance rankings, resolution speed medals, and upgrade leaderboards.</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    score_csv_file = "master_audit_history.csv"
    df_score_raw = get_cached_master_csv(score_csv_file)
    
    if df_score_raw is None or df_score_raw.empty:
        st.info("No master historical records available for scorecard calculation.")
    else:
        df_score = df_score_raw.copy()
        
        # Period Filter for Scorecard
        sc_col1, sc_col2 = st.columns([2, 4])
        with sc_col1:
            all_reports = sorted(list(df_score['Report Date'].dropna().unique()), reverse=True) if 'Report Date' in df_score.columns else []
            sc_period = st.selectbox(
                "Scorecard Time Horizon",
                options=["All-Time Master History", "Last 7 Days (Weekly)", "Current Month (October 2026)", "Bhadra Month (17 Aug - 16 Sep 2026)"],
                index=0
            )
            
        nepal_tz = timezone(timedelta(hours=5, minutes=45))
        now_npt = datetime.now(nepal_tz)
        
        if sc_period == "Last 7 Days (Weekly)":
            cutoff_dt = now_npt - timedelta(days=7)
            def in_last_7(d):
                try:
                    return datetime.strptime(str(d).strip(), "%d %b %Y").date() >= cutoff_dt.date()
                except Exception:
                    return False
            df_filtered_sc = df_score[df_score['Report Date'].apply(in_last_7)].copy() if 'Report Date' in df_score.columns else df_score
        elif "October" in sc_period:
            df_filtered_sc = df_score[df_score['Report Date'].str.contains('Oct 2026', case=False, na=False)].copy() if 'Report Date' in df_score.columns else df_score
        elif "Bhadra" in sc_period:
            df_filtered_sc = df_score[df_score['Report Date'].str.contains('Aug 2026|14 Sep 2026|15 Sep 2026|16 Sep 2026', case=False, na=False)].copy() if 'Report Date' in df_score.columns else df_score
        else:
            df_filtered_sc = df_score.copy()
            
        emp_key = 'Grid Employee Name' if 'Grid Employee Name' in df_filtered_sc.columns else 'Target Employee'
        df_filtered_sc['Technician'] = df_filtered_sc[emp_key].fillna(df_filtered_sc.get('Target Employee', 'Unknown'))
        
        # Helpers for solved and duration
        def is_sc_solved(row):
            st_val = str(row.get("Status", "")).strip().lower()
            rm_val = str(row.get("Grid Remark", "")).strip().lower()
            if st_val in ["completed", "closed"]:
                return True
            if "ms" in rm_val or "assign" in rm_val or "transfer" in rm_val or "forward" in rm_val:
                return True
            return False
            
        def parse_mins(val):
            if not val or pd.isna(val):
                return None
            s = str(val).strip()
            if s in ['N/A', 'Negative Time', '']:
                return None
            total_m = 0
            d_m = re.search(r'(\d+)\s*d', s)
            h_m = re.search(r'(\d+)\s*h', s)
            m_m = re.search(r'(\d+)\s*m', s)
            if d_m: total_m += int(d_m.group(1)) * 1440
            if h_m: total_m += int(h_m.group(1)) * 60
            if m_m: total_m += int(m_m.group(1))
            return total_m if (d_m or h_m or m_m) else None
            
        df_filtered_sc['Is_Solved'] = df_filtered_sc.apply(is_sc_solved, axis=1)
        df_filtered_sc['Is_Wifi6'] = df_filtered_sc.apply(lambda r: classify_work_type(r) == 'WiFi 6 Upgrade', axis=1)
        df_filtered_sc['Duration_Mins'] = df_filtered_sc['Resolution Time (From Assigned)'].apply(parse_mins)
        
        # Aggregate per technician
        scorecard_rows = []
        for tech in sorted(df_filtered_sc['Technician'].dropna().unique()):
            sub = df_filtered_sc[df_filtered_sc['Technician'] == tech]
            tot = len(sub)
            if tot == 0:
                continue
            solved = int(sub['Is_Solved'].sum())
            wifi6 = int(sub['Is_Wifi6'].sum())
            sol_rate = round((solved / tot * 100), 1) if tot > 0 else 0.0
            
            valid_durations = sub['Duration_Mins'].dropna()
            med_time = round(valid_durations.median(), 0) if not valid_durations.empty else None
            avg_time = round(valid_durations.mean(), 0) if not valid_durations.empty else None
            
            def fmt_m(m):
                if m is None: return "N/A"
                m = int(m)
                days = m // 1440
                hours = (m % 1440) // 60
                mins = m % 60
                parts = []
                if days > 0: parts.append(f"{days}d")
                if hours > 0: parts.append(f"{hours}h")
                parts.append(f"{mins}m")
                return " ".join(parts)
                
            scorecard_rows.append({
                "Technician": tech,
                "Total Tickets": tot,
                "Solved Count": solved,
                "Solution Rate %": sol_rate,
                "WiFi 6 Upgrades": wifi6,
                "Median Resolution Time": fmt_m(med_time),
                "Median_Min_Raw": med_time if med_time is not None else 999999
            })
            
        sc_df = pd.DataFrame(scorecard_rows)
        if sc_df.empty:
            st.info("No records match the selected horizon.")
        else:
            # Award Medals and Ranks
            # 1. Volume Champion
            top_volume = sc_df.sort_values(by="Total Tickets", ascending=False).iloc[0]
            # 2. Top Wi-Fi 6 Upgrader
            top_wifi6 = sc_df.sort_values(by="WiFi 6 Upgrades", ascending=False).iloc[0]
            # 3. Fastest Speed Demon (with at least 50 tickets or max)
            speed_pool = sc_df[sc_df["Total Tickets"] >= min(20, sc_df["Total Tickets"].max())]
            top_speed = speed_pool.sort_values(by="Median_Min_Raw", ascending=True).iloc[0] if not speed_pool.empty else sc_df.iloc[0]
            
            # Display Trophy Podium
            p1, p2, p3 = st.columns(3)
            with p1:
                st.markdown(f"""
                <div class="kpi-card" style="border: 1px solid #F9AB00; background: #FEF7E0;">
                    <div class="kpi-title" style="color: #B06000;">🥇 Volume Champion</div>
                    <div class="kpi-value" style="color: #202124; font-size: 1.25rem;">{top_volume['Technician']}</div>
                    <div class="kpi-subtext" style="color: #5F6368;"><b>{top_volume['Total Tickets']} Tickets</b> Handled</div>
                </div>
                """, unsafe_allow_html=True)
            with p2:
                st.markdown(f"""
                <div class="kpi-card" style="border: 1px solid #1A73E8; background: #E8F0FE;">
                    <div class="kpi-title" style="color: #1A73E8;">⚡ Speed Demon (Median TAT)</div>
                    <div class="kpi-value" style="color: #202124; font-size: 1.25rem;">{top_speed['Technician']}</div>
                    <div class="kpi-subtext" style="color: #5F6368;"><b>{top_speed['Median Resolution Time']}</b> Turnaround</div>
                </div>
                """, unsafe_allow_html=True)
            with p3:
                st.markdown(f"""
                <div class="kpi-card" style="border: 1px solid #137333; background: #E6F4EA;">
                    <div class="kpi-title" style="color: #137333;">📡 Wi-Fi 6 Upgrade Master</div>
                    <div class="kpi-value" style="color: #202124; font-size: 1.25rem;">{top_wifi6['Technician']}</div>
                    <div class="kpi-subtext" style="color: #5F6368;"><b>{top_wifi6['WiFi 6 Upgrades']} Upgrades</b> Installed</div>
                </div>
                """, unsafe_allow_html=True)
                
            st.markdown("<br>", unsafe_allow_html=True)
            st.subheader("Official Performance Leaderboard")
            
            # Formatted Leaderboard Table
            leaderboard_table = sc_df.sort_values(by="Total Tickets", ascending=False).drop(columns=["Median_Min_Raw"])
            leaderboard_table.reset_index(drop=True, inplace=True)
            leaderboard_table.index += 1
            leaderboard_table.index.name = "Rank"
            
            st.dataframe(leaderboard_table, use_container_width=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.subheader("Workload & Upgrade Visual Distribution")
            chart_col1, chart_col2 = st.columns(2)
            with chart_col1:
                st.caption("Tickets Handled per Technician")
                st.bar_chart(leaderboard_table.set_index("Technician")["Total Tickets"])
            with chart_col2:
                st.caption("Wi-Fi 6 Upgrades Completed per Technician")
                st.bar_chart(leaderboard_table.set_index("Technician")["WiFi 6 Upgrades"])

with main_mode_tab2:
    st.subheader("Executive Team Report Merger (Manager Tool)")
    st.write("Upload individual employee audit Excel files (`audit_report_*.xlsx`) to merge them into a single **Executive Team Report** for management!")
    
    uploaded_files = st.file_uploader(
        "Upload Individual Employee Audit Excel Reports",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Select multiple files (e.g. Om Neupane, Ajit Shrestha, Shashikant Chaudhary, Sabin Giri, Sunil Chaudhary, Sanjeev Giri, Chandramani Tharu, Rajesh Maharjan, etc.)"
    )
    
    if uploaded_files:
        st.success(f"Received {len(uploaded_files)} Excel report file(s) for team merging.")
        
        all_dfs = []
        for file in uploaded_files:
            try:
                xls = pd.ExcelFile(file)
                if "Audit Details" in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name="Audit Details")
                    
                    emp_name = None
                    if "Grid Employee Name" in df.columns and not df["Grid Employee Name"].dropna().empty:
                        emp_name = df["Grid Employee Name"].dropna().iloc[0]
                    else:
                        match = re.search(r'audit_report_(.+?)_\d', file.name)
                        if match:
                            emp_name = match.group(1).replace('_', ' ')
                        else:
                            emp_name = file.name.replace('.xlsx', '')
                            
                    df["Employee Name"] = emp_name
                    df["Task / Issue Type"] = df.apply(classify_work_type, axis=1)
                    all_dfs.append(df)
            except Exception as read_err:
                st.warning(f"Could not read {file.name}: {read_err}")
                
        if all_dfs:
            master_df = pd.concat(all_dfs, ignore_index=True)
            
            def is_solved(row):
                st_val = str(row.get("Status", "")).strip().lower()
                rm_val = str(row.get("Grid Remark", "")).strip().lower()
                if st_val in ["completed", "closed"]:
                    return True
                if "ms" in rm_val or "assign" in rm_val or "transfer" in rm_val or "forward" in rm_val:
                    return True
                return False

            master_df["Is_Solved_Val"] = master_df.apply(is_solved, axis=1)
            
            # Team Summary Calculation
            summary_rows = []
            for emp, grp in master_df.groupby("Employee Name"):
                tot = len(grp)
                if tot == 0:
                    continue  # Exclude technicians on day off (0 tickets)
                solved = grp["Is_Solved_Val"].sum()
                rate = f"{(solved / tot * 100):.1f}%" if tot > 0 else "0.0%"
                summary_rows.append({
                    "Employee Name": emp,
                    "Total Scraped Tickets": tot,
                    "Solved / Handled Count": solved,
                    "Solution Rate %": rate
                })
                
            team_summary_df = pd.DataFrame(summary_rows)
            
            # Total Row
            tot_tickets_team = len(master_df)
            tot_solved_team = master_df["Is_Solved_Val"].sum()
            team_rate_val = f"{(tot_solved_team / tot_tickets_team * 100):.1f}%" if tot_tickets_team > 0 else "0.0%"
            
            total_team_row = pd.DataFrame([{
                "Employee Name": "GRAND TOTAL (ALL TEAM)",
                "Total Scraped Tickets": tot_tickets_team,
                "Solved / Handled Count": tot_solved_team,
                "Solution Rate %": team_rate_val
            }])
            team_summary_df = pd.concat([team_summary_df, total_team_row], ignore_index=True)
            
            # Task / Issue Breakdown Calculation
            work_summary_df = master_df.groupby("Task / Issue Type", dropna=False).agg(
                Total_Tickets=("Ticket Number", "count"),
                Solved_Count=("Is_Solved_Val", "sum")
            ).reset_index()
            work_summary_df["Solution Rate %"] = (work_summary_df["Solved_Count"] / work_summary_df["Total_Tickets"] * 100).round(1).astype(str) + '%'
            work_summary_df["% Share of Total"] = (work_summary_df["Total_Tickets"] / len(master_df) * 100).round(1).astype(str) + '%'
            
            # Category Summary Calculation
            cat_summary_df = master_df.groupby("Category", dropna=False).agg(
                Total_Tickets=("Ticket Number", "count"),
                Solved_Count=("Is_Solved_Val", "sum")
            ).reset_index()
            cat_summary_df["Solution Rate %"] = (cat_summary_df["Solved_Count"] / cat_summary_df["Total_Tickets"] * 100).round(1).astype(str) + '%'
            
            export_master = master_df.drop(columns=["Is_Solved_Val"], errors="ignore")
            
            st.markdown("### Executive Team Performance Summary")
            
            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("Total Team Tickets Scraped", tot_tickets_team)
            with c2:
                st.metric("Total Team Solved / Handled", tot_solved_team)
            with c3:
                st.metric("Overall Team Solution Rate", team_rate_val)
                
            st.markdown("#### Employee Workload & Performance Comparison")
            st.dataframe(team_summary_df, use_container_width=True)
            
            st.markdown("---")
            st.markdown("### Specific Task & Issue Breakdown (WiFi 6 Upgrades, IPTV, Hardware, etc.)")
            st.dataframe(work_summary_df, use_container_width=True)
            st.bar_chart(work_summary_df.set_index("Task / Issue Type")["Total_Tickets"])
            
            st.markdown("#### Tickets Workload per Employee")
            st.bar_chart(team_summary_df[team_summary_df["Employee Name"] != "GRAND TOTAL (ALL TEAM)"].set_index("Employee Name")["Total Scraped Tickets"])
            
            st.markdown("#### Combined Filterable Master Audit Table")
            st.dataframe(export_master, use_container_width=True, height=400)
            
            excel_bytes = build_executive_team_excel(export_master, team_summary_df, work_summary_df)
            today_filename_str = datetime.now().strftime("%d_%b_%Y")
            exec_file_name = f"EXECUTIVE_TEAM_AUDIT_REPORT_{today_filename_str}.xlsx"
            
            st.download_button(
                label=f"Download Combined Executive Team Report ({exec_file_name})",
                data=excel_bytes,
                file_name=exec_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
