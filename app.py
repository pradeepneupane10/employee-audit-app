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
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

# Set Streamlit Page Configuration for Mobile & Desktop
st.set_page_config(
    page_title="CGNET Employee Audit Portal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Modern, WOW Aesthetics & Mobile Responsiveness
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'DM Sans', sans-serif;
    }
    
    .stApp {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #f8fafc;
    }
    
    /* Header Container */
    .main-header {
        background: linear-gradient(90deg, #1e3a8a 0%, #3b82f6 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        box-shadow: 0 10px 25px -5px rgba(59, 130, 246, 0.3);
        margin-bottom: 2rem;
        text-align: center;
    }
    .main-header h1 {
        color: #ffffff !important;
        font-weight: 700;
        margin: 0;
        font-size: 2.2rem;
    }
    .main-header p {
        color: #e0f2fe !important;
        margin-top: 0.5rem;
        font-size: 1rem;
    }
    
    /* KPI Card Container */
    .kpi-card {
        background: rgba(30, 41, 59, 0.7);
        border: 1px solid rgba(255, 255, 255, 0.1);
        backdrop-filter: blur(10px);
        padding: 1.2rem;
        border-radius: 12px;
        text-align: center;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .kpi-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.3);
    }
    .kpi-title {
        color: #94a3b8;
        font-size: 0.9rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 0.5rem;
    }
    .kpi-value {
        color: #38bdf8;
        font-size: 1.8rem;
        font-weight: 700;
    }
    .kpi-subtext {
        color: #64748b;
        font-size: 0.8rem;
        margin-top: 0.3rem;
    }
    
    /* Download Button */
    .stDownloadButton > button {
        background: linear-gradient(90deg, #10b981 0%, #059669 100%) !important;
        color: white !important;
        font-weight: 700 !important;
        border: none !important;
        padding: 0.75rem 2rem !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 14px rgba(16, 185, 129, 0.4) !important;
        width: 100%;
    }
    
    /* Status Box */
    .status-box {
        background: #1e293b;
        border-left: 4px solid #3b82f6;
        padding: 1rem;
        border-radius: 4px;
        font-family: monospace;
        max-height: 200px;
        overflow-y: auto;
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)

# Auto-ensure Playwright chromium binary is installed in cloud environment
import subprocess
@st.cache_resource
def setup_playwright():
    try:
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False)
    except Exception:
        pass
    return True

setup_playwright()

# Import scraper from audit_automation module
import audit_automation

# Title Banner
st.markdown("""
<div class="main-header">
    <h1>⚡ CGNET Automated Employee Audit Portal</h1>
    <p>Run real-time performance audit scraping, analyze ticket resolutions & download executive reports from any device.</p>
</div>
""", unsafe_allow_html=True)

# Sidebar Inputs
with st.sidebar:
    st.header("📋 Audit Configuration")
    
    # Employee Selection
    emp_options = ["Om Neupane", "Laxman Koirala", "Custom Employee Name..."]
    selected_emp_type = st.selectbox("Employee Name", emp_options, index=0)
    
    if selected_emp_type == "Custom Employee Name...":
        employee_name = st.text_input("Enter Full Employee Name", value="Om Neupane")
    else:
        employee_name = selected_emp_type
        
    st.markdown("---")
    st.subheader("📅 Date Range Selector")
    
    today = datetime.now().date()
    yesterday = today - timedelta(days=1)
    
    date_preset = st.radio("Quick Range", ["Today", "Yesterday & Today (2 Days)", "Last 7 Days", "Custom Range"], index=1)
    
    if date_preset == "Today":
        from_date_obj = today
        to_date_obj = today
    elif date_preset == "Yesterday & Today (2 Days)":
        from_date_obj = yesterday
        to_date_obj = today
    elif date_preset == "Last 7 Days":
        from_date_obj = today - timedelta(days=6)
        to_date_obj = today
    else:
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            from_date_obj = st.date_input("From Date", value=yesterday)
        with col_d2:
            to_date_obj = st.date_input("To Date", value=today)

    from_date_str = from_date_obj.strftime("%d %b %Y")  # e.g., "04 Aug 2026"
    to_date_str = to_date_obj.strftime("%d %b %Y")      # e.g., "05 Aug 2026"

    st.markdown(f"**Target Period:** `{from_date_str}` to `{to_date_str}`")
    st.markdown("---")
    
    run_btn = st.button("🚀 Run Audit Scraper", type="primary", use_container_width=True)

# Main Application Body
if run_btn:
    st.info(f"⏳ Starting automated scraper for **{employee_name}** from `{from_date_str}` to `{to_date_str}`...")
    
    log_container = st.empty()
    logs_list = []
    
    def ui_log(msg, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_line = f"[{timestamp}] [{level}] {msg}"
        logs_list.append(log_line)
        log_html = "<br>".join(logs_list[-8:])
        log_container.markdown(f'<div class="status-box">{log_html}</div>', unsafe_allow_html=True)

    progress_bar = st.progress(0.1, text="Initializing Playwright browser...")

    try:
        # Override CLI args in sys.argv for audit_automation main execution
        sys.argv = [
            "audit_automation.py",
            "--employee", employee_name,
            "--from-date", from_date_str,
            "--to-date", to_date_str,
            "--non-interactive"
        ]
        
        # Ensure Playwright Chromium binary is installed
        import subprocess
        try:
            subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False)
        except Exception:
            pass

        # Execute scraper
        audit_automation.main()
        
        progress_bar.progress(1.0, text="Scraping completed!")
        st.success("✅ Audit Scraper completed successfully!")
        
        # Determine expected output filename
        safe_emp = re.sub(r'[^a-zA-Z0-9]', '_', employee_name)
        safe_from = re.sub(r'[^a-zA-Z0-9]', '_', from_date_str)
        output_file = f"audit_report_{safe_emp}_{safe_from}.xlsx"
        
        if os.path.exists(output_file):
            st.session_state["last_output_file"] = output_file
            st.session_state["last_run_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        else:
            st.warning("Report file generated, but filename format varied.")
            
    except Exception as e:
        st.error(f"❌ Scraper error encountered: {e}")
        st.exception(e)

# Display Dashboard if last output file exists
output_file = st.session_state.get("last_output_file", None)

if output_file and os.path.exists(output_file):
    st.markdown("### 📊 Performance Analytics Dashboard")
    st.caption(f"Last updated: {st.session_state.get('last_run_time', 'Recently')} | Report File: `{os.path.basename(output_file)}`")
    
    try:
        xls = pd.ExcelFile(output_file)
        df_details = pd.read_excel(xls, sheet_name="Audit Details")
        df_summary = pd.read_excel(xls, sheet_name="Summary Report", nrows=5)
        
        # Metrics Display
        col1, col2, col3, col4, col5 = st.columns(5)
        
        def get_val(metric_name):
            row = df_summary[df_summary["Metric"].str.contains(metric_name, case=False, na=False)]
            if not row.empty:
                return str(row.iloc[0]["Value"])
            return "N/A"

        with col1:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">Total Records</div>
                <div class="kpi-value">{get_val("Total Audit Records Scraped")}</div>
                <div class="kpi-subtext">Scraped logs</div>
            </div>
            """, unsafe_allow_html=True)
            
        with col2:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">Solved / Handled</div>
                <div class="kpi-value">{get_val("Total Tickets Solved")}</div>
                <div class="kpi-subtext">Completed or MS Assigned</div>
            </div>
            """, unsafe_allow_html=True)
            
        with col3:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">Solution Rate</div>
                <div class="kpi-value">{get_val("Ticket Solution Rate")}</div>
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

        # Excel Download Section
        st.markdown("### 📥 Download Executive Report")
        with open(output_file, "rb") as f:
            bytes_data = f.read()
            st.download_button(
                label=f"⬇️ Download Excel Audit Report ({os.path.basename(output_file)})",
                data=bytes_data,
                file_name=os.path.basename(output_file),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.markdown("<br>", unsafe_allow_html=True)
        
        # Detailed Data Tabs
        tab1, tab2 = st.tabs(["📝 Scraped Audit Details", "📊 Raw Summary Sheet"])
        
        with tab1:
            st.subheader("Filterable Audit Details Table")
            search_query = st.text_input("🔍 Search records by ticket #, remark, or account name...", "")
            
            if search_query:
                filtered_df = df_details[df_details.astype(str).apply(lambda row: row.str.contains(search_query, case=False).any(), axis=1)]
            else:
                filtered_df = df_details
                
            st.dataframe(filtered_df, use_container_width=True, height=400)
            
        with tab2:
            st.subheader("Summary Report Sheet View")
            df_full_summary = pd.read_excel(xls, sheet_name="Summary Report")
            st.dataframe(df_full_summary, use_container_width=True)

    except Exception as read_err:
        st.error(f"Could not load output preview: {read_err}")
else:
    st.info("👈 Use the sidebar controls to select an Employee and Date Range, then click **Run Audit Scraper** to generate your report.")
