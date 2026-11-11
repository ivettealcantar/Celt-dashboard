import asyncio
from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st

# Import your scraper module
import national_grid_scraper as scraper

st.set_page_config(
    page_title="Utility Automation & Analytics",
    page_icon="⚡",
    layout="wide",
)

st.title("⚡ Utility Scraper & Capital Operations Dashboard")
st.markdown("Upload account ledgers, execute automated web scrapers, and analyze late fees and maintenance charges.")

# Sidebar Controls
st.sidebar.header("📁 Operations Panel")
uploaded_file = st.sidebar.file_uploader(
    "Upload Accounts (.xlsx, .csv)", type=["xlsx", "csv"]
)

# Authentication State Uploader
st.sidebar.divider()
st.sidebar.subheader("🔐 Session Cookie Management")
auth_file_upload = st.sidebar.file_uploader("Update auth_state.json", type=["json"])
if auth_file_upload:
    auth_save_path = Path(__file__).parent / "auth_state.json"
    with open(auth_save_path, "wb") as f:
        f.write(auth_file_upload.getbuffer())
    st.sidebar.success("Session state updated!")

# Multi-Tab Layout
tab_run, tab_analytics = st.tabs(["🚀 Scraper Execution", "📊 Analytics & Reports"])

with tab_run:
    st.header("Run Portal Web Scraper")
    
    if uploaded_file is not None:
        temp_input_path = Path("temp_input.xlsx")
        temp_output_path = Path("temp_output_report.xlsx")
        
        with open(temp_input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        st.success(f"Loaded File: **{uploaded_file.name}**")
        
        try:
            df_input = scraper.load_accounts(temp_input_path)
            st.subheader("Account Queue Preview")
            st.dataframe(df_input, use_container_width=True)
        except Exception as e:
            st.error(f"Error reading file format: {e}")
            
        if st.button("▶️ Start Scraper Process", type="primary"):
            status_box = st.status("Executing Scraper Task...", expanded=True)
            
            try:
                # Run asynchronous scraping pipeline
                asyncio.run(scraper.run_scraper_process(temp_input_path, temp_output_path, status_box))
                
                status_box.update(label="✅ Process Completed Successfully!", state="complete")
                st.balloons()
            except Exception as err:
                status_box.update(label="❌ Scraping Failed", state="error")
                st.error(f"Error: {err}")
    else:
        st.info("👈 Please upload an Excel or CSV file in the left sidebar to begin.")

with tab_analytics:
    st.header("Data Analytics & Reports")
    
    report_path = Path("temp_output_report.xlsx")
    if not report_path.exists():
        report_path = Path("testerfinal_payment_and_late_fee_report.xlsx")
        
    if report_path.exists():
        df_report = pd.read_excel(report_path)
        
        # Summary Metrics Row
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Accounts", len(df_report["Account Number"].unique()))
        
        total_charges = df_report['Current Charges'].sum() if "Current Charges" in df_report else 0
        m2.metric("Total Current Charges", f"${total_charges:,.2f}")
        
        total_late = df_report['Late Fee Amount'].sum() if "Late Fee Amount" in df_report else 0
        m3.metric("Total Late Fees", f"${total_late:,.2f}")
        
        flagged = (df_report["Result"] == "Needs review").sum() if "Result" in df_report else 0
        m4.metric("Flagged Accounts", flagged)
        
        st.divider()
        
        # Charts Row
        col_c1, col_c2 = st.columns(2)
        
        with col_c1:
            st.subheader("Current Charges by Utility Type")
            if "Utility Type" in df_report and "Current Charges" in df_report:
                fig_bar = px.bar(
                    df_report,
                    x="Utility Type",
                    y="Current Charges",
                    color="Result",
                    barmode="group",
                    title="Utility Cost Distribution",
                )
                st.plotly_chart(fig_bar, use_container_width=True)
                
        with col_c2:
            st.subheader("Late Fee Frequency Ratio")
            if "Late Fee Instance" in df_report:
                fig_pie = px.pie(
                    df_report,
                    names="Late Fee Instance",
                    title="Late Fee Occurrences",
                    hole=0.4,
                )
                st.plotly_chart(fig_pie, use_container_width=True)

        # Interactive Data Table & Export
        st.subheader("Detailed Report View")
        st.dataframe(df_report, use_container_width=True)
        
        with open(report_path, "rb") as file:
            st.download_button(
                label="📥 Download Full Excel Report",
                data=file,
                file_name="utility_late_fee_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.warning("No report file available yet. Upload accounts and run the scraper first.")