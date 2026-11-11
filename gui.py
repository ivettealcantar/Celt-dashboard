import asyncio
from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st
import national_grid_scraper as scraper

st.set_page_config(
    page_title="National Grid Utility Scraper & Analytics",
    page_icon="⚡",
    layout="wide",
)

st.title("⚡ National Grid Utility Scraper & Analytics")
st.markdown("Upload account ledgers, execute automated web scrapers, and analyze late fees and maintenance charges.")

# --- Sidebar: User Credentials & Operations ---
st.sidebar.header("🔑 National Grid Login")
user_email = st.sidebar.text_input("Email / Username")
user_password = st.sidebar.text_input("Password", type="password")

st.sidebar.divider()
st.sidebar.header("📁 Operations Panel")
uploaded_accounts = st.sidebar.file_uploader(
    "Upload Accounts (.xlsx, .csv)", type=["xlsx", "csv"]
)

# --- Multi-Tab Interface ---
tab_run, tab_analytics = st.tabs(["🚀 Scraper Execution", "📊 Analytics & Reports"])

with tab_run:
    st.header("Run Portal Web Scraper")
    
    if uploaded_accounts and user_email and user_password:
        st.success("Credentials and account file loaded successfully.")
        
        if st.button("▶️ Start Scraper Process", type="primary"):
            status_box = st.status("Initializing scraping session...", expanded=True)
            
            # Create an isolated temporary folder per user execution
            session_id = st.runtime.scriptrunner.get_script_run_ctx().session_id
            temp_dir = Path(f"temp_{session_id}")
            temp_dir.mkdir(exist_ok=True)
            
            input_path = temp_dir / "input.xlsx"
            output_path = temp_dir / "output.xlsx"
            
            with open(input_path, "wb") as f:
                f.write(uploaded_accounts.getbuffer())
                
            try:
                # Execute async scraper pipeline
                asyncio.run(scraper.run_scraper_process(
                    username=user_email,
                    password=user_password,
                    input_path=input_path, 
                    output_path=output_path, 
                    status_box=status_box
                ))
                
                status_box.update(label="✅ Processing Complete!", state="complete")
                st.balloons()
                
                # Display output preview and download button
                if output_path.exists():
                    df_results = pd.read_excel(output_path)
                    st.subheader("Scraped Accounts Preview")
                    st.dataframe(df_results, use_container_width=True)
                    
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="📥 Download Full Excel Report", 
                            data=f, 
                            file_name="utility_report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
            except Exception as err:
                status_box.update(label="❌ Scraping Failed", state="error")
                st.error(f"Error: {err}")
                
            finally:
                # Clean up session directory immediately
                for item in temp_dir.glob("*"):
                    item.unlink()
                temp_dir.rmdir()

    else:
        if not user_email or not user_password:
            st.warning("👈 Please enter your National Grid Email and Password in the sidebar.")
        if not uploaded_accounts:
            st.info("👈 Please upload an account ledger file (.xlsx or .csv) to begin.")

with tab_analytics:
    st.header("Data Analytics & Reports")
    
    report_path = Path("testerfinal_payment_and_late_fee_report.xlsx")
        
    if report_path.exists():
        df_report = pd.read_excel(report_path)
        
        # Summary Metrics Row
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Accounts", len(df_report["Account Number"].unique()) if "Account Number" in df_report else len(df_report))
        
        total_charges = df_report['Current Charges'].sum() if "Current Charges" in df_report else 0
        m2.metric("Total Current Charges", f"${total_charges:,.2f}")
        
        total_late = df_report['Late Fee Amount'].sum() if "Late Fee Amount" in df_report else 0
        m3.metric("Total Late Fees", f"${total_late:,.2f}")
        
        flagged = (df_report["Result"] == "Needs review").sum() if "Result" in df_report else 0
        m4.metric("Flagged Accounts", flagged)
        
        st.divider()
        
        # Analytics Charts
        col_c1, col_c2 = st.columns(2)
        
        with col_c1:
            st.subheader("Current Charges by Utility Type")
            if "Utility Type" in df_report and "Current Charges" in df_report:
                fig_bar = px.bar(
                    df_report,
                    x="Utility Type",
                    y="Current Charges",
                    color="Result" if "Result" in df_report else None,
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

        st.subheader("Detailed Report View")
        st.dataframe(df_report, use_container_width=True)
    else:
        st.warning("No analytical report generated yet. Run a scraping job to populate data.")