import asyncio
import os
from pathlib import Path
import pandas as pd
import streamlit as st
from playwright.async_api import async_playwright

KERNEL_API_KEY = st.secrets.get("KERNEL_API_KEY", os.getenv("KERNEL_API_KEY"))

def load_accounts(input_path: Path) -> pd.DataFrame:
    """Utility to read Excel/CSV account lists."""
    if input_path.suffix.lower() in [".xlsx", ".xls"]:
        return pd.read_excel(input_path)
    return pd.read_csv(input_path)

async def run_scraper_process(username, password, input_path, output_path, status_box=None):
    """Main scraping pipeline."""
    accounts_df = load_accounts(input_path)
    results = []

    async with async_playwright() as p:
        # 1. Connect to Cloud CDP Browser or Local Fallback
        if KERNEL_API_KEY:
            if status_box:
                status_box.write("Connecting to remote browser engine...")
            browser = await p.chromium.connect_over_cdp(
                f"wss://connect.kernel.sh?api_key={KERNEL_API_KEY}"
            )
        else:
            if status_box:
                status_box.write("Launching browser instance...")
            browser = await p.chromium.launch(headless=True)

        context = await browser.new_context()
        page = await context.new_page()

        try:
            # 2. Dynamic Login Phase
            if status_box:
                status_box.write("Navigating to National Grid login portal...")

            await page.goto("https://www.nationalgridus.com/", wait_until="networkidle")

            if status_box:
                status_box.write("Submitting credentials...")

            # Adjust CSS selectors as needed to match National Grid's portal fields
            await page.fill('input[type="email"], input[name="username"]', username)
            await page.fill('input[type="password"]', password)
            await page.click('button[type="submit"], input[type="submit"]')

            await page.wait_for_load_state("networkidle")

            if status_box:
                status_box.write("Login verified. Scraping account bills...")

            # 3. Account Iteration Phase
            total = len(accounts_df)
            for idx, row in accounts_df.iterrows():
                acct_num = str(row.get("Account Number", "")).strip()
                if status_box:
                    status_box.write(f"Processing account [{idx + 1}/{total}]: {acct_num}")

                # --- Place your page navigation & extraction logic here ---
                # Example dummy record structure:
                results.append({
                    "Account Number": acct_num,
                    "Account Address": row.get("Account Address", ""),
                    "Current Charges": 0.00,
                    "Late Fee Amount": 0.00,
                    "Late Fee Instance": "No",
                    "Result": "Success"
                })

        except Exception as e:
            if status_box:
                status_box.write(f"Error encountered: {e}")
            raise e

        finally:
            await browser.close()

    # Save output to Excel
    out_df = pd.DataFrame(results)
    out_df.to_excel(output_path, index=False)
    return results