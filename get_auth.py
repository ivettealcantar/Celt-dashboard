import asyncio
from pathlib import Path
from playwright.async_api import async_playwright

# Save the auth file in your project folder
AUTH_FILE = Path(__file__).parent / "auth_state.json"
START_URL = "https://accountmanager.nationalgrid.com"

async def generate_auth_state():
    print("=" * 60)
    print("LAUNCHING VISIBLE BROWSER FOR MANUAL LOGIN")
    print("=" * 60)
    
    async with async_playwright() as playwright:
        # Launch visible Chrome/Chromium browser engine
        browser = await playwright.chromium.launch(
            headless=False,  # Must be False so you can see and interact with the page
            channel="chrome"  # Uses installed Chrome browser
        )
        context = await browser.new_context(
            viewport={"width": 1440, "height": 900}
        )
        page = await context.new_page()

        print(f"\n1. Navigating to {START_URL}...")
        try:
            await page.goto(START_URL, timeout=120_000)
        except Exception as exc:
            print(f"Navigation note: {exc}")

        print("\n2. PLEASE MANUALLY:")
        print("   - Enter your Username and Password in the opened browser.")
        print("   - Complete any 2FA / Security Verification codes.")
        print("   - Navigate until you are on the 'Accounts Overview' dashboard page.")
        
        input("\n>>> Once 'Accounts Overview' is fully visible in the browser, press ENTER here to save session... <<<")

        # Capture session cookies and storage state
        await context.storage_state(path=str(AUTH_FILE))
        print(f"\n SUCCESS! Session saved to: {AUTH_FILE.resolve()}")
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(generate_auth_state())