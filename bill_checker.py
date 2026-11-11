import asyncio
import re
import sys
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd
from openpyxl.styles import Font
from playwright.async_api import async_playwright


# ---------------- SETTINGS ----------------

INPUT_FILE = Path("tester.xlsx")
OUTPUT_FILE = Path("testerfinal_payment_and_late_fee_report.xlsx")
SHEET_NAME = 0
START_URL = "https://accountmanager.nationalgrid.com"
MAX_SCROLL_ROUNDS = 40
NAVIGATION_TIMEOUT_MS = 120_000

# Saved login session (per user, in their home folder).
AUTH_FILE = Path.home() / ".bill_checker" / "auth_state.json"
AUTH_FILE.parent.mkdir(parents=True, exist_ok=True)

# Hidden browser only works after a first visible login has been saved.
HEADLESS = False

# "chrome" or "msedge" uses the browser already installed on the computer.
BROWSER_CHANNEL = "chrome"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

OUTPUT_COLUMNS = [
    "Account Number",
    "Account Address",
    "Billing History Row",
    "Service Address",
    "Utility Type",
    "Bill Date",
    "Due Date",
    "Payment Instance",
    "Payment Instance Count",
    "Payment Received Amount",
    "Payment Received Signed Amount",
    "Previous Balance",
    "Balance Forward",
    "Payment Coverage of Current Charges (%)",
    "Current Charges",
    "Amount Past Due",
    "Amount Due",
    "Rate Plan / Tariff",
    "Late Fee Amount",
    "Late Fee Instance",
    "Result",
    "Notes",
]


def pause(message):
    """Overridden by the GUI. Falls back to the console."""
    return input(message)


# ---------------- ACCOUNT INPUT ----------------

ACCOUNT_HEADERS = {
    "account number", "account", "account #", "account no", "account no.",
    "acct", "acct #", "acct number", "accountnumber",
}
ADDRESS_HEADERS = {"account address", "address", "service address", "name"}


def normalize_account(value):
    """Keep account numbers as digits and restore leading zeroes if needed."""
    text = str(value or "").strip()

    if text.endswith(".0"):
        text = text[:-2]

    digits = re.sub(r"\D", "", text)

    # Keep this only if your account numbers are expected to be 10 digits.
    if digits and len(digits) < 10:
        digits = digits.zfill(10)

    return digits


def account_digits(value):
    return re.sub(r"\D", "", str(value or ""))


def load_accounts(path=None):
    """Read accounts from .xlsx/.xls/.csv. Works with or without a header row."""
    path = Path(path or INPUT_FILE)

    if not path.exists():
        raise FileNotFoundError(f"Cannot find {path}")

    # Read everything as text so leading zeroes survive.
    if path.suffix.lower() == ".csv":
        raw = pd.read_csv(path, dtype=str, header=None)
    else:
        raw = pd.read_excel(path, sheet_name=SHEET_NAME, dtype=str, header=None)

    raw = raw.dropna(how="all").reset_index(drop=True)
    if raw.empty:
        raise ValueError("The file is empty.")

    first_row = [str(v).strip().casefold() for v in raw.iloc[0]]
    account_col = next(
        (i for i, v in enumerate(first_row) if v in ACCOUNT_HEADERS), None
    )
    address_col = next(
        (i for i, v in enumerate(first_row) if v in ADDRESS_HEADERS), None
    )

    if account_col is not None:
        data = raw.iloc[1:]  # header row found, skip it
    else:
        account_col, address_col = 0, None  # no header: first column = accounts
        data = raw

    accounts = pd.DataFrame({
        "Account Number": data.iloc[:, account_col].fillna("").astype(str),
        "Account Address": (
            data.iloc[:, address_col].fillna("").astype(str).str.strip()
            if address_col is not None else ""
        ),
    })

    accounts["Account Number"] = accounts["Account Number"].apply(
        normalize_account
    )
    accounts = accounts[accounts["Account Number"] != ""]
    accounts = accounts.drop_duplicates("Account Number").reset_index(drop=True)

    if accounts.empty:
        raise ValueError(
            "No account numbers found. Put them in a column headed "
            "'Account Number', or in the first column."
        )

    return accounts


# ---------------- GENERAL HELPERS ----------------

def clean_text(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def parse_money(value):
    """Convert money text to a signed float."""
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    negative = "-" in text or "−" in text or "(" in text
    numeric = re.sub(r"[^\d.]", "", text)

    if not numeric:
        return None

    try:
        amount = float(numeric)
    except ValueError:
        return None

    return -amount if negative else amount


def find_money_values(text):
    """Find signed money amounts, including comma-separated thousands."""
    if not text:
        return []

    matches = re.findall(
        r"(?<![\w])[-−]?\s*\$?\s*"
        r"(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}"
        r"\s*[-−]?",
        text,
    )

    return [
        value
        for value in (parse_money(item) for item in matches)
        if value is not None
    ]


# ---------------- BILL FIELD PARSING ----------------

def extract_account_number(text):
    match = re.search(
        r"ACCOUNT\s+NUMBER\s+([0-9][0-9-]+)",
        text,
        re.IGNORECASE,
    )
    return match.group(1).strip() if match else ""


def extract_bill_date(text):
    match = re.search(
        r"DATE\s+BILL\s+ISSUED\s+"
        r"([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})",
        text,
        re.IGNORECASE,
    )
    return clean_text(match.group(1)) if match else ""


def extract_due_date(text):
    match = re.search(
        r"PLEASE\s+PAY\s+BY\s+"
        r"([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})",
        text,
        re.IGNORECASE,
    )
    return clean_text(match.group(1)) if match else ""


def extract_service_and_rate(raw_text):
    """
    Extract the address and utility type from the SERVICE FOR block,
    and the rate plan/tariff from the bill.
    """
    result = {
        "Service Address": "",
        "Utility Type": "",
        "Rate Plan / Tariff": "",
    }

    service_match = re.search(
        r"SERVICE\s+FOR\s+(.*?)\s+BILLING\s+PERIOD",
        raw_text,
        re.IGNORECASE | re.DOTALL,
    )

    service_lines = []
    if service_match:
        service_lines = [
            clean_text(line)
            for line in service_match.group(1).splitlines()
            if clean_text(line)
        ]

    # In this bill format, the first line is often the customer name;
    # following lines contain the service location.
    if len(service_lines) >= 2:
        address_lines = service_lines[1:]
    else:
        address_lines = service_lines

    address_lines = [
        re.sub(
            r"\s*,?\s*(GAS|ELECTRIC)\s*$", "", line, flags=re.IGNORECASE
        ).strip(" ,")
        for line in address_lines
    ]
    address_lines = [line for line in address_lines if line]

    result["Service Address"] = ", ".join(address_lines)

    first_page_text = "\n".join(raw_text.splitlines()[:120])
    if re.search(r"\bGAS\b", first_page_text, re.IGNORECASE):
        result["Utility Type"] = "Gas"
    elif re.search(r"\bELECTRIC\b", first_page_text, re.IGNORECASE):
        result["Utility Type"] = "Electric"

    flat_text = re.sub(r"\s+", " ", raw_text)
    rate_match = re.search(
        r"\bRATE\s+(.{1,180}?)"
        r"(?:VOLTAGE\s+DELIVERY\s+LEVEL|"
        r"SERVICE\s+PERIOD|"
        r"NUMBER\s+OF\s+DAYS|"
        r"METER\s+NUMBER|"
        r"NEXT\s+SCHEDULED|$)",
        flat_text,
        re.IGNORECASE,
    )

    if rate_match:
        result["Rate Plan / Tariff"] = clean_text(rate_match.group(1))

    return result


def extract_account_balance(raw_text):
    """
    Read account-balance rows from gas/electric bill formats.

    For multi-column rows, the final amount on the row is treated as the
    Total column. Multiple Payment Received rows are added together.
    """
    flat_text = re.sub(r"\s+", " ", raw_text).strip()

    section_match = re.search(
        r"ACCOUNT\s+BALANCE(.*?)(?:To\s+avoid\s+late\s+payment|"
        r"DETAIL\s+OF\s+CURRENT\s+CHARGES|"
        r"SUMMARY\s+OF\s+CURRENT\s+CHARGES|$)",
        flat_text,
        re.IGNORECASE | re.DOTALL,
    )
    section = section_match.group(1) if section_match else flat_text

    row_pattern = re.compile(
        r"Previous\s+Balance|"
        r"Payment\s*(?:\(s\))?\s+Received|"
        r"Payments\s+Received|"
        r"Balance\s+Forward|"
        r"Amount\s+Past\s+Due|"
        r"Current\s+Charges|"
        r"Amount\s+Due",
        re.IGNORECASE,
    )
    labels = list(row_pattern.finditer(section))

    def label_name(match):
        return re.sub(r"\s+", " ", match.group(0)).strip().casefold()

    def text_after_label(label_match):
        end = len(section)
        for next_label in labels:
            if next_label.start() > label_match.start():
                end = next_label.start()
                break
        return section[label_match.end():end].strip()

    def get_row_total(label_regex):
        for label_match in labels:
            if re.fullmatch(label_regex, label_name(label_match), re.IGNORECASE):
                values = find_money_values(text_after_label(label_match))
                return values[-1] if values else None
        return None

    previous_balance = get_row_total(r"previous\s+balance")
    balance_forward = get_row_total(r"balance\s+forward")
    current_charges = get_row_total(r"current\s+charges")
    amount_past_due = get_row_total(r"amount\s+past\s+due")
    amount_due = get_row_total(r"amount\s+due")

    payment_signed_amounts = []

    for label_match in labels:
        if not re.fullmatch(
            r"payment\s*(?:\(s\))?\s+received|payments\s+received",
            label_name(label_match),
            re.IGNORECASE,
        ):
            continue

        payment_row = text_after_label(label_match)

        if re.search(
            r"no\s+payments?\s+have\s+been\s+received",
            payment_row,
            re.IGNORECASE,
        ):
            continue

        values = find_money_values(payment_row)
        if values and abs(values[-1]) > 0.01:
            # Last amount is the total column when the bill has multiple
            # component columns.
            payment_signed_amounts.append(values[-1])

    payment_instance_count = len(payment_signed_amounts)
    signed_payment_total = sum(payment_signed_amounts)
    payment_received_amount = sum(
        abs(value) for value in payment_signed_amounts
    )

    # Use the bill's printed Balance Forward when present.
    # Otherwise calculate it only when previous balance is known.
    if balance_forward is None and previous_balance is not None:
        balance_forward = previous_balance + signed_payment_total

    # The requested metric is payment received divided by current charges.
    payment_coverage_current_charges = None
    if current_charges is not None and abs(current_charges) > 0.01:
        payment_coverage_current_charges = (
            payment_received_amount / abs(current_charges)
        ) * 100

    return {
        "Payment Instance": "Yes" if payment_instance_count else "No",
        "Payment Instance Count": payment_instance_count,
        "Payment Received Amount": payment_received_amount,
        "Payment Received Signed Amount": signed_payment_total,
        "Previous Balance": previous_balance,
        "Balance Forward": balance_forward,
        "Payment Coverage of Current Charges (%)":
            payment_coverage_current_charges,
        "Current Charges": current_charges,
        "Amount Past Due": amount_past_due,
        "Amount Due": amount_due,
    }


def extract_late_fee(raw_text):
    """Find an actual assessed late fee, excluding future-fee warning text."""
    fee_text = re.sub(
        r"To\s+avoid\s+late\s+payment\s+charges.*?(?:\.|$)",
        " ",
        raw_text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    flat_text = re.sub(r"\s+", " ", fee_text)

    match = re.search(
        r"(?<!Total\s)Late\s+Payment\s+Charges?\s*:?\s*"
        r"\$?\s*([\d,]+\.\d{2})",
        flat_text,
        re.IGNORECASE,
    )

    return float(match.group(1).replace(",", "")) if match else 0.0


def blank_result(note):
    """Return one report row with unavailable values left blank."""
    return {
        "Service Address": "",
        "Utility Type": "",
        "Bill Date": "",
        "Due Date": "",
        "Payment Instance": "",
        "Payment Instance Count": None,
        "Payment Received Amount": None,
        "Payment Received Signed Amount": None,
        "Previous Balance": None,
        "Balance Forward": None,
        "Payment Coverage of Current Charges (%)": None,
        "Current Charges": None,
        "Amount Past Due": None,
        "Amount Due": None,
        "Rate Plan / Tariff": "",
        "Late Fee Amount": None,
        "Late Fee Instance": "",
        "Result": "Needs review",
        "Notes": note,
    }


def parse_bill_pdf(pdf_bytes):
    """Extract the requested fields from one PDF bill."""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as pdf:
            raw_text = "\n".join(
                page.get_text("text") for page in pdf
            )

        if len(raw_text.strip()) < 40:
            return blank_result("PDF has little or no readable text.")

        flat_text = re.sub(r"\s+", " ", raw_text).strip()
        service = extract_service_and_rate(raw_text)
        balances = extract_account_balance(raw_text)
        fee = extract_late_fee(raw_text)

        missing_important = []
        if not service["Service Address"]:
            missing_important.append("Service Address")
        if balances["Current Charges"] is None:
            missing_important.append("Current Charges")
        if balances["Amount Due"] is None:
            missing_important.append("Amount Due")

        result = "Needs review" if missing_important else "Checked"
        notes = (
            "Could not confidently extract: " + ", ".join(missing_important)
            if missing_important
            else ""
        )

        return {
            "Service Address": service["Service Address"],
            "Utility Type": service["Utility Type"],
            "Bill Date": extract_bill_date(flat_text),
            "Due Date": extract_due_date(flat_text),
            **balances,
            "Rate Plan / Tariff": service["Rate Plan / Tariff"],
            "Late Fee Amount": fee,
            "Late Fee Instance": "Yes" if fee > 0.01 else "No",
            "Result": result,
            "Notes": notes,
        }

    except Exception as exc:
        return blank_result(f"Could not read PDF: {exc}")


# ---------------- REPORT OUTPUT ----------------

def save_report(results):
    """Replace the output workbook with exactly the requested report columns."""
    if not results:
        print("No bill rows to save.")
        return

    report = pd.DataFrame(results)

    # Ensure all requested columns exist, in the requested order.
    for column in OUTPUT_COLUMNS:
        if column not in report.columns:
            report[column] = None

    report = report[OUTPUT_COLUMNS]

    while True:
        try:
            with pd.ExcelWriter(
                OUTPUT_FILE,
                engine="openpyxl",
                mode="w",
            ) as writer:
                report.to_excel(
                    writer,
                    sheet_name="Bill Detail",
                    index=False,
                )

                sheet = writer.book["Bill Detail"]
                sheet.freeze_panes = "A2"
                sheet.auto_filter.ref = sheet.dimensions

                for cell in sheet[1]:
                    cell.font = Font(bold=True)

                for cells in sheet.columns:
                    letter = cells[0].column_letter
                    longest = max(
                        len(str(cell.value or ""))
                        for cell in cells
                    )
                    sheet.column_dimensions[letter].width = min(
                        max(longest + 2, 12),
                        42,
                    )

            print(f"Report saved/replaced: {OUTPUT_FILE.resolve()}")
            return

        except PermissionError:
            pause(
                f"\nClose {OUTPUT_FILE.name} in Excel, then press OK/Enter "
                "to retry saving."
            )


# ---------------- PORTAL AUTOMATION ----------------

async def find_search_box(page):
    for placeholder in ("Search Accounts", "Search accounts", "Search"):
        locator = page.get_by_placeholder(placeholder, exact=False)
        if await locator.count():
            return locator.first

    inputs = page.locator("input")
    for index in range(await inputs.count()):
        item = inputs.nth(index)

        try:
            if not await item.is_visible():
                continue

            placeholder = (
                await item.get_attribute("placeholder") or ""
            ).lower()
            aria_label = (
                await item.get_attribute("aria-label") or ""
            ).lower()

            if "search" in placeholder or "search" in aria_label:
                return item
        except Exception:
            pass

    return None


async def search_and_open_account(page, account_number):
    search_box = await find_search_box(page)

    if search_box is None:
        raise RuntimeError(
            "Could not find the account search box. "
            "Make sure Accounts Overview is open."
        )

    await search_box.fill(account_number)
    await page.wait_for_timeout(1000)

    wanted = account_digits(account_number)
    links = page.locator("a")

    for index in range(await links.count()):
        link = links.nth(index)

        try:
            if not await link.is_visible():
                continue

            candidate = account_digits(await link.inner_text())

            if candidate == wanted or (
                candidate and candidate.lstrip("0") == wanted.lstrip("0")
            ):
                await link.click()
                await page.wait_for_timeout(1000)
                return
        except Exception:
            pass

    raise RuntimeError(f"Could not find account link for {account_number}.")


async def open_billing_history(page):
    tab = page.get_by_text("Billing History", exact=True)

    if not await tab.count():
        raise RuntimeError("Could not find the Billing History tab.")

    await tab.first.click()
    await page.wait_for_timeout(1000)


async def visible_pdf_controls(page):
    matches = page.get_by_text("PDF", exact=True)
    controls = []

    for index in range(await matches.count()):
        item = matches.nth(index)
        try:
            if await item.is_visible():
                controls.append(item)
        except Exception:
            pass

    return controls


async def get_bill_history_row(control):
    try:
        return await control.evaluate(
            """
            (element) => {
              let node = element;

              for (let i = 0; i < 10 && node; i++, node = node.parentElement) {
                const text = (node.innerText || "")
                  .replace(/\\s+/g, " ")
                  .trim();

                if (text.includes("PDF") && text.includes("View")) {
                  return text;
                }
              }

              return (element.parentElement?.innerText || "")
                .replace(/\\s+/g, " ")
                .trim();
            }
            """
        )
    except Exception:
        return ""


async def download_pdf(page, control):
    """Click a PDF control and read the downloaded PDF."""
    try:
        async with page.expect_download(timeout=15_000) as event:
            await control.click()

        download = await event.value
        file_path = await download.path()

        if not file_path:
            return None

        pdf_bytes = Path(file_path).read_bytes()
        return pdf_bytes if pdf_bytes.startswith(b"%PDF") else None

    except Exception:
        return None


async def scroll_billing_history(page):
    await page.evaluate(
        """
        () => {
          window.scrollBy(0, 500);

          for (const element of document.querySelectorAll('*')) {
            const style = getComputedStyle(element);
            const canScroll =
              style.overflowY === 'auto' ||
              style.overflowY === 'scroll';

            if (
              canScroll &&
              element.scrollHeight > element.clientHeight + 20
            ) {
              element.scrollTop = Math.min(
                element.scrollTop + Math.max(250, element.clientHeight * 0.7),
                element.scrollHeight
              );
            }
          }
        }
        """
    )
    await page.wait_for_timeout(600)


async def return_to_accounts_overview(page):
    back = page.get_by_text("Back", exact=True)

    if await back.count():
        await back.first.click()
        await page.wait_for_timeout(1000)
        return

    overview = page.get_by_text("Accounts Overview", exact=True)

    if await overview.count():
        await overview.first.click()
        await page.wait_for_timeout(1000)


async def check_account(page, account_number, input_account_address):
    await search_and_open_account(page, account_number)
    await open_billing_history(page)

    results = []
    seen_bill_rows = set()
    empty_rounds = 0

    for scroll_round in range(MAX_SCROLL_ROUNDS):
        controls = await visible_pdf_controls(page)
        new_bills = 0

        for control_index, control in enumerate(controls):
            history_row = await get_bill_history_row(control)
            unique_key = history_row.strip() or f"{scroll_round}-{control_index}"

            if unique_key in seen_bill_rows:
                continue

            seen_bill_rows.add(unique_key)
            new_bills += 1

            if page.is_closed():
                raise RuntimeError("Browser page closed while checking bills.")

            pdf_bytes = await download_pdf(page, control)

            if pdf_bytes is None:
                bill_fields = blank_result(
                    "PDF did not download; this bill was not checked."
                )
            else:
                bill_fields = parse_bill_pdf(pdf_bytes)

            service_address = bill_fields.get("Service Address", "")

            # Use the optional Account Address value from the input workbook.
            # If that is blank, use the full service address parsed from the bill.
            account_address = input_account_address or service_address

            results.append({
                "Account Number": account_number,
                "Account Address": account_address,
                "Billing History Row": history_row,
                **bill_fields,
            })

        if new_bills == 0:
            empty_rounds += 1
        else:
            empty_rounds = 0

        if empty_rounds >= 3:
            break

        await scroll_billing_history(page)

    if not results:
        no_bills = blank_result(
            "No PDF controls were found; bills were not checked."
        )
        results.append({
            "Account Number": account_number,
            "Account Address": input_account_address,
            "Billing History Row": "",
            **no_bills,
        })

    return results


# ---------------- LOGIN / SESSION ----------------

async def login_and_save_state(playwright):
    """Open a visible browser once so the user can sign in, then save the session."""
    print("\n[Action Required] Opening browser for account authentication...")
    browser = await playwright.chromium.launch(
        headless=False, channel=BROWSER_CHANNEL
    )
    context = await browser.new_context()
    page = await context.new_page()

    try:
        await page.goto(
            START_URL,
            wait_until="commit",
            timeout=NAVIGATION_TIMEOUT_MS,
        )
    except Exception as exc:
        print(f"Portal navigation warning: {exc}")
        print("If needed, open the National Grid portal in the browser.")

    pause(
        "\nSign in in the browser and complete verification. "
        "Open Accounts Overview, leave the browser open, "
        "then press OK/Enter here when Accounts Overview is visible."
    )

    await context.storage_state(path=str(AUTH_FILE))
    print(f"Session saved to {AUTH_FILE}")
    await browser.close()


# ---------------- MAIN ----------------

async def main():
    print("=" * 60)
    print(" STARTING NATIONAL GRID BILL & LATE FEE PROCESSOR")
    print(" Loading input file and preparing automation environment...")
    print("=" * 60)

    accounts = load_accounts()
    all_results = []

    print(f"\nLoaded {len(accounts)} account(s) to process:")
    for number in accounts["Account Number"]:
        print(f"  • {number}")
    print("-" * 60)

    async with async_playwright() as playwright:
        # First run (or after the session expires): sign in with a visible browser.
        if not AUTH_FILE.exists():
            await login_and_save_state(playwright)

        print("\nLaunching browser engine...")
        browser = await playwright.chromium.launch(
            headless=HEADLESS, channel=BROWSER_CHANNEL
        )
        context = await browser.new_context(
            accept_downloads=True,
            storage_state=str(AUTH_FILE) if AUTH_FILE.exists() else None,
            user_agent=USER_AGENT,
            viewport={"width": 1440, "height": 900},
        )
        page = await context.new_page()
        page.set_default_navigation_timeout(NAVIGATION_TIMEOUT_MS)

        try:
            print(f"Navigating to {START_URL}...")
            try:
                await page.goto(
                    START_URL,
                    wait_until="domcontentloaded",
                    timeout=NAVIGATION_TIMEOUT_MS,
                )
            except Exception as exc:
                print(f"Portal navigation warning: {exc}")

            await page.wait_for_timeout(5000)

            # If the saved session expired or we landed on the wrong page,
            # the search box won't be there.
            if await find_search_box(page) is None:
                if HEADLESS:
                    AUTH_FILE.unlink(missing_ok=True)
                    raise RuntimeError(
                        "Saved session is no longer valid. "
                        "Run again (with the hidden option off) to sign in."
                    )

                pause(
                    "\nCould not find the account search box.\n"
                    "Sign in if needed and open Accounts Overview in the "
                    "browser, then press OK/Enter."
                )
                try:
                    await context.storage_state(path=str(AUTH_FILE))
                except Exception:
                    pass

            for idx, row in accounts.iterrows():
                account_number = row["Account Number"]
                input_address = str(
                    row.get("Account Address", "") or ""
                ).strip()

                print(f"\nProcessing Account [{idx + 1}/{len(accounts)}]: {account_number}")

                try:
                    account_results = await check_account(
                        page,
                        account_number,
                        input_address,
                    )
                    all_results.extend(account_results)
                    print(
                        f"  --> Extracted {len(account_results)} bill/result row(s)."
                    )

                except Exception as exc:
                    all_results.append({
                        "Account Number": account_number,
                        "Account Address": input_address,
                        "Billing History Row": "",
                        **blank_result(str(exc)),
                    })
                    print(f"  --> Needs review: {exc}")

                # Save progress after each account.
                save_report(all_results)

                if page.is_closed():
                    print("Browser page closed; stopping execution.")
                    break

                try:
                    await return_to_accounts_overview(page)
                except Exception as exc:
                    print(f"Could not return to Accounts Overview: {exc}")
                    break

        finally:
            try:
                await context.close()
            except Exception:
                pass

            try:
                await browser.close()
            except Exception:
                pass

    print("\n" + "=" * 60)
    print(f" PROCESS COMPLETED SUCCESSFULLY!")
    print(f" Final Report Location: {OUTPUT_FILE.resolve()}")
    print(
        " Note: Review rows marked 'Needs review' and verify sample values."
    )
    print("=" * 60)

    return all_results


if __name__ == "__main__":
    print("\nInitialising script execution...")
    asyncio.run(main())