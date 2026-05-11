import os
import time
import traceback
import pyodbc
import pandas as pd
import xlwings as xw
from rich import print
from dotenv import load_dotenv
from datetime import datetime
from pathlib import Path
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fc_utils import chrome, custom_functions, accounts, alert_utils, outlook
from fc_utils.config_utils import get_env, load_config_safe
from fc_utils.schedule_utils import run_on_schedule
from fc_utils.accounts import AMAZON_ACCOUNT_NAMES, AMAZON_URLS
from selenium.common.exceptions import TimeoutException

directory: str = os.getcwd()

load_dotenv()
username: str = os.getenv("AMZN_email")
password: str = os.getenv("AMZN_pass")
sender_email: str = os.getenv("SENDER_EMAIL", "")
to_email: list[str] = [e.strip() for e in os.getenv("TO_EMAIL", "").split(",") if e.strip()]
cc_email: list[str] = [e.strip() for e in os.getenv("CC_EMAIL", "").split(",") if e.strip()]
user_data_dir: str = get_env("CHROME_USER_DATA_DIR", required=True)
table_negative: str = os.getenv("DB_TABLE_NEGATIVE", "FedManNegative")
table_positive: str = os.getenv("DB_TABLE_POSITIVE", "FedManPositive")
for _t in (table_negative, table_positive):
    if not _t.replace("_", "").isalnum():
        raise ValueError(f"Invalid table name: {_t!r}")

_accounts_cfg = load_config_safe(Path.cwd() / "config" / "accounts.json")
_feedback_sheets: dict[str, int] = _accounts_cfg.get("feedback_manager_sheets", {})
_rating_cells: dict[str, str] = _accounts_cfg.get("feedback_manager_rating_cells", {})

body = """
Good morning,<br><br>
Please find attached Feedback Manager report updated for today.<br><br>
If any questions, please let me know.<br><br>
Thanks,<br><br>
"""


def last_update(cursor: object, account: str) -> str:
    """Return the most recent UpdatedAt timestamp for the given account in the negative feedback table.

    Args:
        cursor (object): Active pyodbc cursor.
        account (str): Account display name to filter by.

    Returns:
        str: Timestamp string of the most recent record, or 'None'.
    """
    cursor.execute(
        f"SELECT MAX(UpdatedAt) AS RecentDate FROM {table_negative} WHERE Account = ?",
        (account,)
    )
    return str(cursor.fetchone()[0])


def find_order(cursor: object, order: str) -> int:
    """Return how many times the given Order ID exists in the positive feedback table.

    Args:
        cursor (object): Active pyodbc cursor.
        order (str): Amazon Order ID to look up.

    Returns:
        int: Row count (0 if not found).
    """
    cursor.execute(
        f"SELECT COUNT(OrderID) AS OrderID FROM {table_positive} WHERE OrderID = ?",
        (order,)
    )
    return cursor.fetchone()[0]


def request_report(driver: object, cursor: object, root: str, week_day: str) -> None:
    """Download the negative feedback report and insert rows into the database.

    Args:
        driver (object): Active SeleniumBase WebDriver instance.
        cursor (object): Active pyodbc cursor.
        root (str): Account display name used for logging and DB insert.
        week_day (str): Full weekday name; 'Monday' requests a 7-day report, otherwise 1-day.
    """
    driver.get("https://sellercentral.amazon.com/feedback-manager/index.html#/report")
    driver.switch_to_window(0)

    if week_day == "Monday":
        print(f"[cyan][INFO][/cyan] Getting [cyan]{root}[/cyan] 7-days Feedback Manager Report.")
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".request_report_btn > kat-button:nth-child(1) > button:nth-child(1)"))).click()
    else:
        print(f"[cyan][INFO][/cyan] Getting [cyan]{root}[/cyan] 1-day Feedback Manager Report.")
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "kat-table-row.request_report_row:nth-child(2) > kat-table-cell:nth-child(2) > kat-dropdown:nth-child(1) > div:nth-child(1) > div:nth-child(1)"))).click()
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "1D0"))).click()
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".request_report_btn > kat-button:nth-child(1) > button:nth-child(1)"))).click()

    time.sleep(3)
    report_status: str = WebDriverWait(driver, 60).until(EC.presence_of_element_located((
        By.XPATH,
        "/html/body/div/div[2]/div/my-app/div/div/report/div/kat-table[2]/kat-table-body/kat-table-row[1]/kat-table-cell[4]"
    ))).text

    while report_status != "Ready":
        if report_status == "No Data":
            print(f"[cyan][INFO][/cyan] {report_status}. Moving to Positive Feedback.")
            return

        print(f"[cyan][INFO][/cyan] {report_status}. Waiting for report to be ready.")

        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div/my-app/div/div/report/div/kat-alert[3]")))
            time.sleep(10)
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".report_refresh_btn > button:nth-child(1)"))).click()
        except TimeoutException:
            pass

        report_status = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div/my-app/div/div/report/div/kat-table[2]/kat-table-body/kat-table-row[1]/kat-table-cell[4]"))).text

    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div/my-app/div/div/report/div/kat-table[2]/kat-table-body/kat-table-row[1]/kat-table-cell[5]/kat-link"))).click()
    time.sleep(3)

    print("[cyan][INFO][/cyan] File downloaded successfully. Loading to database.")
    file_path: str = f"{directory}/downloaded_files/report.txt"

    while True:
        try:
            df = pd.read_csv(
                file_path,
                sep="\t",
                header=0,
                dtype={
                    "Date": str,
                    "Rating": str,
                    "Comments": str,
                    "Response": str,
                    "Order ID": str,
                    "Rater Email": str
                }
            )
            break
        except FileNotFoundError:
            print("[bold red][ERROR][/bold red] Failed to read the downloaded file. Waiting 5 seconds and trying again.")
            time.sleep(5)

    df = df.where(pd.notnull(df), None)
    df = df.dropna(subset=["Order ID"])
    if df.empty:
        print(f"[yellow][WARNING][/yellow] No valid rows to insert for [cyan]{root}[/cyan]. Skipping.")
        os.remove(file_path)
        return

    df.insert(0, "Account", root)
    df["UpdatedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for index, row in df.iterrows():
        try:
            cursor.execute(
                f"INSERT INTO {table_negative} (Account, Date, Rating, Comments, Response, OrderID, RaterEmail, UpdatedAt) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (row["Account"], row["Date"], row["Rating"], row["Comments"], row["Response"], row["Order ID"], row["Rater Email"], row["UpdatedAt"])
            )
        except pyodbc.Error:
            print(f"[bold red][ERROR][/bold red] [pyodbc.Error] Error inserting row {index + 1}:\n\n{row}", sep="\n\n")
            traceback.print_exc()
            raise RuntimeError(f"Insert failed at row {index + 1}. See traceback above.")

    cursor.connection.commit()
    print(f"[cyan][INFO][/cyan] Data inserted successfully in table [cyan]{table_negative}[/cyan].")
    os.remove(file_path)


def feedback_manager_ratings(driver: object, account: str, rating_sh: object) -> None:
    """Scrape the feedback ratings summary table and write it to the workbook.

    Args:
        driver (object): Active SeleniumBase WebDriver instance.
        account (str): Account key used to look up the target cell range.
        rating_sh (object): xlwings Sheet object for the ratings sheet.
    """
    print(f"[cyan][INFO][/cyan] Getting [cyan]{AMAZON_ACCOUNT_NAMES[account]}[/cyan] feedback ratings.")
    driver.get("https://sellercentral.amazon.com/feedback-manager/index.html#/")
    driver.switch_to_window(0)

    table_loaded = False
    while not table_loaded:
        try:
            raw_text: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((
                By.XPATH,
                "/html/body/div/div[2]/div/my-app/div/div/home/div/feedback-summary/kat-table"
            ))).text.split("\n")

            pos: list[str] = raw_text[2].split(" ")
            neu: list[str] = raw_text[4].split(" ")
            neg: list[str] = raw_text[6].split(" ")
            totals: list[str] = raw_text[8].split(" ")

            raw_data = [
                int(pos[0]) / 100, int(pos[2]) / 100, int(pos[4]) / 100, int(pos[6]) / 100,
                pos[1].replace("%(", "").replace(")", ""), pos[3].replace("%(", "").replace(")", ""), pos[5].replace("%(", "").replace(")", ""), pos[7].replace("%(", "").replace(")", ""),
                int(neu[0]) / 100, int(neu[2]) / 100, int(neu[4]) / 100, int(neu[6]) / 100,
                neu[1].replace("%(", "").replace(")", ""), neu[3].replace("%(", "").replace(")", ""), neu[5].replace("%(", "").replace(")", ""), neu[7].replace("%(", "").replace(")", ""),
                int(neg[0]) / 100, int(neg[2]) / 100, int(neg[4]) / 100, int(neg[6]) / 100,
                neg[1].replace("%(", "").replace(")", ""), neg[3].replace("%(", "").replace(")", ""), neg[5].replace("%(", "").replace(")", ""), neg[7].replace("%(", "").replace(")", ""),
                totals[0], totals[1], totals[2], totals[3],
            ]

            num_columns = 4
            rows = [raw_data[i:i + num_columns] for i in range(0, len(raw_data), num_columns)]
            df = pd.DataFrame(rows)

            rating_sh.range(_rating_cells[account]).value = df.values
            table_loaded = True

        except Exception:
            print("[bold red][ERROR][/bold red] Error loading website. Retrying.")
            driver.refresh()


def feedback_manager_comments(driver: object, cursor: object, root: str) -> None:
    """Retrieve positive feedback comments and insert new ones into the database.

    Args:
        driver (object): Active SeleniumBase WebDriver instance.
        cursor (object): Active pyodbc cursor.
        root (str): Account display name used for DB insert and logging.
    """
    try:
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#katal-id-1 > button"))).click()
    except TimeoutException:
        pass

    print(f"[cyan][INFO][/cyan] Getting Positive Feedback comments for [cyan]{root}[/cyan].")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".filter-tabs > kat-tabs:nth-child(1) > kat-tab-pane:nth-child(1) > kat-tab-header:nth-child(2)"))).click()
    WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "kat-tab.tab-selected > feedback-list:nth-child(1) > kat-table:nth-child(2) > kat-spinner:nth-child(2)")))

    row = 1
    matched_order = False
    while not matched_order:
        raw_data: list[str] = WebDriverWait(driver, 60).until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            f"kat-tab.tab-selected > feedback-list:nth-child(1) > kat-table:nth-child(2) > kat-table-body:nth-child(2) > kat-table-row:nth-child({row})"
            "#feedback-details"
        ))).text.split("\n")
        raw_data = [item for item in raw_data if item not in ["Choose one", ""]]

        clean_data = []
        clean_data.extend(raw_data[0].split(" "))
        try:
            clean_data.append(raw_data[1])
        except IndexError:
            clean_data.append("N/A")

        current_order: str = clean_data[2]

        if not current_order:
            print(f"[yellow][WARNING][/yellow] Skipping row {row}: empty Order ID.")
            row += 1
            continue

        if find_order(cursor, current_order) != 0:
            print(f"[cyan][INFO][/cyan] Order {current_order} is already in the database.")
            matched_order = True
            continue

        clean_data.insert(0, root)
        df = pd.DataFrame([clean_data], columns=["Account", "Date", "Rating", "Order ID", "Comments"])
        df = df.where(pd.notnull(df), None)

        for index, df_row in df.iterrows():
            try:
                cursor.execute(
                    f"INSERT INTO {table_positive} (Account, Date, Rating, OrderID, Comments) VALUES (?, ?, ?, ?, ?)",
                    (df_row["Account"], df_row["Date"], df_row["Rating"], df_row["Order ID"], df_row["Comments"])
                )
            except pyodbc.Error:
                print(f"[bold red][ERROR][/bold red] [pyodbc.Error] Error inserting current row:", df_row, sep="\n\n")
                traceback.print_exc()
                raise RuntimeError(f"Insert failed at row {index + 1}. See traceback above.")

        cursor.connection.commit()
        print(f"[cyan][INFO][/cyan] Order {current_order} data inserted successfully in table [cyan]{table_positive}[/cyan].")

        row += 1

        if row == 21:
            print("[cyan][INFO][/cyan] Moving to next page.")
            pagination: str = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "kat-tab.tab-selected > feedback-list:nth-child(1) > kat-pagination:nth-child(1)"))).text
            pagination = len(pagination.split("\n"))

            next_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"kat-tab.tab-selected > feedback-list:nth-child(1) > kat-pagination:nth-child(1) > ul:nth-child(1) > li:nth-child({pagination})")))
            driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
            next_btn.click()

            time.sleep(3)
            row = 1


def main() -> None:
    """Run the full Feedback Manager scrape, update workbook, and send email report."""
    driver = None
    try:
        conn = custom_functions.sql_connection("Amazon")
        cursor = conn.cursor()

        week_day: str = datetime.now().strftime("%A")
        curr_date: str = datetime.now().strftime("%Y-%m-%d")
        date_str: str = datetime.now().strftime("%m/%d/%Y")

        feedback_man_path: str = f"{directory}/Amazon/Reports/Feedback-Manager.xlsm"
        all_items_path: str = f"{directory}/Amazon/Reports/All Items.xlsm"

        driver = chrome.start_browser(user_data_dir, "Default", headless=True)

        print("[cyan][INFO][/cyan] Opening Feedback Manager workbook.")
        feedback_man_wb = xw.Book(feedback_man_path)
        rating_sh = feedback_man_wb.sheets(7)
        custom_functions.update_directory(feedback_man_wb)

        refresh_all = feedback_man_wb.macro("Module1.RefreshAll")
        sort_all = feedback_man_wb.macro("Module1.SortAll")
        sort_status = feedback_man_wb.macro("Module1.SortStatus")

        for account, url in AMAZON_URLS.items():
            root = AMAZON_ACCOUNT_NAMES[account]

            print(f"[cyan][INFO][/cyan] Navigating to [cyan]{root}[/cyan] account.")
            driver.get(url)
            time.sleep(2)
            driver.switch_to_window(0)

            try:
                code = None
                while not code:
                    code = accounts.amazon_login(driver, username, username, password)
                    if not code:
                        print("[bold red][ERROR][/bold red] Failed to log in to Amazon. Trying again.")
                        driver.get(url)
                        driver.switch_to_window(0)
            except TimeoutException:
                pass

            print("[cyan][INFO][/cyan] Checking if Feedback Manager's files have been downloaded.")
            if last_update(cursor, root).split(" ")[0] == curr_date:
                print(f"[cyan][INFO][/cyan] Feedback Manager's files for [cyan]{root}[/cyan] has been downloaded today.")
            else:
                request_report(driver, cursor, root, week_day)

            feedback_manager_ratings(driver, account, rating_sh)
            feedback_manager_comments(driver, cursor, root)

        all_items_wb = xw.Book(all_items_path)
        custom_functions.update_directory(all_items_wb)

        print("[cyan][INFO][/cyan] Opening [cyan]All Items[/cyan] and refreshing queries.")
        queries_address = all_items_wb.macro("Module2.QueriesAddress")
        queries_address()
        time.sleep(5)

        refresh_fman = all_items_wb.macro("Module1.RefreshFMan")
        refresh_fman()
        time.sleep(10)

        fman_sh = all_items_wb.sheets(5)
        first_order: int = custom_functions.first_empty_row(fman_sh, "D", "B3")
        last_order: int = int(fman_sh.range(f"B{fman_sh.cells.last_cell.row}").end("up").row)

        if first_order > last_order:
            print("[cyan][INFO][/cyan] No new orders to process.")
            orders = None
            all_items_wb.save()
            all_items_wb.close()
        elif first_order == last_order:
            orders = [fman_sh.range(f"B{first_order}:C{last_order}").value]
        else:
            orders = fman_sh.range(f"B{first_order}:C{last_order}").value

        name_to_url: dict[str, str] = {v: AMAZON_URLS[k] for k, v in AMAZON_ACCOUNT_NAMES.items()}

        if orders is not None:
            for order in orders:
                print(f"[cyan][INFO][/cyan] Navigating to [cyan]{order[0]}[/cyan] account.")
                driver.get(name_to_url[order[0]])
                driver.switch_to_window(0)

                print(f"[cyan][INFO][/cyan] Getting order #{order[1]} details.")
                driver.get(f"https://sellercentral.amazon.com/orders-v3/order/{order[1]}")
                driver.switch_to_window(0)

                WebDriverWait(driver, 60).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "#apploading > span")))

                order_raw: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "myo-list-orders-product-name-cell"))).text.split("\n")
                asin: str = order_raw[1].replace("ASIN: ", "")
                sku: str = order_raw[2].replace("SKU: ", "")

                fman_sh.range(f"D{first_order}").value = [asin, sku]
                first_order += 1

            print("[cyan][INFO][/cyan] Refreshing queries.")
            refresh_fman()
            time.sleep(10)
            all_items_wb.save()
            all_items_wb.close()

        print("[cyan][INFO][/cyan] Refreshing all queries on [cyan]Feedback Manager[/cyan] workbook.")
        refresh_all()
        time.sleep(3)

        sort_status()
        time.sleep(5)

        for account, url in AMAZON_URLS.items():
            root = AMAZON_ACCOUNT_NAMES[account]
            curr_sh = feedback_man_wb.sheets(_feedback_sheets[account])

            first_order = custom_functions.first_empty_row(curr_sh, "J", "B3")
            last_order = int(curr_sh.range(f"E{curr_sh.cells.last_cell.row}").end("up").row)

            if first_order > last_order:
                print(f"[cyan][INFO][/cyan] No new orders to process in [cyan]{root}[/cyan] sheet.")
                continue
            elif first_order == last_order:
                orders = [curr_sh.range(f"E{first_order}:E{last_order}").value]
            else:
                orders = curr_sh.range(f"E{first_order}:E{last_order}").value

            print(f"[cyan][INFO][/cyan] Navigating to [cyan]{root}[/cyan] account.")
            driver.get(url)
            time.sleep(2)
            driver.switch_to_window(0)

            for order in orders:
                print(f"[cyan][INFO][/cyan] Retrieving order #{order} info.")
                driver.get(f"https://sellercentral.amazon.com/messaging/inbox-v3?fi=search&ss={order}")
                driver.switch_to_window(0)

                try:
                    status: str = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "case-status-label"))).text
                    if status.startswith("Resolved on "):
                        status = status.replace("Resolved on ", "Last response on ")
                except TimeoutException:
                    print(f"[cyan][INFO][/cyan] No messages on order #{order}. Moving forward.")
                    status = "N/A"

                curr_sh.range(f"J{first_order}").value = status
                first_order += 1

        driver.quit()
        driver = None

        print("[cyan][INFO][/cyan] Sorting all tables, saving and closing workbook.")
        sort_all()
        time.sleep(3)
        feedback_man_wb.save()
        feedback_man_wb.close()

        print("[cyan][INFO][/cyan] Loading workbook and sending email.")
        time.sleep(60)
        outlook.send_email(
            account=sender_email,
            subject=f"Feedback Manager - {date_str}",
            body=body,
            to=to_email,
            cc=cc_email,
            attachments=[feedback_man_path],
            show=True,
            send=True
        )

        print("[cyan][INFO][/cyan] Email has been sent.")

    except KeyboardInterrupt:
        raise
    except Exception:
        alert_utils.handle_crash(driver, traceback.format_exc(), "Feedback Manager")
    finally:
        if driver:
            driver.quit()


run_on_schedule(main, hour=10, minute=0, day_of_week="mon-fri")
