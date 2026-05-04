import os
import time
import pyodbc
import ctypes
import traceback
import pandas as pd
import xlwings as xw
from rich import print
from dotenv import load_dotenv
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fc_utils import chrome, custom_functions, accounts, outlook
from selenium.common.exceptions import TimeoutException, SessionNotCreatedException

###############################################################################################################################################
#Get current Windows user name and working directory
win_user: str = os.getlogin()
directory: str = os.getcwd()

#Get Seller Central credentials from environment
load_dotenv()
username: str = os.getenv("AMZN_email")
password: str = os.getenv("AMZN_pass")

#Set Chrome User Data Directory
#user_data_dir: str = f"C:/Users/{win_user}/AppData/Local/Google/Chrome/User Data"
user_data_dir: str = f"C:/ChromeAutomationProfile"

#Create SQL Database connection
conn = custom_functions.SQLConnection("Amazon")
cursor = conn.cursor()

##################################################################################################################################################
# Create the body of the email
body = """
Good morning,<br><br>
Please find attached Feedback Manager report updated for today.<br><br>
If any questions, please let me know.<br><br>
Thanks,<br><br>
"""

###############################################################################################################################################
#Set workbook properties
AllItems: str = f"{directory}/Amazon/Reports/All Items.xlsm"
FeedbackMan: str = f"{directory}/Amazon/Reports/Feedback-Manager.xlsm"

###############################################################################################################################################
#Create Amazon accounts links
Accounts = accounts.Amazon()

##################################################################################################################################################
def seconds_until_target(TargetTime: str):
    #Calculate the number of seconds until the target time
    now: datetime = datetime.now()
    TargetTime = datetime.strptime(TargetTime, "%H:%M:%S").replace(year=now.year, month=now.month, day=now.day)

    if TargetTime < now:
        TargetTime += timedelta(days=1)

    return (TargetTime - now).total_seconds()

##################################################################################################################################################
def ShouldRun() -> bool:
    #Check if today is not Saturday or Sunday
    today: str = datetime.now().strftime("%A")
    
    return today not in ["Saturday", "Sunday"]

###############################################################################################################################################
def LastUpdate() -> str:
    """
    Confirms if the Negative Feedback reports have been downloaded today.\n
    Returns the latest date in the table database.

    returns:
        The latest date.
    """
    #Get the most recent date from the "Date" column in the database
    cursor.execute(
        f"""
        SELECT MAX(UpdatedAt) AS RecentDate
        FROM FedManNegative
        WHERE Account = '{root}';
        """
    )

    recent_date = cursor.fetchone()[0]

    return str(recent_date)

###############################################################################################################################################
def FindOrder(order) -> str:
    """
    Confirm if the pending current Order ID is already stored in the database.

    Returns:
        Returns the quantity of matched items.
    """
    #Get the most recent date from the "Date" column in the database
    cursor.execute(
        f"""
        SELECT COUNT(OrderID) AS OrderID
        FROM FedManPositive
        WHERE OrderID = '{order}';
        """
    )

    order_count = cursor.fetchone()[0]

    return order_count

##################################################################################################################################################
def RequestReport() -> None:
    """
    Download Feedback Manager's reports based on the current day and time.
    """
    #Navigate to Feedback Manager Report Dashboard
    driver.get("https://sellercentral.amazon.com/feedback-manager/index.html#/report")
    driver.switch_to_window(0)

    if WeekDay == "Monday":
        print(f"'[INFO]' Getting '{root}' 7-days Feedback Manager Report.")
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".request_report_btn > kat-button:nth-child(1) > button:nth-child(1)"))).click()
    else:
        print(f"'[INFO]' Getting '{root}' 1-day Feedback Manager Report.")
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "kat-table-row.request_report_row:nth-child(2) > kat-table-cell:nth-child(2) > kat-dropdown:nth-child(1) > div:nth-child(1) > div:nth-child(1)"))).click()
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "1D0"))).click()
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".request_report_btn > kat-button:nth-child(1) > button:nth-child(1)"))).click()
    
    #Wait for the report to be generated
    time.sleep(3)
    report_status = WebDriverWait(driver, 60).until(EC.presence_of_element_located((
        By.XPATH,
        "/html/body/div/div[2]/div/my-app/div/div/report/div/kat-table[2]/kat-table-body/kat-table-row[1]/kat-table-cell[4]"
    ))).text
##sc-content-container > div > my-app > div > div > report > div > kat-table:nth-child(11) > kat-table-body > kat-table-row:nth-child(1) > kat-table-cell:nth-child(4)
    while report_status != "Ready":

        if report_status == "No Data":
            print(f"'[INFO]' {report_status}. Moving to Positive Feedback.")
            break

        print(f"'[INFO]' {report_status}. Waiting for report to be ready.")

        #Click refresh button if it takes too long
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div/my-app/div/div/report/div/kat-alert[3]")))
            time.sleep(10)
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".report_refresh_btn > button:nth-child(1)"))).click()
        except:
            pass

        report_status = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div/my-app/div/div/report/div/kat-table[2]/kat-table-body/kat-table-row[1]/kat-table-cell[4]"))).text
    
    #Download file and move to correct location
    if report_status == "No Data":
        pass
    else:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div/my-app/div/div/report/div/kat-table[2]/kat-table-body/kat-table-row[1]/kat-table-cell[5]/kat-link"))).click()
        time.sleep(3)

        #Set the file directory
        print("'[INFO]' File downloaded successfully. Loading to database.")
        file_path = f"{directory}/downloaded_files/report.txt"

        reading_file = True
        while reading_file:
            #If it fails to read the file, wait 5 seconds and try again
            try:
                #Read the downloaded file into a DataFrame
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
                    })

                reading_file = False

            except FileNotFoundError:
                print("'[FileNotFoundError]' Failed to read the downloaded file. Waiting 5 seconds and trying again.")
                time.sleep(5)

        #Replace NaN values with None for all columns
        df = df.where(pd.notnull(df), None)

        #Insert a column to the beginning of the table with the name of the account
        df.insert(0, "Account", root)

        #Insert new column with timestamp
        df["UpdatedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        #Insert data into the SQL Database
        for index, row in df.iterrows():
            try:
                cursor.execute(
                    f"""
                    INSERT INTO FedManNegative (Account, Date, Rating, Comments, Response, OrderID, RaterEmail, UpdatedAt)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    row["Account"],
                    row["Date"],
                    row["Rating"],
                    row["Comments"],
                    row["Response"],
                    row["Order ID"],
                    row["Rater Email"],
                    row["UpdatedAt"]
                )
            except pyodbc.Error:
                print(f"'[pyodbc.Error]' Error inserting row {index+1}:\n\n{row}", sep="\n\n")
                traceback.print_exc()
                quit()

        # Commit the transaction
        conn.commit()
        print(f"'[INFO]' Data inserted successfully in table 'FedManNegative'.")

        #Delete the downloaded file
        os.remove(file_path)

##################################################################################################################################################
def FeedbackManagerRatings() -> None:
    """
    Retrieves Feedback Manager's feedback ratings table.
    """
    #Move to Feedback Manager main Dashboard
    print(f"'[INFO]' Getting '{root}' feedback ratings.")
    driver.get("https://sellercentral.amazon.com/feedback-manager/index.html#/")
    driver.switch_to_window(0)

    table_range = False
    while not table_range:
        try:
            FeedbackRatingRaw: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div/my-app/div/div/home/div/feedback-summary/kat-table"))).text.split("\n")
            
            #Positive feedback
            PosFeedback: list[str] = FeedbackRatingRaw[2].split(" ")

            Pos30: float = int(PosFeedback[0]) / 100
            PosCount30: str = PosFeedback[1].replace("%(", "").replace(")", "")

            Pos90: float = int(PosFeedback[2]) / 100
            PosCount90: str = PosFeedback[3].replace("%(", "").replace(")", "")

            Pos365: float = int(PosFeedback[4]) / 100
            PosCount365: str = PosFeedback[5].replace("%(", "").replace(")", "")

            PosLifetime = PosFeedback[6]
            PosLifetime = int(PosLifetime) / 100
            PosCountLifetime: str = PosFeedback[7].replace("%(", "").replace(")", "")

            #Neutral feedback
            NeuFeedback: list[str] = FeedbackRatingRaw[4].split(" ")

            Neu30: float = int(NeuFeedback[0]) / 100
            NeuCount30: str = NeuFeedback[1].replace("%(", "").replace(")", "")

            Neu90: float = int(NeuFeedback[2]) / 100
            NeuCount90: str = NeuFeedback[3].replace("%(", "").replace(")", "")

            Neu365: float = int(NeuFeedback[4]) / 100
            NeuCount365: str = NeuFeedback[5].replace("%(", "").replace(")", "")

            NeuLifetime: float = int(NeuFeedback[6]) / 100
            NeuCountLifetime: str = NeuFeedback[7].replace("%(", "").replace(")", "")

            #Negative feedback
            NegFeedback: list[str] = FeedbackRatingRaw[6].split(" ")

            Neg30: float = int(NegFeedback[0]) / 100
            NegCount30: str = NegFeedback[1].replace("%(", "").replace(")", "")

            Neg90: float = int(NegFeedback[2]) / 100
            NegCount90: str = NegFeedback[3].replace("%(", "").replace(")", "")

            Neg365: float = int(NegFeedback[4]) / 100
            NegCount365: str = NegFeedback[5].replace("%(", "").replace(")", "")

            NegLifetime: float = int(NegFeedback[6]) / 100
            NegCountLifetime: str = NegFeedback[7].replace("%(", "").replace(")", "")

            #Total Count
            TotalCount: list[str] = FeedbackRatingRaw[8].split(" ")
            
            Count30: str = TotalCount[0]
            Count90: str = TotalCount[1]
            Count365: str = TotalCount[2]
            CountLifetime: str = TotalCount[3]

            table_range = True

        except:
            print("'[ERROR]' Error loading website. Retrying.")
            driver.refresh()

    RawData = [Pos30, Pos90, Pos365, PosLifetime, 
                PosCount30, PosCount90, PosCount365, PosCountLifetime, 
                Neu30, Neu90, Neu365, NeuLifetime, 
                NeuCount30, NeuCount90, NeuCount365, NeuCountLifetime,
                Neg30, Neg90, Neg365, NegLifetime,
                NegCount30, NegCount90, NegCount365, NegCountLifetime,
                Count30, Count90, Count365, CountLifetime]
    
    num_columns = 4
    rows = [RawData[i:i+num_columns] for i in range(0, len(RawData), num_columns)] #Split raw data into rows
    df = pd.DataFrame(rows[0:])

    if account == "AccountA":
        RatingSh.range("C6").value = df.values

    elif account == "AccountB":
        RatingSh.range("I6").value = df.values

    elif account == "AccountD":
        RatingSh.range("C17").value = df.values

    elif account == "AccountC":
        RatingSh.range("I17").value = df.values

    elif account == "Apple":
        RatingSh.range("C28").value = df.values

    elif account == "AccountE":
        RatingSh.range("I28").value = df.values

##################################################################################################################################################
def FeedbackManagerComments() -> None:
    """
    Retrieves Feedback Manager's positive feedback comments.
    """
    #Close pop-up window if opened
    try:
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#katal-id-1 > button"))).click()
    except TimeoutException:
        pass

    #Move to Feedback Manager main Dashboard
    print(f"'[INFO]' Getting Positive Feedback comments for '{root}'.")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".filter-tabs > kat-tabs:nth-child(1) > kat-tab-pane:nth-child(1) > kat-tab-header:nth-child(2)"))).click()
    WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "kat-tab.tab-selected > feedback-list:nth-child(1) > kat-table:nth-child(2) > kat-spinner:nth-child(2)")))

    row = 1
    MatchedOrder = False
    while not MatchedOrder:
        RawData: str = WebDriverWait(driver, 60).until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            f"kat-tab.tab-selected > feedback-list:nth-child(1) > kat-table:nth-child(2) > kat-table-body:nth-child(2) > kat-table-row:nth-child({row})"
            "#feedback-details"
        ))).text.split("\n")
        RawData = [item for item in RawData if item not in ["Choose one", ""]]

        #Clean Data and assign the third item to be the First Order
        CleanData = []
        CleanData.extend(RawData[0].split(" "))
        try:
            CleanData.append(RawData[1])
        except IndexError:
            CleanData.append("N/A")

        CurrentOrder = CleanData[2]

        #Check if the current order is already in the database
        if FindOrder(CurrentOrder) != 0:
            print(f"'[INFO]' Order {CurrentOrder} is already in the database.")
            MatchedOrder = True
            continue

        #Insert the current account name to the first position of the row
        CleanData.insert(0, root)

        #Write review details to a new dataframe
        data_to_insert = [CleanData]
        df = pd.DataFrame(data_to_insert, columns=["Account", "Date", "Rating", "Order ID", "Comments"])

        #Replace NaN values with None for all columns
        df = df.where(pd.notnull(df), None)

        #Insert data into the SQL Database
        for index, df_row in df.iterrows():
            try:
                cursor.execute(
                    f"""
                    INSERT INTO FedManPositive (Account, Date, Rating, OrderID, Comments)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    df_row["Account"],
                    df_row["Date"],
                    df_row["Rating"],
                    df_row["Order ID"],
                    df_row["Comments"]
                )
            except pyodbc.Error:
                print(f"'[pyodbc.Error]' Error inserting current row:", df_row, sep="\n\n")
                traceback.print_exc()
                quit()

        # Commit the transaction
        conn.commit()
        print(f"'[INFO]' Order {CurrentOrder} data inserted successfully in table 'FedManPositive'.")

        row += 1
    
        if row == 21:
            print("'[INFO]' Moving to next page.")
            pagination: str = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "kat-tab.tab-selected > feedback-list:nth-child(1) > kat-pagination:nth-child(1)"))).text
            pagination = len(pagination.split("\n"))

            nextBtn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"kat-tab.tab-selected > feedback-list:nth-child(1) > kat-pagination:nth-child(1) > ul:nth-child(1) > li:nth-child({pagination})")))
            driver.execute_script("arguments[0].scrollIntoView(true);", nextBtn)
            nextBtn.click()
            
            time.sleep(3)
            row = 1

##################################################################################################################################################
#Ask the user if they want to start the process now
BtnPressed = ctypes.windll.user32.MessageBoxW(
    0,
    "Do you want to start the script now?",
    "Feedback Manager",
    4 | 0x20
)

while True:
    #Time to start
    StartTime = "10:00:00"
    StartHour = int(StartTime.split(":")[0])
    StartMin = StartTime.split(":")[1]

    if ShouldRun():
        SleepTime = seconds_until_target(StartTime)
        nowHour = int(datetime.now().strftime("%H"))
        today: str = datetime.now().strftime("%A")
        tomorrow: str = custom_functions.tomorrow()

        #If the user pressed "Yes", then start the process
        if BtnPressed == 7:
            if tomorrow in ["Saturday", "Sunday"]:
                if nowHour >= StartHour:
                    if StartHour > 12:
                        print(f"'[INFO]' Feedback Manager dashboard will be updated on Monday at {StartHour - 12}:{StartMin} PM.")
                    else:
                        print(f"'[INFO]' Feedback Manager dashboard will be updated on Monday at {StartHour}:{StartMin} AM.")
                else:
                    if StartHour > 12:
                        print(f"'[INFO]' Feedback Manager dashboard will be updated today at {StartHour - 12}:{StartMin} PM.")
                    else:
                        print(f"'[INFO]' Feedback Manager dashboard will be updated today at {StartHour}:{StartMin} AM.")
            
            else:
                if nowHour >= StartHour:
                    if StartHour > 12:
                        print(f"'[INFO]' Feedback Manager dashboard will be updated tomorrow {tomorrow} at {StartHour - 12}:{StartMin} PM.")
                    else:
                        print(f"'[INFO]' Feedback Manager dashboard will be updated tomorrow {tomorrow} at {StartHour}:{StartMin} AM.")
                else:
                    if StartHour > 12:
                        print(f"'[INFO]' Feedback Manager dashboard will be updated today at {StartHour}:{StartMin - 12} PM.")
                    else:
                        print(f"'[INFO]' Feedback Manager dashboard will be updated today at {StartHour}:{StartMin} AM.")

            #Sleep until just before the Start time
            time.sleep(max(SleepTime - 1, 0))

            #Loop to ensure that we catch the exact time
            while datetime.now().strftime("%H:%M:%S") != StartTime:
                time.sleep(0.5)

            if tomorrow in ["Saturday", "Sunday"] and nowHour > StartHour:
                continue

        #Reset the value of the button
        BtnPressed = 7

        #Get weekday and full date
        WeekDay: str = datetime.now().strftime("%A")
        currDate: str = datetime.now().strftime("%Y-%m-%d")
        Date: str = datetime.now().strftime("%m/%d/%Y")

        ###############################################################################################################################################
        #Initialize Chrome
        opening_browser = True
        while opening_browser:
            try:
                driver: object = chrome.start_browser(
                    user_data_dir,
                    "Default",
                    headless=True
                )
                opening_browser = False

            except (SessionNotCreatedException, RuntimeError):
                print("'[ERROR]' Failed to open Chrome. It seems Chrome was already open. Killing the application and retrying.")
                custom_functions.kill_app("chrome")
                time.sleep(5)
            
            except PermissionError:
                print("'[ERROR]' Failed to open Chrome. It seems 'uc_driver' was already open. Killing the application and retrying.")
                custom_functions.kill_app("uc_driver")
                time.sleep(5)

        #Open Feedback Manager workbook and update current working directory
        print("'[INFO]' Opening Feedback Manager workbook.")
        FeedbackManWb = xw.Book(FeedbackMan)
        RatingSh = FeedbackManWb.sheets(7)
        custom_functions.update_directory(FeedbackManWb)

        #Create all macros
        RefreshAll = FeedbackManWb.macro("Module1.RefreshAll")
        SortAll = FeedbackManWb.macro("Module1.SortAll")
        SortStatus = FeedbackManWb.macro("Module1.SortStatus")

        #Navigate through each account
        for account, url in Accounts.items():
            if account == "AccountA":
                root = "SellerOrg Corp"
            elif account == "AccountB":
                root = "Account B"
            elif account == "AccountD":
                root = "Account D"
            elif account == "AccountC":
                root = "Account C"
            elif account == "Apple":
                root = "Account F"
            elif account == "AccountE":
                root = "Account E"

            print(f"'[INFO]' Navigating to '{root}' account.")
            driver.get(url)
            time.sleep(2)
            driver.switch_to_window(0)

            try:
                code = None
                while not code:
                    code = accounts.Amazon_login(driver, username, password)

                    if not code:
                        print("'[ERROR]' Failed to log in to Amazon. Trying again.")
                        driver.get(url)
                        driver.switch_to_window(0)

            except TimeoutException:
                pass

            #Confirm if Feedback Manager's files have been already downloaded today
            print("'[INFO]' Checking if Feedback Manager's files have been downloaded.")
            last_update = LastUpdate().split(" ")[0]

            if last_update == currDate:
                print(f"'[INFO]' Feedback Manager's files for '{root}' has been downloaded today.")
            else:
                RequestReport()

            ###############################################################################################################################################
            #Move to Feedback Manager main Dashboard and get ratings table
            FeedbackManagerRatings()

            #Retrieve Feedback Manager's positive feedback comments
            FeedbackManagerComments()

        ###############################################################################################################################################
        #Open All Items workbook and refresh queries
        AllItemsWb = xw.Book(AllItems)
        custom_functions.update_directory(AllItemsWb)

        #Update all queries root folders and refresh queries
        print("'[INFO]' Opening 'All Items' and refreshing queries.")
        QueriesAddress = AllItemsWb.macro("Module2.QueriesAddress")
        QueriesAddress()
        time.sleep(5)

        RefreshFMan = AllItemsWb.macro("Module1.RefreshFMan")
        RefreshFMan()
        time.sleep(10)

        FManSh = AllItemsWb.sheets(5)
        FirstOrder: int = custom_functions.first_empty_row(FManSh, "D", "B3")
        LastOrder = int(FManSh.range(f"B{FManSh.cells.last_cell.row}").end("up").row)

        if FirstOrder > LastOrder:
            print("'[INFO]' No new orders to process.")
            Orders = None
            AllItemsWb.save()
            AllItemsWb.close()
            
        elif FirstOrder == LastOrder:
            Orders = []
            Orders.append(FManSh.range(f"B{FirstOrder}:C{LastOrder}").value)

        else:
            Orders = FManSh.range(f"B{FirstOrder}:C{LastOrder}").value

        if Orders is not None:
            for Order in Orders:
                print(f"'[INFO]' Navigating to '{Order[0]}' account.")

                if Order[0] == "SellerOrg Corp":
                    driver.get(Accounts["AccountA"])
                elif Order[0] == "Account B":
                    driver.get(Accounts["AccountB"])
                elif Order[0] == "Account D":
                    driver.get(Accounts["AccountD"])
                elif Order[0] == "Account C":
                    driver.get(Accounts["AccountC"])
                elif Order[0] == "Account F":
                    driver.get(Accounts["Apple"])
                elif Order[0] == "Account E":
                    driver.get(Accounts["AccountE"])

                driver.switch_to_window(0)

                print(f"'[INFO]' Getting order #{Order[1]} details.")
                driver.get(f"https://sellercentral.amazon.com/orders-v3/order/{Order[1]}")
                driver.switch_to_window(0)

                #Wait until the spinner is not visible
                WebDriverWait(driver, 60).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "#apploading > span")))

                #Extract the ASIN and SKU
                OrderRawData: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "myo-list-orders-product-name-cell"))).text.split("\n")
                ASIN: str = OrderRawData[1].replace("ASIN: ", "")
                SKU: str = OrderRawData[2].replace("SKU: ", "")

                FManSh.range(f"D{FirstOrder}").value = [ASIN, SKU]

                FirstOrder += 1

            print("'[INFO]' Refreshing queries.")
            RefreshFMan()

            time.sleep(10)
            AllItemsWb.save()
            AllItemsWb.close()

        ###############################################################################################################################################
        #Open Feedback Manager workbook, and refresh all queries
        print("'[INFO]' Refreshing all queries on 'Feedback Manager' workbook.")
        RefreshAll()
        time.sleep(3)

        SortStatus()
        time.sleep(5)

        for account, url in Accounts.items():

            if account == "AccountA":
                CurrSh = FeedbackManWb.sheets(1)
                root = "SellerOrg Corp"

            elif account == "AccountB":
                CurrSh = FeedbackManWb.sheets(2)
                root = "Account B"

            elif account == "AccountD":
                CurrSh = FeedbackManWb.sheets(3)
                root = "Account D"

            elif account == "AccountC":
                CurrSh = FeedbackManWb.sheets(4)
                root = "Account C"

            elif account == "Apple":
                CurrSh = FeedbackManWb.sheets(5)
                root = "Account F"

            elif account == "AccountE":
                CurrSh = FeedbackManWb.sheets(6)
                root = "Account E"

            FirstOrder = custom_functions.first_empty_row(CurrSh, "J", "B3")
            LastOrder = int(CurrSh.range(f"E{CurrSh.cells.last_cell.row}").end("up").row)

            if FirstOrder > LastOrder:
                print(f"'[INFO]' No new orders to process in '{root}' sheet.")
                continue

            elif FirstOrder == LastOrder:
                Orders = []
                Orders.append(CurrSh.range(f"E{FirstOrder}:E{LastOrder}").value)

            else:
                Orders = CurrSh.range(f"E{FirstOrder}:E{LastOrder}").value

            print(f"'[INFO]' Navigating to '{root}' account.")
            driver.get(url)
            time.sleep(2)
            driver.switch_to_window(0)

            for Order in Orders:
                print(f"'[INFO]' Retrieving order #{Order} info.")
                driver.get(f"https://sellercentral.amazon.com/messaging/inbox-v3?fi=search&ss={Order}")
                driver.switch_to_window(0)
                
                try:
                    status: str = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "case-status-label"))).text

                    if status.startswith("Resolved on "):
                        status.replace("Resolved on ", "Last response on ")
            
                except TimeoutException:
                    print(f"'[INFO]' No messages on order #{Order}. Moving forward.")
                    status = "N/A"

                CurrSh.range(f"J{FirstOrder}").value = status
                FirstOrder += 1

        driver.quit()

        ###############################################################################################################################################
        #Open workbook, send email, save and close
        print("'[INFO]' Sorting all tables, saving and closing workbook.")
        SortAll()
        time.sleep(3)
        FeedbackManWb.save()
        FeedbackManWb.close()

        print("'[INFO]' Loading workbook and sending email.")
        time.sleep(60)
        outlook.send_email(
            account="user@example.com",
            subject=f"Feedback Manager - {Date}",
            body=body,
            to=["user@example.com"],
            cc=["user@example.com"],
            attachments=[FeedbackMan],
            show=True,
            send=True
        )

        print("'[INFO]' Email has been sent.")

    #Sleep 60 seconds before starting over
    time.sleep(60)