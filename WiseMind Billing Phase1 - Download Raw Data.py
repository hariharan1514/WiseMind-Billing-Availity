import glob
import os.path
import sys
import time
from datetime import datetime
from openpyxl import load_workbook
import shutil
# from pathlib import path
import pandas as pd

from selenium import webdriver
from selenium.webdriver.ie.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

### Set the needed location paths for downloading process ###

systemUsername = os.getenv("USERNAME")
downloadFolderPath = fr"C:\Users\{systemUsername}\Downloads"
parentFolderPath = r"Z:\Wisemind\Charge Entry -Billing\Billing Dates"

pathTemp_Date = datetime.today().strftime('%m%d%Y')

### Create the Template path ###

if not os.path.exists(parentFolderPath + "\\" +pathTemp_Date[4:]):
    os.makedirs(parentFolderPath + "\\" +pathTemp_Date[4:])
if not os.path.exists(parentFolderPath + "\\" + pathTemp_Date[4:] + "\\" + datetime.today().strftime("%m %b'%Y")):
    os.makedirs(parentFolderPath + "\\" + pathTemp_Date[4:] + "\\" + datetime.today().strftime("%m %b'%Y"))
if not os.path.exists(parentFolderPath + "\\" + pathTemp_Date[4:] + "\\" + datetime.today().strftime("%m %b'%Y") + "\\" + pathTemp_Date):
    os.makedirs(parentFolderPath + "\\" + pathTemp_Date[4:] + "\\" + datetime.today().strftime("%m %b'%Y") + "\\" + pathTemp_Date)

### Set the Date range ###
# todayDate = datetime.today()
# startDate = todayDate - relativedelta(months=8)
startDate = input("Enter the Start Date Format(MMDDYYYY) : ")
startDate = startDate[:2] + "-" + startDate[2:4] + "-" + startDate[4:]
endDate = input("Enter the End Date Format(MMDDYYYY) : ")
endDate = endDate[:2] + "-" + endDate[2:4] + "-" + endDate[4:]
print(f"StartDate: {startDate}\nEndDate: {endDate}")

### Capture the Username & Password from the Config Sheet

conficSheetPath = r"Z:\Wisemind\Charge Entry -Billing\Automation Config File\ConfigSheet.xlsx"

configBook = load_workbook(conficSheetPath)
configSheet = configBook['Crendentials Sheet']

username = configSheet['B1'].value
password = configSheet['B2'].value

### Check1 & Check2 ###
attendanceDatafilepath = (parentFolderPath + "\\" +pathTemp_Date[4:] + "\\" +datetime.today().strftime("%m %b'%Y") + "\\" +pathTemp_Date + "\\" +f"Detailed Attendance - {pathTemp_Date}.csv")
allClientsDatafilepath = (parentFolderPath + "\\" +pathTemp_Date[4:] + "\\" +datetime.today().strftime("%m %b'%Y") + "\\" +pathTemp_Date + "\\" +f"All Claims - {pathTemp_Date}.xlsx")

if not os.path.exists(attendanceDatafilepath) or not os.path.exists(allClientsDatafilepath):
    ### Create Chrome Instance & Set Chrome Options ###
    chrome_Options = webdriver.ChromeOptions()
    chrome_Options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_Options)
    driver.maximize_window()

    driver.get("https://app.theranest.com/login")

    try:
        usernameElement = WebDriverWait(driver,60).until(EC.visibility_of_element_located((By.XPATH,"//input[@name='Email']")))
        usernameElement.send_keys(username)
    except:
        print("Login issue !!!")
        driver.quit()
        sys.exit()
    passwordElement = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH,"//input[@name='Password']")))
    passwordElement.send_keys(password)

    loginElement = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,"//button[normalize-space()='Log In']")))
    loginElement.click()

    try:
        bodyElement = WebDriverWait(driver,90).until(EC.visibility_of_all_elements_located((By.XPATH,"//div[@role='group']")))
    except:
        loginElement = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Log In']")))
        print("Please try logging in again after some time. The site is currently experiencing login issues.")
        sys.exit()

    if not os.path.exists(attendanceDatafilepath):
        time.sleep(3)
        driver.get(fr"https://wisemind71.theranest.com/reports/detailed-attendance/fromdate-{startDate}/todate-{endDate}")
        time.sleep(2)

        try:
            # loadElement = WebDriverWait(driver,600).until(EC.invisibility_of_element_located((By.XPATH,"// div[ @ id = 'content_ph']")))
            csvElement = WebDriverWait(driver,600).until(EC.element_to_be_clickable((By.XPATH,"//button[@id='csv']")))
            csvElement.click()
            try:
                portal_issue_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='json-formatter-container']")))
                print("Portal load issue. Kindly run the bot after sometime.!!!")
                sys.exit()
            except:
                pass


            Sec_count = 0
            Previous_files = set(glob.glob(f"{downloadFolderPath}/*.csv"))
            while Sec_count < 300:
                Listof_recentfiles = glob.glob(f"{downloadFolderPath}/*")
                try:
                    latestFile = max(Listof_recentfiles,key=os.path.getmtime)
                except:
                    time.sleep(1)
                    Sec_count +=1
                    continue
                currentFile = set(glob.glob(f"{downloadFolderPath}/*.csv"))
                newFile = currentFile - Previous_files
                if newFile:
                    latestFile = max(newFile,key=os.path.getmtime)
                    if latestFile.endswith('.csv'):
                        break
                else:
                    time.sleep(1)
                    Sec_count += 1
                    continue

            shutil.move(list(newFile)[0], attendanceDatafilepath)
            print(f"Detailed Attendance - {pathTemp_Date} Raw data file downloaded successfully.")

            ### Convert CSV to xlsx ###
            pd.read_csv(attendanceDatafilepath,dtype=str).to_excel(attendanceDatafilepath.replace('.csv', '.xlsx'), index=False)

            #delete original CSV file
            if os.path.exists(attendanceDatafilepath):
                os.remove(attendanceDatafilepath)

        except:
            print("""The site takes more than 10 minutes to load, and the Python script stops the downloading process. Please re-run the script after some time.!!!""")
            driver.close()
            sys.exit()

    if not os.path.exists(allClientsDatafilepath):
        time.sleep(3)
        driver.get("https://wisemind71.theranest.com/tenant/export")
        time.sleep(2)

        try:
            # loadElement = WebDriverWait(driver, 600).until(EC.invisibility_of_element_located((By.XPATH, "// div[ @ id = 'content_ph']")))
            exportAllClientElement = WebDriverWait(driver, 600).until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='Export all clients']")))
            exportAllClientElement.click()
            try:
                portal_issue_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class=json-formatter-container']")))
                print("Portal load issue. Kindly run the bot after sometime.!!!")
                sys.exit()
            except:
                pass

            Sec_count = 0
            Previous_files = set(glob.glob(f"{downloadFolderPath}/*.xlsx"))
            while Sec_count < 300:
                Listof_recentfiles = glob.glob(f"{downloadFolderPath}/*")
                try:
                    latestFile = max(Listof_recentfiles, key=os.path.getmtime)
                except:
                    time.sleep(1)
                    Sec_count += 1
                    continue
                currentFile = set(glob.glob(f"{downloadFolderPath}/*.xlsx"))
                newFile = currentFile - Previous_files
                if newFile:
                    latestFile = max(newFile, key=os.path.getmtime)
                    if latestFile.endswith('.xlsx'):
                        break
                else:
                    time.sleep(1)
                    Sec_count += 1
                    continue

            shutil.move(list(newFile)[0], allClientsDatafilepath)
            print(f"All Claims - {pathTemp_Date} Raw data file downloaded successfully")

        except:
            print("""The site takes more than 10 minutes to load, and the Python script stops the downloading process. Please re-run the script after some time.!!!""")
            driver.close()
            sys.exit()
    print("Download Completed !")
    time.sleep(2)
    driver.get("https://wisemind71.theranest.com/home/logout")
    time.sleep(3)
    driver.close()

else:
    print("The downloading process for the Billing and All Claims data files has already been completed.")
    # //div[@class='noty_message']