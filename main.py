import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from datetime import datetime 
from selenium.webdriver.common.keys import Keys
from functools import partial
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import csv
import json
from openpyxl import Workbook, load_workbook
import pandas as pd
from tqdm import tqdm

executablePath = "C://chromedriver.exe"###need to config
teamsDataLogFile = 'teamsDataLog.txt'
teamsDataTerminalLogFile = 'teamsDataTerminalLog.txt'
global list_of_rows
global row_data

def loadUrl(driver, url):
    firstTimeLoad = True
    while True:
        try:
            driver.set_page_load_timeout(70)
            driver.get(url)
            time.sleep(3)
            try:
                newurl = driver.current_url
            except Exception as e:
                insertToLogFile("Error while accessing to driver", e)
                continue
            page_status = driver.execute_script('return document.readyState;')
            print("page_status : ",page_status)
            if get_response(driver) and 'complete' == page_status:
                if isWebSiteLoadCompletely(driver):
                    for i in range(2):
                        driver = unwantedPopup(driver)
                    time.sleep(1)
                    driver.execute_script("document.body.style.zoom='70%'")
                    time.sleep(1)
                    return
            else:
                driver.refresh()
        except Exception as e:
            insertToLogFile("Error While Loading Url : %s" % url, e)
            try:
                if not firstTimeLoad:
                    driver.refresh()
                    time.sleep(2)
                    try:
                        loadUrl(driver, url)
                        return
                    except:
                        pass
                else:
                    firstTimeLoad = False
            except:
                pass
            time.sleep(3)

def get_response(driver):
    logs = driver.get_log('performance')
    for log in logs:
        if log['message']:
            d = json.loads(log['message'])
            try:
                content_type = 'text/html' in d['message']['params']['response']['headers']['content-type']
                response_received = d['message']['method'] == 'Network.responseReceived'
                if content_type and response_received:
                    print(str(d['message']['params']['response']['status']))
                    if "200" == str(d['message']['params']['response']['status']):
                        return True
                    else:
                        printAndInsertToTerminalLogFile(str(d['message']['params']['response']['status']))
                        return False
            except:
                pass
    return False

def isWebSiteLoadCompletely(driver):
    try:
        body = driver.find_element(By.XPATH, '/html/body')
        tt = body.text
    except:
        return False
    return True

def workOption(option):
    option.add_argument("--start-maximized")
    option.add_argument("disable-infobars")
    option.add_argument("--disable-dev-shm-usage")
    option.add_argument("--disable-gpu")
    option.add_argument("--no-sandbox")
    # option.add_argument('--ignore-certificate-errors')
    # option.add_argument("headless")
    option.add_experimental_option('excludeSwitches', ['enable-logging'])
    chromePrefers = {}
    option.experimental_options["prefs"] = chromePrefers
    chromePrefers["profile.default_content_settings"] = {"images": 2}
    chromePrefers["profile.managed_default_content_settings"] = {"images": 2}
    return option

def unwantedPopup(driver):
    xAskMeLaterButton = './/button[@class="close_btn_thick"]'
    for i in range(4):
        try:
            askMeLater = driver.find_element(By.XPATH, xAskMeLaterButton)
            time.sleep(0.5)
            if askMeLater.is_displayed():
                askMeLater.click()
                time.sleep(1.5)
                return driver
        except:
            pass
    return driver

def findElement(self, xxPath: str, level: str, finds: bool = False, refreshTime: int = 16, Get_None: bool = False,
                time_out=20, text_check=False, timer=True):
    n = 0
    while True:

        n += 1
        if n % time_out == 0 and refreshTime != 16:
            self = loadUrl(self, self.current_url)
            time.sleep(refreshTime)
            if n == 40 and Get_None is True:
                return None
        elif Get_None is True and n == time_out:
            return None

        try:
            if finds:
                element = self.find_elements(By.XPATH, xxPath)
            else:
                element = self.find_element(By.XPATH, xxPath)

            if (element is not None) or (element.is_displayed()):
                self = self
                try:
                    if text_check:
                        if element.text is not None:
                            text = element.text
                    return element
                except Exception as e:
                    insertToLogFile("Error While getting text of  Element", e)
        except Exception as e:
            if n % 15 == 1 and not Get_None:
                insertToLogFile("Error While Finding Element in level : " + level, e)
            time.sleep(1)

        if n % 4 == 0 and timer:
            printAndInsertToTerminalLogFile(int(n / 4))

def start(startUrl="https://wwwapps.ups.com/ctc/request?loc=en_US", needLoad: bool = True):
    global currentTeamLink
    currentTeamLink = startUrl
    option = webdriver.ChromeOptions()
    s = Service(executablePath)
    caps = DesiredCapabilities().CHROME
    caps["pageLoadStrategy"] = "normal"
    caps["goog:loggingPrefs"] = {"performance": "ALL"}
    driver = webdriver.Chrome(service=s, options=workOption(option), desired_capabilities=caps)
    driver.maximize_window()
    if needLoad: loadUrl(driver, startUrl)
    driver.findElement = partial(findElement, driver)
    print("i finish start method")
    time.sleep(6)
    return driver

def insertToLogFile(level, exceptText, element=None):
    exceptText = "\n".join(str(exceptText).split("\n")[:3])
    if element is not None:
        level = "%s ----- > %s" % (level, element)
    report = "\nCurrentTeam : %s \nTime : %s \nLevel : %s \nError : %s \n" % (
        currentTeamLink, datetime.now(), level, exceptText)
    with open(teamsDataLogFile, "a", encoding="utf-8") as f:
        f.write(report)

def printAndInsertToTerminalLogFile(text, end="\n"):
    print(text, end=end)

    with open(teamsDataTerminalLogFile, "a", encoding="utf-8") as f:
        f.write("%s%s" % (str(text), end))

def readEarlierData():
    lines = []
    with open("C://Users//Desktop/ships_transfer/10k_sample_6-23.csv","r") as f :###need to config
        file = csv.reader(f)
        for row in file:
            lines.append(row)
    global list_of_rows
    list_of_rows = lines[1:]
    print("len : ",len(list_of_rows))
    return lines[1:]

def changed_date(input):
    date = pd.to_datetime(input)
    formatted_date = date.strftime('%m/%d/%Y')
    return str(formatted_date)

def validate(i):
    with open("C://Users//Desktop/ships_transfer/validate.txt","a") as f :###need to config
        f.write((",".join(i))+"\n")

def update(driver):
    xUpdate = './/button[@id="ctc_module1_submit"]'
    xEdit = './/div[@id="module1_edit"]/button[@class="ups-link"]'
    try:
        updateButton = driver.findElement(xUpdate,level ="update button")
        updateButton.send_keys(Keys.ENTER)
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, xEdit))
        )
        return driver
    except Exception as e:
        insertToLogFile("Error While Clicking On update Button", e)
    return driver

def append_to_excel(data):
    file_path = "C://Users//Desktop/ships_transfer/final_data.xlsx"###need to config
    try:
        workbook = load_workbook(file_path)
        sheet = workbook.active
    except FileNotFoundError:
        print("excel file not found")
    sheet.append(data)
    workbook.save(file_path)

def collect_details(obj):
    DayOfWeek = obj.text
    if "Saturday" not in DayOfWeek:
        xDaysToDeliver = ".//td[2]/div[1]/div[2]/p"
        xDateOfDeliver = ".//td[2]/div[2]/p[3]"
        temp = None
    else:
        xDaysToDeliver = ".//td[2]/p"
        xDateOfDeliver = ".//td[2]/div[1]/p[3]"
        temp = []
    DaysToDeliver = str(obj.find_element(By.XPATH,xDaysToDeliver).text).replace("*","").replace(":","")
    dateOfDeliver = obj.find_element(By.XPATH,xDateOfDeliver).text
    if temp == []:
        temp = [DaysToDeliver,changed_date(dateOfDeliver)]
        return temp
    global row_data
    row_data.extend([DaysToDeliver,changed_date(dateOfDeliver)])
    return temp

def collector(driver):
    global row_data
    xServices = './/table[@class="dataTable full"]/tbody//tr[td[2]]'
    services = driver.findElement(xServices,level = "searching for services",finds=True)
    f = 2
    if services is not None:
        temp = None
        for n in range(len(services)):
            if f == 0 : break
            if "UPS Ground" in getTextOf(services[n].find_element(By.XPATH,".//td[1]/p")):
                if temp is not None:
                    temp1 = temp
                    temp = collect_details(services[n])
                    row_data.extend(temp1)
                else :
                    temp = collect_details(services[n])
                f-=1
    else :
        row_data.extend(["",""])
    if f == 2 :
        row_data.extend(["",""])
    return driver

def getTextOf(row):
    for i in range(10):
        try:
            rowText = row.text
            if rowText is not None:
                return rowText
        except Exception as e:
            print("i have exception")
            insertToLogFile("Error While getting row text", e)
            time.sleep(3)
    return None

def enter_orig_zip(driver,orig_zip:str):
    xOrigZip = './/input[@name="origPostalCode"]'
    origInput = driver.findElement(xOrigZip,level = "orig inputing")
    origInput.clear()
    origInput.send_keys(orig_zip)
    return driver

def enter_dest_zip(driver,dest_zip:str):
    xOrigZip = './/input[@name="destPostalCode"]'
    origInput = driver.findElement(xOrigZip,level = "orig inputing")
    origInput.clear()
    origInput.send_keys(dest_zip)
    return driver

def enter_input_date(driver,date):
    xDateInput = './/input[@id="ctcDatePicker"]'
    dateInput = driver.findElement(xDateInput,level = "input date")
    dateInput.clear()
    dateInput.send_keys(date)
    time.sleep(0.2)
    dateInput.send_keys(Keys.ENTER)
    return driver

def edit_form(driver):
    try:
        xEdit = './/div[@id="module1_edit"]/button[@class="ups-link"]'
        editButton = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, xEdit)))
        if editButton.is_displayed():
            editButton.send_keys(Keys.ENTER)
            time.sleep(1.5)
    except Exception as e:
        print(e)
    return driver

def main():
    global row_data
    row_data = []
    driver = start()
    while True:
        try:
            driver.find_element(By.XPATH, '/html/body').send_keys(Keys.HOME)
            break
        except:
            loadUrl(driver, driver.current_url)
    print("site load completely")
    readEarlierData()
    list_of_zip_codes = list_of_rows[9500:9750]###need to config
    print("data load completely")
    with tqdm(total=len(list_of_zip_codes),desc="proccesing") as pbar:
        for i in list_of_zip_codes:
            pbar.update(1)
            for ii in range(2):
                input_date = "07/17/2023" if ii ==0 else "07/20/2023"
                if row_data == []:
                    row_data = [i[0],i[1],input_date]
                elif len(row_data) == 5:
                    row_data.extend(["","",input_date])
                else:
                    row_data.append(input_date)
                driver = edit_form(driver)
                driver = enter_orig_zip(driver,str(i[0]).strip())
                driver = enter_dest_zip(driver,str(i[1]).strip())
                driver = enter_input_date(driver,input_date)
                driver = update(driver)
                driver = collector(driver)

            append_to_excel(row_data)
            row_data = []
    driver.close()


def main1():
    global row_data
    row_data = []
    driver = start()
    while True:
        try:
            driver.find_element(By.XPATH, '/html/body').send_keys(Keys.HOME)
            break
        except:
            loadUrl(driver, driver.current_url)
    print("site load completely")
    readEarlierData()
    list_of_zip_codes = list_of_rows[9750:1000]###need to config
    print("data load completely")
    with tqdm(total=len(list_of_zip_codes),desc="proccesing") as pbar:
        for i in list_of_zip_codes:
            pbar.update(1)
            for ii in range(2):
                input_date = "07/17/2023" if ii ==0 else "07/20/2023"
                if row_data == []:
                    row_data = [i[0],i[1],input_date]
                elif len(row_data) == 5:
                    row_data.extend(["","",input_date])
                else:
                    row_data.append(input_date)
                driver = edit_form(driver)
                driver = enter_orig_zip(driver,str(i[0]).strip())
                driver = enter_dest_zip(driver,str(i[1]).strip())
                driver = enter_input_date(driver,input_date)
                driver = update(driver)
                driver = collector(driver)

            append_to_excel(row_data)
            row_data = []
    driver.close()

main()
print("first main finished")
main1()
print("second main finished")
print()

