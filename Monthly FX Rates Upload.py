
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import openpyxl
import time
import os
import pandas as pd
from selenium.webdriver.support.wait import WebDriverWait
import calendar

downloadlink = "***"
uploadlink = "***"
username = "***"
password = "***"

month_number_map = {"Jan":1, 'Feb':2, 'Mar':3, 'Apr':4, 'May':5, 'Jun':6, 'Jul':7, 'Aug':8, 'Sep':9, 'Oct':10, 'Nov':11, 'Dec':12}

upload_month_1 = "Jan"
upload_month_2 = "Feb"
upload_year_1 = 2020
upload_year_2 = 2020
uploadfilename1 = f"FX_{upload_month_1}{str(upload_year_1)[2:]}.csv"
uploadfilename2= f"FX_{upload_month_2}{str(upload_year_2)[2:]}.csv"

download_dir = os.getcwd()
driver_options = Options()
driver_options.add_experimental_option("prefs", {'download.default_directory':download_dir, "download.prompt_for_download":False,
                                                 'download.directory_upgrade':True, "plugins.always_open_pdf_externally":True})
driver = webdriver.Chrome("Driver/chromedriver.exe", options=driver_options)

driver.get(downloadlink)
time.sleep(3)
download = driver.find_element_by_link_text("Month End FX Rates")
download.click()
time.sleep(2)

month_end_FX_Rates = pd.ExcelFile("C:/Users/SXC57/PycharmProjects/Monthly Fx Uplaod/Month End FX Rates.xls")
fx_rates = month_end_FX_Rates.parse("Report")
fx_rates = fx_rates.iloc[:, 0:3]
#for current month rate
month_end_day_currentmonth = calendar.monthrange(upload_year_1,month_number_map[upload_month_1])[1]
month_end_date_current_month = f"{month_end_day_currentmonth}/{month_number_map[upload_month_1]}/{upload_year_1}"
fx_rates.iloc[1,1] = month_end_date_current_month
fx_rates.iloc[5,2] = month_end_date_current_month
fx_rates.iloc[3,0] = f"The following rates are to be applied for accounting purposes and month end reconciliation as at {month_end_date_current_month}"
fx_rates.to_csv(uploadfilename1, index=False)

#for next month rate
month_end_day = calendar.monthrange(upload_year_2,month_number_map[upload_month_2])[1]
month_end_date = f"{month_end_day}/{month_number_map[upload_month_2]}/{upload_year_2}"
fx_rates.iloc[1,1] = month_end_date
fx_rates.iloc[5,2] = month_end_date
fx_rates.iloc[3,0] = f"The following rates are to be applied for accounting purposes and month end reconciliation as at {month_end_date}"
fx_rates.to_csv(uploadfilename2, index=False)

driver = webdriver.Chrome("Driver/chromedriver.exe")
driver.get(uploadlink)
time.sleep(5)
driver.find_element_by_name("username").send_keys(username)
driver.find_element_by_name("password").send_keys(password)
driver.find_element_by_id("btnSiteLogin").click()

# uploadbutton = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "Selectedfile")))
time.sleep(2)
uploadbutton = driver.find_element_by_name("Selectedfile")
uploadbutton.send_keys(f"C:/Users/SXC57/PycharmProjects/Monthly Fx Uplaod/{uploadfilename1}")
uploadbutton.send_keys(f"C:/Users/SXC57/PycharmProjects/Monthly Fx Uplaod/{uploadfilename2}")




