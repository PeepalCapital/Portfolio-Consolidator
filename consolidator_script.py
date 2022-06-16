'''
This program downloads the following
1. One zerodha holding excel sheet
2. Two icici direct holding excel sheet
3. All listed active equities on NSE (not in active code)
4. All listed active equities on BSE (not in active code)
5. Moves all the above files to the current working directory
Credentials are stored in the appropriate json file.
Chromedriver needs to be in the same folder as python file, as well as json file.
Selenium and tqdm needs to be installed
This is built for Windows OS
'''


import json
import time
import glob
import os
import shutil
import sys
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from tqdm import tqdm



'''
Download zerodha holding report
'''


print ("\nZerodha portfolio holding report download in progress...")

for i in tqdm(range(1)):

    with open("zerodha_credentials.json") as credentialsFile:
        data = json.load(credentialsFile)
        username = data['username']
        password = data['password']
        pin = data['pin']

    driver = webdriver.Chrome()
    driver.get('https://kite.zerodha.com/')

    driver.minimize_window()

    try:
        driver.find_element_by_xpath("/html/body/div[1]/div/div/div[1]/div/div/div/form/div[2]/input").send_keys(username)
        driver.find_element_by_xpath("/html/body/div[1]/div/div/div[1]/div/div/div/form/div[3]/input").send_keys(password)
        driver.find_element_by_xpath("/html/body/div[1]/div/div/div[1]/div/div/div/form/div[4]/button").click()
    except:
        print("\nOPERATION FAIL: Login Credentials are either incorrect or need to be changed. Please login manually first!")

    driver.implicitly_wait(20)

    try:
        driver.find_element_by_xpath("//*[@id='container']/div/div/div/form/div[2]/div/input").send_keys(pin)
        driver.find_element_by_xpath("/html/body/div[1]/div/div/div[1]/div/div/div/form/div[3]/button").click()
    except:
        print("\nOPERATION FAIL: PIN is either incorrect or has to be changed. Please login manually first!")

    try:
        driver.find_element_by_xpath("/html/body/div[1]/div[1]/div/div[2]/div[1]/a[3]/span").click()
        driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/section/div/div[1]/div/span[3]/span").click()
    except:
        print("\nOPERATION FAIL: Report link has perhaps changed. Please login manually first!")

    time.sleep(5) #5 seconds to download and close browser

    driver.close()

print ("\nZerodha portfolio holding report has successfully been downloaded!")

time.sleep(0.01)


'''
Downloaded zerodha file to be moved to current working directory and renamed
'''

print ("\n\nZerodha portfolio holding report being renamed & moved to working folder...")

for i in tqdm(range(1)):


    downloaded_files = glob.iglob('C:\\Users\\PeepalCapital\\Downloads\\*') #put the path name of your donwload folder

    latest_file_path_downloaded = max(downloaded_files , key = os.path.getctime)
    shutil.copy(latest_file_path_downloaded, os.getcwd())
    working_file = glob.iglob('C:\\Users\\PeepalCapital\\AppData\\Roaming\\Sublime Text 3\\Packages\\User\\*') #put the path name of your current working folder
    working_file_path_copied = max(working_file , key  = os.path.getctime)
    file_name = os.path.basename (working_file_path_copied)

    try:
        os.rename (file_name, 'zerodha_holdings.csv')
    except WindowsError:
        os.remove('zerodha_holdings.csv')
        os.rename (file_name, 'zerodha_holdings.csv')

print ("\nZerodha portfolio holding report has successfully been renamed and moved!")

time.sleep(0.01)


'''
Download ICICI direct holding report (1)
'''

print ("\n\nICICI Direct portfolio holding report (1) download in progress...")

for i in tqdm(range(1)):

    with open("icici1_credentials.json") as credentialsFile1:
        data1 = json.load(credentialsFile1)
        user1 = data1['user1']
        password1 = data1['password1']
        dob1 = data1['dob1']

    driver = webdriver.Chrome()
    driver.get('https://secure.icicidirect.com/IDirectTrading/Customer/Login.aspx')

    driver.minimize_window()

    try:
        driver.find_element_by_xpath("//*[@id='txtUserId']").send_keys(user1)
        driver.find_element_by_xpath("//*[@id='txtPass']").send_keys(password1)
        driver.find_element_by_xpath("//*[@id='txtDOB']").send_keys(dob1)
        driver.find_element_by_xpath("//*[@id='lbtLogin']").click()
    except:
        print("\nOPERATION FAIL: Login Credentials are either incorrect or need to be changed. Please login manually first!")
        driver.close()
        sys.exit()

    try:
        driver.find_element_by_xpath("//*[@id='hypPF']").click()
        driver.implicitly_wait(20)
        driver.find_element_by_xpath("//*[@id='dvMenu']/ul/li[1]/div/ul/li[1]/a/label[1]").click()
        driver.implicitly_wait(20)
        driver.find_element_by_xpath("//*[@id='hypfilter']").click()
        driver.implicitly_wait(20)
        driver.find_element_by_xpath("//*[@id='dvfilter']/div[2]/ul/li[2]/img[1]").click()
    except:
        print("\nOPERATION FAIL: ICICI is either trying to sell something or wants a confirmation. Please login manually first!")
        driver.close()
        sys.exit()

    time.sleep(5) #5 seconds to download and close browser

    driver.close()

print ("\nICICI direct portfolio holding report (1) has successfully been downloaded!")

time.sleep(0.01)

'''
Downloaded icici direct file (1) to be moved to current working directory and renamed
'''

print ("\n\nICICI direct portfolio holding report (1) being renamed & moved to working folder...")

for i in tqdm(range(1)):


    downloaded_files = glob.iglob('C:\\Users\\PeepalCapital\\Downloads\\*') #put the path name of your donwload folder

    latest_file_path_downloaded = max(downloaded_files , key = os.path.getctime)
    shutil.copy(latest_file_path_downloaded, os.getcwd())
    working_file = glob.iglob('C:\\Users\\PeepalCapital\\AppData\\Roaming\\Sublime Text 3\\Packages\\User\\*') #put the path name of your current working folder
    working_file_path_copied = max(working_file , key  = os.path.getctime)
    file_name = os.path.basename (working_file_path_copied)

    try:
        os.rename (file_name, 'icici_direct_holdings_1.xls')
    except WindowsError:
        os.remove('icici_direct_holdings_1.xls')
        os.rename (file_name, 'icici_direct_holdings_1.xls')

print ("\nICICI direct portfolio holding report (1) has successfully been renamed and moved!")

time.sleep(0.01)


'''
Download ICICI direct holding report (2)
'''

print ("\n\nICICI Direct portfolio holding report (2) download in progress...")

for i in tqdm(range(1)):

    with open("icici2_credentials.json") as credentialsFile2:
        data2 = json.load(credentialsFile2)
        user2 = data2['user2']
        password2 = data2['password2']
        dob2 = data2['dob2']

    driver = webdriver.Chrome()
    driver.get('https://secure.icicidirect.com/IDirectTrading/Customer/Login.aspx')

    driver.minimize_window()

    try:
        driver.find_element_by_xpath("//*[@id='txtUserId']").send_keys(user2)
        driver.find_element_by_xpath("//*[@id='txtPass']").send_keys(password2)
        driver.find_element_by_xpath("//*[@id='txtDOB']").send_keys(dob2)
        driver.find_element_by_xpath("//*[@id='lbtLogin']").click()
    except:
        print("\nOPERATION FAIL: Login Credentials are either incorrect or need to be changed. Please login manually first!")
        driver.close()
        sys.exit()

    try:
        driver.find_element_by_xpath("//*[@id='hypPF']").click()
        driver.implicitly_wait(20)
        driver.find_element_by_xpath("//*[@id='dvMenu']/ul/li[1]/div/ul/li[1]/a/label[1]").click()
        driver.implicitly_wait(20)
        driver.find_element_by_xpath("//*[@id='hypfilter']").click()
        driver.implicitly_wait(20)
        driver.find_element_by_xpath("//*[@id='dvfilter']/div[2]/ul/li[2]/img[1]").click()
    except:
        print("\nOPERATION FAIL: ICICI is either trying to sell something or wants a confirmation. Please login manually first!")
        driver.close()
        sys.exit()

    time.sleep(5) #5 seconds to download and close browser

    driver.close()

print ("\nICICI direct portfolio holding report (2) has successfully been downloaded!")

time.sleep(0.01)

'''
Downloaded icici direct file (2) to be moved to current working directory and renamed
'''

print ("\n\nICICI direct portfolio holding report (2) being renamed & moved to working folder...")

for i in tqdm(range(1)):


    downloaded_files = glob.iglob('C:\\Users\\PeepalCapital\\Downloads\\*') #put the path name of your donwload folder

    latest_file_path_downloaded = max(downloaded_files , key = os.path.getctime)
    shutil.copy(latest_file_path_downloaded, os.getcwd())
    working_file = glob.iglob('C:\\Users\\PeepalCapital\\AppData\\Roaming\\Sublime Text 3\\Packages\\User\\*') #put the path name of your current working folder
    working_file_path_copied = max(working_file , key  = os.path.getctime)
    file_name = os.path.basename (working_file_path_copied)

    try:
        os.rename (file_name, 'icici_direct_holdings_2.xls')
    except WindowsError:
        os.remove('icici_direct_holdings_2.xls')
        os.rename (file_name, 'icici_direct_holdings_2.xls')

print ("\nICICI direct portfolio holding report (2) has successfully been renamed and moved!")

time.sleep(0.01)


'''
Download all active NSE Equity


print ("\n\nAll active equities traded on NSE being downloaded...")

for i in tqdm(range(1)):
    try:
        driver = webdriver.Chrome()
        driver.get('https://www.nseindia.com/corporates/content/securities_info.htm')
        driver.minimize_window()
        driver.find_element_by_xpath("//*[@id='wrapper_btm']/div[1]/div[4]/div/ul/li[1]/a").click()
        time.sleep(5) #5 seconds to download and close browser
        driver.close()
    except:
        print("\nOPERATION FAIL: Not able to download all active equities from NSE")
        driver.close()
        sys.exit()

print ("\nAll active equities traded on NSE has been successfully downloaded!")

time.sleep(0.01)
#reuse code for file movement as above


Download all active BSE Equity


print ("\n\nAll active equities traded on BSE being downloaded...")

for i in tqdm(range(1)):
    try:
        driver = webdriver.Chrome()
        driver.get('https://www.bseindia.com/corporates/List_Scrips.aspx')
        driver.minimize_window()
        select = Select(driver.find_element_by_id('ContentPlaceHolder1_ddSegment'))
        select.select_by_visible_text("Equity")
        select = Select(driver.find_element_by_id('ContentPlaceHolder1_ddlStatus'))
        select.select_by_visible_text("Active")
        driver.find_element_by_xpath("//*[@id='ContentPlaceHolder1_btnSubmit']").click()
        driver.implicitly_wait(20)
        driver.find_element_by_xpath("//*[@id='ContentPlaceHolder1_lnkDownload']/i").click()
        time.sleep(5) #5 seconds to download and close browser
        driver.close()
    except:
        print("\nOPERATION FAIL: Not able to download all active equities from BSE")
        driver.close()
        sys.exit()

print ("\nAll active equities traded on BSE has been successfully downloaded!")

time.sleep(0.01)

#reuse code for file movement as above

'''
