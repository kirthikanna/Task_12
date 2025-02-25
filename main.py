#main.py
"""
Main DDTF Execution Engine for Automation Testing
"""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from data import Data
from locators import WebLocators
from excel_functions import KeerthanaExcelReader

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()
driver.get(Data().URL)
driver.implicitly_wait(10)
excel_reader =  KeerthanaExcelReader(Data().EXCEL_FILE,Data().SHEET_NUMBER)
rows = excel_reader.row_count()

for row in range(2,rows+1):
    username =excel_reader.read_data(row,column_number=6).strip()
    password =excel_reader.read_data(row,column_number=7).strip()

    driver.find_element(by=By.NAME, value=WebLocators().USERNAME_INPUT_BOX).send_keys(username)
    driver.find_element(by=By.NAME, value=WebLocators().PASSWORD_INPUT_BOX).send_keys(password)
    driver.find_element(by=By.CSS_SELECTOR, value=WebLocators().SUBMIT_BUTTON).click()

    """ wait for the page to load and check if the login is successful"""
    driver.implicitly_wait(10)
    if Data().DASHBOARD_URL in driver.current_url:
        print(f"SUCCESS : Login Success with USERNAME={username} and PASSWORD={password}")
        excel_reader.write_data(row, 8,"TEST PASSED")
        driver.find_element(by=By.CLASS_NAME, value=WebLocators().HAMBURGER_BUTTON).click()
        driver.find_element(by=By.XPATH, value=WebLocators().LOGOUT_BUTTON).click()
    elif (Data().URL in driver.current_url):
        print(f"ERROR :Login Failed with USERNAME={username} and PASSWORD={password}")
        excel_reader.write_data(row, 8,"TEST FAIL")
        driver.refresh()

##Close the DDTF Automation Testing
driver.quit()
