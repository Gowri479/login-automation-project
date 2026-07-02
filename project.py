

import time
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

EXCEL_FILE = "accounts.xlsx"


USERNAME_FIELD = (By.ID, "username")
PASSWORD_FIELD = (By.ID, "password")
LOGIN_BUTTON = (By.XPATH, "//button[@type='submit']")
SUCCESS_INDICATOR = (By.XPATH, "//div[@class='flash success']")




def read_accounts(path):
    wb = load_workbook(path)
    sheet = wb.active
    accounts = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        url, username, password = row[0], row[1], row[2]
        if url and username and password:
            accounts.append({"url": url, "username": username, "password": password})
    return accounts


def login(account, driver, wait):
    driver.get(account["url"])

    user_box = wait.until(EC.presence_of_element_located(USERNAME_FIELD))
    user_box.clear()
    user_box.send_keys(account["username"])

    pass_box = driver.find_element(*PASSWORD_FIELD)
    pass_box.clear()
    pass_box.send_keys(account["password"])

    driver.find_element(*LOGIN_BUTTON).click()

    try:
        wait.until(EC.presence_of_element_located(SUCCESS_INDICATOR))
        print(f"Successfully logged in: {account['username']} ({account['url']})")
        return True
    except TimeoutException:
        print(f"Login failed or could not be verified: {account['username']} ({account['url']})")
        return False


def main():
    accounts = read_accounts(EXCEL_FILE)
    if not accounts:
        print("No accounts found in the Excel file.")
        return

    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 15)

    try:
        for account in accounts:
            login(account, driver, wait)
            time.sleep(2)
    finally:
        driver.quit()


if __name__ == "__main__":
    main()