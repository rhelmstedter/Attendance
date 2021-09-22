#! python3
from datetime import date, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
import os
import pyautogui as pag
import time
from tqdm import tqdm


def initializing():
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 10)
    return driver, wait


def login(driver, USERNAME, PASSWORD):
    """ Opens SIS and logs in
    :returns: none
    """
    driver.get("http://venturausd.vcoe.org/q/vcsplash.aspx")
    driver.find_element_by_id("login").click()
    handles = driver.window_handles
    driver.switch_to.window(handles[1])  # switch to login window
    driver.implicitly_wait(2)
    username = driver.find_element_by_id("username")
    username.send_keys(USERNAME)
    password = driver.find_element_by_id("userpassword")
    password.send_keys(PASSWORD)
    password.send_keys(Keys.RETURN)


def check_day():
    """ checks what day of the week the script is run. Then calculates back to the previous monday
    :returns: end of week and beginning of the week formatted as strings
    """
    tday = date.today()
    week_day = tday.weekday()
    today_string = tday.strftime("%m%d%Y")
    beginning_of_week_string = (tday-timedelta(days=week_day)).strftime("%m%d%Y")
    return beginning_of_week_string, today_string


def generate_report(driver, wait, week_start, week_end):
    """navigates to the report page and generates the report. 
    :returns: none
    """
    attendance_report_page = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Class Attendance Spreadsheet Rpt")))
    attendance_report_page.click()
    time.sleep(5)
    track_dropdown = Select(driver.find_element_by_id('optionfilterTrack'))
    track_dropdown.select_by_visible_text("T dean - De Anza Academy of Tech & Arts")
    driver.find_element_by_id("optionDateRangeType_5").click()
    end_date = driver.find_element_by_id("optionEffDateEnd")
    end_date.send_keys(Keys.CONTROL + 'a')
    end_date.send_keys(week_end)
    begin_date = driver.find_element_by_id("optionEffDateStart")
    begin_date.send_keys(Keys.CONTROL + 'a')
    begin_date.send_keys(week_start)
    time.sleep(2)
    wait.until(EC.element_to_be_clickable((By.ID, "ReportMasterPrintBtn"))).click()


def print_report():
    """Uses pyautogui to click the print buttons. Couldn't get it to work with selenium
    :returns: none
    """
    time.sleep(4)
    print_icon = pag.locateCenterOnScreen("./print_icon.png", confidence=0.7)
    pag.moveTo(print_icon)
    pag.click()

    time.sleep(4)
    print_button = pag.locateCenterOnScreen("./print_button.png", confidence=0.7)
    pag.moveTo(print_button)
    pag.click()
    time.sleep(5)


def logout(driver, wait):
    """ logout of SIS and quit browser """
    logout = wait.until(EC.element_to_be_clickable((By.ID, "logoutbtn")))
    logout.click()
    driver.quit()


def main():
    print("Retrieving attendance reports...\n")

    with tqdm(total=100, colour='#98be65') as pbar:

        # Get credentials from environment
        USERNAME = os.environ.get("Q_USERNAME")
        PASSWORD = os.environ.get("Q_PASSWORD")
        # Set Driver

        pbar.set_description('Initializing')
        driver, wait = initializing()
        pbar.update(5)

        pbar.set_description('Logging into Q')
        login(driver, USERNAME, PASSWORD)
        pbar.update(20)
        
        pbar.set_description('Checking Date')
        week_start, week_end = check_day()
        pbar.update(5)
        
        pbar.set_description('Generating Reports')
        generate_report(driver, wait, week_start, week_end)
        pbar.update(40)

        pbar.set_description('Printing Reports')
        print_report()
        pbar.update(30)

        pbar.set_description('Logging Out')
        logout(driver, wait)
        pbar.update(5)

    print("\nAttendence reports have been printed. Please head to the office to sign them.")


if __name__ == "__main__":
    main()

