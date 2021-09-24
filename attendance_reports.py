#! python3
from datetime import date, timedelta
from functools import wraps
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from tqdm import tqdm
import os
import pyautogui as pag
import sys
import time


pbar = tqdm(total=100, colour="#98be65", file=sys.stdout)


def pbar_updater(pbar, description, progress):
    def decorator(function):
        @wraps(function)
        def wrapper(*args, **kwargs):
            pbar.set_description(description)
            result = function(*args, **kwargs)
            pbar.update(progress)
            return result

        return wrapper

    return decorator


@pbar_updater(pbar, "Initializing", 5)
def initialize():
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 10)
    return driver, wait


@pbar_updater(pbar, "Logging Into Q", 20)
def login(driver, USERNAME, PASSWORD, URL):
    """Opens SIS and logs in
    :returns: none
    """
    driver.get(URL)
    driver.find_element_by_id("login").click()
    handles = driver.window_handles
    driver.switch_to.window(handles[1])  # switch to login window
    driver.implicitly_wait(2)
    username = driver.find_element_by_id("username")
    username.send_keys(USERNAME)
    password = driver.find_element_by_id("userpassword")
    password.send_keys(PASSWORD)
    password.send_keys(Keys.RETURN)


@pbar_updater(pbar, "Checking Date", 10)
def check_day():
    """checks the day of the week then calculates back to Monday
    :returns: end of week and beginning of the week formatted as strings
    """
    tday = date.today()
    week_day = tday.weekday()
    today_string = tday.strftime("%m%d%Y")
    beginning_of_week_string = (tday - timedelta(days=week_day)).strftime("%m%d%Y")
    return beginning_of_week_string, today_string


@pbar_updater(pbar, "Generating Report", 35)
def generate_report(driver, wait, week_start, week_end):
    """navigates to the report page in the SIS and generates the report.
    :returns: none
    """
    attendance_report_page = wait.until(
        EC.element_to_be_clickable((By.LINK_TEXT, "Class Attendance Spreadsheet Rpt"))
    )
    attendance_report_page.click()
    time.sleep(5)
    track_dropdown = Select(driver.find_element_by_id("optionfilterTrack"))
    track_dropdown.select_by_visible_text("T dean - De Anza Academy of Tech & Arts")
    driver.find_element_by_id("optionDateRangeType_5").click()
    end_date = driver.find_element_by_id("optionEffDateEnd")
    end_date.send_keys(Keys.CONTROL + "a")
    end_date.send_keys(week_end)
    begin_date = driver.find_element_by_id("optionEffDateStart")
    begin_date.send_keys(Keys.CONTROL + "a")
    begin_date.send_keys(week_start)
    time.sleep(2)
    wait.until(EC.element_to_be_clickable((By.ID, "ReportMasterPrintBtn"))).click()


@pbar_updater(pbar, "Printing Reports", 25)
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


@pbar_updater(pbar, "Logging Out", 5)
def logout(driver, wait):
    """logout of SIS and quit browser"""
    logout = wait.until(EC.element_to_be_clickable((By.ID, "logoutbtn")))
    logout.click()
    driver.quit()


if __name__ == "__main__":

    # Get credentials from environment
    USERNAME = os.environ.get("Q_USERNAME")
    PASSWORD = os.environ.get("Q_PASSWORD")
    URL = os.environ.get("Q_URL")

    driver, wait = initialize()
    login(driver, USERNAME, PASSWORD, URL)
    week_start, week_end = check_day()
    generate_report(driver, wait, week_start, week_end)
    print_report()
    logout(driver, wait)
    pbar.write(
        "\nAttendence reports have been printed. Please head to the office to sign them."
    )
    pbar.close()
