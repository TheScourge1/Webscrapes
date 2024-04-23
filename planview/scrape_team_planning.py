from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.remote.webdriver import WebDriver, NoSuchElementException, WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementClickInterceptedException
import time, os, shutil,sys

# Set the download directory
DOWNLOAD_DIR = os.getcwd() + "\\download"
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("detach", True)
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})


HOME_URL = "https://crelan.pvcloud.com/"

PORTFOLIO = "IT Omnichannels All"
START_DATE = "1-1-2024"
END_DATE = "31-12-2024"

REPORT_NAME = "resources24.xlsx"


def home_page(driver: WebDriver, username: str, password: str) -> WebDriver:
    print(f"logging in into:{HOME_URL}")
    driver.get(HOME_URL)
    driver.find_element(By.ID, "userNameInput").send_keys(username)
    driver.find_element(By.ID, "passwordInput").send_keys(password)
    driver.find_element(By.ID, "submitButton").click()

    if driver.current_url.startswith(HOME_URL):
        print("login successfull")
        return driver
    else:
        raise Exception("Unable to login")


def load_portfolio(driver: WebDriver) -> WebDriver:
    try:
        driver.find_element(By.LINK_TEXT,"Resources").click()
        element = wait_element(driver, (By.LINK_TEXT,PORTFOLIO),10)
        ActionChains(driver).move_to_element(element).perform()
        element = element.find_element(By.XPATH,"following-sibling::*")
        element.click()
        print("Portfolio page loaded.")

        return driver

    except NoSuchElementException:
        raise Exception(f"Unable to locate portfolio with name: {PORTFOLIO}")


def load_report24(driver: WebDriver) -> WebDriver:
    try:
        REPORT_NAME = "pvar-tile-REPORT_RPM_RES24"
        click_element(driver,By.ID,REPORT_NAME)
        print("Report url opened")
        time.sleep(20)
        new_window = driver.window_handles[1]
        driver.switch_to.window(new_window)
        iframe0 = driver.find_element(By.ID,"ctl00_contentBodyMaster_iframeRS")
        driver.switch_to.frame(iframe0)
        iframe1 = driver.find_element(By.TAG_NAME,"iframe")
        driver.switch_to.frame(iframe1)
        driver.maximize_window()
        print("iframe located for report data")
        click_element(driver, By.ID,"ReportViewerControl_ToggleParam")

        # complete report form
        start_date_input = driver.find_element(By.ID,"ReportViewerControl_ctl04_ctl11_txtValue")
        end_date_input = driver.find_element(By.ID,"ReportViewerControl_ctl04_ctl13_txtValue")

        start_date_input.send_keys(Keys.CONTROL + "a", Keys.DELETE)
        start_date_input.send_keys(START_DATE)
        end_date_input.send_keys(Keys.CONTROL + "a", Keys.DELETE)
        end_date_input.send_keys(END_DATE)
        click_element(driver, By.ID,"ReportViewerControl_ctl04_ctl27_ddValue")
        click_element(driver, By.ID,"ReportViewerControl_ctl04_ctl15_txtValue")  # open dropdown
        click_element(driver, By.ID,"ReportViewerControl_ctl04_ctl15_divDropDown_ctl00")  # select all timesheet types
        click_element(driver, By.ID,"ReportViewerControl_ctl04_ctl15_txtValue")  # close dropdown
        time.sleep(1)
        click_element(driver, By.ID,"ReportViewerControl_ctl04_ctl00")  # submit form

        return driver

    except NoSuchElementException:
        raise Exception(f"Unable to load report 24")


def download_report(driver):
    driver.switch_to.default_content()
    iframe0 = driver.find_element(By.ID,"ctl00_contentBodyMaster_iframeRS")
    driver.switch_to.frame(iframe0)
    iframe1 = driver.find_element(By.CSS_SELECTOR,"[title='Report Viewer']")
    driver.switch_to.frame(iframe1)

    print("waiting for report loading")
    wait_element(driver, (By.ID,"ReportViewerControl_ctl05_ctl04_ctl00_Button"),60)

    for i in range(0, 30):
        time.sleep(1)
        save_link = driver.find_element(By.ID,"ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink")
        style = save_link.get_attribute("style")
        if style.rfind("pointer") > 0:

            click_element(driver, By.ID,"ReportViewerControl_ctl05_ctl04_ctl00_Button")
            click_element(driver,By.CSS_SELECTOR,'[title="Excel"]')
            print("downloading excel")
            break
        elif i == 29:
            raise Exception("Failed to load report")


def store_report(report_name: str,download_dir):
    max_download_time = 60

    # Find the latest downloaded file
    cnt = 0
    os.remove(os.path.join(download_dir, report_name))
    files_before_dl = os.listdir(download_dir)
    while cnt < 60:
        cnt += 1
        files = os.listdir(download_dir)
        if len(files) > len(files_before_dl):
            latest_file = set(files).intersection().difference(files_before_dl).pop()
            print("Latest downloaded file:", latest_file)
            file_full_path = os.path.join(download_dir, latest_file)
            target_file_path = os.path.join(download_dir, report_name)
            shutil.copy(file_full_path, target_file_path)
            break
        time.sleep(1)

    if cnt == max_download_time:
        raise Exception(f"failed to retrieve downloaded file in one minute")



def wait_element(driver, query: (str, str), max_timeout=10) -> WebElement:
    wait = WebDriverWait(driver,max_timeout)
    try:
        element = wait.until(EC.visibility_of_element_located(query))
        return element
    except:
        raise Exception("Failed to load element: "+query[1])


def click_element(driver, by_type, query_data: str, timeout=10):
    wait = WebDriverWait(driver, timeout)
    try:
        wait.until(EC.visibility_of_element_located((by_type,query_data)))
    except:
        raise Exception("Failed to load element: "+query_data)

    for _ in range(0, 10):
        try:
            time.sleep(1)
            element = driver.find_element(by_type,query_data)
            ActionChains(driver).move_to_element(element).perform()
            element.click()
            return
        except ElementClickInterceptedException:
            time.sleep(1)

    raise Exception("Click failed due to timeout: "+query_data)


def main(login: str, password: str):
    print("Extracting the team planning file")
    driver = webdriver.Chrome(options=chrome_options)
    driver = home_page(driver, login, password)
    driver = load_portfolio(driver)
    driver = load_report24(driver)
    download_report(driver)
    store_report(REPORT_NAME,DOWNLOAD_DIR)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python script.py <username> <password>")
        sys.exit(1)

    main(sys.argv[1], sys.argv[2])
