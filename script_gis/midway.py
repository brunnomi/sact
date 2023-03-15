import os
import sys
import time
from getpass import getpass
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def web_driver(headless: bool = False, log_lv: int = 3,
               desired_capabilities: str = None,
               user_data_dir: str = None,
               profile: str = None,
               ) -> Chrome:
    # Options
    options = Options()
    options.add_argument(f'log-level={log_lv}')
    if user_data_dir and profile:
        options.add_argument(f"--user-data-dir={user_data_dir}")
        options.add_argument(f"--profile-directory={profile}")
    # SSL Options
    capabilities = DesiredCapabilities.CHROME.copy()
    capabilities['acceptInsecureCerts'] = True
    # Location to download
    if desired_capabilities is not None:
        prefs = {"download.default_directory": desired_capabilities}
        options.add_experimental_option("prefs", prefs)
    # Headless setting
    if headless:
        options.add_argument('--headless')
    return Chrome(
        ChromeDriverManager().install(),
        options=options,
        desired_capabilities=capabilities,
    )


def mw_authentication_check(driver, url):
    start = datetime.now()
    while not (url in driver.current_url):
        end = datetime.now()
        if (end - start).seconds > 30:
            print("Authorization could not be confirmed. Please rerun the program.")
            return False
    print("Midway passed")
    return True


def midway(driver: webdriver, pin, url=None) -> None:
    if url is None:
        url = "https://midway.amazon.com/"
    driver.get(url)
    WebDriverWait(driver,timeout=60,poll_frequency=0.5).until(EC.presence_of_element_located((By.ID,"user_name")))
    driver.find_element(By.ID,"user_name").send_keys(os.getlogin())
    driver.find_element(By.ID,"password").send_keys(pin)
    driver.find_element(By.NAME,"commit").click()
    
    mw_authentication_check(driver, url)


def headless_midway(driver: webdriver, pin, url=None) -> None:
    if url is None:
        url = "https://midway.amazon.com/"
    driver.get(url)
    time.sleep(1)
    if url in driver.current_url:
        print("Midway authorized!")
        return None
    driver.find_element_by_name("commit").click()
    while not (otp_field := driver.find_elements_by_id('otp')):
        pass
    driver.find_element_by_id("user_name").send_keys(os.getlogin())
    driver.find_element_by_id("password").send_keys(pin)
    otp = input("Touch Yubikey: ")
    otp_field[0].send_keys(otp)
    driver.find_element_by_id("verify_btn").click()
    
    mw_authentication_check(driver, url)


if __name__ == "__main__":
    # TEST
    headless_mode = False
    d = web_driver(headless_mode,user_data_dir="C:/Users/brunnomi/AppData/Local/Google/Chrome/User",profile="Profile 2",)
    mw = headless_midway if headless_mode else midway
    pin = getpass("Enter your PIN > ")
    d.get(url="https://gisportal.aka.amazon.com/#/gisv2?clientIdDomain=3PCustomerSales&clientId=247532125543201")
    auth = mw_authentication_check(d,url="https://gisportal.aka.amazon.com/#/gisv2?clientIdDomain=3PCustomerSales&clientId=247532125543201")
    if not auth:
        mw(driver=d, pin=pin,url="https://gisportal.aka.amazon.com/#/gisv2?clientIdDomain=3PCustomerSales&clientId=247532125543201")
    # Please uncomment below if you have set the environment variable PIN.
    # mw(driver=d, pin=os.environ["PIN"])
    input("Did you complete your Midway authentication?")
    d.quit()


