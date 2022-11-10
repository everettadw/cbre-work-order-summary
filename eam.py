from custom_webdriver import CustomChromeWebDriver
from utils import get_file_count
import os


def scrape_work_orders(chrome: CustomChromeWebDriver):
    """
    Downloads a report of PMs and PDMs that will go out of compliance at midnight.
    """

    # pull infor username and password from environment vars
    INFOR_USERNAME = os.getenv("INFOR_USERNAME")
    INFOR_PASSWORD = os.getenv("INFOR_PASSWORD")

    # open infor
    chrome.get("https://eam.amazon.com/web/base/logindisp")

    # find the username, password, and submit button elements
    username_input = chrome.find_element(
        '//input[@name="userid"]', 'clickable')
    password_input = chrome.find_element(
        '//input[@name="password"]', 'clickable')
    submit_button = chrome.find_element('//a[@role="button"]', 'clickable')

    # type in username & password and click submit
    chrome.type_into_element(username_input, INFOR_USERNAME)
    chrome.type_into_element(password_input, INFOR_PASSWORD)
    chrome.click_element(submit_button)

    # find the workorder type & preset filters
    work_order_type_filter = chrome.find_element(
        '//input[@name="filterfields"]', 'clickable')
    preset_filter = chrome.find_element(
        '//input[@name="dataspylist"]', 'clickable')

    # set search filters and search
    chrome.type_into_element(work_order_type_filter,
                             "Work Order",
                             send_enter=True)
    chrome.type_into_element(preset_filter,
                             "All Open WO's",
                             send_enter=True)

    # find the button that exports the results
    save_to_excel = chrome.find_element(
        "(//div[contains(@id, 'toolbar')])[6]/a[4]", 'clickable')

    # 1) count the number of downloaded files
    # 2) click the export results button
    # 3) wait until the number of downloaded files has
    #      increased by one
    file_count = get_file_count(chrome.download_path)
    chrome.click_element(save_to_excel, True)
    chrome.wait_until_downloaded(file_count + 1)
