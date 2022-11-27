from pathlib import Path

from evsauto.utils import get_file_count
from evsauto.webdriver.chrome import CustomChromeWebDriver


def scrape_work_orders(webdriver_wrapper: CustomChromeWebDriver, data_file_name: str, username: str, password: str) -> None:
    """
    Downloads a report of PMs and PDMs that will go out of compliance at midnight.
    """

    # open infor
    webdriver_wrapper.get("https://eam.amazon.com/web/base/logindisp")

    # find the username, password, and submit button elements
    username_input = webdriver_wrapper.find_element(
        '//input[@name="userid"]', 'clickable')
    password_input = webdriver_wrapper.find_element(
        '//input[@name="password"]', 'clickable')
    submit_button = webdriver_wrapper.find_element(
        '//a[@role="button"]', 'clickable')

    # type in username & password and click submit
    webdriver_wrapper.type_into_element(username_input, username)
    webdriver_wrapper.type_into_element(password_input, password)
    webdriver_wrapper.click_element(submit_button)

    # find the workorder type & preset filters
    work_order_type_filter = webdriver_wrapper.find_element(
        '//input[@name="filterfields"]', 'clickable')
    preset_filter = webdriver_wrapper.find_element(
        '//input[@name="dataspylist"]', 'clickable')

    # set search filters and search
    webdriver_wrapper.type_into_element(work_order_type_filter,
                                        "Work Order",
                                        send_enter=True)
    webdriver_wrapper.type_into_element(preset_filter,
                                        "All Open WO's",
                                        send_enter=True)

    # find the button that exports the results
    save_to_excel = webdriver_wrapper.find_element(
        "(//div[contains(@id, 'toolbar')])[6]/a[4]", 'clickable')

    # delete the source file if it exists
    if Path(webdriver_wrapper.download_path, data_file_name).exists():
        Path(webdriver_wrapper.download_path, data_file_name).unlink()

    # 1) count the number of downloaded files
    # 2) click the export results button
    # 3) wait until the number of downloaded files has
    #      increased by one
    file_count = get_file_count(webdriver_wrapper.download_path)
    webdriver_wrapper.click_element(save_to_excel, True)
    webdriver_wrapper.wait_until_downloaded(file_count + 1)
