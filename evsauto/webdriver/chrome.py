from time import sleep

from evsauto.utils import get_file_count
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait


class CustomChromeWebDriver:
    """
    A class that wraps a selenium webdriver in order to
    improve code readability and encourage rapid prototyping.
    """

    def __init__(self, **kwargs):
        # setup variables for instance
        self.driver_path = kwargs.pop('driver_path', None)
        self.download_path = kwargs.pop('download_path', None)
        self.default_timeout = kwargs.pop('default_timeout', 30)
        self.max_timeout = kwargs.pop('max_timeout', 300)

        # ensure a driver_path was supplied
        if not self.driver_path:
            raise AttributeError(
                'Expected a driver path, but did not get one.')

        # setup webdriver options and preferences
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.options.add_argument("--log-level=3")
        self.options.add_experimental_option(
            "excludeSwitches", ['enable-logging'])

        if self.download_path is not None:
            prefs = {
                "download.default_directory": self.download_path,
            }
            self.options.add_experimental_option("prefs", prefs)

    def __enter__(self):
        # configure and assign a webdriver to self
        service = Service(executable_path=self.driver_path)
        self.driver = webdriver.Chrome(service=service, options=self.options)

        self.actions = ActionChains(self.driver)
        self.wait = WebDriverWait(self.driver, self.default_timeout)

        return self

    def __exit__(self, exc_type, exc_value, exc_trace):
        # close the webdriver
        self.driver.quit()
        del self.driver
        del self.actions
        del self.wait

    def get(self, url):
        """
        Navigate to the given URL on the window/frame in focus.
        """
        self.driver.get(url)

    def new_tab(self, url):
        """
        Open the given URL in a new tab and apply focus.
        """
        self.driver.execute_script(f"window.open('{url}');")
        self.driver.switch_to.window(
            self.driver.window_handles[len(self.driver.window_handles) - 1])

    def close_all_tabs(self) -> None:
        """
        Close all tabs except for the first one opened.
        """
        for i in range(len(self.driver.window_handles) - 1, -1, -1):
            if i == 0:
                self.driver.switch_to.window(self.driver.window_handles[i])
                break
            self.driver.switch_to.window(self.driver.window_handles[i])
            self.driver.close()

    def find_element(self, selector, type="presence"):
        element = None
        if type == "presence":
            element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, selector)))
        elif type == "clickable":
            element = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, selector)))
        return element

    def find_elements(self, selector, type="presence"):
        elements = None
        if type == "presence":
            __ = self.wait.until(
                EC.presence_of_element_located((By.XPATH, selector)))
            elements = self.driver.find_elements('xpath', selector)
        elif type == "clickable":
            __ = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, selector)))
            elements = self.driver.find_elements('xpath', selector)
        return elements

    def click_element(self, element, javascript_click=False):
        if javascript_click:
            self.driver.execute_script('arguments[0].click();', element)
        else:
            self.actions.click(element).perform()

    def type_into_element(self, element, content, use_javascript=False, send_enter=False):
        if use_javascript:
            self.driver.execute_script(
                f"arguments[0].setAttribute('value', '{content}');", element)
        else:
            self.actions.click(element).send_keys(content).perform()
            if send_enter:
                self.actions.click(element).send_keys(Keys.RETURN).perform()

    def wait_until_downloaded(self, count) -> bool:
        """
        Waits until 'count' downloaded files
        exist before the timeout is reached. Returns
        False if the timeout is exceeded. Returns True if
        'count' files exist.
        """
        if not self.download_path:
            raise AttributeError(
                "A download path must be specified in order to call wait_until_downloaded.")

        download_timeout = 0
        while download_timeout < self.max_timeout and get_file_count(self.download_path) < count:
            download_timeout += 1
            sleep(1)
        if download_timeout >= self.max_timeout:
            return False
        return True
