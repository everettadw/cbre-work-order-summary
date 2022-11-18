# necessary imports
from time import sleep
from config import WorkOrderSummaryConfig as wos_config
from custom_webdriver import CustomChromeWebDriver
from utils import isin_production
from dotenv import load_dotenv
import scopesuite
import reports
import eam
import sys
import os


# set the root path for the application depending
#   on the current environment
if isin_production():
    ROOT_PATH = os.path.dirname(sys.executable)
else:
    ROOT_PATH = os.path.dirname(
        os.path.dirname(os.path.dirname(sys.executable)))


# application global constants
DRIVER_PATH = os.path.join(ROOT_PATH, 'chromedriver', 'chromedriver.exe')
DOWNLOAD_PATH = os.path.join(ROOT_PATH, 'downloads')
ENV_PATH = os.path.join(ROOT_PATH, '.env')

# load environment variables
load_dotenv(dotenv_path=ENV_PATH)

# setup the webdriver and email instances
CHROME = CustomChromeWebDriver(
    driver_path=DRIVER_PATH,
    download_path=DOWNLOAD_PATH
)
# EMAIL = OutlookClient() # throws an exception is you don't have outlook installed


def main():
    """
    Main function that is run when 'app.py' is directly run.
    """

    # do the webdriver portion of the automation process
    with CHROME:
        # submit_grades(webdriver=CHROME)
        # scrape_work_orders(webdriver=CHROME)
        CHROME.get("https://www.google.com")
        sleep(2)

    # do the excel portion of the automation process
    generate_reports(
        path=DOWNLOAD_PATH,
        config=wos_config)

    # send the reports generated in the last step
    # send_reports(email=EMAIL)


def scrape_work_orders(webdriver):
    """
    Downloads a file of all open work orders.
    """

    try:
        eam.scrape_work_orders(webdriver)
    except Exception as e:
        webdriver.quit()
        print(f"\nError Scraping Work Orders:\n{repr(e)}\n")
        __ = input("Press Enter to continue...")


def submit_grades(webdriver, username="DEVERDAN"):
    """
    Submits a 100% grade for today in ScopeSuite with all strengths.
    """

    try:
        scopesuite.submit_grades(webdriver, username)
    except Exception as e:
        webdriver.quit()
        print(f"\nError Submitting Grades:\n{repr(e)}\n")
        __ = input("Press Enter to continue...")


def generate_reports(path, config):
    """
    Generates and formats several reports for a given date based
    on a source file of open work orders. Those reports are then
    stored in the WebDriver's downloads directory.
    """

    try:
        SOURCE_PATH = os.path.join(
            path,
            config.SOURCE_FILE_NAME
        )
        TARGET_PATH = os.path.join(
            path,
            config.TARGET_FILE_NAME
        )

        reports.generate_work_order_summary(
            SOURCE_PATH,
            TARGET_PATH,
            config.COLUMN_LAYOUT,
            config.TARGET_DATE
        )
        reports.format_work_order_summary(
            TARGET_PATH,
            config.COLUMN_LAYOUT
        )
    except Exception as e:
        print(f"\nError Generating Reports:\n{repr(e)}\n")
        __ = input("Press Enter to continue...")


def fill_shift_template(template_file_path, shift_summary_path, username='DEVERDAN'):
    """Copy work for the shift over to the given shift template."""

    pass


def send_reports(email):
    """
    Send reports located in the WebDriver's downloads directory to where they need to go.
    """

    pass


if __name__ == "__main__":
    main()
