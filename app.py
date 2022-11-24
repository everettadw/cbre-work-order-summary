# necessary imports
from config import WorkOrderSummaryConfig as wos_config
from evsauto.webdriver.chrome import CustomChromeWebDriver
from evsauto.utils import isin_production
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
CHROME_DRIVER = CustomChromeWebDriver(
    driver_path=DRIVER_PATH,
    download_path=DOWNLOAD_PATH
)


def main():
    """Main function that is run when 'app.py' is directly run."""

    # do the webdriver portion of the automation process
    with CHROME_DRIVER as chrome:
        submit_grades(webdriver=chrome)
        scrape_work_orders(webdriver=chrome)

    # do the excel portion of the automation process
    generate_reports(
        path=DOWNLOAD_PATH,
        config=wos_config)


def scrape_work_orders(webdriver):
    """Downloads a file of all open work orders."""
    eam.scrape_work_orders(webdriver)


def submit_grades(webdriver, username="DEVERDAN"):
    """Submits a 100% grade for today in ScopeSuite with all strengths."""
    scopesuite.submit_grades(webdriver, username)


def generate_reports(path, config):
    """
    Generates and formats several reports for a given date based
    on a source file of open work orders. Those reports are then
    stored in the WebDriver's downloads directory.
    """

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


def fill_shift_template(template_file_path, shift_summary_path, username='DEVERDAN'):
    """Copy work for the shift over to the given shift template."""
    pass


def send_reports(email):
    """Send reports located in the WebDriver's downloads directory to where they need to go."""
    pass


if __name__ == "__main__":
    main()
