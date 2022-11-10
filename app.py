from custom_webdriver import CustomChromeWebDriver
from email_driver import OutlookClient
from dotenv import load_dotenv
from datetime import datetime
import sys
import os


import eam
import reports
import scopesuite


# application constants
APPLICATION_PATH = os.path.dirname(sys.executable)
ROOT_PATH = os.path.dirname(os.path.dirname(APPLICATION_PATH))
DRIVER_PATH = os.path.join(ROOT_PATH, 'chromedriver', 'chromedriver.exe')
DOWNLOAD_PATH = os.path.join(ROOT_PATH, 'downloads')
ENV_PATH = os.path.join(ROOT_PATH, '.env')


def main():
    """
    Main function that is run when 'app.py' is directly run.
    """

    # load environment variables
    load_dotenv(dotenv_path=ENV_PATH)

    # setup the webdriver and email instances
    chrome = CustomChromeWebDriver(
        driver_path=DRIVER_PATH,
        download_path=DOWNLOAD_PATH
    )
    email = OutlookClient()

    # do the webdriver portion of the automation process
    # submit_grades(chrome)
    scrape_work_orders(chrome)

    chrome.quit()  # close the webdriver

    # do the excel portion of the automation process
    generate_reports(datetime.today(), DOWNLOAD_PATH)

    # send the reports generated in the last step
    send_reports(email)


def scrape_work_orders(chrome):
    """
    Downloads a file of all open work orders.
    """

    try:
        eam.scrape_work_orders(chrome)
    except Exception as e:
        chrome.quit()
        print(f"\nError Scraping Work Orders:\n{repr(e)}\n")
        __ = input("Press Enter to continue...")


def submit_grades(chrome):
    """
    Submits a 100% grade for today in ScopeSuite with all strengths.
    """

    try:
        scopesuite.submit_grades(chrome)
    except Exception as e:
        chrome.quit()
        print(f"\nError Submitting Grades:\n{repr(e)}\n")
        __ = input("Press Enter to continue...")


def generate_reports(date, path):
    """
    Generates and formats several reports for a given date based
    on a source file of open work orders. Those reports are then
    stored in the WebDriver's downloads directory.
    """

    try:
        SOURCE_PATH = os.path.join(path, 'Sheet1.xlsx')
        MAX_COMPLIANCE_REPORT_TARGET_PATH = os.path.join(
            path, f'Max Compliance {date.strftime("%#m-%#d-%Y")}.xlsx')
        FOLLOW_UPS_REPORT_TARGET_PATH = os.path.join(
            path, f'Follow-Ups {date.strftime("%#m-%#d-%Y")}.xlsx')

        reports.generate_max_compliance(
            SOURCE_PATH, MAX_COMPLIANCE_REPORT_TARGET_PATH, date)
        reports.generate_follow_ups(SOURCE_PATH, FOLLOW_UPS_REPORT_TARGET_PATH)

        reports.format(MAX_COMPLIANCE_REPORT_TARGET_PATH)
        reports.format(FOLLOW_UPS_REPORT_TARGET_PATH,
                       delete_min_max_compliance=True)
    except Exception as e:
        print(f"\nError Generating Reports:\n{repr(e)}\n")
        __ = input("Press Enter to continue...")


def send_reports(email):
    """
    Send reports located in the WebDriver's downloads directory to where they need to go.
    """

    pass


# if this file is directly run, call main()
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nError in Main Function:\n{repr(e)}\n")
        __ = input("Press Enter to continue...")
