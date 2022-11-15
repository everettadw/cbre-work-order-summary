from dotenv import load_dotenv
from datetime import datetime
import sys
import os
from custom_webdriver import CustomChromeWebDriver


import eam
from email_driver import OutlookClient
import reports
import scopesuite


# determine the environment
if os.name == "nt" and os.path.join(
        os.path.basename(os.path.dirname(sys.executable)),
        os.path.basename(sys.executable)) == os.path.join("Scripts", "python.exe"):
    ROOT_PATH = os.path.dirname(
        os.path.dirname(os.path.dirname(sys.executable)))
else:
    ROOT_PATH = os.path.dirname(sys.executable)

# application constants
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
    generate_reports(DOWNLOAD_PATH)

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


def generate_reports(path, date=datetime.today()):
    """
    Generates and formats several reports for a given date based
    on a source file of open work orders. Those reports are then
    stored in the WebDriver's downloads directory.
    """

    sheet_to_columns = {
        "Due by Midnight": [
            'Work Order',
            'Equipment',
            'Description',
            '1 - WO Owner',
            'Type',
            'Sched. Start Date',
            'PM Compliance Max',
            'Reported By'
        ],
        "Scheduled for Shift": [
            'Work Order',
            'Equipment',
            'Description',
            '1 - WO Owner',
            'PM Compliance Max'
        ],
        "Backlog": [
            'Work Order',
            'Equipment',
            'Description',
            '1 - WO Owner',
            'Type',
            'Sched. Start Date',
            'PM Compliance Max',
            'Reported By'
        ]
    }

    try:
        SOURCE_PATH = os.path.join(
            path, 'Sheet1.xlsx')
        TARGET_PATH = os.path.join(
            path, f'WO Summary {date.strftime("%#m-%#d-%Y")}.xlsx')

        reports.generate_work_order_summary(
            SOURCE_PATH, TARGET_PATH, sheet_to_columns, date)
        reports.format_work_order_summary(
            TARGET_PATH, sheet_to_columns)
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
    main()
