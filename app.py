from datetime import datetime, timedelta
import os
import sys

from automations.scopesuite import submit_grades
from automations.reports import generate_work_order_summary, format_work_order_summary
from automations.infor import scrape_work_orders
from evsauto.config import TOMLWrapper
from evsauto.utils import is_executable
from evsauto.webdriver.chrome import CustomChromeWebDriver


def main():
    os.system("cls")
    print(" [ EvsAuto ] > Starting tool")

    print(" [ EvsAuto ] > Generating absolute paths")
    # define the applcation path
    app_path = os.path.dirname(sys.executable)
    if not is_executable():
        app_path = os.path.dirname(app_path)
        app_path = os.path.dirname(app_path)

    print(" [ EvsAuto ] > Loading config")
    # init the config wrapper and get the config as a dict
    config_wrapper = TOMLWrapper(os.path.join(app_path, 'config.toml'))
    config = config_wrapper.get_dict()

    print(" [ EvsAuto ] > Initializing webdriver for chrome")
    # setup webdriver wrapper
    chrome_driver_wrapper = CustomChromeWebDriver(
        driver_path=os.path.join(
            app_path,
            "chromedriver",
            "chromedriver.exe"
        ),
        download_path=os.path.join(app_path, "downloads"),
    )

    # webdriver magic
    settings = config['settings']
    if settings['webdriver'] and (settings['scopesuite'] or settings['scrape-infor']):
        print(" [ EvsAuto ] > Starting webdriver")
        with chrome_driver_wrapper as chrome:
            if settings['scopesuite']:
                print(" [ EvsAuto ] > Starting scopesuite grade submission")
                ss_creds = config['scopesuite-credentials']
                for user in ss_creds:
                    submit_grades(
                        webdriver_wrapper=chrome,
                        username=ss_creds[user]['username'],
                        password=ss_creds[user]['password']
                    )
            if settings['scrape-infor']:
                print(" [ EvsAuto ] > Scraping INFOR")
                scrape_work_orders(
                    webdriver_wrapper=chrome,
                    data_file_name=config['work-order-summary']['data-file-name'],
                    username=config['infor-credentials']['username'],
                    password=config['infor-credentials']['password'],
                )

    # spreadsheet magic
    if settings['reports']:
        generate_reports(
            path=chrome_driver_wrapper.download_path,
            config=config,
        )

    __ = input(" [ EvsAuto ] > Done! Press enter to continue...")
    os.system("cls")


def generate_reports(path, config):
    """
    Generates and formats several reports for a given date based
    on a source file of open work orders. Those reports are then
    stored in the WebDriver's downloads directory.
    """

    wos_config = config['work-order-summary']

    target_date = file_name_date = (datetime.today() if wos_config['date'] == 'today'
                                    else datetime.strptime(wos_config['date'], "%m/%d/%Y"))
    if wos_config['night-shift']:
        target_date = target_date + timedelta(days=1)

    column_layout = {}
    for col in wos_config['columns']:
        column_layout[col['title']] = col['keep-fields']

    target_file_name = wos_config['target-file-name'].replace(
        '{{|d}}', file_name_date.strftime("%#m-%#d-%Y"))

    SOURCE_PATH = os.path.join(path, wos_config['data-file-name'])
    TARGET_PATH = os.path.join(path, target_file_name)

    print(" [ EvsAuto ] > Generating report from source data")
    generate_work_order_summary(
        SOURCE_PATH,
        TARGET_PATH,
        column_layout,
        target_date
    )
    print(" [ EvsAuto ] > Formatting report")
    format_work_order_summary(
        TARGET_PATH,
        column_layout,
        config['infor-credentials']['username']
    )


if __name__ == "__main__":
    main()
