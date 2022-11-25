from datetime import datetime
import os

from automations.scopesuite import submit_grades
from automations.reports import generate_work_order_summary, format_work_order_summary
from automations.infor import scrape_work_orders
from evsauto.config import TOMLWrapper
from evsauto.webdriver.chrome import CustomChromeWebDriver


def main():
    # init the config wrapper and get the config as a dict
    config_wrapper = TOMLWrapper('config.toml')
    config = config_wrapper.get_dict()
    settings = config['settings']

    # setup webdriver wrapper
    chrome_driver_wrapper = CustomChromeWebDriver(
        driver_path=os.path.join(
            os.getcwd(), "chromedriver", "chromedriver.exe"),
        download_path=os.path.join(os.getcwd(), "downloads"),
    )

    # webdriver magic
    if settings['webdriver']:
        with chrome_driver_wrapper as chrome:
            if settings['scopesuite']:
                submit_grades(webdriver_wrapper=chrome, username="DEVERDAN")
            if settings['scrape-infor']:
                scrape_work_orders(webdriver_wrapper=chrome)

    # spreadsheet magic
    if settings['reports']:
        generate_reports(
            path=chrome_driver_wrapper.download_path,
            config=config,
        )


def generate_reports(path, config):
    """
    Generates and formats several reports for a given date based
    on a source file of open work orders. Those reports are then
    stored in the WebDriver's downloads directory.
    """

    wos_config = config['work-order-summary']

    target_date = datetime.today(
    ) if wos_config['date'] == 'today' else datetime.strptime(wos_config['date'], "%m/%d/%Y")
    columns = wos_config['columns']

    column_layout = {}
    for col in columns:
        column_layout[col['title']] = col['keep-fields']

    target_file_name = wos_config['target-file-name'].replace(
        '{{|d}}', target_date.strftime("%#m-%#d-%Y"))

    SOURCE_PATH = os.path.join(path, wos_config['data-file-name'])
    TARGET_PATH = os.path.join(path, target_file_name)

    generate_work_order_summary(
        SOURCE_PATH,
        TARGET_PATH,
        column_layout,
        target_date,
    )
    format_work_order_summary(
        TARGET_PATH,
        column_layout,
        config['infor-credentials']['username']
    )


if __name__ == "__main__":
    main()
