import os
import datetime
from time import sleep
from bs4 import BeautifulSoup
from utils import get_initial_data

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException 
from webdriver_manager.chrome import ChromeDriverManager
from fake_useragent import UserAgent

# Logger setting
from logging import getLogger, FileHandler, DEBUG
logger = getLogger(__name__)
today = datetime.datetime.now()
os.makedirs('./log', exist_ok=True)
handler = FileHandler(f'log/{today.strftime("%Y-%m-%d")}_result.log', mode='a')
handler.setLevel(DEBUG)
logger.setLevel(DEBUG)
logger.addHandler(handler)
logger.propagate = False

### functions ###
def launch_driver():
    url = "https://www.google.co.jp/"

    ua = UserAgent()
    logger.debug(f'\tUserAgent: {ua.chrome}\n')

    options = Options()
    options.add_argument('--headless')
    options.add_argument(f'user-agent={ua.chrome}')

    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    return driver


### functions ###
def get_url_list(sheet):
    url_list = sheet.get_all_values()
    url_list.pop(0)
    url_list.pop(0)
    for u in url_list:
        if u[7] == '○':
            yield [u[0], u[3]]


### main_script ###
if __name__ == '__main__':
    logger.info('\n----- Start getting initial data -----\n')
    projects = get_initial_data()
    logger.info(f'Projects: size: {len(projects)}')

    logger.info('\n----- Start getting screenshot -----\n')

    os.makedirs('./screenshots', exist_ok=True)
    try:
        for index, p in enumerate(projects):
            logger.info(f"{index}: {p['sheet'].title}")
            os.makedirs(f"./screenshots/{p['sheet'].title}", exist_ok=True)

            url_list = list(get_url_list(p['sheet']))
            logger.debug(f"Total: {len(url_list)} URL")
            for i, u in enumerate(url_list):
                logger.debug(f"\t{i}: No.{u[0]}: {u[1]}")
                path = f"./screenshots/{p['sheet'].title}/{u[0]}.png" 
                is_file = os.path.isfile(path)
                if is_file:
                    continue
                driver = launch_driver()
                driver.get(u[1])
                sleep(3)

                S = lambda X: driver.execute_script('return document.body.parentNode.scroll'+X)
                driver.set_window_size(S('Width') + 500,S('Height'))
                el = driver.find_element_by_tag_name('body')
                el.screenshot(path)
                driver.close()
                driver.quit()

        os.remove(f'log/{today.strftime("%Y-%m-%d")}_result.log')
    except Exception as err:
        driver.close()
        driver.quit()
        logger.debug(f'{err}')
        exit(1)
