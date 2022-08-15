import os
import re
import csv
import random
import datetime
import urllib.parse
from time import sleep
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException, TimeoutException
#from selenium.webdriver.support.ui import WebDriverWait
#from selenium.webdriver.support import expected_conditions as EC
from fake_useragent import UserAgent
from webdriver_manager.chrome import ChromeDriverManager

from by_pass_captcha import by_pass_captcha

### functions ###
def launch_driver():
    url = "https://www.google.co.jp/"

    ua = UserAgent()

    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-gpu')
    options.add_argument(f'user-agent={ua.chrome}')

    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    driver.get(url)
    #driver.maximize_window()
    return driver

def check_exists_by_class_name(driver, class_name):
    try:
        driver.find_element_by_class_name(class_name)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_xpath(driver, xpath):
    try:
        driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

def search_all_rank(ssl):
    # ページ解析と結果の出力
    ranking = {}
    rank = 1
    for item in ssl:
        ritem = {'url': '', 'title': ''}
        try:
            snippet = item.select('div.BNeawe.uEec3.AP7Wnd')[0].text
        except IndexError:
            snippet = None
        try:
            site = item.select('div.egMi0.kCrYT > a')[0]
            try:
                site_title = site.select('h3.zBAuLc')[0].text
            except IndexError:
                site_title = site.select('img')[0]['alt']
        except IndexError:
            if snippet and re.search('強調スニペットについて', snippet):
                site = item.select('div.kCrYT > a')[0]
                site_title = site.select('span.UMOHqf.EDgFbc')[0].text
            else:
                continue
        site_url = site['href']
        if not re.match('http', site_url):
            parse_result = urllib.parse.urlparse(site_url)
            query = parse_result.query
            dic = urllib.parse.parse_qs(query)
            site_url = dic['url'][0]
        try:
            ritem['url'] = site_url
        except Exception as err:
            ritem['url'] = 'Unrecognized'
        try:
            ritem['title'] = site_title
        except Exception as err:
            ritem['title'] = 'Unrecognized'
        ranking[f'{rank}'] = ritem
        rank += 1
    while rank <= 100:
        ranking[f'{rank}'] = {'url': None, 'title': None}
        rank += 1
    return ranking

def search_ranking_browser(driver, kw):
    try:
        # Googleから検索結果ページを取得する
        url = f'https://www.google.co.jp/search?hl=ja&num=100&q={kw}'
        driver.get(url)

        sleep(random.randint(1, 2))
        if check_exists_by_class_name(driver, 'g-recaptcha'):
            driver.implicitly_wait(5)
            driver.find_element_by_xpath('//iframe[@title="reCAPTCHA"]').click()
            sleep(random.randint(2, 4))
            ret = by_pass_captcha(driver)
            if ret == False:
                return None

        sleep(random.randint(1, 2))
        if check_exists_by_xpath(driver, '//input[@value="同意する"]'):
            driver.implicitly_wait(5)
            driver.find_element_by_xpath('//input[@value="同意する"]').click()
            sleep(random.randint(2, 4))
            
        soup = BeautifulSoup(driver.page_source, "html.parser")
        search_site_list = soup.select('div.Gx5Zad.fP1Qef.xpd.EtOod.pkphOe')
        return search_all_rank(search_site_list)
    except Exception as err:
        print(f'search_ranking: {err}')
        return None

def search_metadata(driver, data, logger, index, size, trial):
#    wait = WebDriverWait(driver=driver, timeout=30)

    try:
        black_list = ['twitter.com', 'books.google.co.jp', 'www.youtube.com', 'www.facebook.com']
        new = []
        while index < size:
            d = data[index]
            logger.debug(f"\t{index}: {d['url']}")

            if d['domain'] in black_list or re.match(r'.*(\.pdf|\.docx|\.xlsx)$', d['url']) or trial > 9:
                d['title'] = '-'
                d['description'] = '-'
                new.append(d)
                index += 1
                continue

            driver.execute_script('''window.open("","_blank");''')
            driver.switch_to.window(driver.window_handles[1])
            try:
                driver.set_page_load_timeout(12 + trial)
                driver.get(d['url'])
                if driver.title:
                    d['title'] = driver.title
                else:
                    d['title'] = '-'
                if check_exists_by_xpath(driver, "//meta[@name='description']"):
                    d['description'] = driver.find_element_by_xpath("//meta[@name='description']").get_attribute('content')
                elif check_exists_by_xpath(driver, "//meta[@property='og:description']"):
                    d['description'] = driver.find_element_by_xpath("//meta[@property='og:description']").get_attribute('content')
                else:
                    d['description'] = '-'
                new.append(d)
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                index += 1
                trial = 0
                sleep(1)
            except TimeoutException:
                logger.debug(f'\t\tTimeout: {trial + 1} time')
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                sleep(1)
                index, next_new = search_metadata(driver, data, logger, index, size, trial + 1)
                new.extend(next_new)
        
        return index, new
    except Exception as e:
        logger.error(f'search_metadata: {e}')
        return index, new