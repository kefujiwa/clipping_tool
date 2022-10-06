import os
import re
import sys
import json
import datetime
import requests
import gspread
from time import sleep
from urllib.parse import urlparse
from oauth2client.service_account import ServiceAccountCredentials

from search_all_ranking_browser import launch_driver, search_ranking_browser, search_metadata

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
def get_initial_data(gc):
    sheet = gc.worksheet('Projects')
    data = sheet.get_all_values()
    first = data.pop(0)
    cnt = int(re.search(f'\d*$', first[0]).group())
    data.pop(0)

    data = data[0:cnt]
    present = []
    config = {}
    for g in gc.worksheets():
        if g.title == 'Projects':
            config['projects'] = g
        elif g.title == 'Media':
            config['media'] = g
        elif g.title == 'Template':
            config['template'] = g
        else:
            present.append(g)

    projects = []
    for d in data:
        exists = False
        for p in present:
            if d[0] == p.title:
                exists = True
                projects.append({
                    'sheet': p,
                    'keyword': d[1],
                    'date': d[2] if d[2] else '2000-01-01'
                })
                break
        if not exists:
            new = gc.duplicate_sheet(source_sheet_id=config['template'].id, new_sheet_name=d[0])
            new.update('A1', f'■【　{d[0]}　】記事掲載一覧')
            present.append(new)
            projects.append({
                'sheet': new,
                'keyword': d[1],
                'date': d[2] if d[2] else '2000-01-01'
            })

    order = []
    order.append(config['projects'])
    order.extend(present)
    order.append(config['media'])
    order.append(config['template'])
    gc.reorder_worksheets(order)

    return projects


def get_url_list(sheet):
    url_list = sheet.col_values(4)
    url_list.pop(0)
    url_list.pop(0)
    for u in url_list:
        yield f'{urlparse(u).hostname}{urlparse(u).path}'


def write_data(sheet, data):
    last_row = int(sheet.acell('C1').value) + 3

    output = []
    for d in data:
        lst = d['keyword'].split()
        status = '○'
        for e in lst:
            if not re.search(e, d['title']):
                status = '×'
                break
        output.append([
            '=ROW()-2',
            f"=VLOOKUP(\"{d['domain']}\", Media!A1:B, 2, FALSE)",
            d['date'],
            d['url'],
            d['titlelink'],
            d['title'],
            d['description'],
            status
        ])
    sheet.update(f'A{last_row}:H{last_row+len(data)}', output, value_input_option='USER_ENTERED')


### main_script ###
if __name__ == '__main__':
    SPREADSHEET_ID = '1dQuyZcm1RRqBz31hBXIKZbmBZl_kH209aA1nSKAnWR4'
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('../spreadsheet.json', scope)
    gc = gspread.authorize(credentials).open_by_key(SPREADSHEET_ID)

    logger.info('\n----- Start getting initial data -----\n')
    projects = get_initial_data(gc)
    logger.info(f'Projects: size: {len(projects)}')

    logger.info('\n----- Start search ranking -----\n')
    driver = launch_driver()
    try:
        for i, p in enumerate(projects):
            logger.info(f"{i}: {p['sheet'].title}")
            url_list = list(get_url_list(p['sheet']))

            res = search_ranking_browser(driver, p['keyword'], p['date'])
            if not res:
                continue

            data = []
            for v in res.values():
                u = None
                if v['url']:
                    u = f"{urlparse(v['url']).hostname}{urlparse(v['url']).path}"
                if u and not u in url_list:
                    data.append({
                        'keyword': p['keyword'],
                        'date': v['date'],
                        'url': v['url'],
                        'domain': urlparse(v['url']).hostname,
                        'titlelink': v['title']
                    })
            logger.debug(f"Total: {len(data)} URL")

            index, output = search_metadata(driver, data, logger, 0, len(data), 0)
            if len(output) > 0:
                write_data(p['sheet'], output)

        os.remove(f'log/{today.strftime("%Y-%m-%d")}_result.log')
        driver.close()
        driver.quit()
        exit(0)
    except Exception as e:
        driver.close()
        driver.quit()
        logger.error(f'Error: {e}')
        exit(1)