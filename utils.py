import os
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials


### functions ###
def get_initial_data():
    SPREADSHEET_ID = '1dQuyZcm1RRqBz31hBXIKZbmBZl_kH209aA1nSKAnWR4'
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('./spreadsheet.json', scope)
    gc = gspread.authorize(credentials).open_by_key(SPREADSHEET_ID)

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
                    'target': d[2]
                })
                break
        if not exists:
            new = gc.duplicate_sheet(source_sheet_id=config['template'].id, new_sheet_name=d[0])
            new.update('A1', f'■【　{d[0]}　】記事掲載一覧')
            present.append(new)
            projects.append({
                'sheet': new,
                'keyword': d[1],
                'target': d[2]
            })

    order = []
    order.append(config['projects'])
    order.extend(present)
    order.append(config['media'])
    order.append(config['template'])
    gc.reorder_worksheets(order)

    return projects

