import re
from io import StringIO
from datetime import datetime
from html.parser import HTMLParser
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup 
import os
import glob
import csv
from xlsxwriter.workbook import Workbook

DRIVER_PATH = '../chromedriver'
ROOT_URL = 'http://data.stats.gov.cn/easyquery.htm?cn=A01'
ID_NY = 'treeZhiBiao_4_ico' # 能源
ID_NY_ZHUYAO = 'treeZhiBiao_16_ico' # 能源主要产品产量
ALL_NY_ITEMS = {
    'yuanmei': 17,                  # 原煤
    'yuanyou': 18,                  # 原油
    'tianranqi': 19,                # 天然气
    'meicengqi': 20,                # 煤层气
    'yehuatianranqi': 21,           # 液化天然气
    'yuanyoujiagongliang': 22,      # 原油加工量
    'qiyou': 23,                    # 汽油
    'meiyou': 24,                   # 煤油
    'chaiyou': 25,                  # 柴油
    'ranliaoyou': 26,               # 燃料油
    'shinaoyou': 27,                # 石脑油
    'yehuashiyouqi': 28,            # 液化石油气
    'shiyoujiao': 29,               # 石油焦
    'shiyouliqing': 30,             # 石油沥青
    'jiaotan': 31,                  # 焦炭
    'fadianliang': 32,              # 发电量
    'huoli': 33,                    # 火力发电量
    'shuili': 34,                   # 水力发电量
    'heneng': 35,                   # 核能发电量
    'fengli': 36,                   # 风力发电量
    'taiyangneng': 37,              # 太阳能发电量
    'meiqi': 38,                    # 煤气
}

class MLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.reset()
        self.strict = False
        self.convert_charrefs= True
        self.text = StringIO()
    def handle_data(self, d):
        self.text.write(d.replace(" ", ""))
    def get_data(self):
        return self.text.getvalue()

def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return s.get_data().strip()

def parse_table(table):
    contents = strip_tags(table) # remove all html tags of table
    months_start = contents.index('年') - 4 
    months_end = contents.rindex('月') + 1
    months = contents[months_start : months_end] 
    months_arr = months.split() # array of months in table
    n_months = len(months_arr) # number of months in table

    contents = contents[months_end:].strip().replace('\n\n\n', '\n')
    contents = contents.splitlines()

    table_title = contents[0]
    rows = {}
    stats_in_row = 0
    cur_row = ''
    
    for line in contents:
        if line == table_title: # the same title appears in every row, we don't want to see it repeatedly
            continue    
        elif len(line) == 0 and (stats_in_row == n_months or stats_in_row == 0): # useless newlines from innerHTML
            continue
        elif len(line) == 0: # a missing data field, not a useless newline
            rows[cur_row].append('NA')
            stats_in_row += 1
        elif not (line[0].isdigit() or line[0] == '-'): # not a data field (pos or neg number), must be row name
            rows[line] = []
            cur_row = line
            stats_in_row = 0
        else:   # a data field
            rows[cur_row].append(line)
            stats_in_row += 1

    # format as csv
    csv_string = table_title + '\n' + 'Date'
    for key in rows.keys():
        csv_string += ',' + str(key)

    csv_string += '\n'
    for i in range(n_months):
        csv_string += str(months_arr[i]) 
        for key in rows.keys():
            csv_string += ',' + rows[key][i]
        csv_string += '\n'

    # write to output file
    file_name = 'output_' + str(datetime.today())
    csv_file = file_name + '.csv'
    f = open(csv_file, 'a')
    f.write(csv_string)
    f.close()  

    # convert to xlsx
    xlsx_file = file_name + '.xlsx'
    workbook = Workbook(xlsx_file)
    worksheet = workbook.add_worksheet()
    with open(csv_file, 'rt', encoding='utf8') as xf:
        reader = csv.reader(xf)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
    workbook.close()



def scrape_table(item):
    table_id = 'table_container_main'
    wait = WebDriverWait(driver, 10)
    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, table_id)))
    table = driver.find_element_by_class_name(table_id).get_attribute('innerHTML')
    result = BeautifulSoup(table, 'html.parser').prettify()
    parse_table(result)


def get_ny_item(item):
    ny_item_id = 'treeZhiBiao_' + str(ALL_NY_ITEMS[item]) + '_ico'
    wait = WebDriverWait(driver, 10)
    wait.until(EC.element_to_be_clickable((By.ID, ny_item_id)))
    driver.find_element_by_id(ny_item_id).click()
    scrape_table(item)


def open_ny_list():
    # open 能源
    ny_wait = WebDriverWait(driver, 20)
    ny_wait.until(EC.element_to_be_clickable((By.ID, ID_NY)))
    driver.find_element_by_id(ID_NY).click()

    # open 能源主要产品产量
    ny_zhuyao_wait = WebDriverWait(driver, 10)
    ny_zhuyao_wait.until(EC.element_to_be_clickable((By.ID, ID_NY_ZHUYAO)))
    driver.find_element_by_id(ID_NY_ZHUYAO).click()
    # print(driver.page_source)
    print(driver.current_url)
    # driver.quit()


if __name__ == "__main__":
    options = Options()
    options.headless = True
    options.add_argument("--window-size=1920,1200")

    # driver = webdriver.Chrome(options=options, executable_path=DRIVER_PATH)
    driver = webdriver.Chrome(executable_path=DRIVER_PATH)
    driver.get(ROOT_URL)
    open_ny_list()
    get_ny_item('yuanyou')
    driver.quit()
