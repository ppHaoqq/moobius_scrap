from selenium import webdriver
from bs4 import BeautifulSoup as bs
import pandas as pd
import time
import re
from selenium.webdriver.chrome.options import Options
import sys
from pathlib import Path, PurePath


options = Options()
# options.add_argument('--headless')

driver_pass = Path(r'C:\Users\g2945\chromedriver\chromedriver')
try:
    browser = webdriver.Chrome(driver_pass, options=options)
except FileNotFoundError:
    driver_pass = Path(r'C:\chromedriver_win32\chromedriver')
    browser = webdriver.Chrome(driver_pass, options=options)

browser.implicitly_wait(10)
browser.get('https://kibi-cloud.jp/moobius/User/login.aspx')
time.sleep(2)

id1 = browser.find_element_by_css_selector('#txtComID')
id2 = browser.find_element_by_css_selector('#txtLogID')
pw = browser.find_element_by_css_selector('#txtLogPW')
login_button = browser.find_element_by_css_selector('#btnLogin')

id1.send_keys('gaeart-shikoku')
id2.send_keys('g948')
pw.send_keys('g948')
time.sleep(1)

login_button.click()
time.sleep(5)

sekisanform = browser.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[1]/ul/li[7]')
sekisanform.click()
time.sleep(5)

frame = browser.find_element_by_id('mainFram')
browser.switch_to.frame(frame)

_search = browser.find_element_by_id('mainT')
search_btn = _search.find_element_by_id('btnSearchPop')
search_btn.click()
time.sleep(2)

word = '中ゾーン'
word2 = sys.argv
search_box = browser.find_element_by_id('detSerachBox_Name')
if len(word2) == 1:
    search_box.send_keys(word)
else:
    search_box.send_keys(word2[1])
time.sleep(2)
browser.find_element_by_id('btnDetSearch').click()
time.sleep(2)

_elem = browser.find_element_by_class_name('doclist')
elem = _elem.find_element_by_id('tbd-recKouji')
webdriver.ActionChains(browser).double_click(elem).perform()
time.sleep(2)

elem2 = browser.find_element_by_id('tab-menu')
elem3 = elem2.find_elements_by_tag_name('a')
elem3[4].click()
time.sleep(2)

frame2 = browser.find_element_by_id('mainTFram')
browser.switch_to.frame(frame2)

html = browser.page_source
soup = bs(html, 'html.parser')
selector = '#tl_List' + ' tr'
tr = soup.select(selector)
pattern1 = r'<t[h|d].*?>.*?</t[h|d]>'
pattern2 = r'''<(".*?"|'.*?'|[^'"])*?>'''
columns = [re.sub(pattern2, '', s) for s in re.findall(pattern1, str(tr[0]))]
data = [[re.sub(pattern2, '', s) for s in re.findall(pattern1, str(tr[i]))] for i in range(1, len(tr))]
df = pd.DataFrame(data=data, columns=columns)

save_pass = r'C:\Users\g2945\PycharmProjects\moobius_scrap\datas\{}_三宅積算.xlsx'.format(word)


try:
    df.to_excel(save_pass)
except :
    save_pass = 'C:/Users/daiki/PycharmProjects/moobius_scrap/datas/{}_三宅積算.xlsx'.format(word)
    df.to_excel(save_pass)
time.sleep(2)

browser.close()
browser.quit()