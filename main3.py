from selenium import webdriver
from bs4 import BeautifulSoup as bs
import pandas as pd
import time
import re
from selenium.webdriver.chrome.options import Options
import sys


def main():
    kw = input('工事名を入力してください：')
    options = Options()
    #options.add_argument('--headless')
    browser = webdriver.Chrome('C:\chromedriver_win32\chromedriver', options=options)
    browser.implicitly_wait(10)
    browser.get('https://kibi-cloud.jp/moobius/User/login.aspx')
    time.sleep(2)
    #login処理
    login(browser)
    #次ページへ
    sekisanform = browser.find_element_by_xpath('//*[@id="system_select"]/li[6]')
    sekisanform.click()
    time.sleep(5)
    #フレーム更新
    change_frame(browser, 'mainFram')
    #対象検索

    search_def(browser, kw)
    #フレーム更新
    change_frame(browser, 'mainTFram')
    #対象保存
    html = browser.page_source
    save_excel(html)
    #終了処理
    browser.close()
    browser.quit()


def login(browser):
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


def change_frame(browser, id_):
    frame = browser.find_element_by_id(id_)
    browser.switch_to.frame(frame)


def search_def(browser, kw):
    _search = browser.find_element_by_id('mainT')
    search_btn = _search.find_element_by_id('btnSearchPop')
    search_btn.click()
    time.sleep(2)

    search_box = browser.find_element_by_id('detSerachBox_Name')
    search_box.send_keys(kw)
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


#参考サイトhttps://tanuhack.com/selenium/#tablepandasDataFrame
def save_excel(html):
    soup = bs(html, 'html.parser')
    selector = '#tl_List' + ' tr'
    tr = soup.select(selector)
    pattern1 = r'<t[h|d].*?>.*?</t[h|d]>'
    pattern2 = r'''<(".*?"|'.*?'|[^'"])*?>'''
    columns = [re.sub(pattern2, '', s) for s in re.findall(pattern1, str(tr[0]))]
    data = [[re.sub(pattern2, '', s) for s in re.findall(pattern1, str(tr[i]))] for i in range(1, len(tr))]
    df = pd.DataFrame(data=data, columns=columns)
    df.to_excel('C:/Users/daiki/PycharmProjects/moobius_scrap/datas/sekisan.xlsx')


if __name__ == '__main__':
    main()