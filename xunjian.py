from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, WebDriverException
from docx import Document
from docx.shared import Pt, Inches
import os
import time
import re

def sanitize_filename(name):
    name = name.strip()
    return re.sub(r'[\\/*?:"<>|]', "_", name).replace(".", "_")

def she():
    q1 = Options()
    q1.add_argument('--no-sandbox')
    q1.add_argument('--disable-notifications')
    q1.add_experimental_option('detach', True)
    prefs = {
        "translate_whitelists": {
            "zh": "en",
            "ja": "en",
            "ko": "en",
            "fr": "en",
            "de": "en",
            "es": "en",
            "ru": "en",
            "pt": "en",
            "it": "en",
            "ar": "en",
            "nl": "en",
            "tr": "en",
            "vi": "en",
            "pl": "en",
            "th": "en",
            "sv": "en"
        },
        "translate": {
            "enabled": "true"
        }
    }
    q1.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(service=Service('chromedriver.exe'), options=q1)


# 新建基础图片目录
base_dirs = ['xunjian']
for path in base_dirs:
    os.makedirs(path, exist_ok=True)

print("粘贴网站（然后按两次回车）：")
raw_data = []
while True:
    line = input()
    if line.strip() == "":
        break
    raw_data.append(line.strip())

domains, lines = [], []

for row in raw_data:
    parts = row.split("\t")
    if len(parts) >= 1:
        domain = parts[0].strip()
        domains.append(domain)
        lines.append(f"https://www.{domain}/")
    else:
        print(f"⚠️ 行格式错误，已跳过：{row}")

browser = she()

for index, url in enumerate(lines):
    domain = domains[index]
    safe_name = sanitize_filename(domain)
    pictures_dir = os.path.join("xunjian", safe_name)
    os.makedirs(pictures_dir, exist_ok=True)
    try:
        browser.get(url)
        browser.maximize_window()
        time.sleep(1)
    except WebDriverException as e:
        print(f"❌ 无法访问 {url}：{str(e)}")
        continue

    browser.get(url)
    browser.set_window_position(0, 0)
    browser.maximize_window()

    # Domain 截屏
    filename = os.path.join(pictures_dir, f"{safe_name}_a.png")
    browser.execute_script("document.body.style.zoom='50%'")
    time.sleep(2)
    browser.get_screenshot_as_file(filename)


    # Product 页面截一张
    product = ['ALL', 'all', 'All', 'NEW', 'new', 'New', 'SALE', 'sale', 'Sale',
               'SHOP', 'shop', 'Shop', 'SALE', 'sale', 'Sale', 'MORE', 'more', 'More', 'DRESS', 'dress', 'Dress',
               'ARRIVAL',
               'arrival', 'Arrival', 'SELL', 'sell', 'Sell', 'COLLECTION', 'collection', 'Collection', 'PRODUCT',
               'product', 'Product']
    for text in product:
        try:
            elem = browser.find_element(By.PARTIAL_LINK_TEXT, text)
            if elem.tag_name.lower() == 'a' and elem.get_attribute('href'):
                browser.execute_script("arguments[0].click();", elem)

                filename = os.path.join(pictures_dir, f"{safe_name}_b.png")
                browser.execute_script("document.body.style.zoom='50%'")
                time.sleep(2)
                browser.get_screenshot_as_file(filename)

                browser.get(url)
                break
        except NoSuchElementException:
            continue

    # Terms 页面截一张
    term = ['TERM', 'term', 'Term', 'AGREEMENT', 'Agreement', 'agreement', 'GENERAL', 'General', 'general', 'CONDITION',
            'condition', 'Condition']
    for text in term:
        try:
            elem = browser.find_element(By.PARTIAL_LINK_TEXT, text)
            if elem.tag_name.lower() == 'a' and elem.get_attribute('href'):
                browser.execute_script("arguments[0].click();", elem)

                filename = os.path.join(pictures_dir, f"{safe_name}_d.png")
                browser.execute_script("document.body.style.zoom='50%'")
                time.sleep(1)
                browser.get_screenshot_as_file(filename)

                browser.get(url)
                break
        except NoSuchElementException:
            continue


    # CONTACT 页面
    contact = ['CONTACT', 'contact', 'Contact', 'SUPPORT', 'support', 'Support', 'CUSTOMER', 'customer', 'Customer','COUTACT', 'coutact', 'Coutact','US', 'us', 'Us']
    for text in contact:
        try:
            elem = browser.find_element(By.PARTIAL_LINK_TEXT, text)
            if elem.tag_name.lower() == 'a' and elem.get_attribute('href'):
                browser.execute_script("arguments[0].click();", elem)

                filename = os.path.join(pictures_dir, f"{safe_name}_c.png")
                browser.execute_script("document.body.style.zoom='50%'")
                time.sleep(1)
                browser.get_screenshot_as_file(filename)

                browser.get(url)
                break
        except NoSuchElementException:
            continue


browser.quit()