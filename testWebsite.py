from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.keys import Keys
import time

def she():
    options = Options()
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-notifications')
    options.add_experimental_option('detach', True)
    return webdriver.Chrome(service=Service('chromedriver.exe'), options=options)

# 广告页面弹窗处理
def suspondWindowHandler(browser):
    # 第一种广告弹窗
    try:
        suspondWindow = browser.find_element(By.XPATH,"//div[contains(@class, 'identity-dialog')]//*[contains(@class, 'close-icon')]")
        if suspondWindow.tag_name.lower() == 'a' and suspondWindow.get_attribute('href'):
            suspondWindow.click()
        print(f"searchKey: Suspond Page1 had been closed.")
    except Exception as e:
        print(f"searchKey: there is no suspond Page1. e = {e}")
    # 第二种广告弹窗
    # 如果有广告界面弹出，关闭广告。 否则会导致数据无法输入到搜索框
    try:
        suspondWindow = browser.find_element(By.XPATH,"//div[contains(@class,'overlay-box')]//div[contains(@class,'overlay-close')]")
        suspondWindow.click()
        print(f"searchKey: Suspond Page2 had been closed.")
    except Exception as e:
        print(f"searchKey: there is no suspond Page2. e = {e}")

# 主程序
browser = she()
browser.get('gifttn.com')
browser.maximize_window()
print(browser.get_cookies())
browser.delete_all_cookies()
print(browser.get_cookies())
# 等页面稳定后截图
time.sleep(2)
suspondWindowHandler(browser)
browser.save_screenshot("firmoo_clean.png")