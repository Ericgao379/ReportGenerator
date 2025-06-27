from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time


# 启动 Chrome 浏览器
def she():
    q1 = Options()
    q1.add_argument('--no-sandbox')
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


# 获取用户输入的网址
print("粘贴要查询的网址：")
lines = []
unformatted = []
while True:
    line = input()
    if line == "":
        break
    unformatted.append(line)
    formatted = "https://www." + line.strip() + "/"
    lines.append(formatted)

# 部分不需要操作的点击的链接文本列表
links = ['RETURN', 'CONTACT', 'PRIVACY', 'SHIPPING']

browser = she()
# 遍历每个网址
for index, url in enumerate(lines):
    browser.get(url)
    browser.set_window_position(0, 0)
    browser.set_window_size(800, 1200)
    #Domain截屏
    filename = f"{unformatted[index]}_DOMAIN.png"
    browser.execute_script("document.body.style.zoom='50%'")
    time.sleep(1)
    browser.get_screenshot_as_file(filename)
    #滑倒最后截payment
    browser.execute_script("window.scrollTo(0,document.body.scrollHeight);")
    filename = f"{unformatted[index]}_PAYMENT.png"
    browser.execute_script("document.body.style.zoom='50%'")
    time.sleep(1)
    browser.get_screenshot_as_file(filename)
    #Product截屏，两张
    browser.find_element(By.PARTIAL_LINK_TEXT, 'product').click()
    filename = f"{unformatted[index]}_PRODUCT1.png"
    browser.execute_script("document.body.style.zoom='50%'")
    time.sleep(1)
    browser.get_screenshot_as_file(filename)
    browser.execute_script("window.scrollTo(0,document.body.scrollHeight);")
    filename = f"{unformatted[index]}_PRODUCT2.png"
    browser.execute_script("document.body.style.zoom='50%'")
    time.sleep(1)
    browser.get_screenshot_as_file(filename)
    browser.back()
    #Terms截屏，两张
    browser.find_element(By.PARTIAL_LINK_TEXT, 'TERM').click()
    filename = f"{unformatted[index]}_TERMS1.png"
    browser.execute_script("document.body.style.zoom='50%'")
    time.sleep(1)
    browser.get_screenshot_as_file(filename)
    browser.execute_script("window.scrollTo(0,document.body.scrollHeight);")
    filename = f"{unformatted[index]}_TERMS2.png"
    browser.execute_script("document.body.style.zoom='50%'")
    time.sleep(1)
    browser.get_screenshot_as_file(filename)
    browser.back()

    for i, text in enumerate(links):
        try:
            browser.find_element(By.PARTIAL_LINK_TEXT, text).click()
            filename = f"{unformatted[index]}_{text}.png"
            browser.execute_script("document.body.style.zoom='50%'")
            time.sleep(1)
            browser.get_screenshot_as_file(filename)
            browser.back()

        except NoSuchElementException:
            print(f"[{unformatted[index]}] 未找到链接：{text}")

browser.quit()
