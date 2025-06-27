from selenium import webdriver
from selenium.webdriver.chrome.options import Options   #用于设置谷歌浏览器
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time

#设置浏览器，启动浏览器
def she():
    #创建设置浏览器对象
    q1 = Options()
    #禁用沙盘模式
    q1.add_argument('--no-sandbox')
    #保持浏览器打开状态（默认执行完自动关闭）
    q1.add_experimental_option('detach', True)

    #创建并启动浏览器
    a1 = webdriver.Chrome(service=Service('chromedriver.exe'), options=q1)
    return a1

a1 = she()
#打开指定网址
a1.get('https://www.hreinsd.com/')
a1.maximize_window()
#a1.find_element(By.TAG_NAME,'a')[149].click()
#a1.get_screenshot_as_file('1.png')
#a1.back()
#a1.find_element(By.TAG_NAME,'a')[150].click()
#a1.get_screenshot_as_file('2.png')
#a1.back()
#a1.find_element(By.TAG_NAME,'a')[151].click()
#a1.get_screenshot_as_file('3.png')
#a1.back()
#a1.find_element(By.TAG_NAME,'a')[153].click()
#a1.get_screenshot_as_file('4.png')
#a1.back()
#a1.find_element(By.TAG_NAME,'a')[155].click()
#a1.get_screenshot_as_file('5.png')
#a1.back()
time.sleep(5)
a1.find_element(By.PARTIAL_LINK_TEXT,'Return').click()
a1.get_screenshot_as_file('1.png')
a1.back()
a1.find_element(By.PARTIAL_LINK_TEXT,'Contact').click()
a1.get_screenshot_as_file('2.png')
a1.back()
a1.find_element(By.PARTIAL_LINK_TEXT,'Terms').click()
a1.get_screenshot_as_file('3.png')
a1.back()
a1.find_element(By.PARTIAL_LINK_TEXT,'Privacy').click()
a1.get_screenshot_as_file('4.png')
a1.back()
a1.find_element(By.PARTIAL_LINK_TEXT,'Shipping').click()
a1.get_screenshot_as_file('5.png')
a1.back()
#关闭当前标签页
#a1.close()
#退出浏览器并释放驱动
a1.quit()


#---------------------------------
#窗口最大化
#a1.maximize_window()
#窗口最小化
#a1.minimize_window()
#位置
#a1.set_window_position(0,0)
#尺寸
#a1.set_window_size(800,800)
#---------------------------------
#定位元素（定位成功返回结果）
#a2 = a1.find_element(By.id, 'kw')
#定位多个元素（定位成功返回列表）
#a2 = a1.find_elements(By.id, 'kw')
#在谷歌控制台里
#document.getElementById('kw')
#---------------------------------
#元素输入
#a2 = a1.find_element(By.ID, 'kw')
#a2.send_keys('defait')
#元素清空
#a2.clear()
#元素输入
#a2 = a1.find_element(By.ID, 'su')
#点击
#a2.click()
#---------------------------------
#a1.find_element(By.CLASS_NAME,'channel-icons')

