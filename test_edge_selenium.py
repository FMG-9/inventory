from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time

# 測試 Edge WebDriver 是否可正常啟動
try:
    service = EdgeService(executable_path='msedgedriver.exe')
    driver = webdriver.Edge(service=service)
    driver.get('https://www.bing.com')
    print('Edge WebDriver 啟動成功，已開啟 Bing 首頁。')
    time.sleep(3)
    driver.quit()
    print('Edge WebDriver 測試結束，已關閉瀏覽器。')
except Exception as e:
    print('Edge WebDriver 啟動失敗:', e)
