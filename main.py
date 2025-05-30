import requests
from bs4 import BeautifulSoup
import os
from PIL import Image
import pytesseract
import urllib3

# 關閉 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 帳號密碼
username = '411322049'
password = '2004.09.12'

login_url = "https://sys.ndhu.edu.tw/gc/sportcenter/SportsFields/login.aspx"
session = requests.Session()

# Step 1: 取得登入頁面
resp = session.get(login_url, verify=False)
soup = BeautifulSoup(resp.text, 'html.parser')

# 取得必要隱藏欄位
viewstate = soup.find('input', {'id': '__VIEWSTATE'})['value']
eventvalidation = soup.find('input', {'id': '__EVENTVALIDATION'})['value']
viewstategen = soup.find('input', {'id': '__VIEWSTATEGENERATOR'})
viewstategen = viewstategen['value'] if viewstategen else ''
eventtarget = soup.find('input', {'id': '__EVENTTARGET'})
eventtarget = eventtarget['value'] if eventtarget else ''
eventargument = soup.find('input', {'id': '__EVENTARGUMENT'})
eventargument = eventargument['value'] if eventargument else ''
request_verification_token = soup.find('input', {'name': '__RequestVerificationToken'})
request_verification_token = request_verification_token['value'] if request_verification_token else ''

# Step 2: 準備登入資料
payload = {
    '__VIEWSTATE': viewstate,
    '__EVENTVALIDATION': eventvalidation,
    '__VIEWSTATEGENERATOR': viewstategen,
    '__EVENTTARGET': eventtarget,
    '__EVENTARGUMENT': eventargument,
    '__RequestVerificationToken': request_verification_token,
    'ctl00$MainContent$TxtUSERNO': username,
    'ctl00$MainContent$TxtPWD': password,
    'ctl00$MainContent$Button1': '登入'
}s

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Referer': login_url,
    'Origin': 'https://sys.ndhu.edu.tw',
    'Content-Type': 'application/x-www-form-urlencoded'
}

# Step 3: 提交登入表單
login_resp = session.post(login_url, data=payload, headers=headers, verify=False)

print("登入結果:", "成功" if "登出" in login_resp.text else "失敗")

# 將回應內容寫入檔案方便檢查
with open('login_response.html', 'w', encoding='utf-8') as f:
    f.write(login_resp.text)

print('已將登入回應內容輸出到 login_response.html，請用瀏覽器開啟檢查登入失敗原因。')

# 顯示目前 session cookies
print('目前 cookies:', session.cookies.get_dict())

# Step 4: 查詢排球 VOL0B 場地 2025/06/07 是否可申請
# 1. 取得查詢頁面隱藏欄位
soup = BeautifulSoup(login_resp.text, 'html.parser')
viewstate = soup.find('input', {'id': '__VIEWSTATE'})['value']
eventvalidation = soup.find('input', {'id': '__EVENTVALIDATION'})['value']
viewstategen = soup.find('input', {'id': '__VIEWSTATEGENERATOR'})
viewstategen = viewstategen['value'] if viewstategen else ''
request_verification_token = soup.find('input', {'name': '__RequestVerificationToken'})
request_verification_token = request_verification_token['value'] if request_verification_token else ''

# 2. 準備查詢 payload
query_payload = {
    '__VIEWSTATE': viewstate,
    '__EVENTVALIDATION': eventvalidation,
    '__VIEWSTATEGENERATOR': viewstategen,
    '__RequestVerificationToken': request_verification_token,
    '__EVENTTARGET': 'ctl00$MainContent$Button1',  # 查詢按鈕
    '__EVENTARGUMENT': '',
    'ctl00$MainContent$drpkind': '2',  # 2=排球
    'ctl00$MainContent$DropDownList1': 'VOL0B',  # VOL0B場地
    'ctl00$MainContent$TextBox1': '2025/06/07',  # 日期
}

# 3. 送出查詢
reserve_url = login_url  # 查詢同一頁
reserve_resp = session.post(reserve_url, data=query_payload, headers=headers, verify=False)
with open('query_response.html', 'w', encoding='utf-8') as f:
    f.write(reserve_resp.text)
print('已將查詢結果輸出到 query_response.html，請用瀏覽器檢查可申請時段。')
# 你可以用 BeautifulSoup 解析 reserve_resp.text 來自動抓取可申請時段