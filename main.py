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
}

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

# Step 4: 取得查詢頁面隱藏欄位（直接從登入後頁面）
soup = BeautifulSoup(login_resp.text, 'html.parser')
viewstate = soup.find('input', {'id': '__VIEWSTATE'})['value']
eventvalidation = soup.find('input', {'id': '__EVENTVALIDATION'})['value']
viewstategen = soup.find('input', {'id': '__VIEWSTATEGENERATOR'})
viewstategen = viewstategen['value'] if viewstategen else ''
request_verification_token = soup.find('input', {'name': '__RequestVerificationToken'})
request_verification_token = request_verification_token['value'] if request_verification_token else ''

# Step 5: 直接送出「新增申請」POST 請求，優先觸發驗證碼
reserve_date = '2025/06/15'  # 統一申請與查詢日期
add_payload = {
    '__VIEWSTATE': viewstate,
    '__EVENTVALIDATION': eventvalidation,
    '__VIEWSTATEGENERATOR': viewstategen,
    '__RequestVerificationToken': request_verification_token,
    '__EVENTTARGET': 'ctl00$MainContent$Button3',  # 假設「新增申請」按鈕 name/id
    '__EVENTARGUMENT': '',
    'ctl00$MainContent$drpkind': '2',
    'ctl00$MainContent$DropDownList1': 'VOL0B',
    'ctl00$MainContent$TextBox1': reserve_date,
}
add_resp = session.post(login_url, data=add_payload, headers=headers, verify=False)
with open('add_response.html', 'w', encoding='utf-8') as f:
    f.write(add_resp.text)
print('已將新增申請後頁面輸出到 add_response.html，請用瀏覽器檢查申請按鈕。')

# 6. 自動判斷 add_response.html 預約狀態（驗證碼處理等）
with open('add_response.html', 'r', encoding='utf-8') as f:
    add_html = f.read()

if 'captcha' in add_html or '驗證碼' in add_html:
    print('需要處理驗證碼，開始自動辨識...')
    soup = BeautifulSoup(add_html, 'html.parser')
    # 取得 base64 圖片
    captcha_base64 = soup.find('input', {'id': 'hfCaptchaImageBase64'})
    captcha_id = soup.find('input', {'id': 'hfCaptchaId'})
    encrypted_ymdh = soup.find('input', {'id': 'hfEncryptedYMDH'})
    plain_ymdh = soup.find('input', {'id': 'hfPlainYMDH'})
    base64_value = captcha_base64.get('value', '') if captcha_base64 else ''
    captcha_id_value = captcha_id.get('value', '') if captcha_id else ''
    encrypted_ymdh_value = encrypted_ymdh.get('value', '') if encrypted_ymdh else ''
    plain_ymdh_value = plain_ymdh.get('value', '') if plain_ymdh else ''
    if base64_value and captcha_id_value and encrypted_ymdh_value and plain_ymdh_value:
        import base64
        img_data = base64_value.split(',')[-1]
        img_bytes = base64.b64decode(img_data)
        with open('captcha.png', 'wb') as f:
            f.write(img_bytes)
        print('已存檔 captcha.png，進行 OCR...')
        img = Image.open('captcha.png')
        captcha_text = pytesseract.image_to_string(img, config='--psm 8').strip()
        print('OCR 辨識結果:', captcha_text)
        # 取得隱藏欄位
        viewstate = soup.find('input', {'id': '__VIEWSTATE'})['value']
        eventvalidation = soup.find('input', {'id': '__EVENTVALIDATION'})['value']
        viewstategen = soup.find('input', {'id': '__VIEWSTATEGENERATOR'})
        viewstategen = viewstategen['value'] if viewstategen else ''
        request_verification_token = soup.find('input', {'name': '__RequestVerificationToken'})
        request_verification_token = request_verification_token['value'] if request_verification_token else ''
        # 組成 payload
        captcha_payload = {
            '__VIEWSTATE': viewstate,
            '__EVENTVALIDATION': eventvalidation,
            '__VIEWSTATEGENERATOR': viewstategen,
            '__RequestVerificationToken': request_verification_token,
            'ctl00$MainContent$hfCaptchaId': captcha_id_value,
            'ctl00$MainContent$hfEncryptedYMDH': encrypted_ymdh_value,
            'ctl00$MainContent$hfPlainYMDH': plain_ymdh_value,
            'ctl00$MainContent$hfCaptchaImageBase64': base64_value,
            'ctl00$MainContent$hfCaptchaValue': captcha_text,
            'ctl00$MainContent$Button4': '確定送出',
        }
        # 送出申請
        captcha_resp = session.post(login_url, data=captcha_payload, headers=headers, verify=False)
        with open('captcha_submit_response.html', 'w', encoding='utf-8') as f:
            f.write(captcha_resp.text)
        print('已自動送出驗證碼，請檢查 captcha_submit_response.html 結果。')
    else:
        print('找不到驗證碼相關欄位或 value 為空，請手動檢查 add_response.html')
elif '已借用' in add_html:
    print('該時段已被借用，無法預約。')
elif username in add_html:
    print('預約成功，頁面中已出現你的名字！')
else:
    print('預約未成功，請檢查 add_response.html 內容。')

# 7. 查詢排球 VOL0B 場地 reserve_date 是否可申請（用申請流程頁面的隱藏欄位）
query_date = reserve_date
soup = BeautifulSoup(add_resp.text, 'html.parser')
viewstate = soup.find('input', {'id': '__VIEWSTATE'})['value']
eventvalidation = soup.find('input', {'id': '__EVENTVALIDATION'})['value']
viewstategen = soup.find('input', {'id': '__VIEWSTATEGENERATOR'})
viewstategen = viewstategen['value'] if viewstategen else ''
request_verification_token = soup.find('input', {'name': '__RequestVerificationToken'})
request_verification_token = request_verification_token['value'] if request_verification_token else ''

query_payload = {
    '__VIEWSTATE': viewstate,
    '__EVENTVALIDATION': eventvalidation,
    '__VIEWSTATEGENERATOR': viewstategen,
    '__RequestVerificationToken': request_verification_token,
    '__EVENTTARGET': 'ctl00$MainContent$Button1',  # 查詢按鈕
    '__EVENTARGUMENT': '',
    'ctl00$MainContent$drpkind': '2',  # 2=排球
    'ctl00$MainContent$DropDownList1': 'VOL0B',  # VOL0B場地
    'ctl00$MainContent$TextBox1': query_date,  # 日期
}
reserve_url = login_url
reserve_resp = session.post(reserve_url, data=query_payload, headers=headers, verify=False)
with open('query_response.html', 'w', encoding='utf-8') as f:
    f.write(reserve_resp.text)
print(f'已將{query_date}查詢結果輸出到 query_response.html，請用瀏覽器檢查可申請時段。')

# 8. 解析查詢結果，自動選取 16:00~18:00 時段的 button 並取出 encryptedYMDH
import re
selected_btn = None
encrypted_ymdh = None
plain_ymdh = None
for btn in soup.find_all('button', {'type': 'button'}):
    btn_text = btn.text.strip()
    # 改用 in 判斷，避免 font 標籤分割問題
    if '[申請]' in btn_text and ('16' in btn_text or '17' in btn_text):
        print('找到可申請時段按鈕:', btn_text)
        onclick = btn.get('onclick', '')
        m = re.search(r"AppBtnC\('(.+?)'\)", onclick)
        if m:
            encrypted_ymdh = m.group(1)
            plain_ymdh = btn_text
            selected_btn = btn
            break
if not selected_btn:
    print('找不到 16:00~18:00 可申請時段 button，結束流程。')
    exit()

# 9. 取得查詢後頁面的隱藏欄位，準備送出時段申請
viewstate = soup.find('input', {'id': '__VIEWSTATE'})['value']
eventvalidation = soup.find('input', {'id': '__EVENTVALIDATION'})['value']
viewstategen = soup.find('input', {'id': '__VIEWSTATEGENERATOR'})
viewstategen = viewstategen['value'] if viewstategen else ''
request_verification_token = soup.find('input', {'name': '__RequestVerificationToken'})
request_verification_token = request_verification_token['value'] if request_verification_token else ''

# 送出申請時段（模擬 doApp/encryptedYMDH 行為）
apply_payload = {
    '__VIEWSTATE': viewstate,
    '__EVENTVALIDATION': eventvalidation,
    '__VIEWSTATEGENERATOR': viewstategen,
    '__RequestVerificationToken': request_verification_token,
    'ctl00$MainContent$hfEncryptedYMDH': encrypted_ymdh,
    'ctl00$MainContent$hfPlainYMDH': plain_ymdh,
    # 可能還需要其他欄位，根據 add_response.html 裡的表單欄位補齊
}
apply_resp = session.post(reserve_url, data=apply_payload, headers=headers, verify=False)
with open('add_response.html', 'w', encoding='utf-8') as f:
    f.write(apply_resp.text)
print('已將申請時段後頁面輸出到 add_response.html，請用瀏覽器檢查驗證碼。')

# 10. 送出申請時段後，這時才判斷是否需要處理驗證碼
with open('add_response.html', 'r', encoding='utf-8') as f:
    add_html = f.read()

if 'captcha' in add_html or '驗證碼' in add_html:
    print('需要處理驗證碼，開始自動辨識...')
    soup = BeautifulSoup(add_html, 'html.parser')
    captcha_base64 = soup.find('input', {'id': 'hfCaptchaImageBase64'})
    captcha_id = soup.find('input', {'id': 'hfCaptchaId'})
    encrypted_ymdh = soup.find('input', {'id': 'hfEncryptedYMDH'})
    plain_ymdh = soup.find('input', {'id': 'hfPlainYMDH'})
    base64_value = captcha_base64.get('value', '') if captcha_base64 else ''
    captcha_id_value = captcha_id.get('value', '') if captcha_id else ''
    encrypted_ymdh_value = encrypted_ymdh.get('value', '') if encrypted_ymdh else ''
    plain_ymdh_value = plain_ymdh.get('value', '') if plain_ymdh else ''
    if base64_value and captcha_id_value and encrypted_ymdh_value and plain_ymdh_value:
        import base64
        img_data = base64_value.split(',')[-1]
        img_bytes = base64.b64decode(img_data)
        with open('captcha.png', 'wb') as f:
            f.write(img_bytes)
        print('已存檔 captcha.png，進行 OCR...')
        img = Image.open('captcha.png')
        captcha_text = pytesseract.image_to_string(img, config='--psm 8').strip()
        print('OCR 辨識結果:', captcha_text)
        # 取得隱藏欄位
        viewstate = soup.find('input', {'id': '__VIEWSTATE'})['value']
        eventvalidation = soup.find('input', {'id': '__EVENTVALIDATION'})['value']
        viewstategen = soup.find('input', {'id': '__VIEWSTATEGENERATOR'})
        viewstategen = viewstategen['value'] if viewstategen else ''
        request_verification_token = soup.find('input', {'name': '__RequestVerificationToken'})
        request_verification_token = request_verification_token['value'] if request_verification_token else ''
        captcha_payload = {
            '__VIEWSTATE': viewstate,
            '__EVENTVALIDATION': eventvalidation,
            '__VIEWSTATEGENERATOR': viewstategen,
            '__RequestVerificationToken': request_verification_token,
            'ctl00$MainContent$hfCaptchaId': captcha_id_value,
            'ctl00$MainContent$hfEncryptedYMDH': encrypted_ymdh_value,
            'ctl00$MainContent$hfPlainYMDH': plain_ymdh_value,
            'ctl00$MainContent$hfCaptchaImageBase64': base64_value,
            'ctl00$MainContent$hfCaptchaValue': captcha_text,
            'ctl00$MainContent$Button4': '確定送出',
        }
        captcha_resp = session.post(login_url, data=captcha_payload, headers=headers, verify=False)
        with open('captcha_submit_response.html', 'w', encoding='utf-8') as f:
            f.write(captcha_resp.text)
        print('已自動送出驗證碼，請檢查 captcha_submit_response.html 結果。')
    else:
        print('找不到驗證碼相關欄位或 value 為空，請手動檢查 add_response.html')
elif '已借用' in add_html:
    print('該時段已被借用，無法預約。')
elif username in add_html:
    print('預約成功，頁面中已出現你的名字！')
else:
    print('預約未成功，請檢查 add_response.html 內容。')