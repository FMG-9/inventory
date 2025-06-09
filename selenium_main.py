from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from PIL import Image
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'd:\Programs\Vs Code\Microsoft VS Code\scrape\tesseract-5.5.1\tesseract.exe'
import time
import io
import base64

# 帳號密碼與預約參數
USERNAME = '411322049'  # 請改成你的帳號
PASSWORD = '2004.09.12'  # 請改成你的密碼
RESERVE_DATE = '2025/06/15'  # 預約日期
FIELD = 'VOL0B'  # 場地代碼
PERIOD = '06'  # 預設時段

# 場地對應表（依據實際下拉選單，名稱供註解與顯示，輸入時只需代號）
SPORT_FIELD_MAP = {
    '籃球': [
        ('BSK0A', '籃球場A'),
        ('BSK0B', '籃球場B'),
        ('BSK0C', '籃球場C'),
        ('BSK0D', '籃球場D'),
        ('BSK0E', '籃球場E'),
        ('BSK0F', '籃球場F'),
        ('BSK0G', '籃球場I (K書中心)'),
        ('BSK0H', '籃球場J (K書中心)'),
        ('BSK0J', '籃球場L (集賢館場地)'),
        ('BSK0K', '籃球場K (集賢館場地)'),
        ('BSKR1', '籃球場G (原R1)'),
        ('BSKR2', '籃球場H (原R2)'),
    ],
    '排球': [
        ('VOL0A', '排球場A-女'), ('VOL0B', '排球場B-男'), ('VOL0C', '排球場C-女'), ('VOL0D', '排球場D-男'),
        ('VOL0E', '排球場E-女'), ('VOL0F', '排球場F-男'), ('VOL0G', '排球場G-女'), ('VOL0H', '排球場H-男'),
        ('VOL0J', '排球場L-女 (集賢館場地)'), ('VOL0K', '排球場K-男 (集賢館場地)'),
        ('VOLR1', '排球場I-女 (原R1)'), ('VOLR2', '排球場J-男 (原R2)'),
       
    ],
    '操場': [
        ('TRK0A', '田徑場'),
        ('PLA0A', '體育室前廣場'),
        ('GYM0A', '韻律教室'),
        ('ARO0A', '柔道教室A')
    ],
    '體育館': [
        ('XDNCE', '壽豐館-舞蹈教室'),
        ('XGMB1', '壽館場B-羽1'), ('XGMB2', '壽館場B-羽2'), ('XGMB3', '壽館場B-羽3'), ('XGMB4', '壽館場B-羽4'),
        ('XGMC1', '壽館場C-排1'), ('XGMC2', '壽館場C-排2'), ('XGMC3', '壽館場C-排3'), ('XGMC4', '壽館場C-排4'),
        ('XGYMA', '壽館場A-籃球'),
        ('XTKDO', '壽豐館-跆拳道教室'),
        ('XTT0W', '壽豐體育館桌球室全部')
    ],
    '網球場': [
        ('XTNA1', '網球場1'), ('XTNA2', '網球場2'), ('XTNB1', '網球場3'), ('XTNB2', '網球場4'),
        ('XTNB3', '網球場5'), ('XTNB4', '網球場6 (紅土)'), ('XTNB5', '網球場7 (紅土)'),
        ('TNS0G', '網球場G'), ('TNS0H', '網球場H')
    ],
    '戶外大型球場': [
        ('SFT0B', '志學門壘球場1'),
        ('BSB0A', '棒球場'),
        ('BSK02', '高爾夫球場')
    ]
}

# 讓使用者輸入運動類型並顯示可選場地（含註釋），並讓使用者手動輸入場地與日期
sport = input('請輸入運動類型（籃球、排球、操場、體育館、網球場、戶外大型球場）：').strip()
if sport in SPORT_FIELD_MAP:
    print('可選場地：')
    for code, name in SPORT_FIELD_MAP[sport]:
        print(f'  {code}  {name}')
    FIELD = input('請輸入場地代碼（如 VOL0B）：').strip().upper()
    if FIELD not in [code for code, _ in SPORT_FIELD_MAP[sport]]:
        print('場地代碼不在該運動類型可選範圍，請重新執行程式。')
        exit()
else:
    print('運動類型輸入錯誤，請重新執行程式。')
    exit()
# 只需輸入 MMDD 四位數，並自動補上今年年份
while True:
    date_input = input('請輸入預約日期（格式：MMDD，例如 0625）：').strip()
    if len(date_input) == 4 and date_input.isdigit():
        month = date_input[:2]
        day = date_input[2:]
        from datetime import datetime
        year = datetime.now().year
        RESERVE_DATE = f'{year}/{month}/{day}'
        break
    else:
        print('格式錯誤，請重新輸入四位數月份及日期（如 0625）')
# 新增時段輸入，並驗證
while True:
    period = input('請輸入預約時段（06~22，06代表06~08，07代表07~09...）：').strip()
    if period.isdigit() and 6 <= int(period) <= 22:
        PERIOD = period
        break
    else:
        print('時段格式錯誤，請輸入06~22之間的兩位數')

# 啟動 Edge WebDriver
from selenium.webdriver.edge.options import Options
options = Options()
options.add_argument('--ignore-certificate-errors')
service = EdgeService(executable_path='msedgedriver.exe')
driver = webdriver.Edge(service=service, options=options)
driver.get('https://sys.ndhu.edu.tw/gc/sportcenter/SportsFields/Default.aspx')
wait = WebDriverWait(driver, 20)

try:
    # Step 1: 登入
    wait.until(EC.presence_of_element_located((By.ID, 'MainContent_TxtUSERNO'))).send_keys(USERNAME)
    driver.find_element(By.ID, 'MainContent_TxtPWD').send_keys(PASSWORD)
    driver.find_element(By.ID, 'MainContent_Button1').click()
    print('已自動登入')
    time.sleep(2)

    print('登入後網址:', driver.current_url)
    with open('login_response.html', 'w', encoding='utf-8') as f:
        f.write(driver.page_source)
    print('已存檔 login_response.html')
    # 判斷是否登入成功（網址或頁面內容判斷）
    if 'Login' in driver.current_url or 'TxtUSERNO' in driver.page_source:
        print('登入失敗，請檢查帳號密碼或網路/SSL問題')
        driver.quit()
        exit()
    print('登入成功，繼續流程...')
    time.sleep(5)  # 額外等待 JS 載入
    # Step 1.5: 新增申請
    wait.until(EC.element_to_be_clickable((By.ID, 'MainContent_Button2'))).click()
    print('已點擊新增申請')
    time.sleep(1)

    # Step 2: 選擇場地與日期
    wait.until(EC.presence_of_element_located((By.ID, 'MainContent_DropDownList1'))).send_keys(FIELD)
    date_box = driver.find_element(By.ID, 'MainContent_TextBox1')
    driver.execute_script("arguments[0].value = arguments[1];", date_box, RESERVE_DATE)
    driver.find_element(By.ID, 'MainContent_Button1').click()  # 查詢
    print('已查詢場地與日期')
    time.sleep(2)

    # Step 2.5: 檢查是否在開放借用期間
    try:
        btime_label = driver.find_element(By.ID, 'MainContent_BTimeLabel').text
        import re
        match = re.search(r'(\d{4}/\d{2}/\d{2})~(\d{4}/\d{2}/\d{2})', btime_label)
        if match:
            open_start, open_end = match.groups()
            from datetime import datetime
            reserve_date_obj = datetime.strptime(RESERVE_DATE, '%Y/%m/%d')
            open_start_obj = datetime.strptime(open_start, '%Y/%m/%d')
            open_end_obj = datetime.strptime(open_end, '%Y/%m/%d')
            if not (open_start_obj <= reserve_date_obj <= open_end_obj):
                print(f'您選擇的日期 {RESERVE_DATE} 不在開放借用期間：{open_start} ~ {open_end}')
                driver.quit()
                exit()
    except Exception as e:
        print(f'無法取得開放借用期間資訊: {e}')

    # Step 3: 先找所有申請按鈕，再用 Python 過濾時段
    period_range = f"{PERIOD}~{str(int(PERIOD)+2).zfill(2)}"  # 06~08, 07~09, ...
    btns = driver.find_elements(By.XPATH, "//button[contains(., '申請')]")
    target_btn = None
    for btn in btns:
        btn_text = btn.text.replace('\n', '').replace(' ', '')
        if period_range in btn_text:
            target_btn = btn
            break
    if target_btn:
        target_btn.click()
        print(f'已點擊 {RESERVE_DATE} {PERIOD} 時段申請按鈕')
    else:
        print(f'找不到 {PERIOD} 時段的申請按鈕，請確認該時段是否可申請')
        driver.quit()
        exit()
    time.sleep(2)

    # Step 4: 處理驗證碼（增強影像處理與多次嘗試）
    max_attempts = 5
    for attempt in range(1, max_attempts + 1):
        print(f'驗證碼嘗試第 {attempt} 次...')
        if attempt > 1:
            try:
                old_img = driver.find_element(By.ID, 'imgCaptcha')
                wait.until(EC.staleness_of(old_img))
            except Exception:
                # 如果圖片已經消失，直接等待新圖片出現
                pass
        wait.until(EC.presence_of_element_located((By.ID, 'imgCaptcha')))
        time.sleep(1)  # 確保圖片完全載入
        img_base64 = driver.find_element(By.ID, 'imgCaptcha').get_attribute('src')
        img_data = img_base64.split(',')[-1]
        img_bytes = base64.b64decode(img_data)
        img = Image.open(io.BytesIO(img_bytes))
        # 增強影像處理：降低亮度、提升對比、放大、銳化、二值化
        from PIL import ImageEnhance, ImageFilter
        enhancer = ImageEnhance.Brightness(img)
        img = enhancer.enhance(0.7)
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(3.5)
        img = img.convert('L')
        img = img.resize((img.width * 3, img.height * 3))
        threshold = 150
        img = img.point(lambda x: 255 if x > threshold else 0)
        img = img.filter(ImageFilter.SHARPEN)
        img.show()
        captcha_text = input('請輸入圖片中的驗證碼（或輸入 exit 結束）：').strip()
        if captcha_text.lower() == 'exit':
            print('使用者手動結束程式。')
            driver.quit()
            exit()
        captcha_input = driver.find_element(By.ID, 'txtCaptchaValue')
        captcha_input.clear()
        captcha_input.send_keys(captcha_text)
        time.sleep(1)
        # 送出驗證碼，點擊驗證碼區塊下方的「申請」按鈕（通常是最後一個）
        app_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), '申請')]")
        app_buttons[-1].click()
        print('已自動送出驗證碼')
        time.sleep(3)
        page_source = driver.page_source
        # 先判斷是否有驗證碼錯誤訊息
        try:
            captcha_err = driver.find_element(By.ID, 'hfCaptchaErrMsg').get_attribute('value')
        except Exception:
            captcha_err = ''
        if captcha_err and '不通過' in captcha_err:
            print('驗證碼錯誤，重試...')
            continue
        # 再判斷是否預約成功
        if 'MainContent_AppPlaceTB' in page_source and FIELD in page_source and RESERVE_DATE.replace('/', '') in page_source:
            print('預約成功，已加入申請列表！')
            print('驗證碼正確，準備自動填寫借用原因...')
            # 預約成功後等待3秒再填寫借用原因
            time.sleep(3)
            try:
                # 先嘗試用 id
                reason_input = wait.until(EC.element_to_be_clickable((By.ID, 'MainContent_ReasonTextBox1')))
                reason_input.clear()
                reason_input.send_keys('運動')
                driver.execute_script("""
                    arguments[0].value = '運動';
                    arguments[0].dispatchEvent(new Event('input'));
                    arguments[0].dispatchEvent(new Event('change'));
                """, reason_input)
                print('已自動填入借用原因：運動')
            except Exception as e:
                print(f'用 id 找不到借用原因欄位: {e}')
                # 嘗試用 placeholder
                try:
                    inputs = driver.find_elements(By.XPATH, "//input[@type='text']")
                    for inp in inputs:
                        ph = inp.get_attribute('placeholder') or ''
                        if '原因' in ph or '借用' in ph:
                            inp.clear()
                            inp.send_keys('運動')
                            driver.execute_script("""
                                arguments[0].value = '運動';
                                arguments[0].dispatchEvent(new Event('input'));
                                arguments[0].dispatchEvent(new Event('change'));
                            """, inp)
                            print('已自動填入借用原因（用 placeholder 匹配）：運動')
                            break
                    else:
                        print('找不到借用原因欄位，請手動填寫')
                except Exception as e2:
                    print(f'用 placeholder 也找不到借用原因欄位: {e2}')
            # 再等待1秒後自動按下「確定」
            time.sleep(1)
            try:
                ok_btn = wait.until(EC.element_to_be_clickable((By.ID, 'MainContent_Button4')))
                ok_btn.click()
                print('已自動點擊「確定」完成鎖單')
            except Exception as e:
                print(f'找不到「確定」按鈕或點擊失敗: {e}')
            break
        else:
            with open('result_debug.html', 'w', encoding='utf-8') as f:
                f.write(driver.page_source)
            print('已存檔 result_debug.html，請手動檢查內容')
            print('請手動檢查預約結果')
            break
    else:
        print('多次嘗試後仍未成功，請手動輸入驗證碼或檢查流程')
finally:
    input('流程結束，請手動檢查網頁，按 Enter 關閉瀏覽器...')
    driver.quit()
