from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
from bs4 import BeautifulSoup as bs
from dateutil import parser
from datetime import datetime
import re
import pandas as pd
import calendar
import time
import pyautogui
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoAlertPresentException
import os
import shutil
import glob
from selenium.webdriver.common.action_chains import ActionChains
import pygetwindow as gw

NH_BANK_URL = "https://banking.nonghyup.com/servlet/IPMSP0011I.view"

def type_with_keyboard(selector, text, driver, interval=0.3, timeout=15, retries=3, is_secure=False):
    wait = WebDriverWait(driver, timeout)
    element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))

    for attempt in range(retries):
        element.click()
        # element.clear()
        time.sleep(1)  # clear 후 1초 대기
        element.click()
        time.sleep(1)
        pyautogui.write(text, interval=interval)
        if is_secure:
            return True
        else:
            entered_value = element.get_attribute('value').replace('-', '').strip()
            if entered_value == text:
                return True
        time.sleep(1)
    raise ValueError(f"입력 실패 [{selector}]")

def save_page_source(driver, filename='page_source.html'):
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(driver.page_source)


def get_driver(PATH, download_dir):
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-popup-blocking") 
    options.add_argument("--disable-notifications")

    # 크롬 다운로드 설정
    prefs = {
        "download.default_directory": download_dir,  
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    
    service = Service(executable_path=PATH)
    return webdriver.Chrome(service=service, options=options)

def activate_chrome_window():
    # 'Chrome' 창 목록 찾기
    chrome_windows = [w for w in gw.getWindowsWithTitle('Chrome') if w.isVisible]
    
    if chrome_windows:
        chrome_window = chrome_windows[0]
        chrome_window.activate()  # 창 활성화 (강제로 맨 앞으로 올림)
        time.sleep(1)
    else:
        raise RuntimeError("Chrome 창을 찾을 수 없습니다.")
    
def click_excel_button(driver, timeout=30):
    wait = WebDriverWait(driver, timeout)
    try:
        button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='o_print']/span/a")))
        driver.execute_script("arguments[0].click();", button)
    except (ElementClickInterceptedException, TimeoutException) as e:
        raise RuntimeError("인쇄 및 엑셀저장 버튼 클릭 실패:", e)

def switch_to_new_window(driver, timeout=30):
    wait = WebDriverWait(driver, timeout)
    original_window = driver.current_window_handle
    wait.until(EC.number_of_windows_to_be(2))
    new_window = [window for window in driver.window_handles if window != original_window][0]
    driver.switch_to.window(new_window)
    return original_window

def debug_excel_button_xpath(driver):
    original_window = switch_to_new_window(driver)
    time.sleep(3)  # 오즈리포트 창 로딩 시간 고려

    iframes = driver.find_elements(By.TAG_NAME, 'iframe')
    if iframes:
        driver.switch_to.frame(iframes[0])

    excel_buttons = driver.find_elements(By.XPATH, "//img[contains(@src, 'xls') or contains(@alt, '엑셀')]")
    print(f"찾은 버튼 개수: {len(excel_buttons)}")
    for btn in excel_buttons:
        print("src:", btn.get_attribute("src"), "alt:", btn.get_attribute("alt"))

    driver.switch_to.default_content()
    driver.close()  # 디버그 후 오즈 창 닫기
    driver.switch_to.window(original_window)

def click_more_button_until_end(driver, timeout=10):
    wait = WebDriverWait(driver, timeout)
    while True:
        try:
            # 더보기 버튼이 보이고 클릭 가능할 때까지 대기
            more_button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#moreBtnArea span.btn3 a"))
            )
            driver.execute_script("arguments[0].click();", more_button)
            print("더보기 버튼 클릭 완료. 다음 내역을 불러오는 중...")
            time.sleep(6)  # 다음 데이터를 로드할 때까지의 대기시간
        except TimeoutException:
            print("더보기 버튼이 더 이상 없습니다. 모든 내역을 불러왔습니다.")
            break  # 더보기 버튼이 없으면 반복문 종료

def debug_iframe_structure(driver):
    original_window = switch_to_new_window(driver)
    time.sleep(3)

    # 최상위 iframe 확인
    iframes = driver.find_elements(By.TAG_NAME, 'iframe')
    print(f"최상위 iframe 개수: {len(iframes)}")

    for idx, iframe in enumerate(iframes):
        print(f"[최상위 iframe {idx}] src:", iframe.get_attribute("src"))
        driver.switch_to.frame(iframe)
        
        # 중첩 iframe 확인
        child_iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        print(f"  └ 중첩 iframe 개수: {len(child_iframes)}")

        for child_idx, child_iframe in enumerate(child_iframes):
            print(f"    └ 중첩 iframe {child_idx} src:", child_iframe.get_attribute("src"))
            
            # 중첩 iframe 내부의 버튼 여부 확인
            driver.switch_to.frame(child_iframe)
            excel_buttons = driver.find_elements(By.XPATH, "//img[contains(@src, 'xls') or contains(@alt, '엑셀')]")
            print(f"      └ 찾은 엑셀 버튼 개수: {len(excel_buttons)}")
            for btn in excel_buttons:
                print("        └ 버튼 src:", btn.get_attribute("src"), "alt:", btn.get_attribute("alt"))
            driver.switch_to.parent_frame()

        driver.switch_to.default_content()
    
    driver.close()
    driver.switch_to.window(original_window)

def debug_excel_button_xpath(driver):
    original_window = switch_to_new_window(driver)
    time.sleep(3)

    # iframe 확인
    iframes = driver.find_elements(By.TAG_NAME, 'iframe')
    print(f"최상위 iframe 개수: {len(iframes)}")

    for idx, iframe in enumerate(iframes):
        print(f"iframe {idx} src:", iframe.get_attribute('src'))
        driver.switch_to.frame(iframe)

        # iframe 내부의 엑셀 버튼 확인
        excel_buttons = driver.find_elements(By.XPATH, "//img[contains(@src, 'xls') or contains(@alt, '엑셀')]")
        print(f"iframe {idx} 안의 엑셀 버튼 개수: {len(excel_buttons)}")

        for btn in excel_buttons:
            print("버튼 src:", btn.get_attribute("src"), "alt:", btn.get_attribute("alt"))

        driver.switch_to.default_content()

    driver.close()
    driver.switch_to.window(original_window)

def set_input_via_js(selector, value, driver, timeout=15):
    wait = WebDriverWait(driver, timeout)
    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
    driver.execute_script("arguments[0].value = arguments[1];", element, value)

    # 값 변경 후 이벤트 발생
    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));", element
    )
    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", element
    )

def _get_transactions(driver, bank, pw, birthday, start_date, end_date):
    driver.get(NH_BANK_URL)
    
    wait = WebDriverWait(driver, 20)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#InqGjaNbr')))
    time.sleep(5)

    # 계좌번호 입력도 JavaScript로 바꾸면 더욱 확실함
    type_with_keyboard('#InqGjaNbr', bank, driver)
    time.sleep(1)

    # 🔑 비밀번호 입력 (JS 직접입력으로 확실히 입력)
    type_with_keyboard('#GjaSctNbr', pw, driver)
    time.sleep(1)

    # 생년월일 입력도 JS 입력 방식으로 변경 추천
    type_with_keyboard('#rlno1', birthday, driver)
    time.sleep(1)

    Select(driver.find_element(By.CSS_SELECTOR, "select[name='start_year']")).select_by_value(start_date.strftime("%Y"))
    Select(driver.find_element(By.CSS_SELECTOR, "select[name='start_month']")).select_by_value(start_date.strftime("%m"))
    Select(driver.find_element(By.CSS_SELECTOR, "select[name='start_date']")).select_by_value(start_date.strftime("%d"))

    Select(driver.find_element(By.CSS_SELECTOR, "select[name='end_year']")).select_by_value(end_date.strftime("%Y"))
    Select(driver.find_element(By.CSS_SELECTOR, "select[name='end_month']")).select_by_value(end_date.strftime("%m"))
    Select(driver.find_element(By.CSS_SELECTOR, "select[name='end_date']")).select_by_value(end_date.strftime("%d"))

    driver.find_element(By.CSS_SELECTOR, '#btn_search').click()

    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#hiddenResult table.tb_col tbody tr')))
    click_more_button_until_end(driver)
    save_page_source(driver)

    time.sleep(3)
    return driver.find_elements(By.CSS_SELECTOR, '#hiddenResult table.tb_col tbody tr')

def selenium_debug_buttons(driver):
    print("모든 버튼과 링크 요소 탐색 중...")
    elements = driver.find_elements(By.XPATH, "//img | //a | //button")
    print(f"발견된 버튼 및 링크 개수: {len(elements)}")

    for idx, elem in enumerate(elements):
        tag_name = elem.tag_name
        src = elem.get_attribute('src')
        alt = elem.get_attribute('alt')
        href = elem.get_attribute('href')
        onclick = elem.get_attribute('onclick')
        text = elem.text.strip()

        print(f"요소 {idx+1} ---------------------------")
        print(f"태그: {tag_name}")
        if src:
            print(f"src 속성: {src}")
        if alt:
            print(f"alt 속성: {alt}")
        if href:
            print(f"href 속성: {href}")
        if onclick:
            print(f"onclick 속성: {onclick}")
        if text:
            print(f"텍스트 내용: {text}")
        print("------------------------------------")

    print("디버깅 종료. 위 정보에서 엑셀 다운로드와 관련된 요소를 찾아 XPath를 구성하세요.")

def parse_amount(amount_text):
    amount_clean = re.sub(r'[^\d]', '', amount_text)
    return int(amount_clean) if amount_clean else 0

def get_month_date_range(year, month):
    first_day = datetime(year, month, 1)
    last_day_of_month = calendar.monthrange(year, month)[1]
    last_day = datetime(year, month, last_day_of_month)
    today = datetime.today()

    if last_day > today:
        last_day = today

    return first_day, last_day

def convert_xls_to_xlsx(src_file, dest_file):
    # xls 읽기 (xlrd를 통해)
    df = pd.read_excel(src_file, engine='xlrd')
    
    # xlsx로 저장하기 (openpyxl을 통해)
    df.to_excel(dest_file, index=False, engine='openpyxl')

def type_securely(selector, text, driver, timeout=20):
    wait = WebDriverWait(driver, timeout)
    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))

    # JS로 강제 포커스 및 값 입력
    driver.execute_script("arguments[0].focus();", element)
    driver.execute_script("arguments[0].value = arguments[1];", element, text)
    time.sleep(0.5)

    # 필수 입력 이벤트 강제 실행
    driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", element)
    driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", element)
    time.sleep(0.5)

def get_balance(PATH, bank, pw, birthday, start_date, end_date, download_dir):
    driver = get_driver(PATH, download_dir)
    
    transactions = _get_transactions(driver, bank, pw, birthday, start_date, end_date)
    
    existing_files = glob.glob(os.path.join(download_dir, '*'))
    for file in existing_files:
        os.remove(file)

    click_excel_button(driver)
    time.sleep(3)

    download_excel_from_oz_report(driver)
    
    # 다운로드된 파일 기다리기
    downloaded_file = wait_for_file_download(download_dir)

    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    xlsx_filename = os.path.join(download_dir, f'NH_Transactions_{timestamp}.xlsx')
    upload_xlsx_filename = os.path.join(download_dir, f'NH_Transactions_{timestamp}_upload.xlsx')

    # 다운로드된 파일 바로 이동만 수행 (xls → xlsx 변환 없음)
    shutil.move(downloaded_file, xlsx_filename)

    # 업로드 양식 엑셀파일 변환 및 저장
    transformed_df = read_and_transform_downloaded_excel(xlsx_filename)
    save_transformed_excel(transformed_df, upload_xlsx_filename)

    print(f"거래내역 엑셀 파일 저장 완료: {xlsx_filename}")
    print(f"업로드용 엑셀 파일 저장 완료: {upload_xlsx_filename}")

    driver.quit()

    return xlsx_filename, upload_xlsx_filename

def save_transactions_to_excel(transactions, filename):
    df = pd.DataFrame(transactions)
    df['date'] = pd.to_datetime(df['date'], errors='coerce')
    if df['date'].isna().any():
        print("Warning: Some dates could not be parsed correctly.")
        print(df[df['date'].isna()])
        df.dropna(subset=['date'], inplace=True)
    df['date'] = df['date'].dt.strftime('%Y-%m-%d %H:%M:%S')
    df.to_excel(filename, index=False, sheet_name='NH_Transactions')

def wait_for_file_download(download_dir, timeout=60):
    elapsed = 0
    while elapsed < timeout:
        files = glob.glob(os.path.join(download_dir, '*.xls')) + \
                glob.glob(os.path.join(download_dir, '*.xlsx')) + \
                glob.glob(os.path.join(download_dir, '*.crdownload'))

        if any(file.endswith('.crdownload') for file in files):
            time.sleep(1)
            elapsed += 1
        elif files:
            return files[0]  # 첫 번째 다운로드된 파일 반환
        else:
            time.sleep(1)
            elapsed += 1

    raise TimeoutException("파일 다운로드가 완료되지 않았습니다.")

# 다운로드한 파일 이름 변경
def rename_downloaded_file(src_file, dest_file):
    shutil.move(src_file, dest_file)
    
def advanced_debug_for_image(driver, timeout=30):
    original_window = driver.current_window_handle
    wait = WebDriverWait(driver, timeout)

    wait.until(EC.number_of_windows_to_be(2))

    new_window = [window for window in driver.window_handles if window != original_window][0]
    driver.switch_to.window(new_window)

    time.sleep(5)

    iframes = driver.find_elements(By.TAG_NAME, 'iframe')
    if iframes:
        print(f"iframe이 존재함: {len(iframes)}개")
        driver.switch_to.frame(iframes[0])
    else:
        print("iframe이 존재하지 않음.")

    # 모든 요소를 탐색해서 CSS 배경 이미지까지 확인
    elements = driver.find_elements(By.XPATH, "//*")
    print(f"총 탐색된 요소 수: {len(elements)}개")

    target_base64_prefix = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0"

    found = False
    for idx, elem in enumerate(elements, 1):
        tag = elem.tag_name
        src = elem.get_attribute('src')
        style = elem.get_attribute('style')

        if (src and target_base64_prefix in src) or (style and target_base64_prefix in style):
            print(f"\n✅ 원하는 이미지 발견! 요소 번호: {idx}")
            print(f"태그: {tag}")
            if src:
                print(f"src 속성: {src[:100]}...")
            if style:
                print(f"style 속성: {style[:100]}...")
            xpath_script = """
            function getElementXPath(element) {
                if (element && element.id)
                    return '//*[@id="' + element.id + '"]';
                else
                    return getElementTreeXPath(element);
            }

            function getElementTreeXPath(element) {
                var paths = [];
                for (; element && element.nodeType == Node.ELEMENT_NODE; element = element.parentNode) {
                    var index = 0;
                    var hasFollowingSiblings = false;
                    for (var sibling = element.previousSibling; sibling; sibling = sibling.previousSibling) {
                        if (sibling.nodeType == Node.DOCUMENT_TYPE_NODE)
                            continue;

                        if (sibling.nodeName == element.nodeName)
                            ++index;
                    }

                    for (var sibling = element.nextSibling; sibling && !hasFollowingSiblings; sibling = sibling.nextSibling) {
                        if (sibling.nodeName == element.nodeName)
                            hasFollowingSiblings = true;
                    }

                    var tagName = (element.prefix ? element.prefix + ":" : "") + element.localName;
                    var pathIndex = (index || hasFollowingSiblings ? "[" + (index + 1) + "]" : "");
                    paths.splice(0, 0, tagName + pathIndex);
                }

                return paths.length ? "/" + paths.join("/") : null;
            }
            return getElementXPath(arguments[0]);
            """
            xpath = driver.execute_script(xpath_script, elem)
            print(f"✅ 자동 추출된 XPath:\n{xpath}\n")
            found = True
            break

    if not found:
        print("\n❌ 원하는 이미지가 DOM에서 발견되지 않았습니다.")

    input("확인 후 Enter를 눌러 종료하세요.")
    driver.switch_to.default_content()
    driver.close()
    driver.switch_to.window(original_window)

def download_excel_from_oz_report(driver, timeout=30):
    original_window = driver.current_window_handle
    wait = WebDriverWait(driver, timeout)

    wait.until(EC.number_of_windows_to_be(2))

    new_window = [window for window in driver.window_handles if window != original_window][0]
    driver.switch_to.window(new_window)
    
    time.sleep(5)

    try:
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        if iframes:
            driver.switch_to.frame(iframes[0])
            print("iframe이 존재하여 내부로 진입.")
        else:
            print("iframe이 존재하지 않아 현재 창에서 진행합니다.")

        # ✅ 발견된 최종 XPath 적용 (기존 유지)
        excel_button_xpath = "/html/body/div/div[1]/input[4]"
        wait.until(EC.element_to_be_clickable((By.XPATH, excel_button_xpath))).click()

        print("엑셀 버튼 클릭 성공!")
        time.sleep(5)

        driver.switch_to.default_content()

        # 🔵🔵🔵 추가된 부분: 확인 버튼 클릭 🔵🔵🔵
        # 확인 버튼 XPath는 반드시 기존에 찾아둔 XPath를 사용하세요
        confirm_button_xpath = "//button[text()='확인']"  # 예시 XPath, 정확한 XPath로 대체 필요
        wait.until(EC.element_to_be_clickable((By.XPATH, confirm_button_xpath))).click()
        print("확인 버튼 클릭 성공!")
        time.sleep(2)

    except (TimeoutException, ElementClickInterceptedException) as e:
        print("엑셀 버튼 또는 확인 버튼 클릭 실패:", e)
        raise RuntimeError("엑셀 버튼 또는 확인 버튼 클릭 실패")

    finally:
        driver.switch_to.window(new_window)
        driver.close()
        driver.switch_to.window(original_window)

# 엑셀 파일을 읽어서 업로드 양식으로 변환하는 함수
# 엑셀 파일을 읽어서 업로드 양식으로 변환하는 함수
def read_and_transform_downloaded_excel(downloaded_file):
    df = pd.read_excel(downloaded_file, header=None)

    # 데이터 시작 지점을 동적으로 찾기
    data_start_idx = None
    for idx, row in df.iterrows():
        if row.astype(str).str.contains('순번').any():
            data_start_idx = idx
            break
    
    if data_start_idx is None:
        raise ValueError("엑셀 파일에서 '순번'을 찾을 수 없습니다. 엑셀 파일 구조를 확인하세요.")

    df.columns = df.iloc[data_start_idx]
    df = df.iloc[data_start_idx + 1:].reset_index(drop=True)

    # 필요한 컬럼 존재 여부 확인
    required_columns = ['거래일시', '출금금액', '입금금액', '거래후잔액', '거래내용', '거래기록사항']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"필수 컬럼 '{col}'이(가) 엑셀 파일에 없습니다. 엑셀 파일 구조를 확인하세요.")

    df = df[required_columns]

    # 날짜 처리 안전하게 처리하는 함수
    def clean_date(date_str):
        if pd.isna(date_str) or str(date_str).strip().lower() == 'nan':
            return None
        date_str = str(date_str).replace('\n', ' ')
        try:
            date_obj = datetime.strptime(date_str.strip(), "%Y/%m/%d %H:%M:%S")
            return date_obj.strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            return None  # 포맷에 맞지 않으면 None으로 반환

    # 금액 처리 안전하게 처리하는 함수
    def clean_amount(amount):
        if pd.isna(amount) or amount == '':
            return 0
        return int(re.sub(r'[^\d]', '', str(amount)))

    # 날짜가 없는 행 제거
    df['cleaned_date'] = df['거래일시'].apply(clean_date)
    df = df[df['cleaned_date'].notnull()]

    # 💡 여기서 거래기록사항과 거래내용을 서로 바꿔서 반영합니다.
    transformed_df = pd.DataFrame({
        '구분': df.apply(lambda x: '출금' if clean_amount(x['출금금액']) else '입금', axis=1),
        '거래일자': df['cleaned_date'],
        '출금금액(원)': df['출금금액'].apply(clean_amount),
        '입금금액(원)': df['입금금액'].apply(clean_amount),
        '거래 후 잔액(원)': df['거래후잔액'].apply(clean_amount),
        '적요': df['거래기록사항'].fillna(''),   # 거래기록사항 → 적요
        '거래처': df['거래내용'].fillna('')        # 거래내용 → 거래처
    })

    return transformed_df

# 변환된 데이터를 별도의 엑셀 파일로 저장하는 함수
def save_transformed_excel(df, save_path):
    df.to_excel(save_path, index=False, sheet_name='업로드양식')
    print(f"업로드 양식으로 변환된 엑셀 파일 저장 완료: {save_path}")

# 실행 예시 (크롬 드라이버로 변경)
if __name__ == "__main__":
    CHROME_DRIVER_PATH = r"C:\chromedriver-win64\chromedriver.exe"
    DOWNLOAD_DIR = r"C:\NH_Transactions"  # 다운로드 폴더 경로 지정
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    account_number = "3020717230451"
    account_password = "1959"
    birthday = "601030"
    search_year = 2025
    search_month = 5

    transactions = get_balance(
        CHROME_DRIVER_PATH,
        account_number,
        account_password,
        birthday,
        search_year,
        search_month,
        DOWNLOAD_DIR
    )