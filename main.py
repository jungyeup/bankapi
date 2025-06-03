import time
from datetime import datetime
import os
from db_connector import fetch_pending_requests, update_request_status
from NH_BANK import get_balance as nh_personal
from NH_CORP_BANK import corp_get_balance as nh_corp

import shutil
import binascii
import logging
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import padding
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
import re
import subprocess
import getpass
import paramiko
import traceback



CHROME_DRIVER_PATH = r"C:\chromedriver-win64\chromedriver.exe"
BASE_DOWNLOAD_DIR = r"C:\BankLedgers"


def extract_core_error_message(e):
    """핵심 에러 메시지 추출"""
    return f"{type(e).__name__}: {str(e)}"

def sftp_makedirs(sftp, remote_directory):
    """원격 서버의 경로가 없으면 재귀적으로 디렉토리 생성"""
    dirs = remote_directory.split('/')
    path = ''
    for dir_component in dirs:
        if dir_component:
            path += '/' + dir_component
            try:
                sftp.stat(path)
            except FileNotFoundError:
                sftp.mkdir(path)

def upload_file_sftp(local_path, remote_path):
    """SFTP를 통해 원격 서버에 파일 업로드"""
    transport = paramiko.Transport((SFTP_HOST, SFTP_PORT))
    transport.connect(username=SFTP_USER, password=SFTP_PASS)
    sftp = paramiko.SFTPClient.from_transport(transport)

    remote_directory = os.path.dirname(remote_path)
    sftp_makedirs(sftp, remote_directory)  # 디렉토리 없으면 자동 생성

    sftp.put(local_path, remote_path)
    sftp.close()
    transport.close()

    logging.info(f"SFTP 업로드 완료: {local_path} → {remote_path}")

def launch_chrome_with_debugging(user_data_dir=None, debugging_port=9222):
    chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'

    subprocess.call('taskkill /f /t /im chrome.exe', shell=True)
    subprocess.call('taskkill /f /t /im chromedriver.exe', shell=True)
    time.sleep(2)

    username = getpass.getuser()

    if user_data_dir is None:
        user_data_dir = rf"C:\Users\{username}\Desktop\chrome_selenium_profile"

    os.makedirs(user_data_dir, exist_ok=True)

    cmd = (
        f'"{chrome_path}" --remote-debugging-port={debugging_port} '
        f'--user-data-dir="{user_data_dir}" '
        '--no-first-run --no-default-browser-check'
    )

    subprocess.Popen(cmd, creationflags=0x00000008, close_fds=True)
    time.sleep(3)  # 충분한 대기 시간 확보

def not_implemented_bank(*args, **kwargs):
    raise NotImplementedError("해당 은행의 자동화 처리가 아직 구현되지 않았습니다.")

bank_functions = {
    'BANK001': {'personal': nh_personal, 'corp': nh_corp},                 # 농협은행


def ensure_directory_exists(path):
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)

def move_and_return_web_path(src_file, dest_dir, new_filename):
    ensure_directory_exists(dest_dir)
    dest_file = os.path.join(dest_dir, new_filename)

    try:
        shutil.move(src_file, dest_file)
    except Exception as e:
        logging.error(f"파일 이동 중 에러 발생 [{src_file} → {dest_file}] : {e}")
        raise FileNotFoundError(f"파일 이동 중 에러 발생: {e}")

    if not os.path.isfile(dest_file):
        logging.error(f"[파일 이동 실패] {dest_file}")
        raise FileNotFoundError(f"파일 이동 실패: {dest_file}")
    else:
        logging.info(f"[파일 이동 성공] {dest_file}")

    web_path = '/'.join([dest_dir.replace('\\', '/'), new_filename])
    logging.info(f"[웹 경로 생성 완료] {web_path}")
    return web_path

# 🔷 AES128 복호화 함수
def decrypt_aes128(encrypted_hex):
    try:
        backend = default_backend()
        cipher = Cipher(algorithms.AES(AES_KEY), modes.ECB(), backend=backend)
        decryptor = cipher.decryptor()

        encrypted_data = binascii.unhexlify(encrypted_hex)

        decrypted_padded = decryptor.update(encrypted_data) + decryptor.finalize()

        unpadder = padding.PKCS7(algorithms.AES.block_size).unpadder()
        plaintext = unpadder.update(decrypted_padded) + unpadder.finalize()

        return plaintext.decode('utf-8')
    except Exception as e:
        logging.error("AES128 복호화 실패: %s", str(e))
        return None

def is_hex(s):
    return bool(re.fullmatch(r'[0-9a-fA-F]+', s))

def format_birthdate(birthdate):
    """날짜형식을 YYMMDD로 변환하는 함수"""
    try:
        return datetime.strptime(birthdate, '%Y-%m-%d').strftime('%y%m%d')
    except ValueError:
        return birthdate  # 이미 올바른 형식이면 그대로 반환

def check_file_exists(filepath):
    """
    파일의 존재 여부를 명확하게 체크하고, 로그로 기록합니다.
    """
    if os.path.isfile(filepath):
        logging.info(f"[파일 존재 확인] {filepath}")
        return True
    else:
        logging.error(f"[파일 미존재] {filepath}")
        return False
    
def execute_request(request):
    bank_code = request['BANK_SE']
    account_type = request['ACCOUNT_SE']
    req_seq = request['REQ_SEQ']
    ini_hptl_no = request['INI_HPTL_NO']

    update_request_status(req_seq, 'P')

    launch_chrome_with_debugging()

    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    download_dir = os.path.join(BASE_DOWNLOAD_DIR, f"{ini_hptl_no}_{req_seq}_{timestamp}")
    ensure_directory_exists(download_dir)

    try:
        start_date_str = request['SCH_BGNDE']
        end_date_str = request['SCH_ENDDE']

        start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d')

        account_pw_plain = decrypt_aes128(request['ACCOUNT_PW'])
        if not account_pw_plain:
            raise ValueError("ACCOUNT_PW 복호화 실패")

        rprsntv_brthdy_plain = request['RPRSNTV_BRTHDY']
        if rprsntv_brthdy_plain:
            if is_hex(rprsntv_brthdy_plain):
                rprsntv_brthdy_plain = decrypt_aes128(rprsntv_brthdy_plain)
                if not rprsntv_brthdy_plain:
                    raise ValueError("RPRSNTV_BRTHDY 복호화 실패")
            rprsntv_brthdy_plain = format_birthdate(rprsntv_brthdy_plain)

        account_pw2_plain = request['ACCOUNT_PW2']
        if account_pw2_plain and is_hex(account_pw2_plain):
            account_pw2_plain = decrypt_aes128(account_pw2_plain)
            if not account_pw2_plain:
                raise ValueError("ACCOUNT_PW2 복호화 실패")

        account_number = re.sub(r'\D', '', request['ACCOUNT'])

        print(f"[요청처리정보] REQ_SEQ={req_seq}, INI_HPTL_NO={ini_hptl_no}, BANK_CODE={bank_code}, "
              f"ACCOUNT_TYPE={account_type}, START_DATE={start_date_str}, END_DATE={end_date_str}, "
              f"ACCOUNT={account_number}, ACCOUNT_PW={account_pw_plain}, "
              f"RPRSNTV_BRTHDY={rprsntv_brthdy_plain}, ACCOUNT_PW2={account_pw2_plain}")

        if bank_code in ('BANK001', 'BANK011'):
            func = bank_functions[bank_code]['corp' if account_type == '01' else 'personal']
            original_excel, upload_excel = func(
                CHROME_DRIVER_PATH,
                account_number,
                account_pw_plain,
                request['BIZRNO'] if account_type == '01' else rprsntv_brthdy_plain,
                start_date,
                end_date,
                download_dir
            )



        # 엑셀 파일 생성 이후 부분에서...
        ext_original = os.path.splitext(original_excel)[1]
        ext_upload = os.path.splitext(upload_excel)[1]

        api_filename = f"{req_seq}_API_{timestamp}{ext_original}"
        upload_filename = f"{req_seq}_{timestamp}{ext_upload}"


        # SFTP를 통한 파일 업로드 호출
        upload_file_sftp(original_excel, api_remote_path)
        upload_file_sftp(upload_excel, upload_remote_path)

        # DB에 저장할 웹 경로 (리눅스 경로 그대로 사용)
        api_excel_web_path = api_remote_path
        upload_excel_web_path = upload_remote_path

        # 파일 업로드 성공 이후 DB 상태 업데이트
        update_request_status(req_seq, 'S', excel_file=upload_excel_web_path, excel_api_file=api_excel_web_path)

    except Exception as e:
        # 전체 traceback은 로컬 로그로 저장
        logging.error(f"요청 처리 중 에러 발생: {traceback.format_exc()}")

        # 핵심 메시지만 DB에 저장
        core_error_msg = extract_core_error_message(e)
        update_request_status(req_seq, 'E', err_msg=core_error_msg)
        
def main():
    while True:
        request = fetch_pending_requests()
        if request:
            execute_request(request)
        else:
            print(f"{datetime.now()} 처리할 요청이 없습니다.")
        time.sleep(1)

if __name__ == "__main__":
    main()
