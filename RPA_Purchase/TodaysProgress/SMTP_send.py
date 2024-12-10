import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import os
import util as ut
# import xlsxProcess as xlp

# 필요한 파일 다운로드
config = ut.load_config()
current_dir = ut.exedir('py')
download_dir = os.path.join(current_dir, "download")
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

chrome_options = Options()
# chrome_options.add_argument("headless")
chrome_options.add_experimental_option("detach", True)
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option("prefs", {
"download.default_directory": download_dir,  # 다운로드 폴더 지정
"download.prompt_for_download": False,  # 다운로드 시 사용자에게 묻지 않음
"download.directory_upgrade": True,
"safebrowsing.enabled": True,
"credentials_enable_service": False,
"profile.password_manager_enabled": False})
driver = webdriver.Chrome(options=chrome_options)
driver.get("https://iacf.sejong.ac.kr/main_0001_08.act")
driver.maximize_window()
driver.find_element(By.CSS_SELECTOR, config["iacf"]["ID_CSS"]).send_keys(config["iacf"]["ID"])
driver.find_element(By.CSS_SELECTOR, config["iacf"]["PW_CSS"]).send_keys(config["iacf"]["PW"])
driver.find_element(By.CSS_SELECTOR, config["iacf"]["LOGIN_BT_CSS"]).click()
sleep(5)
main_window = driver.window_handles
for i in main_window:
    if i != main_window[0]:
        driver.switch_to.window(i)
        driver.close()
sleep(3)
driver.switch_to.window(main_window[0])
driver.find_element(By.CSS_SELECTOR, "#menu_id_3").click()
iframes = driver.find_elements(By.TAG_NAME, 'iframe')
driver.switch_to.frame(iframes[1])
driver.find_element(By.CSS_SELECTOR, "#START_DATE").clear()
driver.find_element(By.CSS_SELECTOR, "#START_DATE").send_keys(config["iacf"]["ST_DATE"])
driver.find_element(By.CSS_SELECTOR, "#btnSearch").click()
sleep(10)
driver.find_element(By.CSS_SELECTOR, "#gridDonw > img").click()
sleep(3)

# 다운로드 받은 엑셀파일의 pandas 작업
# xlp.xlsxprocess(filename)

# 이메일 정보
sender_name = config["EMAIL"]["Sender_NAME"]
sender_email = config["EMAIL"]["Sender_address"]  # 발신자 이메일 주소
password = config["EMAIL"]["PW"]  # 발신자 이메일 비밀번호
# 받는 사람 목록
recipients = [
    {"name":config["EMAIL_RECEIVER"]["NAME1"], "email":config["EMAIL_RECEIVER"]["address1"]},
    {"name":config["EMAIL_RECEIVER"]["NAME2"], "email":config["EMAIL_RECEIVER"]["address2"]},
    {"name":config["EMAIL_RECEIVER"]["NAME3"], "email":config["EMAIL_RECEIVER"]["address3"]}
]

# 엑셀 파일 경로
excel_file = os.path.join(download_dir, "구매요구목록.xlsx")

# 이메일 메시지 기본 구성
base_msg = MIMEMultipart()
base_msg['From'] = formataddr((sender_name, sender_email))
base_msg['Subject'] = '일일 구매진행현황 보고드립니다.'

# 본문 내용
body = """
\n 안녕하십니까 선생님,
\n 오늘의 구매진행현황 알림드립니다.
\n 즐거운 하루 되세요^^
"""
base_msg.attach(MIMEText(body, 'plain'))

# 엑셀 파일 첨부
with open(excel_file, 'rb') as attachment:
    part = MIMEApplication(attachment.read(), Name="구매진행현황_")
    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(excel_file)}"'
    base_msg.attach(part)

# SMTP 서버에 연결하고 이메일 전송
try:
    with smtplib.SMTP('smtp.sejong.ac.kr', 25) as server:  # SMTP 서버와 포트 번호 입력
        server.starttls()  # TLS 보안 연결 시작
        server.login(sender_email, password)  # 로그인
        for recipient in recipients:
            # 각 수신자에 대해 개별 이메일 메시지 생성
            individual_msg = MIMEMultipart()
            individual_msg['From'] = formataddr((sender_name, sender_email))
            individual_msg['To'] = formataddr((recipient["name"], recipient["email"]))
            individual_msg['Subject'] = base_msg['Subject']
            individual_msg.attach(MIMEText(body, 'plain'))  # 본문 내용
            individual_msg.attach(part)  # 첨부파일 추가

            # 이메일 전송
            server.send_message(individual_msg)
            print(f"이메일이 {recipient['name']}에게 성공적으로 전송되었습니다.")
except Exception as e:
    print(f"이메일 전송 중 오류가 발생했습니다: {e}")