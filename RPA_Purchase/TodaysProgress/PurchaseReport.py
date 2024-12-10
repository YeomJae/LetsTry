from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os
import util as ut
import xlsxProcess as xlp

config = ut.load_config()
current_dir = ut.exedir('py')
download_dir = os.path.join(current_dir, "download")
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

chrome_options = Options()
# chrome_options.add_argument("headless")
chrome_options.add_experimental_option("detach", True)
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # 자동화 탐지 방지
chrome_options.add_argument("--disable-features=IsolateOrigins,site-per-process")  # 일부 보안 기능 비활성화
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
# for i, iframe in enumerate(iframes):
#    print(f"iframe {i}: {iframe.get_attribute('id')}, {iframe.get_attribute('name')}, {iframe.get_attribute('src')}")

driver.switch_to.frame(iframes[1])
driver.find_element(By.CSS_SELECTOR, "#START_DATE").clear()
driver.find_element(By.CSS_SELECTOR, "#START_DATE").send_keys("20240501")
driver.find_element(By.CSS_SELECTOR, "#btnSearch").click()
sleep(10)
driver.find_element(By.CSS_SELECTOR, "#gridDonw > img").click()
sleep(3)

# xlp.xlsxprocess(filename)