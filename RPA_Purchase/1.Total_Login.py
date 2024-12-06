from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os
import util as ut

config = ut.load_config()

chrome_options = Options()
# chrome_options.add_argument("headless")
chrome_options.add_experimental_option("detach", True)
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option("prefs", {
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
driver.find_element(By.CSS_SELECTOR, "#START_DATE").send_keys("20240701")
driver.find_element(By.CSS_SELECTOR, "#btnSearch").click()
sleep(10)

driver.switch_to.default_content()
driver.switch_to.window(driver.window_handles[0])
driver.execute_script('window.open("https://mail.sejong.ac.kr");')

driver.switch_to.window(driver.window_handles[1])
driver.find_element(By.CSS_SELECTOR, config["mail"]["ID_CSS"]).send_keys(config["login"]["ID"])
driver.find_element(By.CSS_SELECTOR, config["mail"]["PW_CSS"]).send_keys(config["login"]["PW"])
driver.find_element(By.CSS_SELECTOR, config["mail"]["LOGIN_BT_CSS"]).click()
driver.find_element(By.CSS_SELECTOR, config["mail"]["RECEIVE_MAIL_CSS"]).click()

driver.switch_to.window(driver.window_handles[0])
driver.execute_script('window.open("http://sjgw.sejong.ac.kr/");')

# driver.find_element(By.CSS_SELECTOR, config["gw"]["ID_CSS"]).send_keys(config["login"]["ID"])
# driver.find_element(By.CSS_SELECTOR, config["gw"]["PW_CSS"]).send_keys(config["login"]["PW"])
# driver.find_element(By.CSS_SELECTOR, config["gw"]["LOGIN_BT_CSS"]).click()
