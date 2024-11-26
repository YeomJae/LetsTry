from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

id = "byj"
pw = "Alsk123!"

chrome_options = Options()
# chrome_options.add_argument("headless")
chrome_options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=chrome_options)
driver.get("https://iacf.sejong.ac.kr/main_0001_08.act")
driver.maximize_window()

user_id = driver.find_element(By.CSS_SELECTOR, "#USER_NM")
user_id.send_keys("751370")
rerppw = driver.find_element(By.CSS_SELECTOR, "#USER_PW")
rerppw.send_keys(pw)
rlogin_btn = driver.find_element(By.CSS_SELECTOR, "#btn_login").click()
sleep(10)

mainw = driver.window_handles
for i in mainw:
    if i != mainw[0] :
        driver.switch_to.window(i)
        driver.close()

driver.switch_to.window(driver.window_handles[0])
driver.execute_script('window.open("https://mail.sejong.ac.kr");')

driver.switch_to.window(driver.window_handles[1])
mail_id = driver.find_element(By.CSS_SELECTOR, "#mailid")
mail_id.send_keys(id)
mqw = driver.find_element(By.CSS_SELECTOR, "#password")
mqw.send_keys(pw)
mlogin_btn = driver.find_element(By.CSS_SELECTOR, "#loginForm > div.login.skin1 > section > div > div.login_content > div.login_form.non_findId > div.input_area > ul > li.btn_line > button")
mlogin_btn.click()
mail_total = driver.find_element(By.CSS_SELECTOR, "#mnu_TOTAL")
mail_total.click()

driver.switch_to.window(driver.window_handles[1])
driver.execute_script('window.open("http://sjgw.sejong.ac.kr/");')