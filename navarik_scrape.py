from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from random import randint
import pandas as pd
from time import sleep
from selenium.webdriver.common.by import By


chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--start-maximized");

driver = webdriver.Chrome(
    ChromeDriverManager().install(),
    options = chrome_options
)
post_url = str('https://www.navarik.net/login/login.php')
driver.get(post_url)

sleep(randint(3,4))

def getHead(driver, dic):
    tableHead = driver.find_element(By.XPATH, '//*[@id="jobOrderContainer"]/thead').find_elements(By.TAG_NAME, 'th')
    for data in tableHead:
        if not data.text == '':
            dic[data.text] = []
    return dic

def getData(driver, dic):    

    tableBody = driver.find_element(By.XPATH, '//*[@id="jobOrderContainer"]/tbody').find_elements(By.TAG_NAME, 'tr')
    for data in tableBody:
        if len(data.find_elements(By.TAG_NAME, 'td')) > 1:
            x = 1
            for key in dic:
                d = data.find_elements(By.TAG_NAME, 'td')
                dic[key].append(d[x].text)
                x += 1
    return dic

def clickOffice(value):
    office = driver.find_element(By.ID,'officeId')
    office.click()

    for data in office.find_elements(By.TAG_NAME, 'option'):
        if data.get_attribute('value') == value:
            data.click()
            break
    
    search = driver.find_element(By.XPATH, '//*[@id="job_order_search"]/div[4]/div/button[1]').click()
    sleep(randint(3,4))

try:
    # Enter email
    email = driver.find_element(By.ID, '1-email')
    email.send_keys('emea.automations@amspecgroup.com')

    # Enter password
    password = driver.find_element(By.XPATH, '//*[@id="login-root"]/div/div/form/div/div/div/div[2]/div[2]/span/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/input')
    password.send_keys('Amspec2022!@')

    # Press log-in
    driver.find_element(By.XPATH, '//*[@id="login-root"]/div/div/form/div/div/div/button').click()
    sleep(randint(10, 15))
    dic = {}
    dic = getHead(driver, dic)
    dic = getData(driver, dic)

    clickOffice('155539')
    dic = getData(driver, dic)


    clickOffice('161082')
    dic = getData(driver, dic)

    clickOffice('161701')
    dic = getData(driver, dic)
    
    df = pd.DataFrame(dic)
    df.to_csv('navarik_scrape.xlsx', index=False)

    
except Exception as e:
    print(e)
    print('Error occured, moving to next account')
    driver.get(post_url)
    sleep(randint(3,4))
        
driver.close()