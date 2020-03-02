import time
import xlsxwriter
import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from datetime import date

#Change this to proper location
loc = ("A:\General\Coding\Python\Ikariam\DateIkariam.xlsx")

#It reads the previous Excecution Date
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
PreviousDate = sheet.cell_value(0, 0)
Today = date.today()
Today = Today.strftime("%m-%d-%Y")
print("Today and Previous Date",Today,"",PreviousDate)
if (Today==PreviousDate):
    print("Same Day No ACTION")
    time.sleep(1)
else:
    #If the day isn;t the same go to Piracy
    driver = webdriver.Chrome()

    #Change link to your World
    driver.get("https://lobby.ikariam.gameforge.com/el_GR/")
    assert "Ikariam" in driver.title
    driver.maximize_window()
    element = driver.find_element_by_name("email")
    element.send_keys("<-------Your Username------->")
    element = driver.find_element_by_name("password")
    element.send_keys("<-------Your Password------->")
    element.send_keys(Keys.ENTER)
    time.sleep(1)
    element = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div[2]/div[2]/div[2]/div/button")
    element.send_keys(Keys.ENTER)
    time.sleep(1)
    driver.switch_to.window(driver.window_handles[0])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(1)

    #Change with your Piracy Fortress Link
    driver.get('<-------Your Link------->')
    element = driver.find_element_by_xpath("//*[@id='pirateHighscoreNext']")
    element.click()
    time.sleep(1)
    element.click()
    actions = ActionChains(driver)
    actions.send_keys(Keys.TAB)
    for i in range(0,114):
        actions.perform()
    actions.send_keys(Keys.ENTER)
    actions.perform()
    time.sleep(2)
    workbook  = xlsxwriter.Workbook('DateIkariam.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0 , Today)
    workbook.close()
    driver.close()
