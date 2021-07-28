import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
import pyautogui
from bs4 import BeautifulSoup
import re
import xlwings as xlw

"""
Media Investor Data Scraped from Airtable
"""

# open excel spreadsheet
wb = xlw.Book("C:\\Users\\{PATH_TO_DIRECTORY}\\Media Investor Data.xlsx")
sheet = wb.sheets['Sheet1']

# open selenium Chrome browser
options = webdriver.ChromeOptions()
driver = webdriver.Chrome("C:\\Users\\{PATH_TO_DIRECTORY}\\chromedriver.exe", options=options)
# get airtable webpage
def get_airtable():
    driver.get("https://airtable.com/shrw3Bd83xkLo2CiJ/tblMjVmr3i7XK7usQ")
    pyautogui.moveTo(600, 700)
    try:
        driver.find_element_by_css_selector("div.focus-visible-opaque.pointer").click()
        time.sleep(1)
    except NoSuchElementException:
        pass
    except StaleElementReferenceException:
        pass
get_airtable()
slider = driver.find_element_by_class_name("antiscroll-scrollbar-shown")
time.sleep(4)
data_1 = {}
# slide only 100px to the right
ActionChains(driver).click_and_hold(slider).move_by_offset(100, 0).release().perform()

# get data from the left side of the airtable
while True:
    bs = BeautifulSoup(driver.page_source, "lxml")
    data = bs.select("div.cell")
    format_data = [line.text.strip() if line.text != "" else "None" for line in data]
    names = format_data[1: format_data.index("")]
    info = format_data[format_data.index("Example Portfolio Co's") + 1:]
    normal = [info[y:y + 4] for y in range(0, len(info), 5)]
    res = {names[i]: normal[i] for i in range(len(names))}
    data_1.update(res)
    time.sleep(1)
    pyautogui.scroll(-100)
    if len(data_1) == 155:
        ActionChains(driver).click_and_hold(slider).move_by_offset(-100, 0).release().perform()
        pyautogui.scroll(4000)
        time.sleep(4)
        break

get_airtable()
slider = driver.find_element_by_class_name("antiscroll-scrollbar-shown")
time.sleep(4)
# slide only 400px to the right
ActionChains(driver).click_and_hold(slider).move_by_offset(400, 0).release().perform()
time.sleep(1)
data_2 = {}
# get data from the right side of the airtable
while True:
    bs = BeautifulSoup(driver.page_source, "lxml")
    data = bs.select("div.cell")
    format_data = [line.text.strip() if line.text != "" else "None" for line in data]
    names = format_data[1: format_data.index("")]
    info = format_data[format_data.index("Example Portfolio Co's") + 1:]
    scrolled = [info[y:y + 6] for y in range(1, len(info), 7)]
    new_res = {names[i]: scrolled[i] for i in range(len(names))}
    data_2.update(new_res)
    time.sleep(1)
    pyautogui.scroll(-100)
    if len(data_2) == 155:
        break

# combine left side and right side data into one dictionary
full_data = {}
for key, val in data_1.items():
    if key in data_2:
        full_data[key] = val + data_2[key]

# add data to excel sheet
for index, (key, value) in enumerate(full_data.items()):
    sheet.range("A"+str(index+2)).value = key
    sheet.range("B"+str(index+2)).value = value[0]
    format_type = re.sub(r"(\))(\w)", r'\1, \2', value[1])
    sheet.range("C"+str(index+2)).value = format_type
    format_tag = re.sub(r"(\w)([A-Z])", r"\1, \2", value[2])
    sheet.range("D"+str(index+2)).value = format_tag
    format_investor = re.sub(r"(\w)([A-Z])", r"\1, \2", value[3])
    sheet.range("E"+str(index+2)).value = format_investor
    format_office = re.sub(r"(\w)([A-Z])", r"\1, \2", value[4])
    sheet.range("F"+str(index+2)).value = format_office
    format_geography = re.sub(r"(\w)([A-Z])", r"\1, \2", value[5])
    sheet.range("G"+str(index+2)).value = format_geography
    sheet.range("H"+str(index+2)).value = value[6]
    format_capital = re.sub(r"(\w)([A-Z])", r"\1, \2", value[7])
    sheet.range("I"+str(index+2)).value = format_capital
    sheet.range("J"+str(index+2)).value = value[8]
    sheet.range("K"+str(index+2)).value = value[9]
