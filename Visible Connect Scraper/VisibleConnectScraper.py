from bs4 import BeautifulSoup, SoupStrainer
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
import time
import xlwings as xlw
import lxml
import cchardet

"""
Visible Connect Investor Data Scraped from Website
"""

# open excel spreadsheet
wb = xlw.Book(r"C:\Users\{PATH_TO_DIRECTORY}\Visible Connect Data.xlsx")
sheet = wb.sheets["Sheet1"]


# open a headless Selenium Chrome Browser (No GUI)
options = Options()
options.headless = True
options.add_argument("--headless")
driver = webdriver.Chrome("C:\\Users\\{PATH_TO_DIRECTORY}\\chromedriver.exe", options=options)
driver.get("https://connect.visible.vc/investors?stages[]=Alt.%20VC&"
           "stages[]=Accelerator&stages[]=Pre-Seed&stages[]=Seed&stages[]=Series%20A&stages[]=Angel&stages[]=Growth")
time.sleep(4)


# write data to excel spreadsheet
def write_data(counter, ele):
    counter += 1
    contents = [v.text.strip() for v in ele.findChildren("div", {'class': 'text-gray-700'})]
    heading = [r.text.strip() for r in ele.findChildren("div", {'class': 'flex flex-col'})]
    for r in heading:
        raw = " ".join(r.split())
        if "Verified investor" in raw:
            arr = " ".join(r.split()).replace("Verified investor", ",").split(" , ")
            sheet.range("A"+str(counter)).value = arr[0]
            sheet.range("C"+str(counter)).value = arr[1]
        else:
            arr = raw.split(" http")
            arr[1] = "http" + arr[1]
            sheet.range("A"+str(counter)).value = arr[0]
            sheet.range("C"+str(counter)).value = arr[1]
    for y in contents:
        sheet.range("B"+str(counter)).value = contents[0]
        word = " ".join(y.split())
        calibrate(word, counter)


# format data from website
def calibrate(word, counter):
    if word.startswith("Stage:"):
        stage = word.replace("Stage:", "").strip()
        sheet.range("D" + str(counter)).value = stage
    if word.startswith("Check size:"):
        size = word.replace("Check size:", "").strip()
        sheet.range("F" + str(counter)).value = size
    if word.startswith("Focus:"):
        focus = word.replace("Focus:", "").strip()
        sheet.range("G" + str(counter)).value = focus
    if word.startswith("Investment geography:"):
        geography = word.replace("Investment geography:", "").strip()
        sheet.range("E" + str(counter)).value = geography


# get specific part of webpage to scrape data from
soup = SoupStrainer('div', {'class': 'mt-2 grid grid-cols-1 gap-5'})
counter = 1
# iterate through all records on webpage
while True:
    bs = BeautifulSoup(driver.page_source, 'lxml', parse_only=soup)
    driver.find_element_by_tag_name("html").send_keys(Keys.END)
    ele = bs.find_all("div", {"border-gray-200"})
    if counter == 5291:
        for e in ele[-5:]:
            counter += 1
            write_data(counter, e)
        break
    # always get the last 10 records everytime the button is pressed
    for e in ele[-10:]:
        counter += 1
        write_data(counter, e)
    try:
        # press button if it exists
        driver.find_element_by_css_selector('div.justify-center.mt-4.flex').click()
        time.sleep(1)
    except NoSuchElementException:
        break
    except StaleElementReferenceException:
        break
