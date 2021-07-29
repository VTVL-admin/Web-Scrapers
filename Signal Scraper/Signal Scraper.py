import time
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, StaleElementReferenceException
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import xlwings as xlw
import lxml
import cchardet

'''
Signal Investor Data Scraped from Website
'''

# open excel spreadsheet
wb = xlw.Book(r"C:\Users\{PATH_TO_DIRECTORY}\Signal Investor Data.xlsx")
# wb = xlw.Book(r"C:\Users\{PATH_TO_DIRECTORY}\Signal Investor Data Series A.xlsx")
sheet = wb.sheets['Sheet1']

# create Counter class to keep track of record number
class Counter:
    def __init__(self):
        self.counter = 0

    def incrementCounter(self):
        self.counter += 1

    def getCounter(self):
        return self.counter

# open selenium browser, fullscreen
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome("C:\\Users\\rosha\\ChromeDriver\\chromedriver.exe", options=options)
driver.get("https://signal.nfx.com/login")

counter = Counter()

# scrape data from table found on website
def scrape_data(source, count):
    bs = BeautifulSoup(source, 'lxml')
    rows = bs.find("table").find("tbody").find_all("tr")
    for r in rows[-8:]:
        count.incrementCounter()
        wrap = r.select("div.sn-investor-name-wrapper")[0]
        group = wrap.find_all("a")
        if len(group) > 1:
            sheet.range("A"+str(count.getCounter())).value = group[0].text
            sheet.range("C"+str(count.getCounter())).value = group[1].text
        else:
            sheet.range("A"+str(count.getCounter())).value = group[0].text
        try:
            role = wrap.find('span', {'class': 'sn-small-link hidden-xs null'})
            sheet.range("B" + str(count.getCounter())).value = role.text
        except AttributeError:
            pass
        sweet = r.select("div.flex-column")[1]
        sheet.range("D"+str(count.getCounter())).value = sweet.text
        if len(r.select("div.sn-clamp")) > 1:
            region = r.select("div.sn-clamp")[0]
            sheet.range("E"+str(count.getCounter())).value = region.text
            focus = r.select("div.sn-clamp")[1]
        else:
            focus = r.select("div.sn-clamp")[0]
        sheet.range("F" + str(count.getCounter())).value = "Series A"
        lists = [line.find('a') for line in focus.findChildren("span")]
        category = []
        for li in lists:
            link = li['href'].split("/")[2].split("-")
            res = link[1:len(link) - 1]
            category.append(" ".join(res).title())
        info = ", ".join(category)
        sheet.range("G"+str(count.getCounter())).value = info


# press button to get new data
def get_data(link, count):
    driver.get(link)
    time.sleep(1)
    while True:
        try:
            driver.find_element_by_tag_name("html").send_keys(Keys.END)
            button = driver.find_element_by_css_selector("button.btn-xs.btn-default.sn-center")
            scrape_data(driver.page_source, count)
            button.click()
            time.sleep(4)
        except NoSuchElementException:
            break
        except ElementNotInteractableException:
            break
        except StaleElementReferenceException:
            break

# alter between getting seed investors and series A investors
seed = driver.find_element_by_id('stage-seed')
series_a = driver.find_element_by_id('stage-series_a')
links = []
for i in series_a.find_elements_by_tag_name("a"):
    links.append(i.get_attribute('href'))
for l in links:
    get_data(l, counter)
