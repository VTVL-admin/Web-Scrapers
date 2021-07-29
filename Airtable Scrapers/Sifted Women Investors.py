import time
from selenium import webdriver
import pyautogui
from bs4 import BeautifulSoup
import xlwings as xlw
import ast

'''
Sifted Women Investor Data Scrape from Airtable
'''

# open excel book and active sheet
wb = xlw.Book(r"C:\Users\{PATH_TO_DIRECTORY}\Sifted Investor Data.xlsx")
sheet = wb.sheets['Sheet1']

# open selenium Chrome browser, fullscreen
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome("C:\\Users\\{PATH_TO_DIRECTORY}\\chromedriver.exe", options=options)
driver.get("https://airtable.com/shrWxNV1xJGDRW0pj/tbl08Q4Ks1CkLBZ1D")
time.sleep(5)
pyautogui.moveTo(600, 700)


# scrape data from each cell in airtable using BeautifulSoup
def automate_scrape(html):
    bs = BeautifulSoup(html, 'html.parser')
    data = bs.select('div.cell')
    format_data = [line.text if line.text != "" else "None" for line in data]
    full_names = format_data[:format_data.index("\n\n\n\n")]
    info = format_data[format_data.index("\n\n\n\n") + 1:]
    full_info = [info[x:x + 4] for x in range(0, len(info), 5)]
    res = {full_names[i]: full_info[i] for i in range(len(full_names))}
    return res


# write data into excel sheet
def write_to_excel(key_val, values, counter_val):
    sheet.range("A" + str(counter_val)).value = key_val
    sheet.range("C" + str(counter_val)).value = values[0]
    sheet.range("E" + str(counter_val)).value = values[1]
    sheet.range("G" + str(counter_val)).value = values[2]
    sheet.range("H" + str(counter_val)).value = values[3]


investors = ""
counter = 2
# Iterate through all rows in airtable and scrape data from cell
while True:
    # get data from text file in order to handle duplicate information
    with open("data.txt", "r") as f:
        str_dict = f.readline()
        if str_dict:
            investors = str_dict

    scrape_dict = automate_scrape(driver.page_source)
    if investors:
        # string dictionary to real dictionary converter
        file_dict = ast.literal_eval(investors)
        for key in scrape_dict.keys():
            if key in file_dict:
                continue
            else:
                investor_values = scrape_dict[key]
                write_to_excel(key, investor_values, counter)
                counter += 1
                # write data into text file
                with open('data.txt', "w") as f:
                    f.write(scrape_dict.__str__())
                pyautogui.scroll(-14)
    else:
        for key in scrape_dict.keys():
            investor_values = scrape_dict[key]
            write_to_excel(key, investor_values, counter)
            counter += 1
            with open('data.txt', "w") as f:
                f.write(scrape_dict.__str__())
            pyautogui.scroll(-14)

    if sheet.range("A196").value:
        driver.close()
        wb.save()
        wb.close()
        break
