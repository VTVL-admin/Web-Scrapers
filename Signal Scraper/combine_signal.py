import  xlwings as xlw

'''
Get data from Series A file and Seed file and combine into one sheet
'''

# open excel spreadsheet
wb1 = xlw.Book(r"C:\Users\{PATH_TO_DIRECTORY}\Signal Investor Data.xlsx")
wb2 = xlw.Book(r"C:\Users\{PATH_TO_DIRECTORY}\Signal Investor Series A Data.xlsx")

# get columns from excel spreadsheet
normal = wb1.sheets[0].range('A2:A3593').value
series_a = wb2.sheets[0].range("A2:A956").value
series_a_type = wb2.sheets[0].range("B2:B956").value
series_a_company = wb2.sheets[0].range("C2:C956").value
series_a_sweet = wb2.sheets[0].range("D2:D956").value
series_a_geog = wb2.sheets[0].range("E2:E956").value
series_a_stage = wb2.sheets[0].range("F2:F956").value
series_a_focus = wb2.sheets[0].range("G2:G956").value
count = 3593
# check if data from series A is found in seed data (remove duplicates)
for i, j in enumerate(series_a):
    if j in normal:
        pos = normal.index(j)
        wb1.sheets[0].range('F'+str(pos+2)).value = "Seed, Series A"
    else:
        count += 1
        wb1.sheets[0].range('A'+str(count)).value = j
        wb1.sheets[0].range('B'+str(count)).value = series_a_type[i]
        wb1.sheets[0].range('C'+str(count)).value = series_a_company[i]
        wb1.sheets[0].range('D'+str(count)).value = series_a_sweet[i]
        wb1.sheets[0].range('E'+str(count)).value = series_a_geog[i]
        wb1.sheets[0].range('F'+str(count)).value = series_a_stage[i]
        wb1.sheets[0].range("G"+str(count)).value = series_a_focus[i]

              
