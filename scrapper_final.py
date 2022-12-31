from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import webdriver_manager.chrome
import time
from openpyxl import Workbook, load_workbook
import datetime

now = datetime.datetime.now()
sheet_name = now.strftime("%A")
#sheet_name = "Monday"

# For searching short and long keywords using chrome driver
def SearchKeyword(keywords):
   my_options = webdriver.ChromeOptions()
   my_options.add_argument('--headless')
   driver = webdriver.Chrome(webdriver_manager.chrome.ChromeDriverManager().install(), options=my_options)
   driver.maximize_window()

   ls_short = []
   ls_long = []

   for keyword in keywords:
      driver.get("https://www.google.com/")
      #keyword="book"
      driver.find_element_by_class_name("gLFyf").send_keys(keyword)
      time.sleep(1.5)

      s = driver.find_elements(By.CSS_SELECTOR, "div.wM6W7d")[1:-1]

      dict_data = {}


      for x in s:

            #print(len(x.text),x.text)
            dict_data[len(x.text)] = x.text

      #wM6W7d
      time.sleep(2)

      #print(dict_data)
      max_value = max(dict_data.keys())
      min_value = min(dict_data.keys())
      long_keyword = dict_data[max_value]
      short_keyword = dict_data[min_value]
      print(keyword+" > "+short_keyword+" & "+long_keyword)
      ls_short.append(short_keyword.capitalize())
      ls_long.append((long_keyword.capitalize()))
      driver.refresh()
   driver.close()
   driver.quit()
   return keyword, ls_short, ls_long

# Selecting Existing Excel File
wb = load_workbook("data.xlsx")

ws = wb.active
wv = wb[sheet_name]

# Grabing input data from Excel File
ls_data = []
for column_data in wv['B'][1:]:
   # Printing the column values of every row
   column_data = column_data.value
   ls_data.append(column_data)

#Feeding Data to the Function
keyword,ls_short,ls_long = SearchKeyword(ls_data)

# Saving output data in xlsx Format
for row in range(0,len(ls_short)):
    for col in range(0,len(ls_short)):
        wv["C"+str(col+2)] = ls_long[col]
        wv["D" + str(col + 2)] = ls_short[col]

wb.save(filename='data.xlsx')