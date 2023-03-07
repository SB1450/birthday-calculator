##########################   Block 0    ##########################

from datetime import datetime
from datetime import date
import openpyxl
import pandas as pd
import sys
import os
import tkinter as tk
from tkinter import messagebox
import xlsxwriter
# import notify2
# import numpy as np
# from openpyxl import Workbook


##########################   Block 1    ##########################

def usage():
  print("\033[1mUsage: Script.py <text_file>\033[0m")
  return exit(1)



## Function to calculate age
def age(birthdate):
  today = date.today()
  age = today.year - birthdate.year - ((today.month, today.day) < (birthdate.month, birthdate.day))
  return age


## Define all lists and variables
exel_file = "DEMO.xlsx"
notf_before_next_bd = 30
path = os.getcwd()
help_call = ["-h", "-H", "--help", "-help"]
try: text_file = sys.argv[1]
except IndexError: usage()
if sys.argv[1] in help_call: usage()
names = []
birthdays = []
ages = []
diff_days = []
lt_30 = {}
top_line= ["Name", "Birthday", "Age", "Days"]
today = date.today().strftime("%d/%m/%Y")

## Check if files and path is correct
for arg in range(1, len(sys.argv)):
  if sys.argv[arg] == path:
    isExist = os.path.exists(sys.argv[arg])
  else:
    if path[-1] != "/":
      path = path + "/"
    isExist = os.path.exists(f"{path}{sys.argv[arg]}")
  if isExist == False:
    print("Invalid argument!\n")
    usage()


## Run threw file lines
with open(f'{path}{text_file}') as f:
  next(f)
  lines = [line.rstrip('\n') for line in f]

  ## Put Name, Birthday into lists
for line in lines:
  name = line.split(',', 1)[0]
  birthday = line.split(',', 1)[1]
  names.append(name)
  birthdays.append(birthday.rstrip('\n'))

## Save all parts of bithday into variables
  t_day = int(today[0:2])   ## today day
  t_month = int(today[3:5]) ## today month
  t_year = int(today[6:])   ## today year
  b_day = str(birthday[0:2])    ## bitrhday day
  b_month = str(birthday[3:5])  ## bitrhday month
  b_year = str(birthday[6:])    ## bitrhday year
  this_year = today[6:]     ## todays year

  ## Put Age into lists
  Age = age(date(int(b_year), int(b_month), int(b_day)))
  ages.append(Age)


##########################   Block 2    ##########################

  ## Calculates the differnece between today and birthday in days
  if t_month > int(b_month):
    this_year= str(int(this_year)+1)
  elif t_month == int(b_month) and t_day > int(b_day):
    this_year= str(int(this_year)+1)

  str_d1 = f"{b_day}/{b_month}/{this_year}"
  str_d2 = today

  ## convert string to date object
  d1 = datetime.strptime(str_d1, "%d/%m/%Y")
  d2 = datetime.strptime(str_d2, "%d/%m/%Y")

  ## diff_dayserence between dates in timedelta
  delta = abs(d2 - d1)
  diff_days.append(delta.days)


##########################   Block 3    ##########################

## Create exel file and give it write premission
workbook = xlsxwriter.Workbook(f"{exel_file}")
worksheet = workbook.add_worksheet()
workbook.close()
os.chmod(f"{path}{exel_file}", 0o755)
## Open the file and make changed
wb = openpyxl.load_workbook(f'{path}{exel_file}')
sheet = wb.active

## Add the headline to the exel file
for index, value in enumerate(top_line):
  sheet.cell(row=1 ,column=index+1 ,value=value)
  
def edit_exel(ROW, COL, data_list):
  for index, value in enumerate(data_list):
    sheet.cell(row=index+ROW ,column=COL ,value=value)
  
# Call the function and insert the data into columns
edit_exel(2, 1, names)
edit_exel(2, 2, birthdays)
edit_exel(2, 3, ages)
edit_exel(2, 4, diff_days)
wb.save(f'{path}{exel_file}')


##########################   Block 4    ##########################

## Check the name of person with the closest birthday
days=int(min(diff_days))
for i in range(2, sheet.max_row+1):
  cell_obj = sheet.cell(row=i, column=4)
  if cell_obj.value == days:
    name = sheet.cell(row=i, column=1).value

# Output to screen with pop-out window
if min(diff_days) >= notf_before_next_bd:
## Pop up windows
  root = tk.Tk()
  root.withdraw()
  messagebox.showwarning('Birthday', f"Closest birthday to {name} in {days} days")


##########################   Block 5 - Extra    ##########################

## Sort exel file by column name
# df = pd.read_excel(f'{path}{exel_file}')
# result = df.sort_values('Column-name')
# print(result)





