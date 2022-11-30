#!/bin/python3

##########################   Block 0    ##########################

from datetime import datetime
from datetime import date
import openpyxl
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import sys
import os
# import numpy as np
# import notify2
# from openpyxl import Workbook


##########################   Block 1    ##########################

def usage():
  print("\033[1mUsage: Script.py <text_file> <exel_file>\033[0m")
  print("Important notes:\n* Exel file must already be exist\n* Must specify full path")#•
  return exit(1)



# Function to calculate age
def age(birthdate):
  today = date.today()
  age = today.year - birthdate.year - ((today.month, today.day) < (birthdate.month, birthdate.day))
  return age


# Define all lists and variables
path = os.getcwd()
text_file = sys.argv[1]
exel_file = sys.argv[2]
names = []
birthdays = []
ages = []
diff_days = []
lt_30 = {}
top_line= ["Name", "Birthday", "Age", "Days"]
today = date.today().strftime("%d/%m/%Y")

# Check if files and path are correct
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


# Run threw file lines
with open(f'{path}{text_file}') as f:
  next(f)
  lines = [line.rstrip('\n') for line in f]

  # Put Name, Birthday into lists
for line in lines:
  name = line.split(',', 1)[0]
  birthday = line.split(',', 1)[1]
  names.append(name)
  birthdays.append(birthday.rstrip('\n'))

# Save all parts of bithday into variables
  t_day = int(today[0:2])   # today day
  t_month = int(today[3:5]) # today month
  t_year = int(today[6:])   # today year
  b_day = str(birthday[0:2])    # bitrhday day
  b_month = str(birthday[3:5])  # bitrhday month
  b_year = str(birthday[6:])    # bitrhday year
  this_year = today[6:]     # todays year

  # Put Age into lists
  Age = age(date(int(b_year), int(b_month), int(b_day)))
  ages.append(Age)


##########################   Block 2    ##########################

  # Calculates the differnece between today and birthday in days
  if t_month > int(b_month):
    this_year= str(int(this_year)+1)
  elif t_month == int(b_month) and t_day > int(b_day):
    this_year= str(int(this_year)+1)

  str_d1 = f"{b_day}/{b_month}/{this_year}"
  str_d2 = today

  # convert string to date object
  d1 = datetime.strptime(str_d1, "%d/%m/%Y")
  d2 = datetime.strptime(str_d2, "%d/%m/%Y")

  # diff_dayserence between dates in timedelta
  delta = abs(d2 - d1)
  diff_days.append(delta.days)


##########################   Block 3    ##########################

os.chmod(f"{path}{exel_file}", 0o644)
wb = openpyxl.load_workbook(f'{path}{exel_file}')
sheet = wb.active

# Add top line of file
for index, value in enumerate(top_line):
  sheet.cell(row=1 ,column=index+1 ,value=value)

# Add\Update names in Exel file
for index, value in enumerate(names):
  sheet.cell(row=index+2 ,column=1 ,value=value)

# Add\Update birthdays in Exel file
for index, value in enumerate(birthdays):
  sheet.cell(row=index+2 ,column=2 ,value=value)

# Add\Update ages in Exel file
for index, value in enumerate(ages):
  sheet.cell(row=index+2 ,column=3 ,value=value)

# Add\Update days until Birthdays in Exel file and save all changes
for index, value in enumerate(diff_days):
  sheet.cell(row=index+2 ,column=4 ,value=value)
wb.save(f'{path}{exel_file}')


##########################   Block 4    ##########################

# Check the name of person with the closest birthday
# for i in diff_days:
#   if i < 30:
#     lt_30[f"{i}"] = ""
# days=int(min(diff_days))
for i in range(2, sheet.max_row+1):
  cell_obj = sheet.cell(row=i, column=4)
  if cell_obj.value == None: break
  if cell_obj.value < 50:
    lt_30[cell_obj.value] = sheet.cell(row=i, column=1).value
lt_30 = dict(sorted(lt_30.items()))

# Build message and output it to screen with pop-out window
msg = ""
for key, value in lt_30.items():
  msg += f"{value} - {key} days\n"
# Only Text Notification
# notify2.init('Basic')
# notify2.Notification('Birthday', f"Closest birthday to {name} in {days} days").show()
# Pop up windows
root = tk.Tk()
root.withdraw()
messagebox.showwarning('Birthday', f"{msg}")


##########################   Block 5 - Extra    ##########################

# Sort exel file by column name
# df = pd.read_excel(f'{path}{exel_file}')
# result = df.sort_values('Days')
# print(result)





