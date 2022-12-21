# Birthday Calculator

## This calcultor builded with python 3.10

This is a python script that get a text file with names and birthdays as argument and calculate and create exel file with all names and their next birthday and in addition generate pop-out window that shows who has the closest birthday.

****Before you run the script you need to Install Tkinter on Ubuntu Linux AND install the requirements.txt file**
1.  Install Tkinter on Ubuntu Linux:
```
sudo apt install python3-tk
```           
<sub>more about Tkinter installation on ubuntu: https://www.pythonguis.com/installation/install-tkinter-linux/#</sub> 

2. Install the requirements.txt file
```
pip install -r requirements.txt
```


The script.py is a python script that get one argument:
Text-file - file with names and birthdays seperated by comma

The script create "Exel file" - file that will contain the results of the script, including those columns: Name, Birthday, Age, Days(until next birthday)

* The order of the Exel-file will be according to the names order in the Text-file. At the end of the script I added code block that allow you to sort the output of Exel-file by the column you want.

**Usage: script.py Text-file**
