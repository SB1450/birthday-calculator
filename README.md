# Birthday Calculator

## This calcultor builded with python 3.10

This is a python script that get a text file with names and birthdays as argument and calculate and create exel file with all names and their next birthday and in addition generate pop-out window that shows who has the closest birthday.

### ****Before you run the script you need to Install Tkinter on Ubuntu Linux (already installed on windows) AND install the requirements.txt file**
1.  Install Tkinter on Ubuntu Linux:
```
sudo apt install python3-tk
```           
<sub>more about Tkinter installation on ubuntu: https://www.pythonguis.com/installation/install-tkinter-linux/#</sub> 

<sub>Tltinker documentation for windows: https://tkdocs.com/tutorial/install.html#installwin</sub>

2. Install the requirements.txt file
```
pip install -r requirements.txt
```
<sub>if you have problem with modules check [here](https://www.quora.com/I-used-pip-to-install-a-library-but-when-I-import-it-it-says-Module-Not-Found-Why-is-that)</sub>

============================================================================================================

The script.py is a python script that get one argument:
Text-file - file with names and birthdays seperated by comma

The script create "Exel file" - file that will contain the results of the script, including those columns: Name, Birthday, Age, Days(until next birthday)

* The order of the Exel-file will be according to person with closest birthday to person with farest birthday

**Usage: script.py Text-file <-all> <-days range->**

There is 2 **optional arguments** that can be given to script in adittion to text-file argument:
1. **-all** - print to screen massage all people in the list
2. **days-range** (in number)- if "-all" flag is given you can choose the days range of birthdays you want to see for example 100 will show only people in list that have birthday within 100 days or less
  
<sub>e.g.: python script.py textfile -all 100</sub>
