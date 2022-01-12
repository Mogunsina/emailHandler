import tkinter as tk
from tkinter import filedialog
import xlrd
import xlsxwriter

root = tk.Tk()
root.withdraw()
filename = filedialog.askopenfilename(title = 'Open Excel File')
loc = filename
wb = xlrd.open_workbook(loc)
emails = wb.sheet_by_index(1)

column1 = emails.col_values(0, start_rowx=2)
column2 = emails.col_values(1, start_rowx=2, end_rowx = None)
def Once(column1, column2):
    # nonmatch = []
    ret1 = set(column1) - set(column2)
    return(ret1)

def nonCommon(column1, column2):
    # nonmatch = []
    return(set(column2) - set(column1))

def Common(column1, column2):
    # match = []
    c1_set = set(column1)
    c2_set = set(column2)
    if(c1_set & c2_set):
        return(c1_set & c2_set)



##### code to write to an excel file #####
workbook = xlsxwriter.Workbook('C:/Users/Morayo/desktop/DASSAULT/sample.xlsx')
worksheet = workbook.add_worksheet("One")
worksheet1 = workbook.add_worksheet("Two")

### Writing a file to excel ###
content = Once(column1, column2)
content1 = nonCommon(column1,column2)

### WORKBOOK ONE ###
row1 = 0
column_1 = 0
row = 0
column = 0
for i in (content):
    worksheet.write(row, column, i)
    row += 1
for j in (content1):
    worksheet1.write(row1, column_1, j)
    row1 += 1
workbook.close()