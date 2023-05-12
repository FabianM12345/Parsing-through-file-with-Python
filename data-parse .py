
#The purpose of this script is to parse through the files in a specific directory
#and transfer the date on the first row and totals into an excel sheet.

import os #This library will allow me to interact with the os.
import openpyxl #Lets me work with excel.
from datetime import date, datetime #Lets you work with dates.

dir = ' ' #Enter the path of the DIR you're working in.
balance_dates = [] #Hold dates
totals = [] #Holds Totals
current_time = datetime.now() #Gets Current Time
all_items = [] #Holds Dates + totals

file = 'File_'+current_time.strftime(str('%m%d%Y'))+'.xlsx' #The name of the excel file that will be created.
wb = openpyxl.Workbook() #Essentially it opens excel
ws = wb.active 

ws.append(['Date', 'Total']) #Creates the first row of the Excel sheet.

def get_file_info_totals(filename):
    #Remember to change dir in terminalS
    #Searches for a string in a file.
    with open(filename, 'r') as fp:
        #Read all lines using realine().
        lines = fp.readlines()
        
        for row in lines:
            #Checks if a string is present in current line.
            Totals_inLine = "Totals"
            if row.find(Totals_inLine) != -1:
                item_total = row[81:97]
                totals.append(str(item_total))


def get_file_info_dates(filename):
    #Remember to change dir in terminalS
    #Searches for a string in a file.
    with open(filename, 'r') as fp:
        #Read all lines using realine().
        lines = fp.readlines()
        
        for row in lines:
            #Checks if a string is present in current line.
            in_line = " " #Search for a string in the line.
            if row.find(in_line) != -1:
                item_date = row[68:77]
                balance_dates.append(item_date)

        

#This function gets the file names in the folder path.
def get_files_in_folder():
    for filename in os.listdir(dir):
      f = os.path.join(dir, filename)
      # checking if it is a file

      if os.path.isfile(f):
         get_file_info_dates(filename)
         get_file_info_totals(filename)


#Creates a string date-total adds to all_tems array.
def to_string_file_info():
    count = 0 #Counts the number of items.
    for i,j in zip(balance_dates,totals):
        new_item = i+'-'+j
        all_items.append(new_item)



get_files_in_folder()
to_string_file_info()

#Adds the items to the excel sheet created.
count = 0 #Counts the total items
for i in all_items:
    items = i.split('-')
    ws.append([items[0], items[1]])
    count = count + 1

#Saves the file.
wb.save(file) #Saves the excel sheet.
print(count)