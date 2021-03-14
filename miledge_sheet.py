import openpyxl
import datetime
import os
import string

operating_system = os.sys.platform

def find_cur_dir():
    cur_dir = os.path.abspath(__file__)
    slash_places = []
    if 'Linux' in operating_system:
        for index,value in enumerate(cur_dir):
            if value == '/':
                slash_places.append(index)
        cur_dir = cur_dir[:slash_places[-1]+1]

    if 'win' in operating_system:
        search_for = string.punctuation[-9]  
        for index,value in enumerate(cur_dir):
            if value == search_for:
                slash_places.append(index)
        cur_dir  = __file__[:slash_places[-1]+1]
    return cur_dir
cur_loc = find_cur_dir()

original_file = f'{cur_loc}blank mileage log.xlsx'
new_file = f'{cur_loc}Gabes_bot.xlsx'

try:
    wb = openpyxl.load_workbook(new_file, read_only=False) #tries to load the new file first before iterating over the old one
except:
    wb = openpyxl.load_workbook(original_file, read_only=False)

sheet = wb["Sheet1"] #only works on the named here

rows = range(17, 40) #range of wanted rows, 40 is the limit so that row 39 can be in the loop as well
columns = ["A", "B", "F"] #wanted columns

def enter_data():
    for row in rows:
        for column in columns:
            address = column + str(row) #combining row and column to give me coordinates
            cell_value = sheet[address].value #spits our cell value from the sheet specified
            if cell_value == None and column == 'A':
                replace = datetime.datetime.now().strftime('%m/%d/%Y') #add the date to the empty cell
                sheet[address].value = str(replace) #replaces value in found cell
            elif cell_value == None and column == 'B':
                to_site = input('Please enter the site your coming from: ')
                from_site =input('Please enter the site your going to: ')
                data = f'{to_site} to {from_site}' #concatenates variables to create 1 main line
                sheet[address].value = str(data) #adds line to file
                wb.save(new_file) #saves file
            else:
                print('[+] DATA ENTERED SUCCESSFULLY')
                pass

#scan gabes bot for if its full
#if its full, notify the user 
#AFTER NOTIFYING, MAKE A NEW FILE, BUT INCREASE THE NUMBER AT THE END BY 1
