# Import the openpyxl module, which is used to interact with Excel files
# With pip installed, install openpyxl by running 'pip install openpyxl' from a prompt
import openpyxl
from openpyxl import Workbook
import sys
import csv

# Create a new workbook to copy CSV file into a new XLSX file
wb = Workbook()
ws = wb.active
if len(sys.argv) < 2:
    sys.exit("You must include a CSV filename as a parameter to run...")
filename = sys.argv[1]
with open(filename, 'r') as f:
    for row in csv.reader(f):
        ws.append(row)
newfile = filename[0:(len(filename)-4)] + '.xlsx'
wb.save(newfile)

# Open the new XLSX file and copy the first worksheet into two new tabs called "Bad Credentials" and "No Connectivty"
wb = openpyxl.load_workbook(newfile)
sheet = wb['Sheet']
creds = wb.copy_worksheet(sheet)
creds.title = "Bad Credentials"
connect = wb.copy_worksheet(sheet)
connect.title = "No Connectivity"
wb.save(newfile)

# For the "Bad Credentials" tab, delete any rows where the Inaccessible tab is not "1", display total number left after loop and save file
rownum = 1
for row in creds:
    if creds['D' + str(rownum)].value != '1':
        creds.delete_rows(rownum, 1)
    else:
        rownum += 1
creds['H1'].value = "Devices with credential issues:"
creds['K1'].value = rownum - 1
wb.save(newfile)

# For the "No Connectivity" tab, delete any rows where the Total IPs tab is not "0", display total number left after loop and save file
rownum = 1
for row in connect:
    if connect['F' + str(rownum)].value != '0':
        connect.delete_rows(rownum, 1)
    else:
        rownum += 1
connect['H1'].value = "Devices with no connectivity:"
connect['K1'].value = rownum - 1
wb.save(newfile)

print('Output file "' + newfile + '" has been created in working directory')