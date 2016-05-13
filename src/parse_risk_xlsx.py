import sys
from openpyxl import Workbook, load_workbook

first_row = 2
title_column = 3
cwss_column = 6
base_column = 7
attack_column = 8
environmental_column = 9


#main

if len(sys.argv) != 2:
    print "Usage: python %s risk_register.xlsx" % sys.argv[0]
    sys.exit(0)

risk_register = sys.argv[1]

print "Reading from risk_register"

# openpyxl give a 'warnings.warn("Discarded range with reserved name")'
# it doesn't affect the sheet we're interested in, so we can safely
# ignore it for now
#warnings.simplefilter("ignore")
wb = load_workbook(risk_register, read_only=True, data_only=True)
print wb.get_sheet_names()

ws = wb['Engagement Findings']


# we start in column C and count from row 2 down until we get None
for x in xrange(first_row,ws.max_row):
    if (ws.cell(row=x, column=title_column).value != None):
        print ws.cell(row=x, column=title_column).value
        print ws.cell(row=x, column=cwss_column).value


