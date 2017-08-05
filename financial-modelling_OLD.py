# Imports
import sys
from pprint import pprint
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Utility functions
def update_worksheet(data):
  global_vars["worksheet"].update_cell(global_vars["row"], global_vars["column"], data)

def update_data_worksheet(data):
  global_vars["data_worksheet"].update_cell(global_vars["row"], global_vars["column"], data)

# All the functions to handle the columns

# Years
def Year_func():
  global global_vars
  if global_vars["state"] == "initialising":
    global_vars["num_years"] = int(raw_input("How many years to run model: "))
  elif global_vars["state"] == "processing":
    update_worksheet(global_vars["row"] + 2015)
  elif global_vars["state"] == "data_sheet":
    update_data_worksheet("Number of years to model : " + str(global_vars["num_years"]))

# Main

# Initialise a bunch of stuff 
global_vars = {}
columns = ["Year"]
global_vars["row"] = 2
global_vars["column"] = 1

# Authorisation
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('financial-modelling-credentials.json', scope)
handle = gspread.authorize(credentials)

# Get spreadsheet 
spreadsheet_url = "https://docs.google.com/spreadsheets/d/1rYsk3jbgctCa2boMILreW3f1a08fR6iHG4IuSc2cdCg/edit#gid=0"
spreadsheet = handle.open_by_url(spreadsheet_url)

# Ask for worksheet name
worksheet_name = raw_input("What is the name of the worksheet: ")

# If worksheet doesn't exist create it and the associated data one, if it does ask whether to overwrite or exit
worksheet_list = spreadsheet.worksheets()
already_exists =  (worksheet_name in (node.title for node in worksheet_list))
if not already_exists:
  print "Creating new worksheet called " + worksheet_name
  spreadsheet.add_worksheet(worksheet_name, 70, 26)
  print "Creating new worksheet called " + worksheet_name + "_data"
  spreadsheet.add_worksheet(worksheet_name + "_data", 70 , 26)
else:
  overwrite_existing_worksheet = raw_input("A worksheet called " + worksheet_name + " already exists. To overwrite type 'yes': ")
  if overwrite_existing_worksheet == "yes":
    print "Clearing worksheets " + worksheet_name + " and " + worksheet_name + "_data"
    spreadsheet.worksheet(worksheet_name).clear() 
    spreadsheet.worksheet(worksheet_name + "_data").clear()
  else:
    print "Not overwriting worksheet " + worksheet_name + " - exiting"
    sys.exit(0)
global_vars["worksheet"] = spreadsheet.worksheet(worksheet_name)
global_vars["data_worksheet"] = spreadsheet.worksheet(worksheet_name + "_data")

# Ask initial questions
global_vars["state"] = "initialising"
for column in columns:
  globals()[column + "_func"]()

# Fill out the data worksheet
global_vars["state"] = "data_sheet"
global_vars["row"] += 1
for column in columns:
  globals()[column + "_func"]()

# Print column headers
col_count = 1
for column in columns:
   global_vars["worksheet"].update_cell(1, col_count, column)
   col_count += 1

# Create all the rows
global_vars["state"] = "processing"
for i in range(1, global_vars["num_years"] + 1):
  for column in columns:
    globals()[column + "_func"]()
    global_vars["column"] += 1
  global_vars["column"] = 1
  global_vars["row"] += 1

# It's all over
print "End of modelling"
