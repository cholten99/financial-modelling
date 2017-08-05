# Imports
import sys
from pprint import pprint
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# All the functions to handle the columns

# Years
def years_func(model_worksheet, model_sheet_row, model_sheet_column):
  global global_vars
  model_worksheet.update_cell(model_sheet_row, model_sheet_column, model_sheet_row + 2015)

# Dave's wage income
def dave_wage_func(model_worksheet, model_sheet_row, model_sheet_column):
  pass

# Annie's wage income
def annie_wage_func(model_worksheet, model_sheet_row, model_sheet_column):
  pass

# Main

# Authorisation
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('financial-modelling-credentials.json', scope)
handle = gspread.authorize(credentials)

# Get spreadsheet 
spreadsheet_url = "https://docs.google.com/spreadsheets/d/1rYsk3jbgctCa2boMILreW3f1a08fR6iHG4IuSc2cdCg/edit#gid=0"
spreadsheet = handle.open_by_url(spreadsheet_url)

# Initialise
global_vars = {}
data_worksheet = spreadsheet.worksheet("Data_for_models")

# Create the list of columns
column_list = []
column_row = 2
while True:
  variable_name = data_worksheet.cell(column_row, 1).value
  if variable_name != "": 
    if data_worksheet.cell(column_row, 2).value == "yes":
      column_list.append(variable_name)
  else:
    break
  column_row += 1

# Set up the global data dictionary and the model worksheets
data_sheet_column = 3
worksheet_list = spreadsheet.worksheets()
while True:
  model_name = data_worksheet.cell(1, data_sheet_column).value
  if model_name != "":
    print "Setting up worksheet for model: " + model_name
    already_exists =  (model_name in (node.title for node in worksheet_list))
    if already_exists:
      spreadsheet.worksheet(model_name).clear()
    else:
      spreadsheet.add_worksheet(model_name, 70, 26)
    model_worksheet = spreadsheet.worksheet(model_name)

    # Create the dict of variables for this model in the global dict and add column headings 
    model_dict = {}
    column_heading_column = 1
    data_sheet_row = 2
    while True:
      variable_name = data_worksheet.cell(data_sheet_row, 1).value
      if variable_name != "":
        model_dict[variable_name] = data_worksheet.cell(data_sheet_row, data_sheet_column).value
        if data_worksheet.cell(data_sheet_row, 2).value == "yes":
          model_worksheet.update_cell(1, column_heading_column, variable_name)
          column_heading_column += 1 
        data_sheet_row += 1
      else:
        global_vars[model_name] = model_dict
        break;

    # Process the model
    print "Processing the model: " + model_name
    model_number_years = global_vars[model_name]["years"]
    for model_sheet_row in range(2, int(model_number_years) + 1):
      model_sheet_column = 1
      for column_name in column_list:
        globals()[column_name + "_func"](model_worksheet, model_sheet_row, model_sheet_column)
        model_sheet_column += 1

    data_sheet_column += 1
  else:
    break 

# It's all over
print "End of modelling"
