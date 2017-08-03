from pprint import pprint
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds']

credentials = ServiceAccountCredentials.from_json_keyfile_name('financial-modelling-credentials.json', scope)

handle = gspread.authorize(credentials)

worksheet = handle.open_by_url("https://docs.google.com/spreadsheets/d/1ZYqRxvxvhLL0SSSXQX6xOl_Ait-HQ6xd8t8ZZnvgfLY/edit#gid=0").sheet1

pprint(worksheet.get_all_values())




