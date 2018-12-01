import gspread
from oauth2client.service_account import ServiceAccountCredentials
import random

# card setup
param_update = 'A'
param_card_title = 'B'
param_midiacode_id = 'A'

# configura as credenciais.
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive',
         'https://www.googleapis.com/auth/spreadsheets']
credentials_filename = 'midiacode-app-ea21e59d0156.json'
credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_filename, scope)
gc = gspread.authorize(credentials)

# abre o arquivo pela URL
sheet_url = 'https://docs.google.com/spreadsheets/d/18mQCXS6lrlfa-IaICUHRirUMIOxdOMxMt9vYDbZLvlA/edit#gid=0'
sh = gc.open_by_url(sheet_url)

# seleciona ou cria a aba de controle do midiacode
try:
    wk_midiacode = sh.worksheet("midiacode")
except gspread.exceptions.WorksheetNotFound:
    wk_midiacode = sh.add_worksheet(title="midiacode", rows="100", cols="20")
    header_line = 1
    label_midiacode_id = param_midiacode_id + str(header_line)
    wk_midiacode.update_acell(label_midiacode_id, "document_id")

# seleciona a aba de dados (db)
wk = sh.worksheet("db")
list_of_lists = wk.get_all_values()
rows_size = len(list_of_lists)


for row_number in range(1, rows_size):
    line_number = str(row_number + 1)
    print("Reading line {}".format(line_number))
    pos_card_title_title = param_card_title + line_number
    field_title = wk.acell(pos_card_title_title).value
    print("Processing title {}".format(field_title))
    label_midiacode_id_for_row = param_midiacode_id + line_number
    midiacode_id = wk_midiacode.acell(label_midiacode_id_for_row).value
    if midiacode_id == '':
        print("Create a new card")
        midiacode_id = random.randint(1,10000)
        wk_midiacode.update_acell(label_midiacode_id_for_row, str(midiacode_id))
    else:
        label_param_update = param_update + line_number
        field_update = wk.acell(label_param_update).value
        if field_update != '':
            print("Update the card id {}.".format(midiacode_id))
            wk.update_acell(label_param_update, "")

print("Done")






