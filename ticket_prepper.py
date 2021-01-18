import webbrowser as wb
import pandas as pd
import os
import time
import datetime as dt

# Importing tickets
file_name = 'input_tickets.xlsx'
path = os.getcwd()

tickets = pd.read_excel(os.path.join(path, file_name), engine='openpyxl')
chrome_path = 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'

lst_of_cats = [
    'Package Delivery',
    'Package Delivery No Hands',
    'Authorize a Visit',
    'Escort & Tour'
]

relevant_tickets = tickets[tickets.Category.isin(lst_of_cats)].copy()
relevant_tickets.reset_index(inplace=True)
relevant_tickets.dropna(1, inplace=True)
relevant_tickets.drop(columns=['index'], inplace=True)

wb.register('chrome', None, wb.BackgroundBrowser(chrome_path))
chrome = wb.get('chrome')

os.startfile(chrome_path, 'open')

url_str_no_ticket = 'https://edgeos.edgeconnex.com/portal/Forms/ShowRequest.aspx?ID_FINDER='

for i in range(len(relevant_tickets['Ticket #'])):
    new_url = url_str_no_ticket + str(relevant_tickets['Ticket #'][i])

    chrome.open_new_tab(new_url)
    time.sleep(0.095)


# Saving the new tickets
relevant_tickets.set_index(['Ticket #'], inplace=True)
date = str(dt.datetime.today().strftime('%d-%m-%Y'))
new_file_name = 'tickets_' + date + '.xlsx'
relevant_tickets.to_excel(os.path.join(path, new_file_name), engine='openpyxl')