import webbrowser as wb
import pandas as pd
import os
import time

tickets = pd.read_excel('Tickets_11.01.21.xlsx', engine='openpyxl')
chrome_path = 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'

wb.register('chrome', None, wb.BackgroundBrowser(chrome_path))
chrome = wb.get('chrome')

os.startfile(chrome_path, 'open')

url_str_no_ticket = 'https://edgeos.edgeconnex.com/portal/Forms/ShowRequest.aspx?ID_FINDER='

for i in range(len(tickets)):
    new_url = url_str_no_ticket + str(tickets['Ticket_no'][i])

    chrome.open_new_tab(new_url)
    time.sleep(0.05)