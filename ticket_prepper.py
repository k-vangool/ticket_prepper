"""
This script has been written to open all the relevant tickets for the SOC of ECX AMS01.

It opens a gui which lets the user open the downloaded tickets from edgeos.
"""

import datetime as dt
import os
import time
import webbrowser as wb
import pandas as pd
import tkinter as tk
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter import messagebox


def open_excel():
    global filepathExcel
    filepathExcel = askopenfilename()

    if not filepathExcel:
        return None

    name, extension = os.path.splitext(filepathExcel)

    if extension != '.xlsx':
        print(extension)
        messagebox.showerror('Wrong file type', 'This script only supports .xlsx files. Select the correct file.')

    if filepathExcel and extension == '.xlsx':
        open_tickets()
        messagebox.showerror('Run Script', 'Script zal nu gesloten worden en tickets geopend.')
        root.withdraw()


# Creating GUI to ask and open a file
def rungui():
    global root
    root = tk.Tk()
    root.title('ECX Ticket Prepper')
    root.geometry('550x320')

    explanation = ('Hi collega,\n\n'
                   '' 
                   'Dit korte script zal ervoor zorgen dat jij niet meer alle tickets handmatig hoeft te openen. \n' 
                   'Selecteer na het klikken op de knop de excel met tickets die jij geopend wilt hebben. \n\n'
                   'Werking van het programma:\n'
                   '    1) Zorg dat je ingelogd bent op EdgeOS in CHROME (!!!) (edgeos.edgeconnex.com).\n'
                   '    2) Klik op de knop "Open ticket Excel".\n'
                   '    3) Selecteer de ticket excel die je wilt openen.\n'
                   '    4) Selecteer de plek waar je de relevante tickets wilt opslaan.\n'
                   '    5) Alles wordt geopend. Relevante tickets worden opgeslagen en het programma sluit.\n'
                   '\nMet vriendelijke groet,\n'
                   'Kars van Gool\n'
                   'AM ECX AMS01\n\n'
                   'PS. Dit script zal enkel de relevante tickets voor het SOC van AMS01 openen.')

    button_open = tk.Button(text='Open ticket Excel', command=open_excel)
    label = tk.Label(root, justify=tk.LEFT, text=explanation)

    label.pack(anchor='nw', padx=20,pady=10)
    button_open.pack(side='bottom', pady=10)

    root.mainloop()


def open_tickets():
    global relevant_tickets, tickets
    tickets = pd.read_excel(filepathExcel, engine='openpyxl')
    chrome_path = 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'

    lst_of_cats = [
        'Package Delivery',
        'Package Delivery No Hands',
        'Authorize a Visit',
        'Escort & Tour'
    ]

    # Only select the relevant tickets for the SOC of AMS01
    relevant_tickets = tickets[(tickets.Category.isin(lst_of_cats)) & (tickets.Site == 'AMS01: Amsterdam 1')].copy()
    relevant_tickets.reset_index(inplace=True)
    relevant_tickets.dropna(1, inplace=True)
    relevant_tickets.drop(columns=['index'], inplace=True)

    # Saving the new tickets
    lst_of_relevantCols = [
        'Ticket #',
        'Category',
        'Name',
        'Phone',
        'Subject',
        'Description',
        'Entered On',
        'Updated On'
    ]
    relevant_tickets_SOC = relevant_tickets[lst_of_relevantCols].copy()
    date = str(dt.datetime.today().strftime('%d-%m-%Y'))
    new_file_name = 'tickets_' + date + '_PROCESSED.xlsx'
    messagebox.showerror('Save Excel', 'Selecteer de locatie waar je de relevante tickets opgeslagen wilt hebben.')
    save_path = os.path.join(askdirectory(), new_file_name)

    # Saving Excel + wrapping text.
    writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
    relevant_tickets_SOC.to_excel(writer,
                                  sheet_name='Relevant tickets SOC',
                                  index=False
                                  )
    workbook = writer.book
    worksheet = writer.sheets['Relevant tickets SOC']
    format = workbook.add_format({'text_wrap': True})
    worksheet.set_column('A:I', None, format)

    # Opening the tickets in a new chrome browser window
    wb.register('chrome', None, wb.BackgroundBrowser(chrome_path))
    chrome = wb.get('chrome')

    os.startfile(chrome_path, 'open')

    url_str_no_ticket = 'https://edgeos.edgeconnex.com/portal/Forms/ShowRequest.aspx?ID_FINDER='

    for i in range(len(relevant_tickets['Ticket #'])):
        new_url = url_str_no_ticket + str(relevant_tickets['Ticket #'][i])

        chrome.open_new_tab(new_url)
        time.sleep(0.095)

rungui()

