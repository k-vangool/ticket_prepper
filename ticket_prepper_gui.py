import tkinter as tk
from tkinter.filedialog import askopenfilename


def open_excel():
    filepathExcel = askopenfilename()
    if not filepathExcel:
        return

    # print(f'\nThe path to the file is: \n{filepath}')


# Creating GUI to ask and open a file
def rungui():
    root = tk.Tk()
    root.title('ECX Ticket Prepper')
    root.geometry('550x250')

    explanation = ('Hi collega,\n\n'
                   '' 
                   'Dit korte script zal ervoor zorgen dat jij niet meer alle tickets handmatig hoeft te openen. \n' 
                   'Selecteer na het klikken op de knop de excel met tickets die jij geopend wilt hebben. \n\n'
                   '    1) Klik op de knop.\n'
                   '    2) Selecteer de ticket excel die je wilt openen.\n'
                   '    3) Alles wordt geopend.\n'
                   '\nMet vriendelijke groet,\n'
                   'Kars van Gool\n'
                   'AM ECX AMS01')

    button_open = tk.Button(text='Open ticket Excel', command=open_excel)
    label = tk.Label(root, justify=tk.LEFT, text=explanation)

    label.pack(anchor='nw', padx=20,pady=10)
    button_open.pack(side='bottom', pady=10)

    root.mainloop()


rungui()
