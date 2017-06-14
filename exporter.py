from tkinter import *
import win32com.client
from bs4 import BeautifulSoup as bs
from openpyxl import Workbook

def go_command():
    wb = Workbook()
    ws = wb.active
    olapp = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
    inbox = olapp.GetDefaultFolder(6)
    subfolder = inbox.Folders(folder_text.get())
    messages = subfolder.Items

    row_list = []
    for last_msg in messages:    
        body_content = last_msg.HTMLbody
        soup = bs(body_content, 'html.parser')
        tbl = soup.table

        for row in tbl.find_all('tr'):
            field_list = []
            for item in row.find_all('td'):
                try:
                    field = item.find('span').text
                except AttributeError:
                    field = ""
                if field != '\xa0':
                    field_list.append(field)
                else:
                    field_list.append("")
            if field_list:
                row_list.append(field_list)

        for index, item in enumerate(row_list):
            for subindex, subitem in enumerate(item):
                ws.cell(row=index+1, column=subindex+1).value = subitem

    wb.save(file_text.get() + ".xlsx")



window = Tk()

window.wm_title("HTML table export")

lab1 = Label(window, text="Inbox subfolder")
lab1.grid(row=0, column=0)

folder_text = StringVar()
ent1 = Entry(window, textvariable=folder_text)
ent1.grid(row=0, column=1)

lab1 = Label(window, text="Output file name:")
lab1.grid(row=1, column=0)

file_text = StringVar()
ent1 = Entry(window, textvariable=file_text)
ent1.grid(row=1, column=1)



b1 = Button(window, text="Go!", width=12, command=go_command)
b1.grid(row=2, column=0)


if __name__ == "__main__":
    window.mainloop()
