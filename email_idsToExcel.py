import email
import imaplib
from openpyxl import *
import os.path

if not (os.path.exists('email_ids.xlsx')):
    wb = Workbook()
    wb.create_sheet("email_ids.xlsx")
    wb.save('email_ids.xlsx')

wb = load_workbook('email_ids.xlsx')
sheet = wb.active
sheet.column_dimensions['A'].width = 30
sheet.column_dimensions['B'].width = 60
sheet.column_dimensions['C'].width = 60
sheet.cell(row=1, column=1).value = 'Date'
sheet.cell(row=1, column=2).value = 'To'
sheet.cell(row=1, column=3).value = 'From'

username = "example@gmail.com"
password = "password"

mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login(username, password)
mail.select('inbox')

result, data = mail.uid('search', None, 'ALL')
inbox_items = data[0].split()


for i in range(len(inbox_items)):
    result2, email_data = mail.uid('fetch', inbox_items[i], '(RFC822)')
    raw_email = email_data[0][1].decode("utf-8")
    msg = email.message_from_string(raw_email)
    current_row = sheet.max_row
    sheet.cell(row=current_row + 1, column=1).value = msg['Date']
    sheet.cell(row=current_row + 1, column=2).value = msg['To']
    sheet.cell(row=current_row + 1, column=3).value = msg['From']

wb.save('email_ids.xlsx')