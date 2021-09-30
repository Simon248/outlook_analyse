import win32com.client as win32
import datetime
import pandas as pd

outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(
    6
)  # "6" refers to the index of a folder - in this case,
# the inbox. You can change that number to reference
# any other folder

messages = inbox.Items
sender = []
dates = []
weeks = []
years = []
cats = []
mails = []
for msg in messages:
    try:
        mail = msg.SenderEmailAddress
    except:
        mail = "unknow"
    mails.append(mail)

    try:
        SenderName = msg.SenderName
    except:
        SenderName = "unknow"
    sender.append(SenderName)

    try:
        date = msg.CreationTime
        week = date.isocalendar()[1]
        year = date.isocalendar()[0]
    except:
        date = "unknow"
        week = "unknow"
        year = "unknow"

    try:
        cat = msg.Sensitivity
    except:
        cat = "unknow"

    cats.append(cat)
    dates.append(date)
    weeks.append(week)
    years.append(year)

df = pd.DataFrame(
    {
        "sender": sender,
        "mail": mails,
        "cat": cats,
        "date": dates,
        "week": weeks,
        "year": years,
    }
)

df["date"] = df["date"].dt.tz_convert(None)

from openpyxl import load_workbook
import os

cwd = os.getcwd()
path = cwd + "\\mail.xlsx"


excelBook = load_workbook(path)

with pd.ExcelWriter(path) as writer:
    writer.book = excelBook
    writer.sheets = dict((ws.title, ws) for ws in excelBook.worksheets)
    df.to_excel(writer, "extract", index=False)
    writer.save()


# #     sender=msg.SenderEmailAdress
# #     date=msg.CreationTime
# #     body_content = msg.Body
# #     subject = msg.Subject
# #     categories = msg.Categories
# #     print(body_content)
# #     print(subject)
# #     print(categories)
# message = messages.GetLast()

# # body_content = message.Body
# subject = message.Subject
# date = message.CreationTime
# sender = message.SenderEmailAddress
# # S = message.SendUsingAccount
# SenderName = message.SenderName

# # print(S)
# print(SenderName)
# print(sender)
# print(date)
# print(subject)
# categories = message.Categories
# print(body_content)
# print(subject)
# print(categories)
# for e in email_items:
#     print(email_items.subject)
