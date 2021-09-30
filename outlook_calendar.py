import datetime as dt
import pandas as pd
import win32com.client
import os
import dateutil.parser

import win32com.client as win32


###PARAMETRES###

temps_de_travail_journalier = 7.75

begin = dt.datetime(2020, 1, 1)
end = dt.date.today()


def get_calendar(begin, end):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort("[Start]")
    restriction = (
        "[Start] >= '"
        + begin.strftime("%m/%d/%Y")
        + "' AND [END] <= '"
        + end.strftime("%m/%d/%Y")
        + "'"
    )
    calendar = calendar.Restrict(restriction)
    return calendar


def get_appointments(calendar):

    appointments = [app for app in calendar]

    tbl_years = []
    tbl_week = []
    tbl_subject = []
    tbl_organizer = []
    tbl_start = []
    tbl_duration = []
    tbl_end = []
    tbl_body = []

    # OUTLOOK_FORMAT = '%m/%d/%Y %H:%M'
    i = 0
    for appointmentItem in appointments:
        # print(i)
        i = i + 1
        try:
            subject = appointmentItem.subject
        except:
            subject = "unknow"

        try:
            organizer = appointmentItem.organizer
        except:
            organizer = "unknow"

        try:
            start = appointmentItem.start
            week = start.isocalendar()[1]
            years = start.isocalendar()[0]

        except Exception as e:
            start = "unknow"
            week = "unknow"
            years = "unknow"

        try:
            duration = appointmentItem.Duration / 60
        except:
            duration = "unknow"

        try:
            end = appointmentItem.end
        except:
            end = "unknow"

        try:
            body = appointmentItem.body
        except:
            body = "unknow"

        tbl_subject.append(subject)
        tbl_organizer.append(organizer)
        tbl_start.append(start)
        tbl_duration.append(duration)
        tbl_end.append(end)
        tbl_body.append(body)
        tbl_week.append(week)
        tbl_years.append(years)

    df = pd.DataFrame(
        {
            "subject": tbl_subject,
            "start": tbl_start,
            "end": tbl_end,
            "body": tbl_body,
            "organisateur": tbl_organizer,
            "duree": tbl_duration,
            "semaine": tbl_week,
            "annee": tbl_years,
        }
    )

    df["start"] = df["start"].dt.tz_convert(None)
    df["end"] = df["end"].dt.tz_convert(None)

    return df


def filter(df):
    import json

    with open("C:\\Users\\srobert\\Desktop\\analyse_outlook\\filtre.json") as f:
        data = json.load(f)

    ne_doit_pas_contenir = data.get("ne_doit_pas_contenir")
    df["keep"] = ~df["subject"].str.contains("|".join(ne_doit_pas_contenir), case=False)
    return df


cal = get_calendar(begin, end)
df = get_appointments(cal)
df_filtred = filter(df)


df_filtred["durée en pourcentage de la semaine"] = df_filtred["duree"] / (
    temps_de_travail_journalier * 5
)


cwd = os.getcwd()
path = cwd + "\\result.xlsx"

from openpyxl import load_workbook

excelBook = load_workbook(path)

with pd.ExcelWriter(path) as writer:
    writer.book = excelBook
    writer.sheets = dict((ws.title, ws) for ws in excelBook.worksheets)
    df_filtred.to_excel(writer, "extract", index=False)
    writer.save()
