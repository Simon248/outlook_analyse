# outlook_analyse
deposit contains two folder to extract and analyse data from your outlook application
 - mail_extract.py
 - outlook_calendar.py 
 
 IN ORDER TO WORK YOU SHOULD HAVE OUTLLOK RUNNING ON YOUR PC DURING EXECUTION.

## mail_extract:
this file allows you to extract all your mails in an Excel document with columns: (sender, content, categorie, date, week, year)
it's up to you to filter and analyse it in Excel.

Other columns could be added, refer to "win32com.client"

## outlook_calendar:
this file allow you to extract all your calendar datas to Excel files. Columns are : (subject, start date, end date, body, organizer, duration, week, year)
You have to fill the dates datas (from - to) at the begining of the code before execution:
```
begin = dt.datetime(2020, 1, 1)
end = dt.date.today()
```
In addition, you can fill the variable "temps de travail" (daily working hours) so the program calculate a ratio of time spend in meeting.
```
temps_de_travail_journalier = 7.75
```

You can filer meeting by fillin filtre.json which will provide substring to filter mettings names.

As for mails, it's up to you to filter and analyse it in Excel.

Other columns could be added, refer to "win32com.client"

