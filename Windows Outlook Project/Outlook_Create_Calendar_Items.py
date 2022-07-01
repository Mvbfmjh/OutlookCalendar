import datetime as dt
import pandas as pd
import win32com.client

wshShell = win32com.client.Dispatch("WScript.Shell")
outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
calendar = namespace.getDefaultFolder(9)		# calendar is now selecting Calendar within Outlook
calendarItems = calendar.Items
calendarItems.includeRecurrences = True

#for k in calendar.Items:
	#print(k.Subject + '\t' + 'k')
# Add 21 hours to get Tokyo Time
# i.e. Jul 2 10AM -> Jul 3 1AM Tokyo Time
begin = dt.datetime(2022,7,2,10)
end = dt.datetime(2022,7,2,18)

# Creates new Calendar Item
# ISSUE: Time set to UTC+0
newItem = calendarItems.Add()
newItem.Subject = 'test'
newItem.Start = begin
newItem.End = end
newItem.Save()
