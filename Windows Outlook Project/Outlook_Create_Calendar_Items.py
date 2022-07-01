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

begin = dt.datetime(2022,7,2)
end = dt.datetime(2022,7,3)

# Creates new Calendar Item
# ISSUE: Time set to UTC+0
newItem = calendarItems.Add(1)
newItem.Subject = 'test'
newItem.Start = begin
newItem.End = end
newItem.Save()
