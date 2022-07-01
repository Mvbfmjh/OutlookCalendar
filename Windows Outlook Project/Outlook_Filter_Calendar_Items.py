import datetime as dt
import pandas as pd
import win32com.client

wshShell = win32com.client.Dispatch("WScript.Shell")
outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
calendar = namespace.getDefaultFolder(9)		# calendar is now selecting Calendar within Outlook
calendarItems = calendar.Items
calendarItems.includeRecurrences = True

begin = dt.datetime(2022,6,1)
end = dt.datetime(2022,7,1)

calendarItems.Sort('[Start]')

# Restriction String acts as a filter for the Items within calendarItems
# Need to look into the Restriction String, and its syntax
restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
calendarItems = calendarItems.Restrict(restriction)

for k in calendarItems:
	print(k.Subject + '\t' + 'k.Start')