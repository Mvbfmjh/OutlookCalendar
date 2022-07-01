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

itemList = [item for item in calendarItems]	# putting all items from calendar into a list - an array of items, with the size len(calendar.Items)

# Test print to see if calendar events were placed into the itemlist
#print(itemList[0].Subject)

event_Subject = [item.Subject for item in calendarItems]
event_Start = [item.Start for item in calendarItems]
event_End = [item.End for item in calendarItems]
event_Body = [item.Body for item in calendarItems]

df = pd.DataFrame({'subject': event_Subject,
				   'start': event_Start,
				   'end': event_End,
				   'body': event_Body})

#print(df.to_string())
# When writing into Excel, needs 'total_seconds'
df.to_excel('test.xlsx')
#wshShell.Run(outlook)