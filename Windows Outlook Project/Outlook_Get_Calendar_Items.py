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

itemList = [item for item in calendar.Items]	# putting all items from calendar into a list - an array of items, with the size len(calendar.Items)

# Test print to see if calendar events were placed into the itemlist
print(itemList[0].Subject)
