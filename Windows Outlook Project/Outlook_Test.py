import datetime as dt
import pandas as pd
import win32com.client

wshShell = win32com.client.Dispatch("WScript.Shell")
outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
calendar = namespace.getDefaultFolder(9) # calendar is now selecting Calendar within Outlook

#for item in calendar.Items:
	#print(item.Subject + '\t' + 'k')

#print(dir(calendar.Items[0]))
#print(calendar.Items[0].Subject)

#print(dir(dt.datetime))
#print(dt.datetime(2022,7,1,13,26,0))