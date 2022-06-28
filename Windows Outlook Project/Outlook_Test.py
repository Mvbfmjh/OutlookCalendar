import datetime as dt
import pandas as pd
import win32com.client

wshShell = win32com.client.Dispatch("WScript.Shell")
outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
calendar = namespace.getDefaultFolder(9) # calendar is now selecting Calendar within Outlook

for item in calendar.Items:
	print(item.Subject + '\t' + 'k')

#wshShell.Run(outlook)