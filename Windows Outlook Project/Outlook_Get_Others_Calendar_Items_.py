import datetime as dt
import pandas as pd
import win32com.client

wshShell = win32com.client.Dispatch("WScript.Shell")
outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')

# Need cooperation from another person to test
otherPerson = namespace.CreateRecipient(-)
otherPerson2 = namespace.CreateRecipient(-)
#print(otherPerson.Resolve)
otherPerson.Resolve
otherPerson2.Resolve
calendarItems1 = namespace.GetSharedDefaultFolder(otherPerson, 9).Items
calendarItems2 = namespace.GetSharedDefaultFolder(otherPerson2, 9).Items
calendarItems1.includeRecurrences = True
calendarItems2.includeRecurrences = True

