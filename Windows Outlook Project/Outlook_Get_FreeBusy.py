import datetime as dt
import pandas as pd
import win32com.client

wshShell = win32com.client.Dispatch("WScript.Shell")
outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')

# Need cooperation from another person to test
#otherPerson = namespace.CreateRecipient(-)
otherPerson = namespace.CreateRecipient(-)
#print(otherPerson.Resolve)
otherPerson.Resolve
today = dt.datetime(2022,7,1)
today2 = dt.datetime(2022,7,2)
data = otherPerson.FreeBusy(today,60,True)
data2 = otherPerson.FreeBusy(today2,60,True)
print(data)
print(data2)

# What is this output....
