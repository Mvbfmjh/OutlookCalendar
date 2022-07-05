import datetime as dt
import pandas as pd
import win32com.client

wshShell = win32com.client.Dispatch("WScript.Shell")
outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')


openWindow = outlook.ActiveExplorer()
if not openWindow:
	wshShell.Run(outlook)
objPane = openWindow.NavigationPane
# Select Calendar Module
calModule = objPane.Modules.GetNavigationModule(1)
# Select Group within Calendar
projectGroup = calModule.NavigationGroups(4)

print(projectGroup.Name)
calList = list()
for folder in projectGroup.NavigationFolders:
	print(folder)
	try:
		print(folder.Folder)
		calList.append(folder.Folder.Items)
	except:
		print("No Access")
		calList.append(None)



getToday = dt.datetime.today()
end = getToday + dt.timedelta(days=7)
restriction = "[Start] >= '" + getToday.strftime('%Y/%m/%d') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
for item in calList:
	if item:
		item.includeRecurrences = True
		item.Sort('[Start]')
		item = item.Restrict(restriction)
		for i in item:
			print(str(i.Subject) + '\t' + i.Start.strftime('%Y/%m/%d'))

