import win32com.client
 
wshShell = win32com.client.Dispatch("WScript.Shell")
wshShell.Run("notepad.exe")