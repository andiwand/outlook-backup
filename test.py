import winreg
import win32api
import win32security

# save registry "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook" (http://www.robvanderwoude.com/regedit.php)
# save "C:\Users\%username%\AppData\Local\Microsoft\Outlook", ignore pst, ost, nst
# save "C:\Users\%username%\AppData\Roaming\Microsoft\Outlook"
# save "C:\Users\%username%\AppData\Roaming\Microsoft\Signatures"

priv_flags = win32security.TOKEN_ADJUST_PRIVILEGES | win32security.TOKEN_QUERY
hToken = win32security.OpenProcessToken(win32api.GetCurrentProcess (), priv_flags)
privilege_id = win32security.LookupPrivilegeValue(None, "SeBackupPrivilege")
win32security.AdjustTokenPrivileges(hToken, 0, [(privilege_id, win32security.SE_PRIVILEGE_ENABLED)])

key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Office\16.0\Outlook", 0, winreg.KEY_ALL_ACCESS)
print(key)
print(winreg.QueryInfoKey(key))

for i in range(0, winreg.QueryInfoKey(key)[0]):
    print(winreg.EnumKey(key, i))

winreg.SaveKey(key, r"office.reg")
