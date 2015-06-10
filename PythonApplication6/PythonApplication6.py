import wmi
import win32com

c=wmi.WMI ()
for s in c.Win32_Service ():
    if s.State == 'Stopped':
        print s.Caption, s.State
