Attribute VB_Name = "AddIn"
Option Explicit
Declare Function WritePrivateProfileString& Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

'====================================================================
'this sub should be executed from the Immediate window
'in order to get this app added to the VBADDIN.INI file
'you must change the name in the 2nd argument to reflecty 
'the correct name of your project
'====================================================================
Sub AddToINI()
    Dim ErrCode As Long
    ErrCode = WritePrivateProfileString("Add-Ins32", "MyAddIn.Connect", "0", "vbaddin.ini")
End Sub

