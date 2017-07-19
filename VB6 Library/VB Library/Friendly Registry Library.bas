Attribute VB_Name = "modFriendlyRegistryLib"
'Robin G Brown, 30th April 1997
'-----------------------------------------------------------------------------
'   This module contains code for _
    handling registry functions
'   This module was originally written by Mark R Matthews, _
    all comments with MRM are original comments, _
    code modified and commented by RGB
'   Find 'Spoiler information' to read how to use this module
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modFriendlyRegistryLib"
Private intCounter As Integer
Private intReturn As Integer
'Constants for EncryptDecryptValue
Global Const conEncryptValue = 0
Global Const conDecryptValue = 1
'-----------------------------Our Gen Decs--------------------------------------------------------
Dim lngHiveKey As Long
'---------------Declares, Types, Constants from WINAPI32.TXT------------------------------------
'DECLARES
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
'Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
'modified:
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As String, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, lpcbSecurityDescriptor As Long) As Long
Private Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Private Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
' Note that if you declare the lpData parameter as String, you must pass it By Value.
'modified:
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long        ' Note that if you declare the lpData parameter as String, you must pass it By Value.
'byval lpData As string
Private Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
'Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
' Note that if you declare the lpData parameter as String, you must pass it By Value.
'modofied
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long        ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'-------------------
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Private Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Boolean
End Type
Private Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As ACL
        Dacl As ACL
End Type
'CONSTANTS
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL = &HFFFF
Public Const REG_NONE = 0                       ' No value PrivateType
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Public Const REG_BINARY = 3                     ' Free form binary
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Public Const REG_LINK = 6                       ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Public Const REG_RESOURCE_LIST = 8              ' Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Public Const REG_CREATED_NEW_KEY = &H1                      ' New Registry Key created
Public Const REG_OPENED_EXISTING_KEY = &H2                      ' Existing Key opened
Public Const REG_WHOLE_HIVE_VOLATILE = &H1                      ' Restore whole hive volatile
Public Const REG_REFRESH_HIVE = &H2                      ' Unwind changes to last flush
Public Const REG_NOTIFY_CHANGE_NAME = &H1                      ' Create or delete (child)
Public Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Public Const REG_NOTIFY_CHANGE_LAST_SET = &H4                      ' Time stamp
Public Const REG_NOTIFY_CHANGE_SECURITY = &H8
Public Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
Public Const REG_OPTION_RESERVED = 0           ' Parameter is reserved
Public Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Public Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Public Const REG_OPTION_CREATE_LINK = 2        ' Created key is a symbolic link
Public Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
Public Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
'Public Const KEY_EXECUTE = (KEY_READ)
Public Const SYNCHRONIZE = &H100000
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
'-----------------------------------------------------------------------------
'The following functions use the API to allow better use of the registry - RGB/30/4/97
'-----------------------------------------------------------------------------
'Spoiler information for using these functions...
'
'   As soon as you use APIOpenKey or APICreateKey all further functions
'   will use the key specified in APIOpenKey/APICreateKey
'   Note that there is no concept of a 'Section' anymore!
'   Usde on of the HKEY_... values for lngHiveKey (see above)

Public Function SetHiveKey(ByVal lngHiveKeyToUse As Long) As Boolean
'Created by Robin G Brown, 8/7/97
'Set the root key for the other registry functions
'Must be one of the following values:
    'HKEY_CLASSES_ROOT = &H80000000
    'HKEY_CURRENT_USER = &H80000001
    'HKEY_LOCAL_MACHINE = &H80000002
    'HKEY_USERS = &H80000003
    'HKEY_PERFORMANCE_DATA = &H80000004
    'HKEY_CURRENT_CONFIG = &H80000005
    'HKEY_DYN_DATA = &H80000006
'Function Declares
Const conFunction = "SetHiveKey"
    'Error Trap
    On Error GoTo SetHiveKey_ErrorHandler
    lngHiveKey = lngHiveKeyToUse
    SetHiveKey = True
Exit Function
'Error Handler
SetHiveKey_ErrorHandler:
    Select Case Err.Number
    Case Else
        'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    SetHiveKey = False
    Exit Function
End Function

Public Function CreateRegistrySetting( _
    ByVal strKeyName As String, _
    ByVal strValueName As String, _
    ByVal strValue As String) As Boolean
'Created by Robin G Brown, 8/7/97
'   To Create a new Registry Setting:
'       APICreateKey
'       APISetKey
'       APICloseKey
'Function Declares
Const conFunction = "CreateRegistrySetting"
Dim lngWasThereAnError As Long
Dim lngHandleToUse As Long
Dim lngLength As Long
Dim lngAPIResult As Long
    'Error Trap
    On Error GoTo CreateRegistrySetting_ErrorHandler
    lngWasThereAnError = RegCreateKeyEx(lngHiveKey, strKeyName, 0, vbNullString, 0, KEY_ALL_ACCESS, vbNullString, lngHandleToUse, lngAPIResult)
    lngLength = Len(strValue)
    lngWasThereAnError = RegSetValueEx(lngHandleToUse, strValueName, 0, REG_SZ, strValue, lngLength)
    lngWasThereAnError = RegCloseKey(lngHandleToUse)
    CreateRegistrySetting = True
Exit Function
'Error Handler
CreateRegistrySetting_ErrorHandler:
    Select Case Err.Number
    Case Else
        'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    CreateRegistrySetting = False
    Exit Function
End Function

Public Function SaveRegistrySetting( _
    ByVal strKeyName As String, _
    ByVal strValueName As String, _
    ByVal strValue As String) As Boolean
'Created by Robin G Brown, 8/7/97
'   To Save a new or current Registry Setting:
'       If APIOpenKey = False then APICreateKey
'       APISetKey
'       APICloseKey
'Function Declares
Const conFunction = "SaveRegistrySetting"
Dim lngWasThereAnError As Long
Dim lngHandleToUse As Long
Dim lngLength As Long
Dim lngAPIResult As Long
    'Error Trap
    On Error GoTo SaveRegistrySetting_Errorhandler
    lngWasThereAnError = RegOpenKeyEx(lngHiveKey, strKeyName, 0, KEY_ALL_ACCESS, lngHandleToUse)
    If lngWasThereAnError <> 0 Then
        lngWasThereAnError = RegCreateKeyEx(lngHiveKey, strKeyName, 0, vbNullString, 0, KEY_ALL_ACCESS, vbNullString, lngHandleToUse, lngAPIResult)
    End If
    lngLength = Len(strValue)
    lngWasThereAnError = RegSetValueEx(lngHandleToUse, strValueName, 0, REG_SZ, strValue, lngLength)
    lngWasThereAnError = RegCloseKey(lngHandleToUse)
    SaveRegistrySetting = True
Exit Function
'Error Handler
SaveRegistrySetting_Errorhandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    SaveRegistrySetting = False
    Exit Function
End Function

Public Function RetrieveRegistrySetting( _
    ByVal strKeyName As String, _
    ByVal strValueName As String, _
    ByRef strValue As String, _
    ByVal strDefault As String) As Boolean
'Created by Robin G Brown, 8/7/97
'       If APIOpenKey = False then APICreateKey
'       APIOpenKey
'       APIQueryKey
'       APICloseKey
'Function Declares
Const conFunction = "RetrieveRegistrySetting"
Dim lngWasThereAnError As Long
Dim lngHandleToUse As Long
Dim lngLength As Long
Dim lngAPIResult As Long
Dim lngMyType As Long
    'Error Trap
    On Error GoTo RetrieveRegistrySetting_ErrorHandler
    lngWasThereAnError = RegOpenKeyEx(lngHiveKey, strKeyName, 0, KEY_ALL_ACCESS, lngHandleToUse)
    If lngWasThereAnError <> 0 Then
        lngWasThereAnError = RegCreateKeyEx(lngHiveKey, strKeyName, 0, vbNullString, 0, KEY_ALL_ACCESS, vbNullString, lngHandleToUse, lngAPIResult)
        lngLength = Len(strDefault)
        If lngLength > 0 Then
            lngWasThereAnError = RegSetValueEx(lngHandleToUse, strValueName, 0, REG_SZ, strDefault, lngLength)
        End If
    End If
    'set to default
    lngMyType = REG_SZ
    lngWasThereAnError = RegQueryValueEx(lngHandleToUse, strValueName, 0, lngMyType, vbNullString, lngLength)
    'next line fills the string to the correct size, Including null terminator character
    strValue = Space$(lngLength)
    'get the data
    lngWasThereAnError = RegQueryValueEx(lngHandleToUse, strValueName, 0, lngMyType, strValue, lngLength)
    lngWasThereAnError = RegCloseKey(lngHandleToUse)
    RetrieveRegistrySetting = True
Exit Function
'Error Handler
RetrieveRegistrySetting_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    RetrieveRegistrySetting = False
    Exit Function
End Function

Public Function DeleteRegistrySetting( _
    ByVal strKeyName As String, _
    ByVal strValueName As String) As Boolean
'Created by Robin G Brown, 8/7/97
'   To Delete a Registry Setting:
'       APIOpenKey
'       APIDeleteKey
'       APICloseKey?
'Function Declares
Const conFunction = "DeleteRegistrySetting"
Dim lngWasThereAnError As Long
Dim lngHandleToUse As Long
Dim lngLength As Long
Dim lngAPIResult As Long
    'Error Trap
    On Error GoTo DeleteRegistrySetting_Errorhandler
    lngWasThereAnError = RegOpenKeyEx(lngHiveKey, strKeyName, 0, KEY_ALL_ACCESS, lngHandleToUse)
    If lngWasThereAnError <> 0 Then
        lngWasThereAnError = RegDeleteValue(lngHandleToUse, strValueName)
        lngWasThereAnError = RegCloseKey(lngHandleToUse)
    End If
    DeleteRegistrySetting = True
Exit Function
'Error Handler
DeleteRegistrySetting_Errorhandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    DeleteRegistrySetting = False
    Exit Function
End Function

