Attribute VB_Name = "modMain"
'Created by Robin G Brown, 22nd April 1997
'-----------------------------------------------------------------------------
'   This module contains code for _
    Main
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modMain"
Private intCounter As Integer
Private intReturn As Integer
Public Const conZTop = 0
Public Const conZBottom = 1
'Layout stuff
Public frmPosition(1 To 6) As Form
Public Const conAutoArrange = 99
Public Const conDefaultLayout = 0
Public Const conCustom1 = 1
Public Const conCustom2 = 2
Public Const conCustom3 = 3
Public Const conCascade = 11
Public Const conTileHorizontal = 12
Public Const conTileVertical = 13
'Display alteration variables
Public Const conDefault = "default"
Public Const conSelected = "check"
Public Const conDeselected = "uncheck"
'Variables required for drag and drop
Public nodDrag As Node
Public booDrag As Boolean
Public tvwCurrent As TreeView
Public Const conDefaultImage = "default"
'Form variables
Public mdiRowSelector As New mdiSelectionChild
Public mdiColSelector As New mdiSelectionChild
Public mdiPageSelector As New mdiSelectionChild
Public strFormDimension As String
Public strRegistrySettings() As String

Sub Main()
'Created by Robin G Brown, 22/4/97
'Sub Declares
Const conSub = "Main"
    'Error Trap
    On Error GoTo Main_ErrorHandler
    'Check for Instances of this program
    'Show splash form
    frmSplash.Show
    'Carry out any slow tasks and initialize
    Load frmMaster
    Call RetrieveApplicationSettings
    Call ChangeTheme(0)
    'Show first form for startup
    frmMaster.Show
    StatusBar ""
    'Wait for a sec or three
    Call Pause
    'Get rid of splash form
    Unload frmSplash
    'Add a new set of files
    Call frmMaster.mnuFileNew_Click
Exit Sub
'Error Handler
Main_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    'Exit the program
    Call ApplicationEnd
    Exit Sub
End Sub

Sub ApplicationEnd()
'Created by Robin G Brown, 22/4/97
'Sub Declares
Const conSub = "ApplicationEnd"
    'Error Trap
    On Error Resume Next
    'CODE_HERE, _
        Close databases, etc _
        Free memory for objects _
        Save registry information _
        and Tidy up
    Call SaveApplicationSettings
    Call KillExcelObject
    'End the application
    End
End Sub

Public Sub RetrieveApplicationSettings()
'Created by Robin G Brown, 30/4/97
'Retrieve the settings for this application from the registry
'Sub Declares
Const conSub = "RetrieveApplicationSettings"
Dim strResult As String
Dim strFullKeyName As String
Dim strValueName As String
Dim strValue As String
    'Error Trap
    On Error GoTo RetrieveApplicationSettings_ErrorHandler
    'If APIOpenKey = False then APICreateKey
    strFullKeyName = "Software\VB and VBA Program Settings\Report Selector"
    intReturn = APIOpenKey(HKEY_CURRENT_USER, strFullKeyName, strResult)
    If intReturn = False Then
        'Key not found - does not exist
        Exit Sub
    End If
    ReDim strRegistrySettings(1 To 3)
    'For AutoArrange
        strValueName = "AutoArrange"
        intReturn = APIQueryKey(strValueName, strValue, strResult)
        'Act on the retrieved information
        strRegistrySettings(1) = strValue
    'For ScratchPad
        strValueName = "ScratchPad"
        intReturn = APIQueryKey(strValueName, strValue, strResult)
        'Act on the retrieved information
        strRegistrySettings(2) = strValue
    'For SourceTree
        strValueName = "SourceTree"
        intReturn = APIQueryKey(strValueName, strValue, strResult)
        'Act on the retrieved information
        strRegistrySettings(3) = strValue
    'All settings completed
    intReturn = APICloseKey(strResult)
Exit Sub
'Error Handler
RetrieveApplicationSettings_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub
Public Sub SaveApplicationSettings()
'Created by Robin G Brown, 30/4/97
'Save the settings for this application into the registry
'Sub Declares
Const conSub = "SaveApplicationSettings"
Dim strResult As String
Dim strFullKeyName As String
Dim strValueName As String
Dim strValue As String
    'Error Trap
    On Error GoTo SaveApplicationSettings_ErrorHandler
    'If APIOpenKey = False then APICreateKey
    strFullKeyName = "Software\VB and VBA Program Settings\Report Selector"
    intReturn = APIOpenKey(HKEY_CURRENT_USER, strFullKeyName, strResult)
    If intReturn = False Then
        intReturn = APICreateKey(HKEY_CURRENT_USER, strFullKeyName, strResult)
    End If
    'AutoArrange
        strValueName = "AutoArrange"
        strValue = strRegistrySettings(1)
        intReturn = APISetKey(strValueName, REG_SZ, strValue, Len(strValue), strResult)
    'ScratchPad
        strValueName = "ScratchPad"
        strValue = strRegistrySettings(2)
        intReturn = APISetKey(strValueName, REG_SZ, strValue, Len(strValue), strResult)
    'Source Tree settings
        strValueName = "SourceTree"
        Call mdiTreeChild.GetTreeSettings
        strValue = strRegistrySettings(3)
        intReturn = APISetKey(strValueName, REG_SZ, strValue, Len(strValue), strResult)
    'All settings completed
    ReDim strRegistrySettings(1 To 1)
    intReturn = APICloseKey(strResult)
Exit Sub
'Error Handler
SaveApplicationSettings_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub
Public Sub StatusBar(ByVal strNewText As String)
'Created by Robin G Brown, 16/4/97
'Sets the statusbar text
'Sub Declares
Const conSub = "StatusBar"
    'Error Trap
    On Error GoTo StatusBar_ErrorHandler
    If strNewText = "" Then
        'Default string
        frmMaster.stbStatus.Panels(3).Text = "Use drag and drop to move trees around"
    Else
        frmMaster.stbStatus.Panels(3).Text = strNewText
    End If
Exit Sub
'Error Handler
StatusBar_ErrorHandler:
    Exit Sub
End Sub
Public Sub ChangeTheme(Optional ByVal varTheme As Variant)
'Created by Robin G Brown, 28/4/97
'Load the imagelist control with the images we want to use
'Sub Declares
Const conSub = "LoadImages"
Dim intTheme As Integer
Dim imgLoad As ListImage
Dim strRootPath As String
    'Error Trap
    On Error GoTo LoadImages_ErrorHandler
    If IsMissing(varTheme) Then
        intTheme = 1
    Else
        intTheme = CInt(varTheme)
    End If
    If intTheme = 0 Then
        'Set up the default theme, without referencing the child forms
        strRootPath = App.Path & "\Theme1\"
        'Load the three dimension images
        frmMaster.imlDimension.ListImages.Clear
        frmMaster.imlDimension.ImageHeight = 16
        frmMaster.imlDimension.ImageWidth = 16
        Set imgLoad = frmMaster.imlDimension.ListImages.Add _
            (1, "rows", LoadPicture(strRootPath & "rows.ico"))
        Set imgLoad = frmMaster.imlDimension.ListImages.Add _
            (2, "columns", LoadPicture(strRootPath & "columns.ico"))
        Set imgLoad = frmMaster.imlDimension.ListImages.Add _
            (3, "pages", LoadPicture(strRootPath & "pages.ico"))
        'Load the three node images
        frmMaster.imlNode.ListImages.Clear
        frmMaster.imlNode.ImageHeight = 16
        frmMaster.imlNode.ImageWidth = 16
        Set imgLoad = frmMaster.imlNode.ListImages.Add _
            (1, "default", LoadPicture(strRootPath & "default.ico"))
        Set imgLoad = frmMaster.imlNode.ListImages.Add _
            (2, "check", LoadPicture(strRootPath & "check.ico"))
        Set imgLoad = frmMaster.imlNode.ListImages.Add _
            (3, "uncheck", LoadPicture(strRootPath & "uncheck.ico"))
    Else
        'Change the theme
        Call mdiTreeChild.DefaultImages
        Call mdiRowSelector.DefaultImages
        Call mdiColSelector.DefaultImages
        Call mdiPageSelector.DefaultImages
        strRootPath = App.Path & "\Theme" & intTheme & "\"
        'Load the three dimension images
        frmMaster.imlDimension.ListImages.Clear
        frmMaster.imlDimension.ImageHeight = 16
        frmMaster.imlDimension.ImageWidth = 16
        Set imgLoad = frmMaster.imlDimension.ListImages.Add _
            (1, "rows", LoadPicture(strRootPath & "rows.ico"))
        Set imgLoad = frmMaster.imlDimension.ListImages.Add _
            (2, "columns", LoadPicture(strRootPath & "columns.ico"))
        Set imgLoad = frmMaster.imlDimension.ListImages.Add _
            (3, "pages", LoadPicture(strRootPath & "pages.ico"))
        'Load the three node images
        frmMaster.imlNode.ListImages.Clear
        frmMaster.imlNode.ImageHeight = 16
        frmMaster.imlNode.ImageWidth = 16
        Set imgLoad = frmMaster.imlNode.ListImages.Add _
            (1, "default", LoadPicture(strRootPath & "default.ico"))
        Set imgLoad = frmMaster.imlNode.ListImages.Add _
            (2, "check", LoadPicture(strRootPath & "check.ico"))
        Set imgLoad = frmMaster.imlNode.ListImages.Add _
            (3, "uncheck", LoadPicture(strRootPath & "uncheck.ico"))
        Call mdiTreeChild.ChangeImages
        Call mdiRowSelector.ChangeImages
        Call mdiColSelector.ChangeImages
        Call mdiPageSelector.ChangeImages
    End If
Exit Sub
'Error Handler
LoadImages_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

