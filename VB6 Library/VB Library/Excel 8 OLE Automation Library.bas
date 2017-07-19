Attribute VB_Name = "modExcelOLEAutomation"
'Created by Robin G Brown, 30th April 1997
'-----------------------------------------------------------------------------
'   This module contains code for _
    handling OLE automation of Excel
'   History : Originally created by RGB and Mark R Matthews from May 1995 and onwards _
        Updated and modified by RGB in April/May 1997
'   Item : OLE declarations and _
           Xl Constants are in the XL_CONST.TXT file
'   These may not be needed as a reference to Excel is now available?
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modExcelOLEAutomation"
Private intCounter As Integer
Private intReturn As Integer
'OLE bits
Global objExcelApplication As Object
Global objWorkSheet As Object
Global strDataRange As String
'-----------------------------------------------------------------------------
'   Spoiler information for the general routine required to start Excel
'       run a macro and save the results, e.g. Produce a formatted Excel book
'
'   CreateExcelObject
'   OpenExcelFile/CreateNewExcelFile
'   RunExcelMacro
'   SaveExcelFile
'   SwitchToExcel
'   KillExcelObject
'
'   Emd of spoiler information
'-----------------------------------------------------------------------------

Function ActivateExcelFile(strFileName As String) As Boolean
'Updated by Robin G Brown, 30/4/97, Original Date : RGB/6/5/95
'Activate a file in Excel
'Sub Declares
Const conSub = "ActivateExcelFile"
    'Error Trap
    On Error GoTo ActivateExcelFile_ErrorHandler
    'Activate file
    objExcelApplication.Windows(strFileName).Activate
    ActivateExcelFile = True
Exit Function
ActivateExcelFile_ErrorHandler:
    ActivateExcelFile = False
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Function
End Function

Function AddNewExcelSheet(ByVal strSheetName As String) As Boolean
'Updated by Robin G Brown, 30/4/97, Original : RGB/6/5/95
'Add a new sheet to the active book and name it
'Sub Declares
Const conSub = "AddNewExcelSheet"
    'Error Trap
    Dim iResult As Integer
    On Error GoTo AddNewExcelSheet_ErrorHandler
    objExcelApplication.Worksheets.Add
    objExcelApplication.ActiveSheet.Name = strSheetName
    AddNewExcelSheet = True
Exit Function
AddNewExcelSheet_ErrorHandler:
    AddNewExcelSheet = False
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Function
End Function
Function CopyExcelSheet(strSourceSheet As String, strTargetBook As String) As Boolean
'Updated by Robin G Brown, 30/4/97, Original : RGB/6/5/95
'Copy the source sheet to the target book
'Sub Declares
Const conSub = "CopyExcelSheet"
    'Error Trap
    Dim iResult As Integer
    On Error GoTo CopyExcelSheet_ErrorHandler
    'iResult = SelectExcelFile(strTargetSheet)
    objExcelApplication.Sheets(strSourceSheet).Select
    objExcelApplication.Sheets(strSourceSheet).Copy objExcelApplication.Workbooks(strTargetBook).Sheets(1)
    'objExcelApplication.Windows(strSourceSheet).Activate
    'objExcelApplication.activesheet.Copy
    'objExcelApplication.Windows(strTargetSheet).Activate
    'objExcelApplication.activesheet.Paste
    'Sheets("Sheet2").Copy Workbooks("Book1").Sheets(1)
    CopyExcelSheet = True
Exit Function
CopyExcelSheet_ErrorHandler:
    CopyExcelSheet = False
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Function
End Function

Sub ClearExcelLinks()
'Updated by Robin G Brown, 30/4/97, Original : RGB/11/5/95
'This sub deletes all the formulas in the current sheet
'Sub Declares
Const conSub = "ClearExcelLinks"
    'Error Trap
Dim intReturn As Integer
    On Error Resume Next
    objExcelApplication.Cells.Select
    objExcelApplication.Selection.Copy
    intReturn = objExcelApplication.Selection.PasteSpecial(XLVALUES, XLNONE, False, False)
    On Error GoTo 0
End Sub

Sub CloseExcelFile()
'Updated by Robin G Brown, 30/4/97, Original : RGB/11/5/95
'This sub closes the current file in excel
'Sub Declares
Const conSub = "CloseExcelFile"
    'Error Trap
    On Error Resume Next
    objExcelApplication.ActiveWorkbook.[Close] (False)
    On Error GoTo 0
End Sub

Function CreateExcelObject() As Boolean
'Updated by Robin G Brown, 30/4/97, Original : RGB/6/5/95
'This sub creates and initialises an excel object
'Sub Declares
Const conSub = "CreateExcelObject"
    'Error Trap
    On Error GoTo CreateExcelObject_ErrorHandler
    'Get top-level Application object from Excel, objExcelApplication is already defined with module level scope
    Set objExcelApplication = CreateObject("Excel.Application")
    CreateExcelObject = True
Exit Function
CreateExcelObject_ErrorHandler:
    Set objExcelApplication = Nothing
    CreateExcelObject = False
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Function
End Function

Function CreateNewExcelFile() As Boolean
'Updated by Robin G Brown, 30/4/97, Original : RGB/6/5/95
'CreateNew a file in Excel
'Sub Declares
Const conSub = "CreateNewExcelFile"
    'Error Trap
    On Error GoTo CreateNewExcelFile_ErrorHandler
    'CreateNew file
    objExcelApplication.Workbooks.Add
    CreateNewExcelFile = True
Exit Function
CreateNewExcelFile_ErrorHandler:
    CreateNewExcelFile = False
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Function
End Function

Sub HideExcelToolbars()
'Updated by Robin G Brown, 30/4/97, Original : RGB/unknown
'Hide all the Excel Toolbars
'Sub Declares
Const conSub = "HideExcelToolbars"
Dim intToolbarCount As Integer
    'Error Trap
    On Error Resume Next
    For intToolbarCount = 1 To objExcelApplication.Toolbars.Count
        objExcelApplication.Toolbars(intToolbarCount).Visible = False
    Next
    On Error GoTo 0
End Sub

Sub KillExcelObject()
'Updated by Robin G Brown, 30/4/97, Original : RGB/6/5/95
'Get rid of the Excel object and free memory
'Sub Declares
Const conSub = "KillExcelObject"
    'Error Trap
    On Error Resume Next
    objExcelApplication.Quit
    Set objExcelApplication = Nothing
    On Error GoTo 0
End Sub

Function OpenExcelFile(strCSVFileName As String) As Boolean
'Updated by Robin G Brown, 30/4/97, Original : RGB/6/5/95
'Open a file in Excel
'Sub Declares
Const conSub = "OpenExcelFile"
    'Error Trap
    On Error GoTo OpenExcelFile_ErrorHandler
    'Open file
    objExcelApplication.Workbooks.Open (strCSVFileName)
    'objExcelApplication.Windows(strCSVFileName).Activate
    OpenExcelFile = True
Exit Function
OpenExcelFile_ErrorHandler:
    OpenExcelFile = False
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Function
End Function

Function PrintExcelWorkbook() As Boolean
'Updated by Robin G Brown, 30/4/97, Original : RGB/11/5/95
'Print the current file - will probably only print selected sheets
'Sub Declares
Const conSub = "PrintExcelWorkbook"
    'Error Trap
    On Error GoTo PrintExcelWorkbook_ErrorHandler
    objExcelApplication.ActiveWorkbook.PrintOut  'Copies:=1
    PrintExcelWorkbook = True
Exit Function
PrintExcelWorkbook_ErrorHandler:
    PrintExcelWorkbook = False
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Function
End Function

Function RenameSheet(strNewName As String) As Boolean
'Updated by Robin G Brown, 30/4/97, Original : RGB/6/7/95
'Rename the current worksheet according to the supplied parameter
'Sub Declares
Const conSub = "RenameSheet"
    'Error Trap
    On Error GoTo RenameSheet_ErrorHandler
    objExcelApplication.ActiveWorkbook.ActiveSheet.Name = strNewName
    RenameSheet = True
Exit Function
RenameSheet_ErrorHandler:
    RenameSheet = False
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Function
End Function

Function RunExcelMacro(strMacroName As String) As Boolean
'Updated by Robin G Brown, 30/4/97, Original : RGB/6/5/95
'Open a file in Excel
'Sub Declares
Const conSub = "RunExcelMacro"
    'Error Trap
    On Error GoTo RunExcelMacro_ErrorHandler
    'Open file
    objExcelApplication.Run (strMacroName)
    RunExcelMacro = True
Exit Function
RunExcelMacro_ErrorHandler:
    RunExcelMacro = False
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Function
End Function

Function SaveExcelFile(strFileName) As Boolean
'Updated by Robin G Brown, 30/4/97, Original : RGB/11/5/95
'Save the current excel file with the given filename, deleting any file with the previous name
'Sub Declares
Const conSub = "SaveExcelFile"
    'Error Trap
    On Error GoTo SaveExcelFile_ErrorHandler
    intReturn = objExcelApplication.ActiveWorkbook.SaveAs(strFileName, XLNORMAL)
    SaveExcelFile = True
Exit Function
SaveExcelFile_ErrorHandler:
    SaveExcelFile = False
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Function
End Function

Function SaveExcelFileasCSV(strCSVFileName) As Boolean
'Updated by Robin G Brown, 30/4/97, Original : RGB/11/5/95
'Save the current excel file with the given filename in CSV format, deleting any file with the previous name
'Sub Declares
Const conSub = "SaveExcelFileasCSV"
    'Error Trap
    On Error GoTo SaveExcelFileasCSV_ErrorHandler
    intReturn = objExcelApplication.ActiveWorkbook.SaveAs(strCSVFileName, XLCSV)
    SaveExcelFileasCSV = True
Exit Function
SaveExcelFileasCSV_ErrorHandler:
    SaveExcelFileasCSV = False
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Function
End Function

Function SelectExcelFile(strFileName As String) As Boolean
'Updated by Robin G Brown, 30/4/97, Original : RGB/11/5/95
'Select the file with the given filename
'Sub Declares
Const conSub = "SelectExcelFile"
    'Error Trap
    On Error GoTo SelectExcelFile_ErrorHandler
    intReturn = objExcelApplication.Windows(strFileName).Select
    SelectExcelFile = True
Exit Function
SelectExcelFile_ErrorHandler:
    SelectExcelFile = False
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Function
End Function

Sub Set5ExcelUserDefaults(strTitle As String, strSubject As String, strAuthor As String, strKeywords As String, strComments As String)
'Updated by Robin G Brown, 30/4/97, Original : RGB/9/6/95
'Set the default values for user, title etc.
'Sub Declares
Const conSub = "Set5ExcelUserDefaults"
    'Error Trap
    On Error Resume Next
    objExcelApplication.ActiveWorkbook.Title = strTitle
    objExcelApplication.ActiveWorkbook.Subject = strSubject
    objExcelApplication.ActiveWorkbook.Author = strAuthor
    objExcelApplication.ActiveWorkbook.Keywords = strKeywords
    objExcelApplication.ActiveWorkbook.Comments = strComments
    On Error GoTo 0
End Sub

Sub SwitchToExcel()
'Updated by Robin G Brown, 30/4/97, Original : RGB/6/5/95
'This sub causes Excel to appear full size, provided the object is already initialised
'Sub Declares
Const conSub = "SwitchToExcel"
    'Error Trap
    On Error Resume Next
    'Show Excel full screen for editing etc.
    objExcelApplication.Visible = True
    objExcelApplication.WindowState = XLMAXIMIZED
    On Error GoTo 0
End Sub

