Attribute VB_Name = "modExcelOLEAutomation"
Option Explicit
'---Comment Box by Robin G Brown---------------------------------
'   Item : OLE declarations,
'           Xl Constants are in the XL_CONST.TXT file
'   Date : 6/5/95
'----------------------------------------------------------------
'OLE bits
Global objExcelApplication As Object
Global objWorkSheet As Object
Global strDataRange As String
'----------------------------------------------------------------

Function ActivateExcelFile(strFileName As String) As Integer
'---Comment Box by Robin G Brown---------------------------------
'   Item : Activate a file in Excel
'   Date : 6/5/95
'----------------------------------------------------------------

    On Error GoTo ActivateExcelFile_ErrorHandler
    'Activate file
    objExcelApplication.Windows(strFileName).Activate
    ActivateExcelFile = True

Exit Function
ActivateExcelFile_ErrorHandler:
    ActivateExcelFile = False
    Exit Function


End Function

Function AddExcelSheet(strSourceSheet As String, strTargetBook As String) As Integer
'---Comment Box by Robin G Brown---------------------------------
'   Item : Add a sheet to a workbook in Excel
'   Date : 6/5/95
'----------------------------------------------------------------
    Dim iResult As Integer

    On Error GoTo AddExcelSheet_ErrorHandler
    'iResult = SelectExcelFile(strTargetSheet)

    objExcelApplication.Sheets(strSourceSheet).Select
    objExcelApplication.Sheets(strSourceSheet).Copy objExcelApplication.Workbooks(strTargetBook).Sheets(1)
    
    'objExcelApplication.Windows(strSourceSheet).Activate
    'objExcelApplication.activesheet.Copy
    'objExcelApplication.Windows(strTargetSheet).Activate
    'objExcelApplication.activesheet.Paste
    'Sheets("Sheet2").Copy Workbooks("Book1").Sheets(1)
    AddExcelSheet = True

Exit Function
AddExcelSheet_ErrorHandler:
    AddExcelSheet = False
    Exit Function

End Function

Sub ClearExcelLinks()
'---Comment Box by Robin G Brown---------------------------------
'   Item : This sub deletes all the formulas in the current sheet
'   Date : 11/5/95
'----------------------------------------------------------------

Dim intReturn As Integer

    On Error Resume Next
    objExcelApplication.Cells.Select
    objExcelApplication.Selection.Copy
    intReturn = objExcelApplication.Selection.PasteSpecial(XLVALUES, XLNONE, False, False)
    On Error GoTo 0

End Sub

Sub CloseExcelFile()
'---Comment Box by Robin G Brown---------------------------------
'   Item : This sub closes the current file in excel
'   Date : 11/5/95
'----------------------------------------------------------------

    On Error Resume Next
    objExcelApplication.ActiveWorkbook.[Close] (False)
    On Error GoTo 0

End Sub

Function CreateExcelObject() As Integer
'---Comment Box by Robin G Brown---------------------------------
'   Item : This sub creates and initialises an excel object
'   Date : 6/5/95
'----------------------------------------------------------------
    
    On Error GoTo CreateExcelObject_ErrorHandler
    'Get top-level Application object from Excel, objExcelApplication is already defined with module level scope
    Set objExcelApplication = CreateObject("Excel.Application")
    CreateExcelObject = True

Exit Function
CreateExcelObject_ErrorHandler:
    Set objExcelApplication = Nothing
    CreateExcelObject = False
    Exit Function

End Function

Function CreateNewExcelFile() As Integer
'---Comment Box by Robin G Brown---------------------------------
'   Item : CreateNew a file in Excel
'   Date : 6/5/95
'----------------------------------------------------------------

    On Error GoTo CreateNewExcelFile_ErrorHandler
    'CreateNew file
    objExcelApplication.Workbooks.Add
    CreateNewExcelFile = True

Exit Function
CreateNewExcelFile_ErrorHandler:
    CreateNewExcelFile = False
    Exit Function

End Function

Sub HideExcelToolbars()

Dim intToolbarCount As Integer

    On Error Resume Next
    For intToolbarCount = 1 To objExcelApplication.Toolbars.Count
        objExcelApplication.Toolbars(intToolbarCount).Visible = False
    Next
    On Error GoTo 0

End Sub

Sub KillExcelObject()
'---Comment Box by Robin G Brown---------------------------------
'   Item : Get rid of the Excel object and free memory
'   Date : 6/5/95
'----------------------------------------------------------------

    On Error Resume Next
    objExcelApplication.Quit
    Set objExcelApplication = Nothing
    On Error GoTo 0

End Sub

Function OpenExcelFile(strCSVFileName As String) As Integer
'---Comment Box by Robin G Brown---------------------------------
'   Item : Open a file in Excel
'   Date : 6/5/95
'----------------------------------------------------------------

    On Error GoTo OpenExcelFile_ErrorHandler
    'Open file
    objExcelApplication.Workbooks.Open (strCSVFileName)
    'objExcelApplication.Windows(strCSVFileName).Activate
    OpenExcelFile = True

Exit Function
OpenExcelFile_ErrorHandler:
    OpenExcelFile = False
    Exit Function

End Function

Function PrintExcelWorkbook() As Integer
'---Comment Box by Robin G Brown---------------------------------
'   Item : Print the current file - will probably only print selected sheets
'   Date : 11/5/95
'----------------------------------------------------------------

    On Error GoTo PrintExcelWorkbook_ErrorHandler
    objExcelApplication.ActiveWorkbook.PrintOut  'Copies:=1
    PrintExcelWorkbook = True

Exit Function
PrintExcelWorkbook_ErrorHandler:
    PrintExcelWorkbook = False
    Exit Function

End Function

Function RenameSheet(strNewName As String) As Integer
'---Comment Box by Tony Christian---------------------------------
'   Item : Rename the current worksheet according to
'          the supplied parameter
'   Date : 6/7/95
'----------------------------------------------------------------
Dim intReturn As Integer

    On Error GoTo RenameSheet_ErrorHandler
    objExcelApplication.ActiveWorkbook.ActiveSheet.Name = strNewName
    RenameSheet = True

Exit Function
RenameSheet_ErrorHandler:
    RenameSheet = False
    Exit Function


End Function

Function RunExcelMacro(strMacroName As String) As Integer
'---Comment Box by Robin G Brown---------------------------------
'   Item : Open a file in Excel
'   Date : 6/5/95
'----------------------------------------------------------------

    On Error GoTo RunExcelMacro_ErrorHandler
    'Open file
    objExcelApplication.Run (strMacroName)
    RunExcelMacro = True

Exit Function
RunExcelMacro_ErrorHandler:
    RunExcelMacro = False
    Exit Function


End Function

Function SaveExcelFile(strFileName) As Integer
'---Comment Box by Robin G Brown---------------------------------
'   Item : Save the current excel file with the given filename, deleting any file with the previous name
'   Date : 11/5/95
'----------------------------------------------------------------
Dim intReturn As Integer

    On Error GoTo SaveExcelFile_ErrorHandler
    intReturn = objExcelApplication.ActiveWorkbook.SaveAs(strFileName, XLNORMAL)
    SaveExcelFile = True

Exit Function
SaveExcelFile_ErrorHandler:
    SaveExcelFile = False
    Exit Function

End Function

Function SaveExcelFileasCSV(strCSVFileName) As Integer
'---Comment Box by Robin G Brown---------------------------------
'   Item : Save the current excel file with the given filename
'           in CSV format, deleting any file with the previous name
'   Date : 11/5/95
'----------------------------------------------------------------
Dim intReturn As Integer

    On Error GoTo SaveExcelFileasCSV_ErrorHandler
    intReturn = objExcelApplication.ActiveWorkbook.SaveAs(strCSVFileName, XLCSV)
    SaveExcelFileasCSV = True

Exit Function
SaveExcelFileasCSV_ErrorHandler:
    SaveExcelFileasCSV = False
    Exit Function

End Function

Function SelectExcelFile(strFileName As String) As Integer
'---Comment Box by Robin G Brown---------------------------------
'   Item : Select the file with the given filename
'   Date : 11/5/95
'----------------------------------------------------------------
Dim intReturn As Integer

    On Error GoTo SelectExcelFile_ErrorHandler
    intReturn = objExcelApplication.Windows(strFileName).Select
    SelectExcelFile = True

Exit Function
SelectExcelFile_ErrorHandler:
    SelectExcelFile = False
    Exit Function

End Function

Sub Set5ExcelUserDefaults(strTitle As String, strSubject As String, strAuthor As String, strKeywords As String, strComments As String)

'---Comment Box by Robin G Brown---------------------------------
'   Item :  Set the default values for user, title etc.
'   Date :  9/6/95
'----------------------------------------------------------------

    On Error Resume Next
    objExcelApplication.ActiveWorkbook.Title = strTitle
    objExcelApplication.ActiveWorkbook.Subject = strSubject
    objExcelApplication.ActiveWorkbook.Author = strAuthor
    objExcelApplication.ActiveWorkbook.Keywords = strKeywords
    objExcelApplication.ActiveWorkbook.Comments = strComments
    On Error GoTo 0

End Sub

Sub SwitchToExcel()
'---Comment Box by Robin G Brown---------------------------------
'   Item : This sub causes Excel to appear full size,
'           provided the object is already initialised
'   Date : 6/5/95
'----------------------------------------------------------------

    On Error Resume Next
    'Show Excel full screen for editing etc.
    objExcelApplication.Visible = True
    objExcelApplication.WindowState = XLMAXIMIZED
    On Error GoTo 0

End Sub

