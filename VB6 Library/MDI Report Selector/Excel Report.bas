Attribute VB_Name = "modExcelReport"
'Created by Robin G Brown, 30th April 1997
'-----------------------------------------------------------------------------
'   This module contains code for _
    running ActiveX/OLE automation reports
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modExcelReport"
Private intCounter As Integer
Private intReturn As Integer
Public trpTriApp As Triangulator.Application

Public Sub ShowTriangulator()
'Created by Robin G Brown, 2/5/97
'Calls the Triangulator app to show the report
'Sub Declares
Const conSub = "ShowTriangulator"
    'Error Trap
    On Error GoTo ShowTriangulator_ErrorHandler
    Set trpTriApp = CreateObject("Triangulator.Application")
    Call trpTriApp.Visible(True)
    'Set trpTriApp = Nothing
Exit Sub
'Error Handler
ShowTriangulator_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Sub ExcelReport(ByVal strReportName As String)
'Created by Robin G Brown, 30/4/97
'Start Excel, send some data across, save it and show the user what's happened
'Sub Declares
Const conSub = "ExcelReport"
Dim strExcelFileName As String
    'Error Trap
    On Error GoTo ExcelReport_ErrorHandler
    intReturn = CreateExcelObject()
    If intReturn = True Then
        'intReturn = CreateNewExcelFile()
        intReturn = OpenExcelFile(App.Path & "\trireport.xls")
        'Call CreateFakeTriangulationReport
        'RunExcelMacro
        'strExcelFileName = "Report.xls"
        'If CheckForDuplicate(strExcelFileName) = True Then
        '    intReturn = SaveExcelFile(strExcelFileName)
        'End If
        'Call CreateFakeTriangulationReport(strReportName)
        'Call FormatTriangulationReport("Triangulation Report")
        Call SwitchToExcel
        'Call KillExcelObject - this is in modMain.ApplicationEnd!
    End If
Exit Sub
'Error Handler
ExcelReport_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Sub CreateFakeTriangulationReport(ByVal strReportName As String)
'Created by Robin G Brown, 30/4/97
'Creates a fake triangulation report in the active workbook
'Sub Declares
Const conYears = 30
Dim strTriangulationSheet As String
Dim intColCounter As Integer
Dim intYearCounter As Integer
Dim intRowCounter As Integer
Dim intOffset As Integer
    'Error Trap
    On Error Resume Next
    frmFlood.Show
    Call frmFlood.SetFlood(0)
    strTriangulationSheet = "Triangulation Report"
    intReturn = AddNewExcelSheet(strTriangulationSheet)
    If intReturn = True Then
        With objExcelApplication.Worksheets(strTriangulationSheet)
            'Set the number of cols and rows
            'Write the col headers
            .Cells(1, 1) = Trim$(StripCRLF(strReportName))
            intOffset = 1
            For intYearCounter = 1 To conYears
                    .Cells(1 + intOffset, ((intYearCounter - 1) * 4) + 2) = "UWYear:" & Format$(intYearCounter, "00")
                For intColCounter = 1 To 4
                    .Cells(2 + intOffset, ((intYearCounter - 1) * 4) + intColCounter + 1) = "Acc" & Format$(intColCounter, "00")
                Next
            Next
            intOffset = 3
            For intRowCounter = 1 To conYears
                Call frmFlood.SetFlood((intRowCounter / (conYears + 2)) * 100)
                'Write the row header
                .Cells(intRowCounter + intOffset, 1) = "Year:" & Format$((intRowCounter), "00")
                'Write the data
                For intYearCounter = 1 To conYears
                    If intYearCounter > (conYears + 1 - intRowCounter) Then
                        Exit For
                    Else
                        For intColCounter = 1 To 4
                            .Cells(intRowCounter + intOffset, ((intYearCounter - 1) * 4) + intColCounter + 1) = Format$(intYearCounter * intRowCounter * intColCounter, "00.00")
                        Next
                    End If
                Next
            Next
        End With
    End If
    Unload frmFlood
End Sub

Public Sub FormatTriangulationReport(ByVal strTriangulationSheet As String)
'Created by Robin G Brown, 30/4/97
'Format a triangulation report sheet
'Sub Declares
Const conSub = "FormatTriangulationReport"
    'Error Trap
    On Error GoTo FormatTriangulationReport_ErrorHandler
    'ColumnWidth property for Excel is in CHARACTERS!
    For intCounter = 2 To 122
        'ColumnWidth property for Excel is in CHARACTERS!
        objExcelApplication.Worksheets(strTriangulationSheet).Columns(intCounter).Select
        objExcelApplication.Selection.ColumnWidth = 5
        If (intCounter + 2) Mod 4 = 0 Then
            With objExcelApplication.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End If
    Next
    objExcelApplication.Worksheets(strTriangulationSheet).Rows(3).Select
    With objExcelApplication.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    objExcelApplication.Worksheets(strTriangulationSheet).Cells(1, 1).Select
    With objExcelApplication.Selection.Font
        .Bold = True
    End With
Exit Sub
'Error Handler
FormatTriangulationReport_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

