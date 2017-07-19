Attribute VB_Name = "modCommonMenus"
'Created by Robin G Brown, 27th August 1997
'-----------------------------------------------------------------------------
'   This module contains code for _
    DESCRIPTION_OF_MODULE_FUNCTIONALITY_HERE
'-----------------------------------------------------------------------------
'   Special Instructions
'   Add mnuImmediateError_Click to Form modules
'   calling appropriate functions, DelayedErrorMenuRoutine|ImmediateErrorMenuRoutine
'   as required
'   This allows a common set of code to be created which can be called from multiple menus
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modCommonMenus"
Private lngCounter As Long
Private lngReturn As Long

Public Sub NewReportMenuRoutine()
'Created by Robin G Brown, 28/8/97
'Create a new report window
'Sub declares
Dim frmNew As frmGridReport
    'Error Trap
    On Error Resume Next
    'Load the form, but don't do anything with it
    Set frmNew = New frmGridReport
    Load frmNew
    'Set up the data array for the form
    Call frmNew.DataRandomInitialize(frmMaster.cmbStartYear.Text)
    'Now show the form
    frmNew.Show
    Set frmNew = Nothing
End Sub

Public Sub NewGraphMenuRoutine()
'Created by Robin G Brown, 28/8/97
'Create a new graph window
'Sub declares
    'Error Trap
    On Error Resume Next
    'Only one copy of frmchartreport is allowed at any time
    If FormInstances("frmChartReport") > 0 Then
        Beep
        Exit Sub
    End If
    'Load the form, but don't do anything with it
    Load frmChartReport
    If Err.Number <> 0 Then
        Exit Sub
    End If
    'Now show the form
    frmChartReport.Show
End Sub

Public Sub FileOpenMenuRoutine()
'Created by Robin G Brown, 28/8/97
'Retrieve the contents of a file into the current grid
'sub declares
Dim strFileAndPath As String
Dim frmNew As frmGridReport
    'Error Trap
    On Error GoTo FileOpenMenuRoutine_ErrorHandler
    With frmMaster.dlgMaster
        .CancelError = True
        .Flags = cdlOFNHideReadOnly
        .Filter = conAnyTriFilter
        .FilterIndex = 2
        .ShowOpen
        strFileAndPath = .filename
    End With
    Select Case GetFileExtension(strFileAndPath)
    Case "trs"
        'Session file
        Call frmMaster.SessionFileOpenMenuRoutine(strFileAndPath)
    Case "trg"
        'Graph file
        'Only one copy of frmchartreport is allowed at any time
        If FormInstances("frmChartReport") > 0 Then
            Beep
            Exit Sub
        End If
        'Load the form, but don't do anything with it
        Load frmChartReport
        If Err.Number <> 0 Then
            Exit Sub
        End If
        Call frmChartReport.GraphFileOpenMenuRoutine(strFileAndPath)
        frmChartReport.Show
    Case "trr"
        'Report file
        Set frmNew = New frmGridReport
        Load frmNew
        Call frmNew.GridFileOpenMenuRoutine(strFileAndPath)
        frmNew.Show
        Set frmNew = Nothing
    Case Else
        'Other file = Error
        MsgBox "Not a valid filetype, cannot continue.", vbOKOnly, App.ProductName
    End Select
Exit Sub
FileOpenMenuRoutine_ErrorHandler:
    Exit Sub
End Sub




