Attribute VB_Name = "modMain"
'Created by Robin G Brown, 2nd May 1997
'-----------------------------------------------------------------------------
'   This module contains code for _
    starting the TriReporter ActiveX exe _
    and global variable declarations for the application
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modMain"
Private lngCounter As Long
Private lngReturn As Long
'Global Drag/Drop variables
Public gbooColumnDragInProgress As Boolean
'Database constants
Private Const conDefaultDBName = ""
'Triangulation Selection constants
Public Const conSelectYear = 100
Public Const conSelectPeriod = 101
Public Const conSelectAccount = 102
'Triangulation file filter
Public Const conTriFilter = "Triangulator Files (*.tr*)|*.tr*"
Public Const conTriReportFilter = "Triangulator Report (*.trr)|*.trr"
Public Const conTriGraphFilter = "Triangulator Graph (*.trg)|*.trg"
Public Const conTriSessionFilter = "Triangulator Session (*.trs)|*.trs"
Public Const conAnyTriFilter = conAnyFilter & "|" & conTriFilter & "|" & conTriReportFilter & "|" & conTriGraphFilter & "|" & conTriSessionFilter

Sub Main()
'Created by Robin G Brown, 2/5/97
'Default behaviour
'Sub Declares
Const conSub = "Main"
    'Error Trap
    On Error GoTo Main_ErrorHandler
    'ApplicationInitialise is always the first code to run - RGB/15/9/97
    Call ApplicationInitialise
    'This code only runs when the application is started as an exe, _
        it does not run when creating a TriReporter.TriReport object
    'Show splash form and give it time to sort itself out
    Load frmSplash
    'Set the picture to show
    'frmSplash.PictureFile = ""
    'Show the splash form
    frmSplash.Show
    DoEvents
    'Carry out any slow tasks...
    Call PauseWithEvents
    'Show master form
    frmMaster.Show
    'No forms  should be loaded at the start - RGB/15/9/97
    'Get rid of splash form
    Unload frmSplash
Exit Sub
'Error Handler
Main_ErrorHandler:
    Select Case Err
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Error$(Err), vbExclamation, conModule & " : " & conSub
    End Select
    'Exit the program
    Call ApplicationEnd
    Exit Sub
End Sub

Public Sub ApplicationInitialise()
'Created by Robin G Brown, 6/5/97
'We are using a seperate initialise so that it can be called from more than one place, _
    depending on the way in which this application was started.
'IMPORTANT - There must be no form based activity in this routine
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Check for Instances of this program
    If App.PrevInstance = True Then
        Err.Raise vbObjectError, App.ProductName, "Only one instance of this program is allowed, cannot continue."
        End
    End If
    'Any database connections required should be done here
    'Call ConnectSelectedDatabase(conDefaultDBName)
End Sub

Sub ApplicationEnd()
'Created by Robin G Brown, 2/5/97
'Default behaviour
'Sub Declares
Const conSub = "ApplicationEnd"
    'Error Trap
    On Error Resume Next
    'CODE_HERE, _
        Close databases, etc _
        Free memory for objects _
        Save registry information _
        and Tidy up
    Set frmMaster = Nothing
    'Call DisconnectDatabase
End Sub

