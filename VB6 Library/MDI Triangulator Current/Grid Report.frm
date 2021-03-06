VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmGridReportOld 
   Caption         =   "Report"
   ClientHeight    =   5475
   ClientLeft      =   1155
   ClientTop       =   1560
   ClientWidth     =   9105
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Grid Report.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5475
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "-"
      Top             =   60
      Width           =   555
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1020
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid grdReport 
      DragIcon        =   "Grid Report.frx":0442
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7858
      _Version        =   327680
      Rows            =   20
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   2
      BackColor       =   -2147483645
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      BorderStyle     =   0
   End
End
Attribute VB_Name = "frmGridReportOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 2nd May 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    showing a report grid
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmBoilerPlate"
Private intCounter As Integer
Private intReturn As Integer
Private intColCounter As Integer
Private intRowCounter As Integer
Public dtaDataArray As TriDataArray

Private Sub Form_Initialize()
'Created by Robin G Brown, 27/5/97
'Initialize the data array
'Sub Declares
Const conSub = "Form_Initialize"
    'Error Trap
    On Error Resume Next
    'Make sure the datagrid array is set up
    Set dtaDataArray = New TriDataArray
    If Err.Number <> 0 Then
        MsgBox "Data failed to initialise due to error : " & Err.Description, vbExclamation, conModule & " : " & conSub
        Err.Clear
    End If
End Sub

Private Sub Form_Load()
'Created by Robin G Brown, 2/5/97
'Default Behaviour
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    'Data will be set by whatever causes this form to load
Exit Sub
'Error Handler
Form_Load_ErrorHandler:
    Exit Sub
End Sub

Private Sub Form_Resize()
'Created by Robin G Brown, 2/5/97
'Default Behaviour
'Sub Declares
Const conSub = "Form_Resize"
    'Error Trap
    On Error GoTo Form_Resize_ErrorHandler
    'Resize the grid to fill the form
    If Me.WindowState <> 1 Then
        grdReport.Height = Me.Height - 420
        grdReport.Width = Me.Width - 120
    End If
Exit Sub
'Error Handler
Form_Resize_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub grdReport_DblClick()
'Created by Robin G Brown, 14/5/97
'Select items depending on where the dblclick took place
'Sub Declares
Const conSub = "grdReport_DblClick"
Dim intAccount As Integer
Dim frmChart As frmChartReport
    'Error Trap
    On Error GoTo grdReport_DblClick_ErrorHandler
    If gfgcCurrent.GridRow >= grdReport.FixedRows _
    And gfgcCurrent.GridCol >= grdReport.FixedCols Then
        'Show this column (or account, if and other years is selected) _
            in a chart report form
        Set frmChart = New frmChartReport
        Load frmChart
        intAccount = dtaDataArray.GetAccountFromColumn(gfgcCurrent.GridCol)
        If frmMaster.stbStatus.Panels("panYears").Tag = "OFF" Then
            'Single column
            Call dtaDataArray.FillChart(frmChart.graReport, grdReport, intAccount, gfgcCurrent.GridCol)
        Else
            'Multiple column
            Call dtaDataArray.FillChart(frmChart.graReport, grdReport, intAccount, 0)
        End If
        Call PositionLegend(frmChart.graReport, conAlignLegendBottomCentre, 90, 10)
        Call MaximizePlot(frmChart.graReport)
        frmChart.Show
    End If
    Set frmChart = Nothing
Exit Sub
'Error Handler
grdReport_DblClick_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub grdReport_DragDrop(Source As Control, X As Single, Y As Single)
'Created by Robin G Brown, 27/5/97
'Cancel a drag drop op
'Sub Declares
Const conSub = "DelayedErrorSub"
    'Error Trap
    On Error Resume Next
    grdReport.Drag vbCancel
    gbooColumnDragInProgress = False
End Sub

Private Sub grdReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Created by Robin G Brown, 13/5/97
'Make sure that gfgcCurrent is set
'Sub Declares
    'Error Trap
    On Error Resume Next
    'First we need to find out where the mouse is _
        and what keys were pressed at the time
    gmstMouse.Button = Button
    gmstMouse.Shift = Shift
    gmstMouse.X = X
    gmstMouse.Y = Y
    Call GetCellFromMouse(grdReport, gmstMouse, gfgcCurrent)
    'Then popup a menu if required
    If Button = vbRightButton Then
        'CODE_HERE
    End If
End Sub

Private Sub grdReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Created by Robin G Brown, 23/5/97
'Start a drag/drop operation
'Sub Declares
    'Error Trap
    On Error Resume Next
    Text1.Text = Button
    'If IsLeftDown(Button) Then
    If Button = 1 Then
        grdReport.Drag vbBeginDrag
        gbooColumnDragInProgress = True
    End If
End Sub

Private Sub grdReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Created by Robin G Brown, 13/5/97
'Make sure that gfgcCurrent is set
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Cancel a drag/drop operation
    grdReport.Drag vbCancel
    gbooColumnDragInProgress = False
End Sub

Public Sub DataRandomInitialize(Optional ByVal varStartYear As Variant)
'Created by Robin G Brown, 27/5/97
'Initialise the data using the given start year and creating random data
'Sub Declares
Const conSub = "DataRandomInitialize"
Dim intStartYear As Integer
    'Error Trap
    On Error Resume Next
    If IsMissing(varStartYear) Then
        'Create a default start year
        intStartYear = 1990
    Else
        intStartYear = CInt(varStartYear)
        If Err.Number <> 0 Then
            intStartYear = 1990
            MsgBox "Error : " & Err.Description, vbExclamation, conModule & " : " & conSub
            Err.Clear
        End If
    End If
    'Create some random data
    Call dtaDataArray.CreateData(intStartYear)
    dtaDataArray.Name = "Report"
    Me.Caption = dtaDataArray.Name
    'Fill the grid from dtaDataArray
    Call DrawGrid
End Sub

Public Sub DrawGrid()
'Created by Robin G Brown, 14/5/97
'Redisplay the grid with the current settings
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Fill the grid from dtaDataArray
    grdReport.Redraw = False
    Call dtaDataArray.FillGrid(grdReport)
    Call AutoSizeGrid(grdReport)
    Call AlignAllText(grdReport, conAlignTextCenterCenter)
    grdReport.Redraw = True
End Sub


