VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form mdiReportGrid 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Report"
   ClientHeight    =   3705
   ClientLeft      =   4650
   ClientTop       =   5280
   ClientWidth     =   4605
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
   Icon            =   "ReportGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3705
   ScaleWidth      =   4605
   Begin MSGrid.Grid grdReport 
      Height          =   3675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3915
      _Version        =   65536
      _ExtentX        =   6906
      _ExtentY        =   6482
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   50
      Cols            =   50
   End
End
Attribute VB_Name = "mdiReportGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 22nd April 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    displaying results in a grid
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "mdiReportGrid"
Private intCounter As Integer
Private intReturn As Integer
Private Sub Form_Load()
'Created by Robin G Brown, 22/4/97
'Default behaviour
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    'CODE_HERE
Exit Sub
'Error Handler
Form_Load_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Created by Robin G Brown, 25/4/97
'do not allow the child to be closed unless the program is ending
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    If UnloadMode = vbFormControlMenu Then
        MsgBox "You may not close this window."
        Cancel = True
    End If
End Sub
Private Sub Form_Resize()
'Created by Robin G Brown, 22/4/97
'Default behaviour
'Sub Declares
Const conSub = "Form_Resize"
    'Error Trap
    On Error GoTo Form_Resize_ErrorHandler
    If Me.WindowState <> 1 Then
        If Me.Height < 1000 Then Me.Height = 1000
        If Me.Width < 1000 Then Me.Width = 1000
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
Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, 22/4/97
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    'CODE_HERE
End Sub
Public Sub FillReportGrid()
'Created by Robin G Brown, 15/11/97
'Fills the grid with a selection of reports, _
    dictated by the contents of the various tvwSelection controls
'Sub Declares
Const conSub = "FillReportGrid"
Const conIndent = 3
Const conRowHeight = 200
Dim intNodeCounter1 As Integer
Dim intNodeCounter2 As Integer
Dim intNodeCounter3 As Integer
Dim intNodeCounter4 As Integer
Dim intNodeCounter5 As Integer
Dim intWriteCounter2 As Integer
Dim intWriteCounter3 As Integer
Dim intWriteCounter4 As Integer
Dim intWriteCounter5 As Integer
Dim nodDimension1 As Node
Dim nodDimension2 As Node
Dim nodDimension3 As Node
Dim nodDimension4 As Node
Dim nodDimension5 As Node
Dim intRowDimensions As Integer
Dim intColDimensions As Integer
Dim intColCounter As Integer
Dim intRowCounter As Integer
Dim intIndentation As Integer
    'No error trap, just make sure this completes
    On Error Resume Next
    'Find out how many dimensions we have to cope with, _
        enquiring of each selection form as to how many dimensions there are in use, _
        assuming that there are only as many controls as dimensions and no 'empty' controls
    intRowDimensions = mdiRowSelector.GetDimensions()
    intColDimensions = mdiColSelector.GetDimensions()
    'Set the number of fixed cols and rows
    grdReport.FixedCols = 0
    grdReport.FixedRows = 0
    'Set up the header cols and rows
    grdReport.Rows = intColDimensions
    grdReport.Cols = intRowDimensions
    'Read from the  treeview controls to populate the headers, then fill in the grid with info from the headers
    'Read from mdiColSelector.tvwSelection(X) into col headers
    'For each node in mdiColSelector.tvwSelection(1) we do the following
    For intNodeCounter1 = 2 To mdiColSelector.tvwSelection(1).Nodes.Count
        Set nodDimension1 = mdiColSelector.tvwSelection(1).Nodes(intNodeCounter1)
        If AddHeader(nodDimension1) = True Then
            grdReport.Cols = grdReport.Cols + 1
            grdReport.Col = grdReport.Cols - 1
            grdReport.Row = 0
            intIndentation = GetIndentation(nodDimension1)
            grdReport.Text = Space$(intIndentation * conIndent) & nodDimension1.Text
            If grdReport.RowHeight(0) < (intIndentation + 1 * conRowHeight) Then
                grdReport.RowHeight(0) = intIndentation + 1 * conRowHeight
            End If
            'Now multiply out by the second dimension
            If intColDimensions > 1 Then
                intWriteCounter2 = 0
                For intNodeCounter2 = 2 To mdiColSelector.tvwSelection(2).Nodes.Count
                    Set nodDimension2 = mdiColSelector.tvwSelection(2).Nodes(intNodeCounter2)
                    If AddHeader(nodDimension2) = True Then
                        If intWriteCounter2 > 0 Then
                            grdReport.Cols = grdReport.Cols + 1
                        End If
                        grdReport.Col = grdReport.Cols - 1
                        grdReport.Row = 1
                        intIndentation = GetIndentation(nodDimension2)
                        grdReport.Text = Space$(intIndentation * conIndent) & nodDimension2.Text
                        If grdReport.RowHeight(1) < (intIndentation + 1 * conRowHeight) Then
                            grdReport.RowHeight(1) = intIndentation + 1 * conRowHeight
                        End If
                        intWriteCounter2 = intWriteCounter2 + 1
                        'Now multiply out by the third dimension
                        If intColDimensions > 2 Then
                            intWriteCounter3 = 0
                            For intNodeCounter3 = 2 To mdiColSelector.tvwSelection(3).Nodes.Count
                                Set nodDimension3 = mdiColSelector.tvwSelection(3).Nodes(intNodeCounter3)
                                If AddHeader(nodDimension3) = True Then
                                    If intWriteCounter3 > 0 Then
                                        grdReport.Cols = grdReport.Cols + 1
                                    End If
                                    grdReport.Col = grdReport.Cols - 1
                                    grdReport.Row = 2
                                    intIndentation = GetIndentation(nodDimension3)
                                    grdReport.Text = Space$(intIndentation * conIndent) & nodDimension3.Text
                                    If grdReport.RowHeight(2) < (intIndentation + 1 * conRowHeight) Then
                                        grdReport.RowHeight(2) = intIndentation + 1 * conRowHeight
                                    End If
                                    intWriteCounter3 = intWriteCounter3 + 1
                                    'Now multiply out by the fourth dimension
                                    If intColDimensions > 3 Then
                                        intWriteCounter4 = 0
                                        For intNodeCounter4 = 2 To mdiColSelector.tvwSelection(4).Nodes.Count
                                            Set nodDimension4 = mdiColSelector.tvwSelection(4).Nodes(intNodeCounter4)
                                            If AddHeader(nodDimension4) = True Then
                                                If intWriteCounter4 > 0 Then
                                                    grdReport.Cols = grdReport.Cols + 1
                                                End If
                                                grdReport.Col = grdReport.Cols - 1
                                                grdReport.Row = 3
                                                intIndentation = GetIndentation(nodDimension4)
                                                grdReport.Text = Space$(intIndentation * conIndent) & nodDimension4.Text
                                                If grdReport.RowHeight(3) < (intIndentation + 1 * conRowHeight) Then
                                                    grdReport.RowHeight(3) = intIndentation + 1 * conRowHeight
                                                End If
                                                intWriteCounter4 = intWriteCounter4 + 1
                                                'Now multiply out by the fifth dimension
                                                If intColDimensions > 4 Then
                                                    intWriteCounter5 = 0
                                                    For intNodeCounter5 = 2 To mdiColSelector.tvwSelection(5).Nodes.Count
                                                        Set nodDimension5 = mdiColSelector.tvwSelection(5).Nodes(intNodeCounter5)
                                                        If AddHeader(nodDimension5) = True Then
                                                            If intWriteCounter5 > 0 Then
                                                                grdReport.Cols = grdReport.Cols + 1
                                                            End If
                                                            grdReport.Col = grdReport.Cols - 1
                                                            grdReport.Row = 4
                                                            intIndentation = GetIndentation(nodDimension5)
                                                            grdReport.Text = Space$(intIndentation * conIndent) & nodDimension5.Text
                                                            If grdReport.RowHeight(4) < (intIndentation + 1 * conRowHeight) Then
                                                                grdReport.RowHeight(4) = intIndentation + 1 * conRowHeight
                                                            End If
                                                            intWriteCounter5 = intWriteCounter5 + 1
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        End If
    Next
    'Set the height of the 1st row depending on the number of indents
    'Read from mdiRowSelector.tvwSelection(X) into row headers
    'For each node in mdiRowSelector.tvwSelection(1) we do the following
    For intNodeCounter1 = 2 To mdiRowSelector.tvwSelection(1).Nodes.Count
        Set nodDimension1 = mdiRowSelector.tvwSelection(1).Nodes(intNodeCounter1)
        If AddHeader(nodDimension1) = True Then
            grdReport.Rows = grdReport.Rows + 1
            grdReport.Row = grdReport.Rows - 1
            grdReport.Col = 0
            grdReport.Text = Space$(GetIndentation(nodDimension1) * conIndent) & nodDimension1.Text
            'Now multiply out by the second dimension
            If intRowDimensions > 1 Then
                intWriteCounter2 = 0
                For intNodeCounter2 = 2 To mdiRowSelector.tvwSelection(2).Nodes.Count
                    Set nodDimension2 = mdiRowSelector.tvwSelection(2).Nodes(intNodeCounter2)
                    If AddHeader(nodDimension2) = True Then
                        If intWriteCounter2 > 0 Then
                            grdReport.Rows = grdReport.Rows + 1
                        End If
                        grdReport.Row = grdReport.Rows - 1
                        grdReport.Col = 1
                        grdReport.Text = Space$(GetIndentation(nodDimension2) * conIndent) & nodDimension2.Text
                        intWriteCounter2 = intWriteCounter2 + 1
                        'Now multiply out by the third dimension
                        If intRowDimensions > 2 Then
                            intWriteCounter3 = 0
                            For intNodeCounter3 = 2 To mdiRowSelector.tvwSelection(3).Nodes.Count
                                Set nodDimension3 = mdiRowSelector.tvwSelection(3).Nodes(intNodeCounter3)
                                If AddHeader(nodDimension3) = True Then
                                    If intWriteCounter3 > 0 Then
                                        grdReport.Rows = grdReport.Rows + 1
                                    End If
                                    grdReport.Row = grdReport.Rows - 1
                                    grdReport.Col = 2
                                    grdReport.Text = Space$(GetIndentation(nodDimension3) * conIndent) & nodDimension3.Text
                                    intWriteCounter3 = intWriteCounter3 + 1
                                    'Now multiply out by the fourth dimension
                                    If intRowDimensions > 3 Then
                                        intWriteCounter4 = 0
                                        For intNodeCounter4 = 2 To mdiRowSelector.tvwSelection(4).Nodes.Count
                                            Set nodDimension4 = mdiRowSelector.tvwSelection(4).Nodes(intNodeCounter4)
                                            If AddHeader(nodDimension4) = True Then
                                                If intWriteCounter4 > 0 Then
                                                    grdReport.Rows = grdReport.Rows + 1
                                                End If
                                                grdReport.Row = grdReport.Rows - 1
                                                grdReport.Col = 3
                                                grdReport.Text = Space$(GetIndentation(nodDimension4) * conIndent) & nodDimension4.Text
                                                intWriteCounter4 = intWriteCounter4 + 1
                                                'Now multiply out by the fifth dimension
                                                If intRowDimensions > 4 Then
                                                    intWriteCounter5 = 0
                                                    For intNodeCounter5 = 2 To mdiRowSelector.tvwSelection(5).Nodes.Count
                                                        Set nodDimension5 = mdiRowSelector.tvwSelection(5).Nodes(intNodeCounter5)
                                                        If AddHeader(nodDimension5) = True Then
                                                            If intWriteCounter5 > 0 Then
                                                                grdReport.Rows = grdReport.Rows + 1
                                                            End If
                                                            grdReport.Row = grdReport.Rows - 1
                                                            grdReport.Col = 4
                                                            grdReport.Text = Space$(GetIndentation(nodDimension5) * conIndent) & nodDimension5.Text
                                                            intWriteCounter5 = intWriteCounter5 + 1
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        End If
    Next
    'Fill in the rest of the grid?
    'For intRowCounter = intColDimensions To grdReport.Rows - 1
    '    For intColCounter = intRowDimensions To grdReport.Cols - 1
    '        grdReport.Col = intColCounter
    '        grdReport.Row = intRowCounter
    '        grdReport.Text = "Report"
    '    Next
    'Next
    'Set all the column widths
    For intCounter = 0 To grdReport.Cols - 1
        grdReport.ColWidth(intCounter) = 1500
    Next
    'Set the number of fixed cols and rows
    grdReport.FixedCols = intRowDimensions
    grdReport.FixedRows = intColDimensions
    'Position the current row/col
    grdReport.Col = intRowDimensions
    grdReport.Row = intColDimensions
End Sub
Public Function AddHeader(ByRef nodTest As Node) As Boolean
'Created by Robin G Brown, 21/4/97
'Determines wether or not to add this node, called by FillReportGrid
'Function Declares
    'Error Trap
    On Error Resume Next
    'Do we want to show this node in the header?
    AddHeader = False
    If nodTest.Image = "check" Then
        AddHeader = True
    Else
        If nodTest.Parent.Expanded = True Then
            AddHeader = True
        End If
    End If
    'Make sure we never show 'unchecked' nodes
    If nodTest.Image = "uncheck" Then
        AddHeader = False
    End If
End Function

Public Function GetReportTitle() As String
'Created by Robin G Brown, 25/4/97
'Return a report title based on the current cell in the grid
'Function Declares
Dim intCurrentRow As Integer
Dim intcurrentcol As Integer
Dim strReportTitle As String
Dim intSubCounter As Integer
Dim strArray() As String
Dim intArrayCounter As Integer
    'Error Trap
    'On Error Resume Next
    strReportTitle = ""
    If grdReport.FixedRows = 0 Or grdReport.FixedCols = 0 Then
        Exit Function
    End If
    intCurrentRow = grdReport.Row
    intcurrentcol = grdReport.Col
    If intCurrentRow > (grdReport.FixedRows - 1) _
    And intcurrentcol > (grdReport.FixedCols - 1) _
    Then
        grdReport.Row = intCurrentRow
        intArrayCounter = 0
        ReDim strArray(1 To 1)
        For intCounter = grdReport.FixedCols - 1 To 0 Step -1
            grdReport.Col = intCounter
            If grdReport.Text = "" Then
                Do While grdReport.Text = ""
                    grdReport.Row = grdReport.Row - 1
                Loop
                intArrayCounter = intArrayCounter + 1
                ReDim Preserve strArray(1 To intArrayCounter)
                strArray(intArrayCounter) = grdReport.Text
            Else
                intArrayCounter = intArrayCounter + 1
                ReDim Preserve strArray(1 To intArrayCounter)
                strArray(intArrayCounter) = Trim$(grdReport.Text)
            End If
        Next
        For intSubCounter = UBound(strArray, 1) To 1 Step -1
            strReportTitle = strReportTitle & strArray(intSubCounter) & ","
        Next
        strReportTitle = Left$(strReportTitle, Len(strReportTitle) - 1)
        strReportTitle = strReportTitle & vbLf & " by " & vbLf
        grdReport.Col = intcurrentcol
        intArrayCounter = 0
        ReDim strArray(1 To 1)
        For intCounter = grdReport.FixedRows - 1 To 0 Step -1
            grdReport.Row = intCounter
            If grdReport.Text = "" Then
                Do While grdReport.Text = ""
                    grdReport.Col = grdReport.Col - 1
                Loop
                intArrayCounter = intArrayCounter + 1
                ReDim Preserve strArray(1 To intArrayCounter)
                strArray(intArrayCounter) = grdReport.Text
            Else
                intArrayCounter = intArrayCounter + 1
                ReDim Preserve strArray(1 To intArrayCounter)
                strArray(intArrayCounter) = Trim$(grdReport.Text)
            End If
        Next
        For intSubCounter = UBound(strArray, 1) To 1 Step -1
            strReportTitle = strReportTitle & strArray(intSubCounter) & ","
        Next
        strReportTitle = Left$(strReportTitle, Len(strReportTitle) - 1)
    End If
    GetReportTitle = strReportTitle
End Function

Private Sub grdReport_DblClick()
'Created by Robin G Brown, 25/4/97
'Pop up a menu
    'Error Trap
    On Error Resume Next
    'Make sure this is now the active form
    Me.SetFocus
    Me.PopupMenu frmMaster.mnuReport
End Sub

Private Sub grdReport_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'Created by Robin G Brown, 25/4/97
'Pop up a menu if required
    'Error Trap
    On Error Resume Next
    'Now popup a menu if the right button was pressed
    If Button = vbRightButton Then
        'Make sure this is now the active form
        Me.SetFocus
        Me.PopupMenu frmMaster.mnuReport
    End If
End Sub

