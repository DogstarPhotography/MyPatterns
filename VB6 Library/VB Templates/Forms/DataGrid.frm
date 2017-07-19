VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDataGrid 
   ClientHeight    =   4590
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleMode       =   0  'User
   ScaleWidth      =   6150
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   6150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6150
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   330
         Left            =   4398
         TabIndex        =   5
         Tag             =   "&Close"
         Top             =   0
         Width           =   1437
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Filter"
         Height          =   330
         Left            =   2924
         TabIndex        =   4
         Tag             =   "&Filter"
         Top             =   0
         Width           =   1462
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "&Sort"
         Height          =   330
         Left            =   1462
         TabIndex        =   3
         Tag             =   "&Sort"
         Top             =   0
         Width           =   1462
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   330
         Left            =   0
         TabIndex        =   2
         Tag             =   "&Refresh"
         Top             =   0
         Width           =   1462
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2505
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2175
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid grdDataGrid 
      Align           =   1  'Align Top
      Bindings        =   "DataGrid.frx":0000
      Height          =   3645
      Left            =   0
      OleObjectBlob   =   "DataGrid.frx":00CE
      TabIndex        =   0
      Top             =   330
      Width           =   6150
   End
End
Attribute VB_Name = "frmDataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim msSortCol As String
Dim mbCtrlKey As Integer


Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFilter_Click()
    On Error GoTo FilterErr

    Dim recRecordset1 As Recordset, recRecordset2 As Recordset
    Dim sFilterStr As String

    If Data1.RecordsetType = vbRSTypeTable Then
        Beep
        MsgBox "You Cannot Filter a Table Recordset!", 48
        Exit Sub
    End If
    
    Set recRecordset1 = Data1.Recordset                        'copy the recordset
    
    sFilterStr = InputBox("Enter Filter Expression:")
    If Len(sFilterStr) = 0 Then Exit Sub

    Screen.MousePointer = vbHourglass
    recRecordset1.Filter = sFilterStr
    Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type) 'establish the filter
    Set Data1.Recordset = recRecordset2                        'assign back to original recordset object

    Screen.MousePointer = vbDefault
    Exit Sub

FilterErr:
    Screen.MousePointer = vbDefault
    MsgBox "Error:" & Err & " " & Err.Description
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo RefErr
    
    Data1.Recordset.Requery
    
    Exit Sub
    
RefErr:
    MsgBox "Error:" & Err & " " & Err.Description
End Sub


Private Sub cmdSort_Click()
    On Error GoTo SortErr

    Dim recRecordset1 As Recordset, recRecordset2 As Recordset
    Dim SortStr As String

    If Data1.RecordsetType = vbRSTypeTable Then
        Beep
        MsgBox "You Cannot Sort a Table Recordset!", 48
        Exit Sub
    End If

    Set recRecordset1 = Data1.Recordset                        'copy the recordset
    
    If Len(msSortCol) = 0 Then
        SortStr = InputBox("Enter Sort Column:")
        If Len(SortStr) = 0 Then Exit Sub
    Else
        SortStr = msSortCol
    End If

    Screen.MousePointer = vbHourglass
    recRecordset1.Sort = SortStr
    
    'establish the Sort
    Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type)
    Set Data1.Recordset = recRecordset2
    
    Screen.MousePointer = vbDefault
    Exit Sub

SortErr:
    Screen.MousePointer = vbDefault
    MsgBox "Error:" & Err & " " & Err.Description
End Sub


Private Sub Form_Load()
    Dim bParmQry As Integer
    Dim qdfTmp As QueryDef
    
    On Error GoTo LoadErr

    'To Do
    'gsDatabase is a global string that needs
    'to be set by the startup sub for the app
    Data1.DatabaseName = gsDatabase
    'gsRecordSource is a global string that needs
    'to be set by the sub routine that loads this form
    Data1.RecordSource = gsRecordsource
    Data1.RecordsetType = 1     'dynaset
    Data1.Options = 0
    Data1.Refresh

    If Len(Data1.RecordSource) > 50 Then
        Me.Caption = "SQL Statement"
    Else
        Me.Caption = Data1.RecordSource
    End If

    Exit Sub

LoadErr:
    MsgBox "Error:" & Err & " " & Err.Description
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> 1 Then
        grdDataGrid.Height = Me.Height - (425 + picButtons.Height)
    End If
End Sub

Private Sub grdDataGrid_BeforeDelete(Cancel As Integer)
    If MsgBox("Delete Current Row?", vbYesNo + vbQuestion) <> vbYes Then
        Cancel = True
    End If
End Sub

Private Sub grdDataGrid_BeforeUpdate(Cancel As Integer)
    If MsgBox("Commit changes?", vbYesNo + vbQuestion) <> vbYes Then
        Cancel = True
    End If
End Sub

Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
    'let's sort on this column
    If Data1.RecordsetType = vbRSTypeTable Then Exit Sub
    
    'check for the use of the ctrl key for descending sort
    If mbCtrlKey Then
        msSortCol = "[" & Data1.Recordset(ColIndex).Name & "] desc"
        mbCtrlKey = 0 'reset it
    Else
        msSortCol = "[" & Data1.Recordset(ColIndex).Name & "]"
    End If
    cmdSort_Click
    msSortCol = vbNullString 'reset it
    
End Sub

Private Sub grdDataGrid_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    mbCtrlKey = Shift
End Sub

