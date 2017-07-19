VERSION 5.00
Begin VB.Form frmQuerys 
   Caption         =   "Querys"
   ClientHeight    =   4185
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5100
   Tag             =   "Querys"
   Begin VB.ListBox lstQueryDefs 
      Height          =   1260
      Left            =   96
      TabIndex        =   0
      Top             =   274
      Width           =   3392
   End
   Begin VB.TextBox txtSQLStatement 
      BackColor       =   &H00FFFFFF&
      Height          =   2159
      Left            =   96
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1921
      Width           =   4931
   End
   Begin VB.CommandButton cmdRemoveQuery 
      Caption         =   "&Remove"
      Height          =   370
      Left            =   3572
      TabIndex        =   3
      Tag             =   "&Remove"
      Top             =   1277
      Width           =   1443
   End
   Begin VB.CommandButton cmdSaveQueryDef 
      Caption         =   "&Save"
      Height          =   370
      Left            =   3572
      TabIndex        =   2
      Tag             =   "&Save"
      Top             =   775
      Width           =   1443
   End
   Begin VB.CommandButton cmdExecuteSQL 
      Caption         =   "&Execute"
      Enabled         =   0   'False
      Height          =   370
      Left            =   3572
      TabIndex        =   1
      Tag             =   "&Execute"
      Top             =   274
      Width           =   1443
   End
   Begin VB.Label lblSQL 
      Caption         =   "SQL Statement:"
      Height          =   251
      Index           =   1
      Left            =   132
      TabIndex        =   6
      Tag             =   "SQL Statement:"
      Top             =   1682
      Width           =   2189
   End
   Begin VB.Label lblSQL 
      Caption         =   "Saved Querys:"
      Height          =   251
      Index           =   0
      Left            =   108
      TabIndex        =   5
      Tag             =   "Saved Querys:"
      Top             =   24
      Width           =   2189
   End
End
Attribute VB_Name = "frmQuerys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================================================================
'this template requires the following code (or it's equivalent)
'to be present in the application as well as a reference to DAO 3.50
'and the DataGrid template
'
'Global gsDatabase As String
'Global gsRecordsource As String
'
'Sub Main()
'  gsDatabase = "c:\vb5\biblio.mdb"
'  frmQuerys.Show
'End Sub
'====================================================================


Dim mdbDatabase As Database
Private Sub Form_Load()
    Set mdbDatabase = OpenDatabase(gsDatabase)
    RefreshQuerys
    Me.Left = GetSetting(App.Title, "Settings", "QueryLeft", 0)
    Me.Top = GetSetting(App.Title, "Settings", "QueryTop", 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "QueryLeft", Me.Left
        SaveSetting App.Title, "Settings", "QueryTop", Me.Top
    End If
End Sub



Private Sub cmdSaveQueryDef_Click()
    On Error GoTo SQDErr

    Dim sQueryName As String
    Dim sTmp As String
    Dim qdNew As QueryDef

    If lstQueryDefs.ListIndex >= 0 Then
        'a querydef is selected so user may want to update it's SQL
        If MsgBox("Update '" & lstQueryDefs.Text & "'?", vbYesNo + vbQuestion) = vbYes Then
            'store the SQL from the SQL Window in the currently
            'selected querydef
            mdbDatabase.QueryDefs(lstQueryDefs.Text).SQL = Me.txtSQLStatement.Text
            Exit Sub
        End If
    End If
    
    'either there is no current querydef selected or the user
    'didn't want to update the current one so we need
    'to propmpt for a new name
    sQueryName = InputBox("Enter New Query Name:")
    If Len(sQueryName) = 0 Then Exit Sub
    
    'add the new querydef
    Set qdNew = mdbDatabase.CreateQueryDef(sQueryName)
    'prompt for passthrough querydef
    If MsgBox("Is this a SQLPassThrough QueryDef?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        sTmp = InputBox("Enter Connect property value:")
        If Len(sTmp) > 0 Then
            qdNew.Connect = sTmp
            If MsgBox("Is the Query Row Returning?", vbYesNo + vbQuestion) = vbNo Then
                qdNew.ReturnsRecords = False
            End If
        End If
    End If
    qdNew.SQL = txtSQLStatement.Text
    
    mdbDatabase.QueryDefs.Refresh
    RefreshQuerys

    Exit Sub

SQDErr:
    MsgBox Err.Description
End Sub

Private Sub lstQueryDefs_Click()
    txtSQLStatement.Text = mdbDatabase.QueryDefs(lstQueryDefs.Text).SQL
End Sub

Private Sub lstQueryDefs_DblClick()
    cmdExecuteSQL_Click
End Sub

Private Sub txtSQLStatement_Change()
    If Len(txtSQLStatement.Text) > 0 Then
        cmdExecuteSQL.Enabled = True
    Else
        cmdExecuteSQL.Enabled = False
    End If
End Sub

Private Sub cmdExecuteSQL_Click()
    Dim rsTmp As Recordset
    Dim dbTmp As Database
    Dim qdfTmp As QueryDef
    Dim bSavedQDF As Boolean
    Dim sSQL As String
    
    If Len(txtSQLStatement.Text) = 0 Then Exit Sub
    
    Set dbTmp = OpenDatabase(gsDatabase)
    
    If lstQueryDefs.ListIndex >= 0 Then
        sSQL = dbTmp.QueryDefs(lstQueryDefs.Text).SQL
        If sSQL = txtSQLStatement.Text Then
            Set qdfTmp = dbTmp.QueryDefs(lstQueryDefs.Text)
            bSavedQDF = True
            If Not SetQryParams(qdfTmp) Then Exit Sub
        Else
            'just create a temp querydef
            Set qdfTmp = dbTmp.CreateQueryDef(vbNullString, txtSQLStatement.Text)
        End If
    Else
        'just create a temp querydef
        Set qdfTmp = dbTmp.CreateQueryDef(vbNullString, txtSQLStatement.Text)
    End If

    Screen.MousePointer = vbHourglass
    If UCase(Mid(txtSQLStatement, 1, 6)) = "SELECT" And InStr(UCase(txtSQLStatement.Text), " INTO ") = 0 Then
        On Error GoTo SQLErr
MakeDynaset:
        Dim f As New frmDataGrid
        Set rsTmp = qdfTmp.OpenRecordset()
        Set f.Data1.Recordset = rsTmp
        If bSavedQDF Then
            f.Caption = qdfTmp.Name
        Else
            f.Caption = Left(txtSQLStatement.Text, 32) & "..."
        End If
        f.Show
    Else
        On Error GoTo SQLErr
        qdfTmp.Execute
    End If

    Screen.MousePointer = vbDefault
    Exit Sub

SQLErr:
    If Err = 3065 Or Err = 3078 Then 'row returning or name not found so try to create recordset
        Resume MakeDynaset
    End If
    MsgBox Err.Description

SQLEnd:

End Sub

Private Sub Form_Resize()
    On Error Resume Next

    If WindowState <> 1 Then
        If Me.Width < 5220 Then Me.Width = 5220
        If Me.Height < 2784 Then Me.Height = 2784
        
        txtSQLStatement.Width = Me.Width - 320
        txtSQLStatement.Height = Me.Height - 2424
    End If
End Sub

Sub RefreshQuerys()
    Dim qdf As QueryDef
    
    lstQueryDefs.Clear
    
    For Each qdf In mdbDatabase.QueryDefs
        lstQueryDefs.AddItem qdf.Name
    Next
    
End Sub

Private Function SetQryParams(rqdf As QueryDef) As Boolean
    On Error GoTo SPErr
    
    Dim prm As Parameter
    Dim sTmp As String
    Dim i As Integer
    
    For Each prm In rqdf.Parameters
        'get the value from the user
        sTmp = InputBox("Enter Value for Parameter '" & prm.Name & "':")
        If Len(sTmp) = 0 Then
            'bail out if the user doesn't enter one of the params
            SetQryParams = False
            Exit Function
        End If
        'store the value
        prm.Value = CVar(sTmp)
    Next
    
    SetQryParams = True
    Exit Function
        
SPErr:
    MsgBox Err.Description
End Function

