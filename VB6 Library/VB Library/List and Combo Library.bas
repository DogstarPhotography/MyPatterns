Attribute VB_Name = "modListComboLib"
'Created by Robin G Brown, 8th July 1997
'-----------------------------------------------------------------------------
'   This module contains code for _
    handling list and combo controls
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modListComboLib"
Private intCounter As Integer
Private intReturn As Integer

Public Function FillListOrComboFromRecordset(dynaSource As Recordset, strField As String, ctlTarget As Control) As Integer
'----------------------------------------------------------------
'this function fills a list or combo box with the contents of a field from a recordset
'dynaSource, strField, ctlTarget
'----------------------------------------------------------------
'RGB
    On Error GoTo FillListOrComboFromRecordset_ErrorHandler
    dynaSource.MoveFirst
    If TypeOf ctlTarget Is ListBox Then
        'listbox
        Do While Not dynaSource.EOF
            ctlTarget.AddItem dynaSource.Fields(strField).Value
            dynaSource.MoveNext
        Loop
        FillListOrComboFromRecordset = True
    ElseIf TypeOf ctlTarget Is ComboBox Then
        'combobox
        Do While Not dynaSource.EOF
            ctlTarget.AddItem dynaSource.Fields(strField).Value
            dynaSource.MoveNext
        Loop
        FillListOrComboFromRecordset = True
    Else
        'not a list or combo box so ignore it
        FillListOrComboFromRecordset = False
    End If
Exit Function
FillListOrComboFromRecordset_ErrorHandler:
    FillListOrComboFromRecordset = False
    Exit Function
End Function

