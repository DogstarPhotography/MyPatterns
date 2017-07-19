Attribute VB_Name = "modHTMLLib"
'Created by Robin G Brown, 25/7/97
'-----------------------------------------------------------------------------
'   This module contains code for _
    creating, reading and handling HTML in programs
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modHTMLLib"
Private lngCounter As Long
Private lngReturn As Long

Public Function BoilerInsertTAG() As String
'Created by Robin G Brown, INSERT_DATE
'Insert TAG
'Function Declares
    'Error Trap
    On Error Resume Next
    BoilerInsertTAG = "<TAG>" & vbCrLf
End Function

Public Function BoilerWrapTAG(ByVal strContents As String) As String
'Created by Robin G Brown, INSERT_DATE
'Insert TAG
'Function Declares
    'Error Trap
    On Error Resume Next
    BoilerWrapTAG = "<TAG>" & vbCrLf & strContents & "</TAG>" & vbCrLf
End Function

Public Function CreateTableFromArray(ByRef strTableArray() As String) As String
'Created by Robin G Brown, 25/7/97
'Create a simple HTML table from a 1 or 2 dimensional array, _
    the 1st dimension should be rows and the 2nd should be columns
'Function Declares
Dim intRowCounter As Integer
Dim intCellCounter As Integer
Dim strHTML As String
    'Error Trap
    On Error Resume Next
    intCellCounter = UBound(strTableArray, 2)
    strHTML = "<TABLE>" & vbCrLf
    If Err.Number <> 0 Then
        '1 dimension only
        For intRowCounter = 1 To UBound(strTableArray, 1)
            strHTML = strHTML & WrapTableRow(WrapTableCell(strTableArray(intRowCounter)))
        Next
    Else
        For intRowCounter = 1 To UBound(strTableArray, 1)
            strHTML = strHTML & "<TR>"
            For intCellCounter = 1 To UBound(strTableArray, 2)
                strHTML = strHTML & WrapTableCell(strTableArray(intRowCounter, intCellCounter))
            Next
            strHTML = strHTML & "</TR>"
        Next
    End If
    strHTML = strHTML & "</TABLE>" & vbCrLf
    CreateTableFromArray = strHTML
End Function

Public Function WrapHTML(ByVal strContents As String) As String
'Created by Robin G Brown, 25/7/97
'Insert HTML
'Function Declares
    'Error Trap
    On Error Resume Next
    WrapHTML = "<HTML>" & vbCrLf & strContents & "</HTML>" & vbCrLf
End Function

Public Function WrapBODY(ByVal strContents As String) As String
'Created by Robin G Brown, 25/7/97
'Insert BODY
'Function Declares
    'Error Trap
    On Error Resume Next
    WrapBODY = "<BODY>" & vbCrLf & strContents & "</BODY>" & vbCrLf
End Function

Public Function WrapTable(ByVal strContents As String) As String
'Created by Robin G Brown, 25/7/97
'Insert TABLE
'Function Declares
    'Error Trap
    On Error Resume Next
    WrapTable = "<TABLE>" & vbCrLf & strContents & "</TABLE>" & vbCrLf
End Function

Public Function WrapTableRow(ByVal strContents As String) As String
'Created by Robin G Brown, 25/7/97
'Insert TR
'Function Declares
    'Error Trap
    On Error Resume Next
    WrapTableRow = "<TR>" & vbCrLf & strContents & "</TR>" & vbCrLf
End Function

Public Function WrapTableCell(ByVal strContents As String) As String
'Created by Robin G Brown, 25/7/97
'Insert TD
'Function Declares
    'Error Trap
    On Error Resume Next
    WrapTableCell = "<TD>" & vbCrLf & strContents & "</TD>" & vbCrLf
End Function

Public Function WrapParagraph(ByVal strContents As String) As String
'Created by Robin G Brown, 25/7/97
'Insert TAG
'Function Declares
    'Error Trap
    On Error Resume Next
    WrapParagraph = "<P>" & vbCrLf & strContents & "</P>" & vbCrLf
End Function

Public Function WrapIndent(ByVal strContents As String) As String
'Created by Robin G Brown, 25/7/97
'Insert BLOCKQUOTE
'Function Declares
    'Error Trap
    On Error Resume Next
    WrapIndent = "<BLOCKQUOTE>" & vbCrLf & strContents & "</BLOCKQUOTE>" & vbCrLf
End Function

Public Function InsertRule() As String
'Created by Robin G Brown, 25/7/97
'Insert TAG
'Function Declares
    'Error Trap
    On Error Resume Next
    InsertRule = "<HR>" & vbCrLf
End Function

Public Function InsertAddress(ByVal strLinkText As String, ByVal strURL As String) As String
'Created by Robin G Brown, 25/7/97
'Insert an address
'Function Declares
    'Error Trap
    On Error Resume Next
    InsertAddress = "<A HREF=""" & strURL & """>" & strLinkText & "</A>"
End Function

Public Sub ConvertToTag(ByRef strTag As String)
'Created by Robin G Brown, 25/7/97
'Convert a string to a Tag by adding < and > to either end
'Sub Declares
    'Error Trap
    On Error Resume Next
    strTag = "<" & strTag & ">"
End Sub

Public Sub ConvertToEndTag(ByRef strTag As String)
'Created by Robin G Brown, 25/7/97
'Convert a string to a Tag by adding < and > to either end
'Sub Declares
    'Error Trap
    On Error Resume Next
    strTag = "</" & strTag & ">"
End Sub


