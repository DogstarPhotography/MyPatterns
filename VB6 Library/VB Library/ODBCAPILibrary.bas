Attribute VB_Name = "modODBCAPILibrary"
'ODBC Declares
Declare Function SQLAllocEnv Lib "odbc.dll" (env As Long) As Integer
Declare Function SQLFreeEnv Lib "odbc.dll" (ByVal env As Long) As Integer
Declare Function SQLAllocConnect Lib "odbc.dll" (ByVal env As Long, hdbc _
   As Long) As Integer
Declare Function SQLConnect Lib "odbc.dll" (ByVal hdbc As Long, ByVal _
   Server As String, ByVal serverlen As Integer, ByVal uid As String, ByVal _
   uidlen As Integer, ByVal pwd As String, ByVal pwdlen As Integer) As Integer
Declare Function SQLFreeConnect Lib "odbc.dll" (ByVal hdbc As Long) _
   As Integer
Declare Function SQLDisconnect Lib "odbc.dll" (ByVal hdbc As Long) As Integer
Declare Function SQLAllocStmt Lib "odbc.dll" (ByVal hdbc As Long, hstmt As Long) As Integer
Declare Function SQLFreeStmt Lib "odbc.dll" (ByVal hstmt As Long, ByVal EndOption As Integer) As Integer
Declare Function SQLExecDirect Lib "odbc.dll" (ByVal hstmt As Long, ByVal sqlString As String, ByVal sqlstrlen As Long) As Integer
Declare Function SQLNumResultCols Lib "odbc.dll" (ByVal hstmt As Long, NumCols As Integer) As Integer
Declare Function SQLFetch Lib "odbc.dll" (ByVal hstmt As Long) As Integer
Declare Function SQLGetData Lib "odbc.dll" (ByVal hstmt As Long, ByVal Col As Integer, ByVal wConvType As Integer, ByVal lpbBuf As String, ByVal dwbuflen As Long, lpcbout As Long) As Integer
Declare Function SQLError Lib "odbc.dll" (ByVal env As Long, ByVal hdbc As Long, ByVal hstmt As Long, ByVal SQLState As String, NativeError As Long, ByVal Buffer As String, ByVal Buflen As Integer, OutLen As Integer) As Integer
' ODBC Constants
Global Const SQL_SUCCESS = 0
Global Const SQL_SUCCESS_WITH_INFO = 1
Global Const SQL_ERROR = -1
Global Const SQL_NO_DATA_FOUND = 100
Global Const SQL_CLOSE = 0
Global Const SQL_DROP = 1
Global Const SQL_MAX_MESSAGE_LENGTH = 512
Global Const SQL_CHAR = 1
'Global constant for declaring fixed length
'buffer variables:
Global Const gblnBUFFERLEN = 256
'Windows API constants for message boxes, mouse
'pointers:
Global Const MB_ICONEXCLAMATION = 48
Global Const DEFAULTCURSOR = 0
'Global ODBC environment, database connection,
'and statement handles:
Global gblHenv As Long
Global gblHdbc As Long
Global gblHstmt As Long
  
  

Function AllocateODBChEnv(hEnv As Long)
'Listing 1. Declare Yourself. You'll need these declares and constant 'declarations for the foundation of your ODBC API work. Place this 'code inside the declarations section of a BAS module in your project.
   Dim Result As Integer
   AllocateODBChEnv = SQL_SUCCESS
   Result = SQLAllocEnv(hEnv)
   If Result <> SQL_SUCCESS Then
      MsgBox "Cannot allocate environment handle.", MB_ICONEXCLAMATION, "ODBC Error"
      Screen.MousePointer = DEFAULTCURSOR
      AllocateODBChEnv = Result
      Exit Function
   End If
End Function

  
  

Function ConnectToDatasource(hEnv, hdbc As Long, hstmt As Long, DataSource As String, UserId As String, Password As String) As Integer
'Listing 2. Get a Handle on the Environment. Function 'AllocateODBChEnv() allocates an ODBC environment handle. It returns 'either SQL_SUCCESS or an error code from SQLAllocEnv().
'henv parameter must be a valid ODBC environment 'handle.
' hdbc and hstmt parameters are set by this function.
'Function returns SQL_SUCCESS or error code returned from SQLAllocConnect,
' SQLConnect, or SQLAllocStmt
   Dim Result As Integer
   ConnectToDatasource = SQL_SUCCESS
   Result = SQLAllocConnect(hEnv, hdbc)
   If Result <> SQL_SUCCESS Then
      MsgBox "Cannot allocate connection handle.", MB_ICONEXCLAMATION, "ODBC Error"
      Screen.MousePointer = DEFAULTCURSOR
      ConnectToDatasource = Result
      Exit Function
   End If
   Result = SQLConnect(hdbc, DataSource, Len(DataSource), UserId, Len(UserId), Password, Len(Password))
   If Result <> SQL_SUCCESS Then
      MsgBox "Cannot establish datasource connnection.", MB_ICONEXCLAMATION, "ODBC Error"
      Screen.MousePointer = DEFAULTCURSOR
      ConnectToDatasource = Result
      Exit Function
   End If
   Result = SQLAllocStmt(hdbc, hstmt)
   If Result <> SQL_SUCCESS Then
      MsgBox "Cannot allocate statement handle.", MB_ICONEXCLAMATION, "ODBC Error"
      Screen.MousePointer = DEFAULTCURSOR
      ConnectToDatasource = Result
      Exit Function
   End If
End Function

  
  

Function DisconnectFromDatasource(hdbc As Long, hstmt As Long) As Integer
'Listing 3. Make the Connection. ConnectToDatasource() takes an 'environment handle, a statement handle, and data source parameters 'and establishes an ODBC connection. Call this function after calling 'AllocateODBChEnv().
'Function Returns False if any API call fails; True otherwise
   Dim Result As Integer
   DisconnectFromDatasource = True
   If hstmt <> 0 Then
      Result = SQLFreeStmt(hstmt, SQL_DROP)
      If Result <> SQL_SUCCESS Then DisconnectFromDatasource = False
   End If
   If hdbc <> 0 Then
      Result = SQLDisconnect(hdbc)
      If Result <> SQL_SUCCESS Then DisconnectFromDatasource = False
   End If
   If hdbc <> 0 Then
      Result = SQLFreeConnect(hdbc)
      If Result <> SQL_SUCCESS Then DisconnectFromDatasource = False
   End If
End Function

  
  

Function FreeODBChEnv(hEnv As Long) As Integer
'Listing 4. Break the Connection. After DisconnectFromDatasource() 'disconnects, it frees connection and environment handles. You must 'error-check for each step.
'Returns False if unsuccessful; True otherwise
   Dim Result As Integer
   FreeODBChEnv = True
   If hEnv <> 0 Then
      Result = SQLFreeEnv(hEnv)
      If Result <> SQL_SUCCESS Then FreeODBChEnv = False
   End If
End Function

  
  

Sub DisplayODBCError(hdbc As Long, hstmt As Long, WindowCaption As String)
'Listing 5. Free Your Handle. FreeODBChEnv frees the previously 'allocated environment handle. You should call it after 'DisconnectFromDatasource().
'Displays one or more message boxes with error
'message for the given connection and statement
'handles
   Dim SQLState As String * 16
   Dim ErrorMsg As String * SQL_MAX_MESSAGE_LENGTH
   Dim ErrMsgSize As Integer
   Dim ErrorCode As Long
   Dim ErrorCodeStr As String
   Dim Result As Integer
   SQLState = String$(16, 0)
   ErrorMsg = String$(SQL_MAX_MESSAGE_LENGTH - 1, 0)
   Do
      Result = SQLError(0, hdbc, hstmt, SQLState, ErrorCode, ErrorMsg, Len(ErrorMsg), ErrMsgSize)
      Screen.MousePointer = DEFAULTCURSOR
      If Result = SQL_SUCCESS Or Result = SQL_SUCCESS_WITH_INFO Then
         If ErrMsgSize = 0 Then
            MsgBox "SQL_SUCCESS Or " & _
               "SQL_SUCCESS_WITH_INFO Error No additional information available.", MB_ICONEXCLAMATION, WindowCaption
         Else
            If ErrorCode = 0 Then
               ErrorCodeStr = ""
            Else
               ErrorCodeStr = Trim$(Str(ErrorCode)) & ""
            End If
            MsgBox ErrorCodeStr & Left$(ErrorMsg, ErrMsgSize), MB_ICONEXCLAMATION, WindowCaption
         End If
      End If
   Loop Until Result <> SQL_SUCCESS
End Sub

  
  

Function LoadControl(CtlName As Control, Query As String, hstmt As Long, ItemDataFill As Integer, Delimeter As String) As Integer
'Listing 6. Display Error Messages. DisplayODBCError() is a wrapper 'function that displays ODBC error messages. Note the Do . . . Loop 'construct used to display multiple messages resulting from one SQL 'statement.
   Dim Result As Integer
   Dim Temp As Integer
   Dim RowCnt As Integer
   Dim NumCols As Integer
   Dim ColCnt As Integer
   Dim Buffer As String * gblnBUFFERLEN
   Dim ItemText As String
   Dim ItemDataString As String
   Dim OutLen As Long
   LoadControl = SQL_SUCCESS
   ' Make sure referenced control is a listbox,
   'combobox, or grid:
   If TypeOf CtlName Is ListBox Then
   ElseIf TypeOf CtlName Is ComboBox Then
   ' Un-comment next line if grid.vbx in project
   'ElseIf TypeOf CtlName Is Grid Then
   Else
      LoadControl = -3
      Exit Function
   End If
   ' Do the query:
   Result = SQLExecDirect(hstmt, Query, Len(Query))
   If Result <> SQL_SUCCESS Then
      Call DisplayODBCError(gblHdbc, gblHstmt, "SQL          Statement Error During LoadControl")
      Temp = SQLFreeStmt(hstmt, SQL_CLOSE)
      If (Temp <> SQL_SUCCESS) Then
         Screen.MousePointer = DEFAULTCURSOR
         MsgBox "Cannot free statement handle.", MB_ICONEXCLAMATION, "Unexpected ODBC             Driver Function failure"
      End If
      LoadControl = Result
      Exit Function
   End If
   'Check number of columns returned:
   Result = SQLNumResultCols(hstmt, NumCols)
   If Result <> SQL_SUCCESS Then
      Temp = SQLFreeStmt(hstmt, SQL_CLOSE)
      Screen.MousePointer = DEFAULTCURSOR
      LoadControl = Result
      Exit Function
   End If
   'Set return value to SQL_NO_DATA_FOUND if no rows
   'returned:
   If NumCols = 0 Then
      Temp = SQLFreeStmt(hstmt, SQL_CLOSE)
      Screen.MousePointer = DEFAULTCURSOR
      LoadControl = SQL_NO_DATA_FOUND
      Exit Function
   End If
   'Un-comment next three lines if grid.vbx in
   'project
   'If TypeOf CtlName Is Grid Then
   '   CtlName.Cols = NumCols + 1
   '   CtlName.Rows = 2
   'End If
   Buffer = String$(gblnBUFFERLEN, 0)
   'Fill in the control
   RowCnt = 0
   Do
      ' Get next row
      Result = SQLFetch(hstmt)
      If Result <> SQL_SUCCESS Then
         If Result = SQL_NO_DATA_FOUND Then
            Temp = SQLFreeStmt(hstmt, SQL_CLOSE)
            Screen.MousePointer = DEFAULTCURSOR
            If RowCnt > 0 Then
               Exit Do
            Else
               LoadControl = Result
               Exit Function
            End If
         Else
            Temp = SQLFreeStmt(hstmt, SQL_CLOSE)
            Screen.MousePointer = DEFAULTCURSOR
            LoadControl = Result
            Exit Function
         End If
      End If
      RowCnt = RowCnt + 1
      ' Un-comment next three lines if using grid.vbx
      'If TypeOf CtlName Is Grid Then
      '   CtlName.Row = RowCnt
      'End If
      ItemText = ""
      ItemDataString = ""
      ' Get each column.
      ' For grids, place column value in control
      'column.
      ' For lists and combo boxes, concatenate all
      'column values for row and place the
      'concatenated string into control.
      For ColCnt = 1 To NumCols
         Result = SQLGetData(hstmt, ColCnt, SQL_CHAR, Buffer, gblnBUFFERLEN, OutLen) ', "Call to             SQLGetData Failed"
         If Result <> SQL_SUCCESS Then
            Temp = SQLFreeStmt(hstmt, SQL_CLOSE)
            Screen.MousePointer = DEFAULTCURSOR
            LoadControl = Result
            Exit Function
         End If
         'Un-comment next seven lines and
         'End If further down if using grid.vbx
         'If TypeOf CtlName Is Grid Then
         '   CtlName.Col = ColCnt
         '   If OutLen > 0 Then
         '      CtlName.Text = Left$(Buffer, OutLen)
         '      ' Add code here to set column width (if
               ' desired)
         '   End If
         'Else
            If ItemDataFill And ColCnt = 1 Then
               If OutLen > 0 Then
                  ItemDataString = Left$(Buffer, OutLen)
               Else
                  ItemDataString = ""
               End If
            Else
               If OutLen > 0 Then
                  If ItemText = "" Then
                     ItemText = Left$(Buffer, OutLen)
                  Else
                     ItemText = ItemText & Delimeter & Left$(Buffer, OutLen)
                  End If
               Else
                  ItemText = ItemText & Delimeter
               End If
            End If
         'End If
      Next ColCnt
      ' For combos and lists, do the AddItem, and set
      'ItemData if ItemDataFill = True:
      ' (If we're loading a grid, szTempStr = "")
      If ItemText <> "" Then
         On Error Resume Next
         CtlName.AddItem ItemText
         If Err = 0 Then
            If ItemDataString <> "" Then
               CtlName.ItemData(CtlName.NewIndex) = Val(ItemDataString)
            End If
         Else
            MsgBox "Result Set too large to fit in                control", MB_ICONEXCLAMATION, "LoadControl"
            Exit Do
         End If
         On Error GoTo 0
      End If
      ' Un-comment next three lines if grid.vbx in
      'project
      'If TypeOf CtlName Is Grid Then
      '   CtlName.Rows = CtlName.Rows + 1
      'End If
   Loop
   LoadControl = SQL_SUCCESS
   Screen.MousePointer = DEFAULTCURSOR
End Function

  
  

Function GetNthField(Control As Control, FieldNumber As Integer, Delimeter As String) As String
'Listing 7. Fill In the Data. LoadControl() can fill a list box, combo box, 'or grid with data. Note that you must comment out the grid code to 'avoid runtime errors for projects that don't contain GRID.VBX.
   Dim Count As Integer
   Dim DelimeterPos As Integer
   Dim SecondDelimPos As Integer
   Dim SearchStr As String
   GetNthField = ""
   SearchStr = Delimeter & _
      Control.List(Control.ListIndex) & Delimeter
   If TypeOf Control Is ListBox Then
   Else
      If TypeOf Control Is ComboBox Then
      Else
         Exit Function
      End If
   End If
   DelimeterPos = 0
   For Count = 1 To FieldNumber
      DelimeterPos = InStr(DelimeterPos + 1, SearchStr, Delimeter)
      If DelimeterPos = 0 Then
         DelimeterPos = -1
         Exit For
      End If
   Next Count
   If DelimeterPos <> -1 Then
      SecondDelimPos = InStr(DelimeterPos + 1, SearchStr, Delimeter)
   End If
   If SecondDelimPos > DelimeterPos And SecondDelimPos <> 0 Then
      GetNthField = Mid$(SearchStr, DelimeterPos + 1, SecondDelimPos - (DelimeterPos + 1))
   End If
End Function
