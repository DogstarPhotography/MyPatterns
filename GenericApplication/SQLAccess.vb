' Amendment History:
' 19/10/2010 JG  Created.
' 05/02/2010 AP  Added bulk copy support.
' 12/07/2010 JG  Added Execute() method to run stored procedures/scripts that return no result.
' 28/07/2010 JG  Added support for LastReturnCode property to hold RETURN code after all data operations.
' 20/07/2010 JG  Support for output parameters.

Imports System.Data.SqlClient

''' <summary>
''' SQLAccess: To provide a common interface for the reading and writing of SQL data.
'''
''' Data fetch:        GetDatatable(), GetDataTableIntoDataset(), GetResult(), GetScalar();
''' Data modification: Insert(), Update(), Delete();
''' Transactional:     BeginTransaction(), CommitTransaction(), RollbackTransaction().
''' Bulk Copy:         CopyDataWithColumnMapping().
''' Logging:           Optional, only to text files currently.
'''
''' See also Sirius.IO.ConnectionManager.
''' </summary>
''' <remarks></remarks>
Public Class SQLAccess
#Region " Private global variables "
    Private _Connection As SqlConnection           ' Used for non-transactional operations.
    Private _Name As String                        ' A useful identifier for the SQLAccess instance.
    Private _TransactionID As String = ""          ' If not "", then a transaction is in progress.
    Private _ErrorLog As String = ""               ' For logging of errors/operations.

    Private _LogOperations As Boolean = False      ' Logs all SQL operations.
    Private _LogErrors As Boolean = True           ' Logs SQL errors.
    Private _LastReturnCode As Integer = 0         ' The last RETURN code from the SQL operation.
    Private _LastError As Exception = Nothing      ' Retains the last SQL error until next operations.
    Private _CrashOnError As Boolean = True        ' Performs Environment.Exit() if SQL operation fails.
    Private _NotifyBulkProgressCount As Integer = 1000 ' Number of rows before Bulk copy raises a notification event.

    ' Separate adapter and connection for transactions, allowing interleaving 
    ' of non-transactional reads while a transaction is already begun.
    Private _Transaction As SqlTransaction
    Private _TransactionConnection As SqlConnection
    Private Const DefaultTimeOut As Integer = 30
    Private _CurrentTimeOut As Integer = DefaultTimeOut
#End Region

#Region " Properties "

    Public Property TimeOut() As Integer
        Get
            Return _CurrentTimeOut
        End Get
        Set(ByVal value As Integer)
            _CurrentTimeOut = value
        End Set
    End Property
    Public Sub ResetCurrentTimeOut()
        _CurrentTimeOut = DefaultTimeOut
    End Sub

    ''' <summary>
    '''  The connection string to be used when connecting to the SQL database.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ConnectionString() As String
        Get
            Return _Connection.ConnectionString
        End Get
    End Property

    ''' <summary>
    ''' A suitable identifying name for the connection eg, "Common" which will be used for logging etc.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Name() As String
        Get
            Return _Name
        End Get
    End Property

    Public ReadOnly Property TransactionID() As String
        Get
            Return _TransactionID
        End Get
    End Property

    ''' <summary>
    ''' The integer expression returned by the SQL RETURN statement of the last operation that was performed.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property LastReturnCode() As Integer
        Get
            Return _LastReturnCode
        End Get
    End Property

    ''' <summary>
    ''' If true, logs all SQL access operations to the logfile.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LogOperations() As Boolean
        Get
            Return _LogOperations
        End Get
        Set(ByVal value As Boolean)
            _LogOperations = value
        End Set
    End Property

    ''' <summary>
    ''' If true, logs all errors (and exceptions) to the logfile.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LogErrors() As Boolean
        Get
            Return _LogErrors
        End Get
        Set(ByVal value As Boolean)
            _LogErrors = True
        End Set
    End Property

    ''' <summary>
    ''' If true then the application is closed, otherwse control is passed back to the caller.
    ''' The caller should test the property LastError.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CrashOnError() As Boolean
        Get
            Return _CrashOnError
        End Get
        Set(ByVal value As Boolean)
            _CrashOnError = value
        End Set
    End Property

    ''' <summary>
    ''' The value of the last exception to have occurred.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property LastSQLError() As Exception
        Get
            Return _LastError
        End Get
    End Property

    Public ReadOnly Property ErrorLog() As String
        Get
            Return _ErrorLog
        End Get
    End Property

    Public Property NotifyBulkProgressCount() As Integer
        Get
            Return _NotifyBulkProgressCount
        End Get
        Set(ByVal value As Integer)
            _NotifyBulkProgressCount = value
        End Set
    End Property
#End Region

#Region " Constructors "
    Public Sub New(ByVal Name As String, ByVal ConnectionString As String)
        _Connection = New SqlConnection(ConnectionString)
        _TransactionConnection = New SqlConnection(ConnectionString)
        _Name = Name
        SetupLogging()
    End Sub

    Public Sub New(ByVal Name As String, ByVal Connection As SqlConnection)
        _Connection = Connection
        _TransactionConnection = New SqlConnection(Connection.ConnectionString)
        _Name = Name
        SetupLogging()
    End Sub

    Private Sub SetupLogging()
        If _Name <> "" Then
            _ErrorLog = Environment.GetEnvironmentVariable("TMP") & "\SQLAccess_" & _Name & ".log"
        Else
            _ErrorLog = Environment.GetEnvironmentVariable("TMP") & "\SQLAccess.log"
        End If
    End Sub
#End Region

#Region " Public Transaction Handling Operations "
    ''' <summary>
    ''' Begins a SQL transaction.
    ''' </summary>
    ''' <param name="TransactionID">Non-null transaction identifier.</param>
    ''' <remarks></remarks>
    Public Sub BeginTransaction(ByVal TransactionID As String)
        If TransactionID = "" Then
            If _LogErrors Then Log(LogType.Fatal, "XBegin: TransactionID cannot be null.")
            Throw New ApplicationException("TransactionID cannot be null.")
        End If

        _LastError = Nothing
        If _TransactionID = "" Then
            If _LogOperations Then Log(LogType.Information, "Begin transaction: " & TransactionID)
            _TransactionConnection.Open()
            _TransactionID = TransactionID
            _Transaction = _TransactionConnection.BeginTransaction(_TransactionID)
        Else
            If _LogErrors Then Log(LogType.Fatal, "XBegin: " & _TransactionID & " already in progress.")
            Throw New ApplicationException("Transaction '" & _TransactionID & " already in progress.")
        End If
    End Sub

    ''' <summary>
    ''' Commits the SQL transaction.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CommitTransaction()
        If _Transaction IsNot Nothing Then
            _LastError = Nothing
            If _LogOperations Then Log(LogType.Information, "Commit transaction: " & TransactionID)
            _Transaction.Commit()
            _TransactionConnection.Close()
            _TransactionID = ""
            _Transaction = Nothing
        Else
            If _LogErrors Then Log(LogType.Fatal, "XCommit: " & "No transaction is currently in progress.")
            Throw New ApplicationException("No transaction is currently in progress.")
        End If
    End Sub

    ''' <summary>
    ''' Rolls the failed SQL transaction back.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RollbackTransaction()
        If _Transaction IsNot Nothing Then
            _LastError = Nothing
            If _LogOperations Then Log(LogType.Information, "Rollback transaction: " & TransactionID)
            _Transaction.Rollback()
            _TransactionConnection.Close()
            _TransactionID = ""
            _Transaction = Nothing
        Else
            If _LogErrors Then Log(LogType.Fatal, "XRollback: " & "No transaction is currently in progress.")
            Throw New ApplicationException("No transaction is currently in progress.")
        End If
    End Sub
#End Region

#Region " Public Data Fetch Operations "
    ''' <summary>
    ''' Returns a datatable of results for the SQL operation.
    ''' </summary>
    ''' <param name="SQLOperation">Storedprocedure or text</param>
    ''' <param name="SQLcommand">Command text</param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetDatatable(ByVal SQLOperation As CommandType, ByVal SQLcommand As String) As DataTable
        Return GetDatatable(SQLOperation, SQLcommand, Nothing)
    End Function

    ''' <summary>
    ''' Returns a datatable of results for the SQL operation.
    ''' </summary>
    ''' <param name="SQLOperation">Stored procedure or text</param>
    ''' <param name="SQLcommand">Command text</param>
    ''' <param name="DBParameters">Array of parameters</param>
    ''' <returns>Datatable</returns>
    ''' <remarks></remarks>
    Public Function GetDatatable(ByVal SQLOperation As CommandType, ByVal SQLcommand As String, ByVal DBParameters As DBParameter()) As DataTable
        Dim dt As New DataTable
        Dim sqlconn As New SqlConnection(Me.ConnectionString)

        _LastError = Nothing

        Try
            Dim sqlcomm As New SqlCommand(SQLcommand, sqlconn)
            If _LogOperations Then Log(LogType.Information, "GetDatatable: " & SQLcommand)

            sqlcomm.CommandType = SQLOperation
            sqlcomm.CommandTimeout = _CurrentTimeOut

            AddReturnValueParameter(sqlcomm)
            AddSqlParameters(sqlcomm, DBParameters)

            Dim sqlda As New SqlDataAdapter(sqlcomm)
            sqlda.Fill(dt)

            _LastReturnCode = GetReturnCode(sqlcomm)
            StoreOutputParameters(sqlcomm, DBParameters)

            If _LogOperations Then Log(LogType.Information, "Number of records read=" & dt.Rows.Count)

        Catch ex As Exception
            _LastError = ex
            If _LogErrors Then Log(LogType.Fatal, "GetDatatable: " & ex.Message)
            ShowErrorMessage(ex)
        End Try

        Return dt
    End Function

    ''' <summary>
    ''' Updates an existing dataset with a datatable given the SQL operation provided.
    ''' </summary>
    ''' <param name="Dataset">An existing dataset.</param>
    ''' <param name="TableName">The name that the table should have in the dataset.</param>
    ''' <param name="SQLOperation">Stored procedure or text</param>
    ''' <param name="SQLcommand">Command text</param>
    ''' <returns>True if the operation was successful, else false.</returns>
    ''' <remarks>If the datatable already exists in the dataset it will be removed befor ethe new one is added.</remarks>
    Public Function GetDatatableIntoDataset(ByVal Dataset As DataSet, ByVal TableName As String, _
                                                ByVal SQLOperation As CommandType, ByVal SQLcommand As String) As Boolean
        Return GetDatatableIntoDataset(Dataset, TableName, SQLOperation, SQLcommand, Nothing)
    End Function

    ''' <summary>
    ''' Updates an existing dataset with a datatable given the SQL operation provided. 
    ''' </summary>
    ''' <param name="Dataset">An existing dataset.</param>
    ''' <param name="TableName">The name that the table should have in the dataset.</param>
    ''' <param name="SQLOperation">Stored procedure or text</param>
    ''' <param name="SQLcommand">Command text</param>
    ''' <param name="DBParameters">Array of parameters</param>
    ''' <returns>True if the operation was successful, else false.</returns>
    ''' <remarks>If the datatable already exists in the dataset it will be removed befor ethe new one is added.</remarks>
    Public Function GetDatatableIntoDataset(ByVal Dataset As DataSet, ByVal TableName As String, _
                                            ByVal SQLOperation As CommandType, ByVal SQLcommand As String, _
                                            ByVal DBParameters As DBParameter()) As Boolean
        Dim dt As DataTable = GetDatatable(SQLOperation, SQLcommand, DBParameters)
        If dt Is Nothing Then Return False
        dt.TableName = TableName
        If Dataset.Tables.Contains(TableName) Then Dataset.Tables.Remove(TableName)
        Dataset.Tables.Add(dt)
        Return True
    End Function

    ''' <summary>
    ''' Returns the first value from the result set of the SQL query/stored procedure.
    ''' </summary>
    ''' <param name="SQLOperation">Stored procedure or text.</param>
    ''' <param name="SQLcommand">Command text</param>
    ''' <returns>Any data type that can be held on a SQL Server.</returns>
    ''' <remarks></remarks>
    Public Function GetScalar(ByVal SQLOperation As CommandType, ByVal SQLcommand As String) As Object
        Return GetScalar(SQLOperation, SQLcommand, Nothing)
    End Function


    ''' <summary>
    ''' Returns the first value from the result set of the SQL query/stored procedure.
    ''' </summary>
    ''' <param name="SQLOperation">Stored procedure or text.</param>
    ''' <param name="SQLcommand">Command text</param>
    ''' <param name="DBParameters">Array of parameters</param>
    ''' <returns>Any data type that can be held on a SQL Server.</returns>
    ''' <remarks></remarks>
    Public Function GetScalar(ByVal SQLOperation As CommandType, ByVal SQLcommand As String, ByVal DBParameters As DBParameter()) As Object
        Dim res As Object = Nothing
        _LastError = Nothing
        Dim sqlconn As New SqlConnection(ConnectionString)

        Try
            Dim sqlcomm As New SqlCommand(SQLcommand, sqlconn)
            If _LogOperations Then Log(LogType.Information, "GetScalar: " & SQLcommand)

            sqlcomm.CommandType = SQLOperation
            sqlcomm.CommandTimeout = _CurrentTimeOut
            AddReturnValueParameter(sqlcomm)
            AddSqlParameters(sqlcomm, DBParameters)

            sqlconn.Open()
            res = sqlcomm.ExecuteScalar

            _LastReturnCode = GetReturnCode(sqlcomm)
            StoreOutputParameters(sqlcomm, DBParameters)

            If _LogOperations Then Log(LogType.Information, "res=" & res.ToString)

        Catch ex As Exception
            _LastError = ex
            If _LogErrors Then Log(LogType.Fatal, "GetScalar: " & ex.Message)
            ShowErrorMessage(ex)
        Finally
            sqlconn.Close()
        End Try

        Return res
    End Function

    ''' <summary>
    ''' Returns the RESULT of a stored procedure.
    ''' </summary>
    ''' <param name="StoredProcedure"></param>
    ''' <returns>Any data type that can be held on a SQL Server.</returns>
    ''' <remarks>Use this function if the stroed procedure returns a result. Use GetScalar if the stored procedure
    ''' uses SELECT to return a single item or the first item of a result set.</remarks>
    Public Function GetResult(ByVal StoredProcedure As String) As Object
        Return GetResult(StoredProcedure, Nothing)
    End Function

    ''' <summary>
    ''' Returns the RESULT of a stored procedure.
    ''' </summary>
    ''' <param name="StoredProcedure"></param>
    ''' <param name="DBParameters"></param>
    ''' <returns>Any data type that can be held on a SQL Server.</returns>
    ''' <remarks>Use this function if the stroed procedure returns a result. Use GetScalar if the stored procedure
    ''' uses SELECT to return a single item or the first item of a result set.</remarks>
    Public Function GetResult(ByVal StoredProcedure As String, ByVal DBParameters As DBParameter()) As Object
        _LastError = Nothing
        Dim sqlconn As New SqlConnection(ConnectionString)

        Try
            If _LogOperations Then Log(LogType.Information, "GetResult: " & StoredProcedure)
            Dim sqlcomm As New SqlCommand(StoredProcedure, sqlconn)

            AddReturnValueParameter(sqlcomm)

            With sqlcomm
                .CommandTimeout = _CurrentTimeOut
                .CommandType = CommandType.StoredProcedure
                If DBParameters IsNot Nothing Then
                    For Each param As DBParameter In DBParameters
                        .Parameters.Add(param.Name, param.DBType, param.Length)
                        .Parameters(param.Name).Value = param.ParameterValue
                        If _LogOperations Then
                            Log(LogType.Information, "Parameter: " & param.Name & "=" & NullStr(param.ParameterValue))
                        End If
                    Next
                End If
            End With

            sqlconn.Open()
            sqlcomm.ExecuteNonQuery()

            _LastReturnCode = GetReturnCode(sqlcomm)
            StoreOutputParameters(sqlcomm, DBParameters)

            If _LogOperations Then Log(LogType.Information, "res=" & _LastReturnCode.ToString)
        Catch ex As Exception
            _LastError = ex
            If _LogErrors Then Log(LogType.Fatal, "GetResult: " & ex.Message)
            ShowErrorMessage(ex)
        Finally
            sqlconn.Close()
        End Try

        Return _LastReturnCode
    End Function

    ''' <summary>
    ''' Given sql text, will return a blank row for the given sql SELECT statement.
    ''' Will throw an error if the sql contains "INSERT", "UPDATE" or "DELETE FROM"
    ''' DO NOT INCLUDE a WHERE Clause
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="sql"></param>
    ''' <returns>DataRow</returns>
    ''' <remarks></remarks>
    Public Function GetNewRow(ByVal tableName As String, ByVal sql As String) As System.Data.DataRow
        If sql.ToUpper.Contains("DELETE FROM") Then Throw New ArgumentException("Can not process SQL statement, it contains a DELETE statement" & vbCrLf & sql)
        If sql.ToUpper.Contains("UPDATE") Then Throw New ArgumentException("Can not process SQL statement, it contains an UPDATE statement" & vbCrLf & sql)
        If sql.ToUpper.Contains("INSERT") Then Throw New ArgumentException("Can not process SQL statement, it conatins an INSERT statement" & vbCrLf & sql)
        Dim dt As DataTable = Me.GetDatatable(CommandType.Text, sql & " WHERE 0=1", Nothing)
        dt.TableName = tableName
        dt.Rows.Add(dt.NewRow)

        Return dt.Rows(0)
    End Function

#End Region

#Region " Public Data Update/Insert/Delete Operations "
    ''' <summary>
    ''' To be used for inserting rows. Expects the result returned to be the identity of the newly
    ''' inserted row. 
    ''' </summary>
    ''' <param name="SQLOperation"></param>
    ''' <param name="SQLcommand"></param>
    ''' <param name="DBParameters"></param>
    ''' <returns>ID of row inserted</returns>
    ''' <remarks></remarks>
    Public Function Insert(ByVal SQLOperation As CommandType, ByVal SQLcommand As String, ByVal DBParameters As DBParameter()) As Integer
        Dim res As Integer = 0
        Dim sqlconn As SqlConnection
        _LastError = Nothing

        If _TransactionID <> "" Then
            sqlconn = _TransactionConnection
        Else
            sqlconn = New SqlConnection(Me.ConnectionString)
        End If

        Try
            Dim sqlcomm As New SqlCommand(SQLcommand, sqlconn)
            If _LogOperations Then Log(LogType.Information, "Insert " & SQLcommand)

            sqlcomm.CommandType = SQLOperation
            AddReturnValueParameter(sqlcomm)
            AddSqlParameters(sqlcomm, DBParameters)
            sqlcomm.CommandTimeout = _CurrentTimeOut

            If sqlconn IsNot _TransactionConnection Then
                sqlconn.Open()
            Else
                sqlcomm.Transaction = _Transaction
            End If

            res = CInt(sqlcomm.ExecuteScalar)

            _LastReturnCode = GetReturnCode(sqlcomm)
            StoreOutputParameters(sqlcomm, DBParameters)

            If _LogOperations Then Log(LogType.Information, "id=" & res.ToString)

        Catch ex As Exception
            _LastError = ex
            If _LogErrors Then Log(LogType.Fatal, "Insert: " & ex.Message)

            If _TransactionID <> "" Then
                ' For transactions, the caller must handle fatal errors to call rollback.
                Throw ex
            Else
                ShowErrorMessage(ex)
            End If

        Finally
            If sqlconn IsNot _TransactionConnection Then sqlconn.Close()
        End Try

        Return res
    End Function

    ''' <summary>
    ''' To be used for updating rows. The result returned is the number of rows updated.
    ''' </summary>
    ''' <param name="SQLOperation"></param>
    ''' <param name="SQLcommand"></param>
    ''' <param name="DBParameters"></param>
    ''' <returns>Number of rows updated</returns>
    ''' <remarks></remarks>
    Public Function Update(ByVal SQLOperation As CommandType, ByVal SQLcommand As String, ByVal DBParameters As DBParameter()) As Integer
        Dim res As Integer = 0
        Dim sqlconn As SqlConnection
        _LastError = Nothing

        If _TransactionID <> "" Then
            sqlconn = _TransactionConnection
        Else
            sqlconn = New SqlConnection(ConnectionString)
        End If

        Try
            If _LogOperations Then Log(LogType.Information, "Update/Delete: " & SQLcommand)

            Dim sqlcomm As New SqlCommand(SQLcommand, sqlconn)
            sqlcomm.CommandType = SQLOperation
            sqlcomm.CommandTimeout = _CurrentTimeOut
            AddReturnValueParameter(sqlcomm)
            AddSqlParameters(sqlcomm, DBParameters)

            If sqlconn IsNot _TransactionConnection Then
                sqlconn.Open()
            Else
                sqlcomm.Transaction = _Transaction
            End If

            res = sqlcomm.ExecuteNonQuery
            _LastReturnCode = GetReturnCode(sqlcomm)
            StoreOutputParameters(sqlcomm, DBParameters)

            If _LogOperations Then Log(LogType.Information, "Number of rows=" & res.ToString)
        Catch ex As Exception
            _LastError = ex
            If _LogErrors Then Log(LogType.Fatal, "Update/Delete: " & ex.Message)

            If _TransactionID <> "" Then
                ' For transactions, the caller must handle fatal errors to call rollback.
                Throw ex
            Else
                'Throw New ApplicationException("Insert error: " & ex.Message)
                ShowErrorMessage(ex)
            End If
        Finally
            If sqlconn IsNot _TransactionConnection Then sqlconn.Close()
        End Try

        Return res
    End Function

    ''' <summary>
    ''' To be used for deleting rows. The result returned is the number of rows deleted.
    ''' </summary>
    ''' <param name="SQLOperation"></param>
    ''' <param name="SQLcommand"></param>
    ''' <param name="DBParameters"></param>
    ''' <returns>Number of rows deleted</returns>
    ''' <remarks>Allows the upper layer to call Update as a Delete operation.</remarks>
    Public Function Delete(ByVal SQLOperation As CommandType, ByVal SQLcommand As String, ByVal DBParameters As DBParameter()) As Integer
        Return Update(SQLOperation, SQLcommand, DBParameters)
    End Function

    Public Sub Execute(ByVal SQLOperation As CommandType, ByVal SQLcommand As String)
        Execute(SQLOperation, SQLcommand, Nothing)
    End Sub

    ''' <summary>
    ''' To be used for executing an ad-hoc SQL command or stored procedure.
    ''' </summary>
    ''' <param name="SQLOperation"></param>
    ''' <param name="SQLcommand"></param>
    ''' <param name="DBParameters"></param>
    ''' <remarks></remarks>
    Public Sub Execute(ByVal SQLOperation As CommandType, ByVal SQLcommand As String, ByVal DBParameters As DBParameter())
        Dim res As Integer = 0
        Dim sqlconn As SqlConnection
        _LastError = Nothing

        If _TransactionID <> "" Then
            sqlconn = _TransactionConnection
        Else
            sqlconn = New SqlConnection(Me.ConnectionString)
        End If

        Try
            Dim sqlcomm As New SqlCommand(SQLcommand, sqlconn)
            If _LogOperations Then Log(LogType.Information, "Execute " & SQLcommand)

            sqlcomm.CommandType = SQLOperation
            AddReturnValueParameter(sqlcomm)
            AddSqlParameters(sqlcomm, DBParameters)
            sqlcomm.CommandTimeout = _CurrentTimeOut

            If sqlconn IsNot _TransactionConnection Then
                sqlconn.Open()
            Else
                sqlcomm.Transaction = _Transaction
            End If

            sqlcomm.ExecuteNonQuery()
            _LastReturnCode = GetReturnCode(sqlcomm)
            StoreOutputParameters(sqlcomm, DBParameters)

        Catch ex As Exception
            _LastError = ex
            If _LogErrors Then Log(LogType.Fatal, "Execute: " & ex.Message)

            If _TransactionID <> "" Then
                ' For transactions, the caller must handle fatal errors to call rollback.
                Throw ex
            Else
                ShowErrorMessage(ex)
            End If

        Finally
            If sqlconn IsNot _TransactionConnection Then sqlconn.Close()
        End Try
    End Sub
#End Region

#Region " Public Bulk copy Methods "
    ''' <summary>
    ''' Bulk copy with append to destination table's data.
    ''' </summary>
    ''' <param name="sourceTable">The table of data to be copied.</param>
    ''' <param name="sourceColumnList">The list of columns from which to take data.</param>
    ''' <param name="destinationTableName">The table of data to be bulk copied.</param>
    ''' <param name="destinationColumnList">The list of columns to copy the data to.</param>
    ''' <param name="AppendToDestinationTable">The name  of the destination table.</param>
    ''' <remarks>Rows in the destination tyable are deleted if any are present.</remarks>
    Public Sub CopyDataWithColumnMapping(ByVal sourceTable As DataTable, _
                                         ByVal sourceColumnList As List(Of String), _
                                         ByVal destinationTableName As String, _
                                         ByVal destinationColumnList As List(Of String), _
                                         ByVal AppendToDestinationTable As Boolean)

        If sourceColumnList.Count <> destinationColumnList.Count Then
            Throw New Exception("Source and detsination column counts differ.")
        End If

        Using s As New SqlBulkCopy(_Connection)
            s.DestinationTableName = destinationTableName

            If _NotifyBulkProgressCount > 0 Then
                s.NotifyAfter = _NotifyBulkProgressCount
                AddHandler s.SqlRowsCopied, AddressOf s_SqlRowsCopied
            End If

            For i As Integer = 0 To sourceColumnList.ToArray().Count() - 1
                s.ColumnMappings.Add(sourceColumnList(i).ToString, destinationColumnList(i).ToString)
            Next

            Try
                _Connection.Open()

                If Not AppendToDestinationTable Then
                    Dim cmd As New SqlCommand("delete from " & destinationTableName, _Connection)
                    cmd.ExecuteNonQuery()
                End If

                s.WriteToServer(sourceTable)

            Catch ex As Exception
                _LastError = ex
                If _LogErrors Then Log(LogType.Fatal, "CopyDataWithColumnMapping: " & ex.Message)
                ShowErrorMessage(ex)
            Finally
                _Connection.Close()
                s.Close()
            End Try
        End Using
    End Sub

    ''' <summary>
    ''' Bulk copy with append to destination table's data.
    ''' </summary>
    ''' <param name="sourceTable">The table of data to be copied.</param>
    ''' <param name="sourceColumnList">The list of columns from which to take data.</param>
    ''' <param name="destinationTableName">The table of data to be bulk copied.</param>
    ''' <param name="destinationColumnList">The list of columns to copy the data to.</param>
    ''' <remarks>Rows in the destination table are preserved.</remarks>
    Public Sub CopyDataWithColumnMapping(ByVal sourceTable As DataTable, _
                                         ByVal sourceColumnList As List(Of String), _
                                         ByVal destinationTableName As String, _
                                         ByVal destinationColumnList As List(Of String))
        CopyDataWithColumnMapping(sourceTable, sourceColumnList, destinationTableName, destinationColumnList, True)
    End Sub

    Public Event CopyNotify(ByVal sender As Object, ByVal e As SqlRowsCopiedEventArgs)

    Private Sub s_SqlRowsCopied(ByVal sender As Object, ByVal e As SqlRowsCopiedEventArgs)
        Debug.WriteLine("-- Copied " & e.RowsCopied & " rows")
        RaiseEvent CopyNotify(sender, e)
    End Sub
#End Region

#Region " Private Methods "

    Private Enum LogType
        Information = 1
        Warning = 2
        Fatal = 3
    End Enum

    Private Sub Log(ByVal LogType As LogType, ByVal text As String)
        Dim logw As System.IO.StreamWriter = Nothing

        ' Although the Log file is created in the TMP directory (which is pretty much useless if
        ' you can't write to it), Windows can set the TMP variable to a read-only location, eg.   
        ' for a Scheduled task. We could also get here from disk full. 
        '
        ' Catch the error and crash ... 
        Try
            logw = New System.IO.StreamWriter(_ErrorLog, True)
            logw.Write(Now.ToString("yyyyMMdd@HHmmss: "))

            Select Case LogType
                Case SQLAccess.LogType.Information : logw.Write("Info  ")
                Case SQLAccess.LogType.Warning : logw.Write("Warn  ")
                Case SQLAccess.LogType.Fatal : logw.Write("Fatal ")
            End Select

            If _TransactionID <> "" Then
                logw.Write("(" & _TransactionID & ") ")
            End If
            logw.WriteLine(text)
        Catch unauthex As System.UnauthorizedAccessException
            Throw New ApplicationException("Unable to log to " & _ErrorLog & " (access is denied)." & _
                                           "Log message is " & ControlChars.CrLf & text)
        Catch ex As Exception
            Throw New ApplicationException("Unable to log to " & _ErrorLog & ex.Message & _
                                           ControlChars.CrLf & text)
        Finally
            logw.Close()
        End Try
    End Sub

    Private Sub ShowErrorMessage(ByVal ex As Exception)
        ' Console.WriteLine(ex.Message)
        If _CrashOnError AndAlso _TransactionID = "" Then Throw ex
    End Sub

    ''' <summary>
    ''' The @RETURN_VALUE output parameter is added to the parameters list.
    ''' </summary>
    ''' <param name="sqlcommand"></param>
    ''' <remarks>Use sqlcommand.Parameters("@RETURN_VALUE").Value to obtain the returned result.</remarks>
    Private Sub AddReturnValueParameter(ByVal sqlcommand As SqlCommand)
        sqlcommand.Parameters.Add(New SqlParameter( _
            "@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, _
            10, 0, Nothing, System.Data.DataRowVersion.Current, False, Nothing, "", "", ""))
    End Sub

    Private Function GetReturnCode(ByVal sqlcommand As SqlCommand) As Integer
        Return CInt(sqlcommand.Parameters("@RETURN_VALUE").Value)
    End Function

    Private Sub StoreOutputParameters(ByVal sqlcommand As SqlCommand, ByVal dbParameters As DBParameter())
        If dbParameters IsNot Nothing Then
            For Each parameter In dbParameters
                If parameter.ForOutput Then
                    parameter.ParameterValue = sqlcommand.Parameters(parameter.Name).Value
                End If
            Next
        End If
    End Sub
    ''' <summary>
    ''' Used to guarantee safe entry of all parameter values into the log.
    ''' </summary>
    ''' <param name="val"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function NullStr(ByVal val As Object) As String
        If IsDBNull(val) Then Return "<DBNull>"
        If val Is Nothing Then Return "<Nothing>"
        If val.ToString = "" Then Return "''"
        Return val.ToString
    End Function

    Private Sub AddSqlParameters(ByVal sqlcomm As SqlCommand, ByVal params() As DBParameter)
        If params Is Nothing Then Exit Sub

        For Each param As DBParameter In params
            If sqlcomm.CommandType = CommandType.Text Then
                If sqlcomm.CommandText.ToLower.IndexOf(param.Name.ToLower) < 1 Then Continue For
            End If

            Dim sqlParam As SqlParameter = Nothing
            If param.ForOutput Then
                sqlParam = New SqlParameter(param.Name, param.DBType, param.Length, _
                                            ParameterDirection.Output, True, 18, 2, _
                                            Nothing, DataRowVersion.Default, param.ParameterValue)
            Else
                sqlParam = New SqlParameter(param.Name, param.DBType, param.Length)
                sqlParam.Value = param.ParameterValue
            End If

            sqlcomm.Parameters.Add(sqlParam)

            If _LogOperations Then
                Log(LogType.Information, "Parameter: " & param.Name & "=" & NullStr(param.ParameterValue))
            End If
        Next
    End Sub
#End Region

#Region "Examine database"
    ''' <summary>
    ''' Return a list of all the stored procedures in the database
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExamineStoredProcedureNames() As List(Of String)
        Dim procs As New List(Of String)
        Dim commandtext As String = "SELECT [SPECIFIC_NAME] FROM information_schema.routines WHERE ROUTINE_TYPE = 'PROCEDURE' ORDER BY [SPECIFIC_NAME]"
        Dim newTable As DataTable = Me.GetDatatable(CommandType.Text, commandtext)
        For Each row As DataRow In newTable.Rows
            procs.Add(CStr(row("SPECIFIC_NAME")))
        Next
        Return procs
    End Function
#End Region
End Class

