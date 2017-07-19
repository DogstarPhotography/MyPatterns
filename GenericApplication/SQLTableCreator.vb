Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Data.SqlClient
Imports System.Data

Public Class SqlTableCreator

    Private Shared Function GetCharsToTrim() As Char()
        Dim totrim() As Char = {CChar(","), CChar(" ")}
        Return totrim
    End Function

#Region "Instance Variables"
    Private _connection As SqlConnection
    Public Property Connection() As SqlConnection
        Get
            Return _connection
        End Get
        Set(ByVal value As SqlConnection)
            _connection = value
        End Set
    End Property

    Private _transaction As SqlTransaction
    Public Property Transaction() As SqlTransaction
        Get
            Return _transaction
        End Get
        Set(ByVal value As SqlTransaction)
            _transaction = value
        End Set
    End Property

    Private _tableName As String
    Public Property TableName() As String
        Get
            Return _tableName
        End Get
        Set(ByVal value As String)
            _tableName = value
        End Set
    End Property
#End Region

#Region "Constructor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal connection As SqlConnection)
        _connection = connection
    End Sub

    Public Sub New(ByVal connection As SqlConnection, ByVal transaction As SqlTransaction)
        _connection = connection
        _transaction = transaction
    End Sub
#End Region

#Region "Instance Methods"
    Public Function Create(ByVal schema As DataTable) As Object
        Return Create(schema, 0)
    End Function

    Public Function Create(ByVal schema As DataTable, ByVal numKeys As Integer) As Object
        Dim primaryKeys(numKeys) As Integer
        For i As Integer = 0 To numKeys
            primaryKeys(i) = i
        Next
        Return Create(schema, primaryKeys)
    End Function

    Public Function Create(ByVal schema As DataTable, ByVal primaryKeys As Integer()) As Object
        Dim sql As String = GetCREATETABLESQL(_tableName, schema, primaryKeys)
        Dim cmd As SqlCommand
        If _transaction IsNot Nothing And _transaction.Connection IsNot Nothing Then
            cmd = New SqlCommand(sql, _connection, _transaction)
        Else
            cmd = New SqlCommand(sql, _connection)
        End If
        Return cmd.ExecuteNonQuery()
    End Function

    Public Function CreateFromDataTable(ByVal table As DataTable) As Object
        Dim sql As String = GetCREATETABLESQLFromDataTable(_tableName, table)
        Dim cmd As SqlCommand
        If _transaction IsNot Nothing And _transaction.Connection IsNot Nothing Then
            cmd = New SqlCommand(sql, _connection, _transaction)
        Else
            cmd = New SqlCommand(sql, _connection)
        End If
        Return cmd.ExecuteNonQuery()
    End Function
#End Region

#Region "Shared Methods"

    Public Shared Function GetCREATETABLESQL(ByVal tableName As String, ByVal schema As DataTable, ByVal primaryKeys As Integer()) As String

        Dim sql As String = "CREATE TABLE [" & tableName & "] ( "
        ' columns
        For Each column As DataRow In schema.Rows
            If Not (schema.Columns.Contains("IsHidden") And CBool(column("IsHidden")) = True) Then
                sql &= " [" & column("ColumnName").ToString() & "] " & GetSQLColumnType(column)
                If schema.Columns.Contains("AllowDBNull") And CBool(column("AllowDBNull")) = False Then
                    sql &= " NOT NULL"
                End If
                sql &= ", "
            End If
        Next
        sql = sql.TrimEnd(GetCharsToTrim()) & " "
        ' primary keys
        Dim pk As String = ", CONSTRAINT PK_" & tableName & " PRIMARY KEY CLUSTERED ("
        Dim hasKeys As Boolean = primaryKeys IsNot Nothing And primaryKeys.Length > 0
        If hasKeys = True Then
            ' user defined keys
            For Each key As Integer In primaryKeys
                pk &= schema.Rows(key)("ColumnName").ToString() & ", "
            Next
        Else
            ' check schema for keys
            Dim keys As String = String.Join(", ", GetPrimaryKeys(schema))
            pk &= keys
            hasKeys = (keys.Length > 0)
        End If
        pk = pk.TrimEnd(GetCharsToTrim()) & ") "
        If hasKeys = True Then sql &= pk
        sql &= ")"
        Return sql

    End Function

    Public Shared Function GetCREATETABLESQLFromDataTable(ByVal tableName As String, ByVal table As DataTable, Optional ByVal forceSimple As Boolean = False) As String

        Dim sql As String = "CREATE TABLE [" & tableName & "] ( " & vbCrLf
        ' columns
        For Each column As DataColumn In table.Columns
            If forceSimple = True Then
                sql &= "[" & column.ColumnName & "] " & GetSimpleColumnSQL(column.ColumnName, table) & ", " & vbCrLf
            Else
                sql &= "[" & column.ColumnName & "] " & GetSQLColumnType(column) & ", " & vbCrLf
            End If
        Next
        sql = sql.TrimEnd(GetCharsToTrim()) & " "
        ' primary keys
        If (table.PrimaryKey.Length > 0) Then
            sql &= "CONSTRAINT (PK_" & tableName & ") PRIMARY KEY CLUSTERED ("
            For Each column As DataColumn In table.PrimaryKey
                sql &= "[" & column.ColumnName & "],"
            Next
            sql = sql.TrimEnd(GetCharsToTrim()) & ")) "
        End If
        'if not ends with ")"
        If (table.PrimaryKey.Length = 0 And Not sql.EndsWith(")")) Then
            sql &= ")"
        End If
        Return sql

    End Function

    Public Shared Function GetColumnsFromDataTable(ByRef table As DataTable) As String

        Dim sql As String = ""
        ' columns
        For Each column As DataColumn In table.Columns
            sql &= "[" & column.ColumnName & "], "
        Next
        sql = sql.TrimEnd(GetCharsToTrim()) & " "
        Return sql

    End Function

    Public Shared Function GetColumnListFromDataTable(ByRef table As DataTable) As List(Of String)

        Dim sql As New List(Of String)
        ' columns
        For Each column As DataColumn In table.Columns
            sql.Add(column.ColumnName)
        Next
        Return sql

    End Function

    Public Shared Function GetMaxColumnWidth(ByVal column As String, ByRef table As DataTable) As String

        Dim max As Integer = 0
        For Each row As DataRow In table.Rows
            If row(column).ToString.Length > max Then
                Debug.Print(row(column).ToString)
                max = row(column).ToString.Length
            End If
        Next
        'Now group the result
        Select Case max
            Case Is <= 12
                Return max.ToString
            Case Is <= 20
                Return "20"
            Case Is <= 50
                Return "50"
            Case Is <= 255
                Return "255"
            Case Is <= 1024
                Return "1024"
            Case Else
                Return "MAX"
        End Select

    End Function

    Public Shared Function GetSimpleColumnSQL(ByVal column As String, ByRef table As DataTable) As String

        Dim max As Integer = 1
        Dim numeric As Boolean = True
        Dim isInteger As Boolean = True
        Dim parse As Integer
        Dim value As String
        For Each row As DataRow In table.Rows
            value = row(column).ToString
            If numeric = True AndAlso IsNumeric(value) = False Then
                numeric = False
                isInteger = False
            End If
            If numeric = True AndAlso isInteger = True AndAlso Integer.TryParse(value, parse) = False Then
                isInteger = False
            End If
            If value.Length > max Then
                Debug.Print(value)
                max = value.Length
            End If
        Next
        If numeric = True Then
            If isInteger = True Then
                Return "INT"
            Else
                Return "DECIMAL(18,6)"
            End If
        Else
            'Now group the result
            Select Case max
                Case Is <= 12
                    Return "VARCHAR(" & max.ToString & ") NULL"
                Case Is <= 20
                    Return "VARCHAR(20) NULL"
                Case Is <= 50
                    Return "VARCHAR(50) NULL"
                Case Is <= 255
                    Return "VARCHAR(255) NULL"
                Case Is <= 1024
                    Return "VARCHAR(1024) NULL"
                Case Else
                    Return "VARCHAR(MAX) NULL"
            End Select
        End If

    End Function

    Public Shared Function GetPrimaryKeys(ByVal schema As DataTable) As String()

        Dim keys As List(Of String) = New List(Of String)
        For Each column As DataRow In schema.Rows
            If schema.Columns.Contains("IsKey") And CType(column("IsKey"), Boolean) Then
                keys.Add(column("ColumnName").ToString())
            End If
        Next
        Return keys.ToArray()

    End Function

    ' Return T-SQL data type definition, based on schema definition for a column
    Public Shared Function GetSQLColumnType(ByVal type As Object, ByVal columnSize As Integer, ByVal numericPrecision As Integer, ByVal numericScale As Integer) As String

        'Switch(Type.ToString())
        Select Case type.ToString()
            Case "System.String"
                If columnSize = -1 Then
                    Return "VARCHAR(255)"
                ElseIf columnSize > 8000 Then
                    Return "VARCHAR(MAX)"
                Else
                    Return "VARCHAR(" & columnSize.ToString & ")"
                End If

            Case "System.Decimal"
                If numericScale > 0 Then
                    Return "REAL"
                ElseIf numericPrecision > 10 Then
                    Return "BIGINT"
                Else
                    Return "INT"
                End If

            Case "System.Double", "System.Single"
                Return "REAL"

            Case "System.Int64"
                Return "BIGINT"

            Case "System.Int16", "System.Int32"
                Return "INT"

            Case "System.DateTime"
                Return "DATETIME"

            Case "System.Boolean"
                Return "BIT"

            Case "System.Byte"
                Return "TINYINT"

            Case "System.Guid"
                Return "UNIQUEIDENTIFIER"

            Case Else
                Throw New Exception(type.ToString() & " not implemented.")

        End Select
    End Function

    ' Overload based on row from schema table
    Public Shared Function GetSQLColumnType(ByVal schemaRow As DataRow) As String
        Return GetSQLColumnType(schemaRow("DataType"), _
             Integer.Parse(schemaRow("ColumnSize").ToString()), _
             Integer.Parse(schemaRow("NumericPrecision").ToString()), _
             Integer.Parse(schemaRow("NumericScale").ToString()))
    End Function

    ' Overload based on DataColumn from DataTable type
    Public Shared Function GetSQLColumnType(ByVal column As DataColumn) As String
        Return GetSQLColumnType(column.DataType, column.MaxLength, 10, 2)
    End Function
#End Region

End Class
