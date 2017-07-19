Imports System.Data.Common
Imports System.Data.OleDb
''' <summary>
''' Holds a set of data for reading in a disconnected manner 
''' to reduce database load and improve speed 
''' when referencing the same set of tables again and again
''' </summary>
''' <remarks>
''' Consumers should first call LoadData with a connection string and list of tables/views to be loaded.
''' Then consumers can use the various Get... functions to fetch data as required.
''' </remarks>
Public Class Data

    Public ConnectionString As String = ""
    Public Data As DataSet
    Public Tables As Generic.List(Of String)

    Public Sub LoadData(ByVal OleDBConnectionString As String, ByVal TablesToLoad As Generic.List(Of String))
        Dim conn As OleDbConnection
        'Save data
        Tables = TablesToLoad
        ConnectionString = OleDBConnectionString
        'Open DB
        conn = New OleDbConnection(ConnectionString)
        conn.Open()
        'Fetch all data from DB
        Data = New DataSet
        Try
            For Each curTablename As String In Tables
                Dim Adapter As New OleDbDataAdapter("SELECT * FROM " & curTablename, conn)
                Adapter.Fill(Data, curTablename)
            Next
        Catch ex As Exception
            Throw ex
        End Try
        'Close DB
        conn.Close()
    End Sub

#Region "Public Procedures"
    Public Function GetTable(ByVal Name As String) As DataTable
        Return Data.Tables.Item(Name)
    End Function

    Public Function GetTable(ByVal Name As String, ByVal WhereClause As String) As DataTable
        Try
            Dim TheTable As DataTable = GetTable(Name)
            Dim DV As DataView = New DataView(TheTable, WhereClause, "", DataViewRowState.CurrentRows)
            Return DV.ToTable
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetView(ByVal Name As String, ByVal WhereClause As String) As DataView
        Try
            Dim TheTable As DataTable = GetTable(Name)
            Dim DV As DataView = New DataView(TheTable, WhereClause, "", DataViewRowState.CurrentRows)
            Return DV
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetReader(ByVal Name As String) As DataTableReader
        Return GetTable(Name).CreateDataReader
    End Function

    Public Function GetReader(ByVal Name As String, ByVal WhereClause As String) As DataTableReader
        Try
            Dim TheTable As DataTable = GetTable(Name)
            Dim DV As DataView = New DataView(TheTable, WhereClause, "", DataViewRowState.CurrentRows)
            Return DV.ToTable.CreateDataReader
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Sample procedures"
    Public Function GetSPECIFIC_TABLEReader() As DbDataReader
        Return GetReader("SPECIFIC_TABLE")
    End Function

    Public Function GetSPECIFIC_TABLEReaderByCOLUMN(ByVal Value As String) As DbDataReader
        Return GetReader("SPECIFIC_TABLE", "COLUMN = " & Value)
    End Function

    Public Function GetItemList(Of T As New)(ByVal Name As String, ByVal WhereClause As String) As Generic.List(Of T)
        'Create list
        Dim Items As Generic.List(Of T) = New Generic.List(Of T)
        Dim rdrItems As DbDataReader = Nothing
        Try
            'Fetch data
            rdrItems = GetReader(Name, WhereClause)
            'Copy data to list
            If rdrItems.HasRows = True Then
                Do While rdrItems.Read = True
                    Dim newItem As T = New T
                    'Add item specific properties
                    'TODO
                    'EXAMPLE: newItem.PROPERTY_NAME = rdrItems.GetValue(0)
                    Items.Add(newItem)
                Loop
            End If
        Catch ex As Exception
            Throw ex
        Finally
            rdrItems.Close()
        End Try
        Return Items
    End Function
#End Region
End Class
