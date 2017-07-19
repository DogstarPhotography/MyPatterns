Imports System.Data.SQLite

Public Class DataComponent(Of Connection As {System.Data.Common.DbConnection, New}, Command As {System.Data.Common.DbCommand, New}, DataReader As {System.Data.Common.DbDataReader})

#Region "Constructor"
    Private myDataStore As DataStore(Of Connection, Command, DataReader)

    Private Sub New()
        'Prevent anyone from instantiating this class by making New private
    End Sub

    Public Sub New(ByVal DataStore As DataStore(Of Connection, Command, DataReader))
        myDataStore = DataStore
    End Sub

    Public Property DataStore() As DataStore(Of Connection, Command, DataReader)
        Get
            Return myDataStore
        End Get
        Set(ByVal value As DataStore(Of Connection, Command, DataReader))
            myDataStore = value
        End Set
    End Property
#End Region

#Region "Single Item"
    Public Function Add(ByVal ID As Integer, ByVal Item As DataItem) As DataItem
        'Check for duplicate
        If Me.ContainsRecord(ID) = True Then
            'Update
            UpdateRecord(ID, Item)
        Else
            'Add the given LatLon to the database
            CreateRecord(ID, Item)
        End If
        'Return the entry in the database
        Return RetrieveRecord(Item.ID)
    End Function

    Public Function Add(ByVal Item As DataItem) As DataItem
        'Call the add function using the items id property
        Return Me.Add(Item.ID, Item)
    End Function

    Public Property Item(ByVal ID As Integer) As DataItem
        Get
            'Fetch data
            Return RetrieveRecord(ID)
        End Get
        Set(ByVal value As DataItem)
            'Save data
            UpdateRecord(ID, value)
        End Set
    End Property

    Public Sub Remove(ByVal LatLonID As Integer)
        DeleteRecord(LatLonID)
    End Sub

    Public Sub Clear()
        DeleteAllRecords()
    End Sub
#End Region

#Region "ID"
    Public Function ContainsRecord(ByVal ID As Integer) As Boolean
        Dim conn As New Connection
        conn.ConnectionString = myDataStore.ConnectionString
        Dim newItem As New DataItem
        Try
            conn.Open()
            Dim SQL As String
            SQL = "SELECT ID FROM ItemTable WHERE ID = " & ID.ToString
            'TODO Log the SQL
            Dim cmd As New Command()
            cmd.Connection = conn
            cmd.CommandText = SQL
            cmd.CommandType = CommandType.Text
            Dim rdr As DataReader = CType(cmd.ExecuteReader(), DataReader)
            'Do we have any resuls?
            If rdr.HasRows = True Then
                rdr.Close()
                Return True
            Else
                rdr.Close()
                Return False
            End If
        Catch ex As Exception
            Throw ex
        Finally
            conn.Close()
        End Try
        Return False
    End Function

    Public Function GetMaxID() As Integer
        Dim conn As New Connection
        conn.ConnectionString = myDataStore.ConnectionString
        'Dim newItem As New DataItem
        Dim Result As Integer
        Try
            conn.Open()
            Dim SQL As String
            SQL = "SELECT Max(ID) as ID FROM ItemTable"
            'TODO Log the SQL
            Dim cmd As New Command()
            cmd.Connection = conn
            cmd.CommandText = SQL
            cmd.CommandType = CommandType.Text
            Dim rdr As DataReader = CType(cmd.ExecuteReader(), DataReader)
            'Add properties
            If rdr.HasRows = True Then
                'Read the first row
                rdr.Read()
                If Not DBNull.Value.Equals(rdr("ID")) Then
                    Result = CInt(rdr("ID"))
                Else
                    Result = 0
                End If
            Else
                Result = 0
            End If
            rdr.Close()
        Catch ex As Exception
            Throw ex
        Finally
            conn.Close()
        End Try
        Return Result
    End Function
#End Region

#Region "CRUD"
    Private Sub CreateRecord(ByVal ID As Integer, ByVal Item As DataItem)
        Dim conn As New Connection
        conn.ConnectionString = myDataStore.ConnectionString
        Try
            conn.Open()
            Dim SQL As String
            SQL = "INSERT INTO ItemTable (ID, Field1, Field2, Field3) " & _
                    "VALUES (" & ID & ", '" & Item.Field1 & "', '" & Item.Field2 & "', '" & Item.Field3 & "')"
            'TODO Log the SQL
            Dim cmd As New Command()
            cmd.Connection = conn
            cmd.CommandText = SQL
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            conn.Close()
        End Try
    End Sub

    Private Function RetrieveRecord(ByVal ID As Integer) As DataItem
        Dim conn As New Connection
        conn.ConnectionString = myDataStore.ConnectionString
        Dim newItem As New DataItem
        Try
            conn.Open()
            Dim SQL As String
            SQL = "SELECT ID, Field1, Field2, Field3 FROM ItemTable WHERE ID = " & ID.ToString
            'TODO Log the SQL
            Dim cmd As New Command()
            cmd.Connection = conn
            cmd.CommandText = SQL
            cmd.CommandType = CommandType.Text
            Dim rdr As DataReader = CType(cmd.ExecuteReader(), DataReader)
            'Read the first row
            rdr.Read()
            'Add properties
            newItem.ID = CInt(rdr("ID"))
            newItem.Field1 = CStr(rdr("Field1"))
            newItem.Field2 = CStr(rdr("Field2"))
            newItem.Field3 = CStr(rdr("Field3"))
            rdr.Close()
        Catch ex As Exception
            Throw ex
        Finally
            conn.Close()
        End Try
        Return newItem
    End Function

    Private Sub UpdateRecord(ByVal ID As Integer, ByVal Item As DataItem)
        Dim conn As New Connection
        conn.ConnectionString = myDataStore.ConnectionString
        Try
            conn.Open()
            Dim SQL As String
            SQL = "UPDATE ItemTable SET Field1 = '" & Item.Field1 & "', Field2 = '" & Item.Field2 & "', Field3 = '" & Item.Field3 & "' " & _
                                        "WHERE ID = " & ID
            'TODO Log the SQL
            Dim cmd As New Command()
            cmd.Connection = conn
            cmd.CommandText = SQL
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub DeleteRecord(ByVal ID As Integer)
        Dim conn As New Connection
        conn.ConnectionString = myDataStore.ConnectionString
        Try
            conn.Open()
            Dim SQL As String
            SQL = "DELETE FROM ItemTable WHERE ID = " & ID
            'TODO Log the SQL
            Dim cmd As New Command()
            cmd.Connection = conn
            cmd.CommandText = SQL
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub DeleteAllRecords()
        Dim conn As New SQLiteConnection(myDataStore.ConnectionString)
        Try
            conn.Open()
            Dim SQL As String
            SQL = "DELETE FROM ItemTable"
            'TODO Log the SQL
            Dim cmd As New Command()
            cmd.Connection = conn
            cmd.CommandText = SQL
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            conn.Close()
        End Try
    End Sub
#End Region
End Class
