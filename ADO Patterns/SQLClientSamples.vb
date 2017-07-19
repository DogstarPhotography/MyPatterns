Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic

Public Class SQLClientSamples

    Public SQLConnectionString As String = ""

    Public Function GetDataCollection() As SortedList(Of String, String)

        Dim conn As New SqlConnection(SQLConnectionString)
        Dim newlist As New SortedList(Of String, String)

        Try
            conn.Open()

            'TODO - Create SQL
            Dim cmd As New SqlCommand("COMMAND_TEXT_HERE", conn) 'SELECT KEY_COLUMN_NAME, DATA_COLUMN_NAME FROM TABLE_NAME
            cmd.CommandType = CommandType.Text
            Dim rdr As SqlDataReader = cmd.ExecuteReader()

            While rdr.Read
                'TODO - Fix column names
                newlist.Add(CStr(rdr("KEY_COLUMN_NAME_HERE")), CStr(rdr("DATA_COLUMN_NAME_HERE")))
            End While

            rdr.Close()

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

        Return newlist

    End Function

    Public Function GetObjectCollection() As SortedList(Of String, Object)

        Dim conn As New SqlConnection(SQLConnectionString)
        Dim newlist As New SortedList(Of String, Object)

        Try
            conn.Open()

            'TODO - Create SQL
            Dim cmd As New SqlCommand("COMMAND_TEXT_HERE", conn) 'SELECT KEY_COLUMN_NAME, PROPERTY_COLUMN_NAME, PROPERTY_COLUMN_NAME, PROPERTY_COLUMN_NAME FROM TABLE_NAME
            cmd.CommandType = CommandType.Text
            Dim rdr As SqlDataReader = cmd.ExecuteReader()

            While rdr.Read
                'Create object
                Dim newObject As New Object
                'Add properties
                'TODO - fix property names
                'newObject.PROPERTY_COLUMN_NAME = CStr(rdr("PROPERTY_COLUMN_NAME"))
                'newObject.PROPERTY_COLUMN_NAME = CStr(rdr("PROPERTY_COLUMN_NAME"))
                'newObject.PROPERTY_COLUMN_NAME = CStr(rdr("PROPERTY_COLUMN_NAME"))
                'TODO - add remaining properties
                'Add to list
                newlist.Add(CStr(rdr("KEY_COLUMN_NAME_HERE")), newObject)
            End While

            rdr.Close()

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

        Return newlist

    End Function

    Public Function GetDatasource() As SqlDataReader

        Dim conn As New SqlConnection(SQLConnectionString)

        Try
            conn.Open()

            Dim cmd As New SqlCommand("COMMAND_TEXT_HERE", conn)
            cmd.CommandType = CommandType.Text
            Dim rdr As SqlDataReader = cmd.ExecuteReader()

            Return rdr

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

    End Function

    Public Sub ExecuteCommand()

        Dim conn As New SqlConnection(SQLConnectionString)

        Try
            conn.Open()

            Dim cmd As New SqlCommand("COMMAND_TEXT_HERE", conn) 'DELETE* FROM TABLE_NAME
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

    End Sub

    Public Sub RunStoredProcedure()

        Dim conn As New SqlConnection(SQLConnectionString)

        Try
            conn.Open()

            Dim cmd As New SqlCommand("STORED_PROCEDURE_NAME_HERE", conn)
            cmd.CommandType = CommandType.StoredProcedure

            Dim param As New SqlParameter("PARAMETER_NAME_HERE", SqlDbType.VarChar)
            param.Value = "PARAMETER_VALUE_HERE"
            cmd.Parameters.Add(param)
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

    End Sub

    Public Function GetDataset() As DataSet

        Dim conn As New SqlConnection(SQLConnectionString)
        Dim ds As New DataSet

        Try
            conn.Open()

            Dim adapter As New SqlDataAdapter("SQL_HERE", conn) 'SELECT * FROM TABLE_NAME
            adapter.Fill(ds)

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

        Return ds

    End Function

    Public Sub InsertRecord(ByVal Record As Object)

        Dim conn As New SqlConnection(SQLConnectionString)

        Try
            conn.Open()

            Dim cmd As New SqlCommand("INSERT_COMMAND_HERE", conn) 'INSERT INTO TABLE_NAME (COLUMN_NAME, COLUMN_NAME, COLUMN_NAME) VALUES (VALUE, VALUE, VALUE)
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

    End Sub

    Public Sub UpdateRecord(ByVal Record As Object)

        Dim conn As New SqlConnection(SQLConnectionString)

        Try
            conn.Open()

            Dim cmd As New SqlCommand("UPDATE_COMMAND_HERE", conn) 'UPDATE TABLE_NAME SET COLUMN_NAME = VALUE, COLUMN_NAME = VALUE, COLUMN_NAME = VALUE WHERE COLUMN_NAME = FILTER
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

    End Sub
End Class
