Imports System.Data
Imports System.Data.OleDb
Imports System.Collections.Generic

Public Class OleDbClientSamples

    Public OleDbConnectionString As String = ""

    Public Function GetDataCollection() As SortedList(Of String, String)

        Dim conn As New OleDbConnection(OleDbConnectionString)
        Dim newlist As New SortedList(Of String, String)

        Try
            conn.Open()

            Dim cmd As New OleDbCommand("COMMAND_TEXT_HERE", conn) 'SELECT KEY_COLUMN_NAME, DATA_COLUMN_NAME FROM TABLE_NAME
            cmd.CommandType = CommandType.Text
            Dim rdr As OleDbDataReader = cmd.ExecuteReader()

            While rdr.Read
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

        Dim conn As New OleDbConnection(OleDbConnectionString)
        Dim newlist As New SortedList(Of String, Object)

        Try
            conn.Open()

            'TODO - Create SQL
            Dim cmd As New OleDbCommand("COMMAND_TEXT_HERE", conn) 'SELECT KEY_COLUMN_NAME, PROPERTY_COLUMN_NAME, PROPERTY_COLUMN_NAME, PROPERTY_COLUMN_NAME FROM TABLE_NAME
            cmd.CommandType = CommandType.Text
            Dim rdr As OleDbDataReader = cmd.ExecuteReader()

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

    Public Function GetDatasource() As OleDbDataReader

        Dim conn As New OleDbConnection(OleDbConnectionString)

        Try
            conn.Open()

            Dim cmd As New OleDbCommand("COMMAND_TEXT_HERE", conn)
            cmd.CommandType = CommandType.Text
            Dim rdr As OleDbDataReader = cmd.ExecuteReader()

            Return rdr

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

    End Function

    Public Sub ExecuteCommand()

        Dim conn As New OleDbConnection(OleDbConnectionString)

        Try
            conn.Open()

            Dim cmd As New OleDbCommand("COMMAND_TEXT_HERE", conn) 'DELETE* FROM TABLE_NAME
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

    End Sub

    Public Sub RunSP()

        Dim conn As New OleDbConnection(OleDbConnectionString)

        Try
            conn.Open()

            Dim cmd As New OleDbCommand("STORED_PROCEDURE_NAME_HERE", conn)
            cmd.CommandType = CommandType.StoredProcedure

            Dim param As New OleDbParameter("PARAMETER_NAME_HERE", OleDbType.VarChar)
            param.Value = "PARAMETER_VALUE_HERE"
            cmd.Parameters.Add(param)
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

    End Sub

    Public Function GetDataset() As DataSet

        Dim conn As New OleDbConnection(OleDbConnectionString)
        Dim ds As New DataSet

        Try
            conn.Open()

            Dim adapter As New OleDbDataAdapter("OleDb_HERE", conn) 'SELECT * FROM TABLE_NAME
            adapter.Fill(ds)

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

        Return ds

    End Function

    Public Function GetDatasetFromSP() As DataSet

        Dim conn As New OleDbConnection(OleDbConnectionString)
        Dim ds As New DataSet

        Try
            conn.Open()

            Dim cmd As New OleDbCommand("STORED_PROCEDURE_NAME_HERE", conn)
            cmd.CommandType = CommandType.StoredProcedure

            Dim param As New OleDbParameter("PARAMETER_NAME_HERE", OleDbType.VarChar)
            param.Value = "PARAMETER_VALUE_HERE"
            cmd.Parameters.Add(param)

            Dim adapter As New OleDbDataAdapter()
            adapter.SelectCommand = cmd
            adapter.Fill(ds)

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

        Return ds

    End Function

    Public Sub InsertRecord(ByVal Record As Object)

        Dim conn As New OleDbConnection(OleDbConnectionString)

        Try
            conn.Open()

            Dim cmd As New OleDbCommand("INSERT_COMMAND_HERE", conn) 'INSERT INTO TABLE_NAME (COLUMN_NAME, COLUMN_NAME, COLUMN_NAME) VALUES (VALUE, VALUE, VALUE)
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

    End Sub

    Public Sub UpdateRecord(ByVal Record As Object)

        Dim conn As New OleDbConnection(OleDbConnectionString)

        Try
            conn.Open()

            Dim cmd As New OleDbCommand("UPDATE_COMMAND_HERE", conn) 'UPDATE TABLE_NAME SET COLUMN_NAME = VALUE, COLUMN_NAME = VALUE, COLUMN_NAME = VALUE WHERE COLUMN_NAME = FILTER
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            conn.Close()

        End Try

    End Sub
End Class
