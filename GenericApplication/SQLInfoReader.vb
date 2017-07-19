Imports Sirius.Core

Public Class SQLInfoReader
    Private _dbAccess As SQLAccess

    Public Function GetTablesandViews() As List(Of TableOrView)
        Dim list = New List(Of TableOrView)

        For Each row As DataRow In _dbAccess.GetDatatable(CommandType.Text, _
                "SELECT Name, 'Table' AS [Type] FROM sysobjects WHERE xtype='u' " & _
                "UNION " & _
                "SELECT name, 'View' AS [Type] FROM sysobjects WHERE xtype='v' " & _
                "ORDER BY name").Rows
            list.Add(New TableOrView(CStr(row("name")), CStr(row("type"))))
        Next
        Return list
    End Function

    Public Function GetTableColumns(ByVal tableName As String) As List(Of TableColumn)
        Dim list = New List(Of TableColumn)

        For Each row As DataRow In _dbAccess.GetDatatable(CommandType.Text, _
               "SELECT column_name,table_name,data_type,character_maximum_length,IS_NULLABLE " & _
               "FROM information_schema.columns " & _
               "WHERE table_name=@TableName " & _
               "ORDER BY ordinal_position", _
                New DBParameter() _
               {New DBParameter("@TableName", SqlDbType.VarChar, 50, tableName)}).Rows
            list.Add(New TableColumn(CStr(row("column_name")), CStr(row("table_name")), _
                                     CStr(row("data_type")), NullInt(row("character_maximum_length")), _
                                     CStr(row("IS_NULLABLE"))))
        Next
        Return list
    End Function

    Public Sub New(ByVal db As String)
        _dbAccess = ConnectionManager.GetInstance.GetSQLAccess(db)
    End Sub
End Class
