Imports System.Data.SqlClient
Imports System.Runtime.CompilerServices

Public Module BackingStoreExtensions

    <Extension> Public Function SafeGetString(reader As IDataReader, colIndex As Integer) As String
        If reader.GetValue(colIndex) IsNot DBNull.Value Then
            Return reader.GetString(colIndex)
        Else
            Return String.Empty
        End If
    End Function

    '<Extension> Public Function SafeGetInt32(reader As IDataReader, colIndex As Integer) As Integer
    '    If reader.GetInt32(colIndex) IsNot DBNull.Value Then
    '        Return reader.GetString(colIndex)
    '    Else
    '        Return String.Empty
    '    End If
    'End Function

End Module
