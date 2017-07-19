Public Class TableOrView
    Private _Name As String
    Private _Type As DatabaseType

    Public Enum DatabaseType
        Table
        View
    End Enum

    Public ReadOnly Property Name() As String
        Get
            Return _Name
        End Get
    End Property

    Public ReadOnly Property [Type]() As DatabaseType
        Get
            Return _Type
        End Get
    End Property

    Public Sub New(ByVal name As String, ByVal [type] As String)
        _Name = name
        If [type] = "Table" Then
            _Type = DatabaseType.Table
        Else
            _Type = DatabaseType.View
        End If
    End Sub
End Class
