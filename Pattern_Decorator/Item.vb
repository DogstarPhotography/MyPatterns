Public Class Item
    'Concrete instance of base class
    Inherits Component

    Private MyValue As String

    Public Sub New()
        MyValue = "ItemInstance"
    End Sub

    Public Overloads Sub MySub()
        'Do our own thing here
    End Sub

    Public Overloads Function MyFunction() As String
        'Return our value
        Return MyValue
    End Function

End Class
