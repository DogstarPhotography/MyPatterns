Public Class Instance
    Implements IComponent

    Private MyValue As String

    Public Sub New()
        MyValue = "Instance"
    End Sub

    Public Function MyFunction() As String Implements IComponent.MyFunction
        Return MyValue
    End Function

    Public Sub MySub() Implements IComponent.MySub

    End Sub

End Class
