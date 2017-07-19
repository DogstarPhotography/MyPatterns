Public Class Decorator
    Implements IComponent

    Public Function MyFunction() As String Implements IComponent.MyFunction

    End Function

    Public Sub MySub() Implements IComponent.MySub

    End Sub
End Class
