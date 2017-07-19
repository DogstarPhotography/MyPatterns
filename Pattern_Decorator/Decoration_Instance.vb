Public Class Decoration_Instance
    Inherits Decorator

    'Instance of a Component that we are wrapping
    Public MyDecorator As Decorator

    Public Sub New(ByVal NewDecorator As Decorator)
        'Wrap the instance with this decoration
        MyDecorator = NewDecorator
    End Sub

    Public Overloads Sub MySub()
        'Call base class method
        MyBase.MySub()
        'Add our own actions
        'CODE_HERE
    End Sub

    Public Overloads Function MyFunction() As String
        'Get value from base and add our own value
        Return MyBase.MyFunction & "decoration"
    End Function

End Class
