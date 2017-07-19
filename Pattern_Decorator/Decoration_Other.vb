Public Class Decoration_Other
    'Inherits Decorator
    Inherits Component

    'Instance of a Component that we are wrapping
    Public MyDecorator As Component

    Public Sub New(ByVal NewDecorator As Component)
        'Wrap the Component with this decoration
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
        Return MyBase.MyFunction & "other"
    End Function

End Class
