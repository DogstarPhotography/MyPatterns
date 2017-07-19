''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class ObjectAdapter
    Implements ITarget
    'Private instance of adaptee
    Private MyAdaptee As New ObjectAdaptee
    ''' <summary>
    ''' Adapt ITarget.Request into Adaptee.Use
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Request() Implements ITarget.Request
        MyAdaptee.Use()
    End Sub
End Class
