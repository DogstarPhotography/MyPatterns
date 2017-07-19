''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class ClassAdapter
    Implements ITarget
    Implements IWork
    'Private instance of adaptee
    Private MyAdaptee As IWork = New ClassAdaptee
    ''' <summary>
    ''' Adapt ITarget.Request into Adaptee.Use
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Request() Implements ITarget.Request
        DoWork()
    End Sub

    Public Sub DoWork() Implements IWork.DoWork
        MyAdaptee.DoWork()
    End Sub
End Class
