Public Class Facade
    'Facade is composed of various components
    Private MyComponentOne As New ComponentOne
    Private MyComponentTwo As New ComponentTwo
    ''' <summary>
    ''' Carry out a simple operation
    ''' </summary>
    ''' <remarks>The 'simple' operation consists of a complex set of instructions to components</remarks>
    Public Sub SimpleOperation()
        MyComponentOne.Operate()
        MyComponentTwo.Execute()
    End Sub
End Class
