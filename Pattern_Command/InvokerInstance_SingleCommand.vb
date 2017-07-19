Public Class InvokerInstance_SingleCommand

    Private MyCommand As ICommand

    ''' <summary>
    ''' standard constructor
    ''' </summary>
    ''' <remarks>We deliberately create a NoCommand object internally so that calls to DoCommand and Undo will not fail</remarks>
    Public Sub New()
        MyCommand = New NoCommand
    End Sub

    Public Sub SetCommand(ByVal TheCommand As ICommand)
        MyCommand = TheCommand
    End Sub

    Public Sub DoCommand()
        MyCommand.Execute()
    End Sub

    Public Sub Undo()
        MyCommand.Undo()
    End Sub

End Class
