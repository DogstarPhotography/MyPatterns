Public Class CommandInstance
    Implements ICommand

    Private MyReceiver As ReceiverInstance

    Public Sub New(ByVal TheReceiver As ReceiverInstance)
        MyReceiver = TheReceiver
    End Sub

    ''' <summary>
    ''' Do our thing, whatever it is
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Execute() Implements ICommand.Execute
        MyReceiver.Action()
    End Sub
    ''' <summary>
    ''' Oops, undo our thing, whatever it was
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Undo() Implements ICommand.Undo
        MyReceiver.undo()
    End Sub

End Class
