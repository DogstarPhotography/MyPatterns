''' <summary>
''' This class deliberately does nothing
''' </summary>
''' <remarks></remarks>
Public Class NoCommand
    Implements ICommand

    Public Sub Execute() Implements ICommand.Execute

    End Sub

    Public Sub Undo() Implements ICommand.Undo

    End Sub
End Class
