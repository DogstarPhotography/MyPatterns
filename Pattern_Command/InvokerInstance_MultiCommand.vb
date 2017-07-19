Public Class InvokerInstance_MultiCommand
    Private MyCommand As Dictionary(Of Integer, ICommand)
    Private LastKey As Integer
    ''' <summary>
    ''' standard constructor
    ''' </summary>
    ''' <remarks>We deliberately create a NoCommand object internally so that calls to DoCommand and Undo will not fail</remarks>
    Public Sub New()
        MyCommand = New Dictionary(Of Integer, ICommand)
        MyCommand.Add(0, New NoCommand)
        LastKey = 0
    End Sub

    Public Sub SetCommand(ByVal Key As Integer, ByVal TheCommand As ICommand)
        MyCommand.Add(Key, TheCommand)
        LastKey = Key
    End Sub

    Public Sub DoCommand(ByVal Key As Integer)
        MyCommand.Item(Key).Execute()
        LastKey = Key
    End Sub

    Public Sub Undo()
        MyCommand.Item(LastKey).Undo()
    End Sub
End Class
