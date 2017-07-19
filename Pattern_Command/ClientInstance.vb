Public Class ClientInstance

    Private MyReceiver As ReceiverInstance
    Private MyCommand As CommandInstance
    Private MyInvoker As InvokerInstance_SingleCommand

    Public Sub Test()

        'Create a receiver to do the work
        MyReceiver = New ReceiverInstance 'In the real world we might use a factory method here to create a ReceiverInstance
        'Create a command for the receiver
        MyCommand = New CommandInstance(MyReceiver)

        'Create an invoker
        MyInvoker = New InvokerInstance_SingleCommand
        'Give the invoker the command
        MyInvoker.SetCommand(MyCommand)

        'Invoker runs the command
        MyInvoker.DoCommand()

    End Sub

End Class
