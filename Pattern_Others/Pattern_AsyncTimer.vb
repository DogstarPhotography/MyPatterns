''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class Pattern_AsyncTimer

    'This procedure will run once
    Public Sub Method_V0()
        'Do the work here
        'CODE_HERE
    End Sub

    'This function will run again and again
    Public Overloads Sub Method_V0Async()

        ' Start the asynchronous operation.
        Launch()

    End Sub

#Region "Timer"
    Private MyTimer As System.Threading.Timer
    Private MyPeriod As Integer

    Public Property Period() As Integer
        Get
            Return MyPeriod
        End Get
        Set(ByVal value As Integer)
            MyPeriod = value
        End Set
    End Property

    Public Overridable Sub Launch()
        'Create timer
        MyTimer = New Timer(New TimerCallback(AddressOf DoWork), Nothing, Timeout.Infinite, Timeout.Infinite)
        'Start the timer
        MyTimer.Change(0, MyPeriod) 'Calls StartWorker now and on every period
    End Sub

    Public Overridable Sub Cancel()
        Try
            'Stop the timer
            MyTimer.Change(Timeout.Infinite, Timeout.Infinite)
        Catch ex As Exception
            'Ignore
        End Try
    End Sub

    Protected Overridable Sub DoWork(ByVal State As Object)
        'Call the method
        Method_V0()
    End Sub
#End Region
End Class
