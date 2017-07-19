Public Class FreeClient

    Public Event WorkDone(ByVal Name As String)

    Public Name As String
    Public ClientController As Controller

    Private Sub New()
        'Disabled default constructor
    End Sub

    Public Sub New(ByVal ClientController As Controller, ByVal Name As String)

        Me.ClientController = ClientController
        Me.Name = Name

    End Sub

    Public Sub DoWork()

        'Simulate a blocking call
        Thread.Sleep(3000)
        RaiseEvent WorkDone(Me.Name)

    End Sub

    Public Sub DoWorkAsync()

        ' Start the asynchronous operation.
        MyBackgroundWorker = New BackgroundWorker
        MyBackgroundWorker.RunWorkerAsync()

    End Sub

#Region "BackgroundWorker"
    'Background worker will do the actual work on a different thread
    Private WithEvents MyBackgroundWorker As BackgroundWorker

    Private Sub MyBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles MyBackgroundWorker.DoWork

        ' This method performs the actual computation.
        ' It is executed on the worker thread.

        'Do the work here by calling the blocking procdure
        DoWork()

        'If we need to communicate progress with the consumer (optional feature) then use the MyBackgroundWorker.ReportProgress method
        'MyBackgroundWorker.ReportProgress(50)

        'Once the work is done MyBackgroundWorker_RunWorkerCompleted will be raised so there is no need to do anything here
    End Sub

    Private Sub MyBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles MyBackgroundWorker.ProgressChanged

        'Optionally, raise an event or do something else here
        '   Note that events raised here will be raised on the original thread, not the worker one
        'TODO

    End Sub

    Private Sub MyBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles MyBackgroundWorker.RunWorkerCompleted

        'Optionally, raise an event or do something else here
        '   Note that events raised here will be raised on the original thread, not the worker one
        'TODO

    End Sub
#End Region
End Class

Public Class LockClient

    Public Event WorkDone(ByVal Name As String)

    Public Name As String
    Public ClientController As Controller

    Private Sub New()
        'Disabled default constructor
    End Sub

    Public Sub New(ByVal ClientController As Controller, ByVal Name As String)

        Me.ClientController = ClientController
        Me.Name = Name

    End Sub

    Public Sub DoWork()

        'This client needs access to a shared resource so has to apply a synclock
        SyncLock ClientController.ControlKey
            'Simulate a blocking call
            Thread.Sleep(3000)
            RaiseEvent WorkDone(Me.Name)
        End SyncLock

    End Sub

    Public Sub DoWorkAsync()

        ' Start the asynchronous operation.
        MyBackgroundWorker = New BackgroundWorker
        MyBackgroundWorker.RunWorkerAsync()

    End Sub

#Region "BackgroundWorker"
    'Background worker will do the actual work on a different thread
    Private WithEvents MyBackgroundWorker As BackgroundWorker

    Private Sub MyBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles MyBackgroundWorker.DoWork

        ' This method performs the actual computation.
        ' It is executed on the worker thread.

        'Do the work here by calling the blocking procdure
        DoWork()

        'If we need to communicate progress with the consumer (optional feature) then use the MyBackgroundWorker.ReportProgress method
        'MyBackgroundWorker.ReportProgress(50)

        'Once the work is done MyBackgroundWorker_RunWorkerCompleted will be raised so there is no need to do anything here
    End Sub

    Private Sub MyBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles MyBackgroundWorker.ProgressChanged

        'Optionally, raise an event or do something else here
        '   Note that events raised here will be raised on the original thread, not the worker one
        'TODO

    End Sub

    Private Sub MyBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles MyBackgroundWorker.RunWorkerCompleted

        'Optionally, raise an event or do something else here
        '   Note that events raised here will be raised on the original thread, not the worker one
        'TODO

    End Sub
#End Region
End Class
