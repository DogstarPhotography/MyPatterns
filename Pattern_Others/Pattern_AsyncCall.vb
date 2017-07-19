''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class Pattern_AsyncCall

    'This procedure will block
    Public Sub Method_V0()
        'Do the work here
        'CODE_HERE
    End Sub

    'This function will not block
    Public Overloads Sub Method_V0Async()

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
        Method_V0()

        'If we need to communicate progress with the consumer (optional feature) then use the MyBackgroundWorker.ReportProgress method
        MyBackgroundWorker.ReportProgress(50)

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
