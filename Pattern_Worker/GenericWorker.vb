Public Class GenericWorker(Of J As IJob(Of T, S, O), T As ITask(Of S, O), S As IState(Of O), O)
    Implements IWork(Of J, T, S, O)

    Public Event Completed(ByVal Worker As IWork(Of J, T, S, O)) Implements IWork(Of J, T, S, O).Completed

    Private MyJobs As Generic.List(Of J)
    Private CompleteFlag As Boolean = False

    Public Property Jobs() As System.Collections.Generic.List(Of J) Implements IWork(Of J, T, S, O).Jobs
        Get
            Return MyJobs
        End Get
        Set(ByVal value As System.Collections.Generic.List(Of J))
            MyJobs = value
        End Set
    End Property

    Public ReadOnly Property Complete() As Boolean Implements IWork(Of J, T, S, O).Complete
        Get
            Return CompleteFlag
        End Get
    End Property

    Public Sub StartWork() Implements IWork(Of J, T, S, O).StartWork
        DoWork()
    End Sub

    Public Sub StopWork() Implements IWork(Of J, T, S, O).StopWork

    End Sub

    Private Sub DoWork()
        'Do each of the jobs in the list in no particular order
        For Each curJob As IJob(Of T, S, O) In MyJobs
            AddHandler curJob.Completed, AddressOf JobCompletedHandler
            curJob.StartJob()
        Next
    End Sub

    Private Sub JobCompletedHandler(ByVal TheJob As J)
        'Check each job
        For Each curJob As IJob(Of T, S, O) In MyJobs
            If curJob.Complete = False Then Exit Sub
        Next
        'If all jobs are complete set the flag and raise the event
        CompleteFlag = True
        RaiseEvent Completed(Me)
    End Sub
End Class
