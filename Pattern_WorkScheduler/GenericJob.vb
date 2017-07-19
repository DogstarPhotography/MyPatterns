Public Class GenericJob(Of T As ITask(Of S, O), S As IState(Of O), O)
    Implements IJob(Of T, S, O)

    Public Event Completed(ByVal TheJob As IJob(Of T, S, O)) Implements IJob(Of T, S, O).Completed
    Public Event Updated(ByVal TheJob As IJob(Of T, S, O)) Implements IJob(Of T, S, O).Updated

    Private MyTasks As Generic.Queue(Of T)
    'Private MyRunTasks As Generic.Queue(Of T)
    Private MyStatus As IJob(Of T, S, O).JobStatus

    Public Sub New()
        MyStatus = IJob(Of T, S, O).JobStatus.Waiting
        MyTasks = New Generic.Queue(Of T)
    End Sub

    Public Sub StartJob() Implements IJob(Of T, S, O).StartJob
        'BuildRunQueue()
        MyStatus = IJob(Of T, S, O).JobStatus.InProgress
        RunNextTask()
    End Sub

    Public Sub StopJob() Implements IJob(Of T, S, O).StopJob

    End Sub

    'Private Sub RunNextTask()
    '    Dim curTask As T
    '    'Grab the next task in the queue and run it
    '    curTask = MyRunTasks.Dequeue
    '    AddHandler curTask.Completed, AddressOf TaskCompletedHandler
    '    curTask.RunTask()
    '    MyTasks.Enqueue(curTask)
    'End Sub

    Private Sub RunNextTask()
        Dim curTask As T
        'Grab the next task in the queue and run it
        curTask = MyTasks.Dequeue
        AddHandler curTask.Completed, AddressOf TaskCompletedHandler
        curTask.RunTask()
    End Sub

    Private Sub TaskCompletedHandler(ByVal Task As ITask(Of S, O))
        'Any more tasks to do?
        If MyTasks.Count > 0 Then
            RunNextTask()
        Else
            MyStatus = IJob(Of T, S, O).JobStatus.Complete
            RaiseEvent Completed(Me)
        End If
    End Sub

#Region "Properties"
    Public Property Tasks() As System.Collections.Generic.Queue(Of T) Implements IJob(Of T, S, O).Tasks
        Get
            Return MyTasks
        End Get
        Set(ByVal value As System.Collections.Generic.Queue(Of T))
            MyTasks = value
        End Set
    End Property

    Public ReadOnly Property Status() As IJob(Of T, S, O).JobStatus Implements IJob(Of T, S, O).Status
        Get
            Return MyStatus
        End Get
    End Property

    Public Property Description() As String Implements IJob(Of T, S, O).Description
        Get
            Return Me.ToString
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Property Name() As String Implements IJob(Of T, S, O).Name
        Get
            Return Me.ToString
        End Get
        Set(ByVal value As String)

        End Set
    End Property
#End Region
End Class
