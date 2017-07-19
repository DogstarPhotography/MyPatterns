Public Class GenericJob(Of T As ITask(Of S, O), S As IState(Of O), O)
    Implements IJob(Of T, S, O)

    Public Event Completed(ByVal TheJob As IJob(Of T, S, O)) Implements IJob(Of T, S, O).Completed

    Private MyTasks As Generic.Queue(Of T)
    'Private MyRunTasks As Generic.Queue(Of T)
    Private CompleteFlag As Boolean = False

    Public Sub New()
        MyTasks = New Generic.Queue(Of T)
    End Sub

    Public Sub StartJob() Implements IJob(Of T, S, O).StartJob
        'BuildRunQueue()
        RunNextTask()
    End Sub

    Public Property Tasks() As System.Collections.Generic.Queue(Of T) Implements IJob(Of T, S, O).Tasks
        Get
            Return MyTasks
        End Get
        Set(ByVal value As System.Collections.Generic.Queue(Of T))
            MyTasks = value
        End Set
    End Property

    Public ReadOnly Property Complete() As Boolean Implements IJob(Of T, S, O).Complete
        Get
            Return CompleteFlag
        End Get
    End Property

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
            CompleteFlag = True
            RaiseEvent Completed(Me)
        End If
    End Sub

    'Private Sub BuildRunQueue()
    '    Do While MyTasks.Count > 0
    '        MyRunTasks.Enqueue(MyTasks.Dequeue)
    '    Loop
    'End Sub

End Class
