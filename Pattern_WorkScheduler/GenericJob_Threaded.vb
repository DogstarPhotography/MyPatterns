Imports System.ComponentModel
''' <summary>
''' 
''' </summary>
''' <typeparam name="T"></typeparam>
''' <typeparam name="S"></typeparam>
''' <typeparam name="O"></typeparam>
''' <remarks></remarks>
Public Class GenericJob_Threaded(Of T As ITask(Of S, O), S As IState(Of O), O)
    Implements IJob(Of T, S, O)

    Public Event Completed(ByVal TheJob As IJob(Of T, S, O)) Implements IJob(Of T, S, O).Completed
    Public Event Updated(ByVal TheJob As IJob(Of T, S, O)) Implements IJob(Of T, S, O).Updated

    Private MyTasks As Generic.Queue(Of T)
    Private MyStatus As IJob(Of T, S, O).JobStatus
    Private MyProgress As Integer

    Public Sub New()
        MyStatus = IJob(Of T, S, O).JobStatus.Waiting
        MyProgress = 0
        MyTasks = New Generic.Queue(Of T)
    End Sub

    Public Sub StartJob() Implements IJob(Of T, S, O).StartJob
        'Set status
        MyStatus = IJob(Of T, S, O).JobStatus.InProgress
        'Create and start a BackgroundWorker
        MyBackgroundWorker = New BackgroundWorker
        MyBackgroundWorker.RunWorkerAsync()
    End Sub

    Public Sub StopJob() Implements IJob(Of T, S, O).StopJob
        'Cancel the background worker
        MyBackgroundWorker.CancelAsync()
    End Sub

#Region "BackgroundWorker"
    'Background worker will do the actual work on a different thread
    Private WithEvents MyBackgroundWorker As BackgroundWorker

    Private Sub MyBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles MyBackgroundWorker.DoWork

        Dim numOriginalTasks As Integer
        Dim numCompletedTasks As Integer
        Dim curTask As T

        'Get number of original tasks and completed tasks for progress report
        numCompletedTasks = 0
        numOriginalTasks = MyTasks.Count

        'Process each of the tasks in the queue
        Do While MyTasks.Count > 0
            curTask = MyTasks.Dequeue
            'AddHandler curTask.Completed, AddressOf TaskCompletedHandler
            curTask.RunTask() 'Blocking call
            'Report progress
            numCompletedTasks = numCompletedTasks + 1
            MyBackgroundWorker.ReportProgress(CInt((numCompletedTasks / numOriginalTasks) * 100))
        Loop

        'Once we are done MyBackgroundWorker_RunWorkerCompleted will be raised so there is no need to do anything here
    End Sub

    Private Sub MyBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles MyBackgroundWorker.ProgressChanged
        'Report back up to our consumer
        MyStatus = IJob(Of T, S, O).JobStatus.InProgress
        MyProgress = e.ProgressPercentage
        RaiseEvent Updated(Me)
    End Sub

    Private Sub MyBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles MyBackgroundWorker.RunWorkerCompleted
        'All done
        MyStatus = IJob(Of T, S, O).JobStatus.Complete
        RaiseEvent Completed(Me)
    End Sub
#End Region

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

    Public ReadOnly Property Progress() As Integer
        Get
            Return MyProgress
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
