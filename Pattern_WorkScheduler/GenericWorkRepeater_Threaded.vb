Imports System.Threading
Imports System.ComponentModel
''' <summary>
''' GenericWorker that repeats the jobs it is given on a regular basis
''' </summary>
''' <typeparam name="J">IJob(Of T, S, O)</typeparam>
''' <typeparam name="T">ITask(Of S, O)</typeparam>
''' <typeparam name="S">IState(Of O)</typeparam>
''' <typeparam name="O">Object</typeparam>
''' <remarks>Multithreaded version of GenericWorkRepeater for improved responsiveness. Change the Period property before calling StartWork to set the repeat time in milliseconds</remarks>
Public Class GenericWorkRepeater_Threaded(Of J As IJob(Of T, S, O), T As ITask(Of S, O), S As IState(Of O), O)
    Implements IWork(Of J, T, S, O)

    Public Event JobCompleted(ByVal Worker As IWork(Of J, T, S, O), ByVal Job As J) Implements IWork(Of J, T, S, O).JobCompleted
    Public Event WorkCompleted(ByVal Worker As IWork(Of J, T, S, O)) Implements IWork(Of J, T, S, O).WorkCompleted

    Private MyJobList As Generic.List(Of J)
    Private MyCurrentJobs As Generic.List(Of J)
    Private MyCompleteFlag As Boolean = False
    Private MyTimer As Timer
    Private MyPeriod As Integer

    Public Sub New()
        MyCompleteFlag = False
        MyJobList = New Generic.List(Of J)
        MyCurrentJobs = New Generic.List(Of J)
        'Create a timer to schedule the jobs
        MyPeriod = Timeout.Infinite
        MyTimer = New Timer(New TimerCallback(AddressOf RepeatWork), Nothing, Timeout.Infinite, Timeout.Infinite)
    End Sub

    Public Sub StartWork() Implements IWork(Of J, T, S, O).StartWork
        'Now start the timer
        MyTimer.Change(0, MyPeriod) 'Calls DoWork now and on every period
    End Sub

    Public Sub StopWork() Implements IWork(Of J, T, S, O).StopWork
        Try
            'Stop the timer
            MyTimer.Change(Timeout.Infinite, Timeout.Infinite)
            'Stop the backgroundworker
            MyBackgroundWorker.CancelAsync()
        Catch ex As Exception
            'Ignore - the timer or background worker might not be instantiated
        End Try
    End Sub

    Private Sub RepeatWork(ByVal State As Object)
        'This is called by the timer in order to start and repeat the jobs to be done
        'Can't repeat work while previous work is running
        If Not MyBackgroundWorker Is Nothing Then
            If MyBackgroundWorker.IsBusy = True Then
                'Raise event/throw exception?
                Exit Sub
            End If
        End If
        'Reset
        MyCompleteFlag = False
        'Load mycurrentjobs from myjoblist
        For Each curJob As J In MyJobList
            MyCurrentJobs.Add(curJob)
        Next
        'Start the asynchronous operation.
        MyBackgroundWorker = New BackgroundWorker
        MyBackgroundWorker.RunWorkerAsync()
    End Sub

#Region "BackgroundWorker"
    'Background worker will do the actual work on a different thread
    Private WithEvents MyBackgroundWorker As BackgroundWorker
    Private numCompletedJobs As Integer

    Private Sub MyBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles MyBackgroundWorker.DoWork
        'Count number of completed jobs for reporting progress
        numCompletedJobs = 0
        'Do each of the jobs in the list in no particular order
        For Each curJob As IJob(Of T, S, O) In MyCurrentJobs
            AddHandler curJob.Completed, AddressOf JobCompletedHandler
            curJob.StartJob() 'Blocking or async call
        Next
        'Once the work is done MyBackgroundWorker_RunWorkerCompleted will be raised so there is no need to do anything here
    End Sub

    Private Sub JobCompletedHandler(ByVal TheJob As IJob(Of T, S, O))
        'Report progress
        numCompletedJobs = numCompletedJobs + 1
        MyBackgroundWorker.ReportProgress(CInt((numCompletedJobs / MyCurrentJobs.Count) * 100), TheJob)
    End Sub

    Private Sub MyBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles MyBackgroundWorker.ProgressChanged
        'Report progress to consumer
        RaiseEvent JobCompleted(Me, CType(e.UserState, J))
    End Sub

    Private Sub MyBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles MyBackgroundWorker.RunWorkerCompleted
        'As we never really complete, because we are a repeater, we never set flags or raise events here
    End Sub
#End Region

#Region "Properties"
    Public ReadOnly Property Complete() As Boolean Implements IWork(Of J, T, S, O).Complete
        Get
            Return MyCompleteFlag
        End Get
    End Property

    Public Property Jobs() As System.Collections.Generic.List(Of J) Implements IWork(Of J, T, S, O).Jobs
        Get
            Return MyJobList
        End Get
        Set(ByVal value As System.Collections.Generic.List(Of J))
            MyJobList = value
        End Set
    End Property

    Public Property Period() As Integer
        Get
            Return MyPeriod
        End Get
        Set(ByVal value As Integer)
            If value < 0 Then value = Timeout.Infinite
            MyPeriod = value
        End Set
    End Property

    Public Property Description() As String Implements IWork(Of J, T, S, O).Description
        Get
            Return Me.ToString
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Property Name() As String Implements IWork(Of J, T, S, O).Name
        Get
            Return Me.ToString
        End Get
        Set(ByVal value As String)

        End Set
    End Property
#End Region
End Class
