Imports System.Threading
''' <summary>
''' GenericWorker that repeats the jobs it is given on a regular basis
''' </summary>
''' <typeparam name="J">IJob(Of T, S, O)</typeparam>
''' <typeparam name="T">ITask(Of S, O)</typeparam>
''' <typeparam name="S">IState(Of O)</typeparam>
''' <typeparam name="O">Object</typeparam>
''' <remarks>Change the period property before calling startwork to set the repeat time in milliseconds</remarks>
Public Class GenericWorkRepeater(Of J As IJob(Of T, S, O), T As ITask(Of S, O), S As IState(Of O), O)
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
        MyTimer = New Timer(New TimerCallback(AddressOf DoWork), Nothing, Timeout.Infinite, Timeout.Infinite)
    End Sub

    Public Sub StartWork() Implements IWork(Of J, T, S, O).StartWork
        'Now start the timer
        MyTimer.Change(0, MyPeriod) 'Calls DoWork now and on every period
    End Sub

    Public Sub StopWork() Implements IWork(Of J, T, S, O).StopWork
        MyTimer.Change(Timeout.Infinite, Timeout.Infinite)
    End Sub

    Private Sub DoWork(ByVal State As Object)
        'Load mycurrentjobs from myjoblist
        For Each curJob As J In MyJobList
            MyCurrentJobs.Add(curJob)
        Next
        'Do each of the jobs in MyCurrentJobs in no particular order
        For Each curJob As J In MyCurrentJobs
            AddHandler curJob.Completed, AddressOf JobCompletedHandler
            curJob.StartJob()
        Next
    End Sub

    Private Sub JobCompletedHandler(ByVal TheJob As IJob(Of T, S, O))
        'Inform the consumer
        RaiseEvent JobCompleted(Me, CType(TheJob, J))
        'Check each J
        For Each curJob As J In MyCurrentJobs
            If curJob.Status <> IJob(Of T, S, O).JobStatus.Complete Then Exit Sub
        Next
        'If all jobs are complete check to see if we have scheduled jobs waiting then set the flag and raise the event
        MyCompleteFlag = True
        RaiseEvent WorkCompleted(Me)
    End Sub

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
