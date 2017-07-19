Imports System.Threading

Public Class GenericScheduler(Of W As {IWork(Of J, T, S, O), New}, J As IJob(Of T, S, O), T As ITask(Of S, O), S As IState(Of O), O)
    Implements ILaunch(Of W, J, T, S, O)

    Private MyWorker As W
    Private MyConfigFile As String
    Private MyJobs As Generic.List(Of J)

    Private MyTimer As Timer
    Private MyPeriod As Integer
    Private MySettings As Config

    Public Sub New()
        MyJobs = New Generic.List(Of J)
        'Create a timer to schedule the work
        MyTimer = New Timer(New TimerCallback(AddressOf StartWorker), Nothing, Timeout.Infinite, Timeout.Infinite)
    End Sub

    Public Sub Prepare() Implements ILaunch(Of W, J, T, S, O).Prepare
        'Fetch jobs and create worker
        MyWorker = New W
        'Fetch and read config file
        MySettings = New Config
        MySettings.LoadXML(MyConfigFile)
        'Build jobs and tasks from config
        'Get jobs
        'TODO
        'For each job get tasks
        'TODO
    End Sub

    Public Sub Launch() Implements ILaunch(Of W, J, T, S, O).Launch
        'Start the timer
        MyTimer.Change(0, MyPeriod) 'Calls StartWorker now and on every period
    End Sub

    Public Sub Cancel() Implements ILaunch(Of W, J, T, S, O).Cancel
        'Stop the timer
        MyTimer.Change(Timeout.Infinite, Timeout.Infinite)
        'Stop the worker
        MyWorker.StopWork()
    End Sub

    Private Sub StartWorker()
        'Each time this comes around start a new worker and pass it the list of jobs to be done
        MyWorker = New W
        MyWorker.Jobs = MyJobs
        MyWorker.StartWork()
    End Sub

    Public Property ConfigFile() As String Implements ILaunch(Of W, J, T, S, O).ConfigFile
        Get
            Return MyConfigFile
        End Get
        Set(ByVal value As String)
            MyConfigFile = value
        End Set
    End Property

    Public Property Jobs() As System.Collections.Generic.List(Of J) Implements ILaunch(Of W, J, T, S, O).Jobs
        Get
            Return MyJobs
        End Get
        Set(ByVal value As System.Collections.Generic.List(Of J))
            MyJobs = value
        End Set
    End Property

    Public Property Worker() As W Implements ILaunch(Of W, J, T, S, O).Worker
        Get
            Return MyWorker
        End Get
        Set(ByVal value As W)
            MyWorker = value
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
End Class
