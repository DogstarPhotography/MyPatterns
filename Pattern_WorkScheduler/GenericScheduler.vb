Imports System.Threading

Public Class GenericScheduler(Of W As {IWork(Of J, T, S, O), New}, J As IJob(Of T, S, O), T As ITask(Of S, O), S As IState(Of O), O)
    Implements IManage(Of W, J, T, S, O)

    Public Event Completed() Implements IManage(Of W, J, T, S, O).Completed
    Public Event Update(ByVal Message As String) Implements IManage(Of W, J, T, S, O).Update

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

    Public Sub Launch() Implements IManage(Of W, J, T, S, O).Launch
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
        'Start the timer
        MyTimer.Change(0, MyPeriod) 'Calls StartWorker now and on every period
    End Sub

    Public Sub Cancel() Implements IManage(Of W, J, T, S, O).Cancel
        'Stop the timer
        MyTimer.Change(Timeout.Infinite, Timeout.Infinite)
        'Stop the worker
        MyWorker.StopWork()
    End Sub

    Private Sub StartWorker(ByVal State As Object)
        'Each time this comes around start a new worker and pass it the list of jobs to be done
        MyWorker = New W
        AddHandler MyWorker.JobCompleted, AddressOf OnJobCompleted
        AddHandler MyWorker.WorkCompleted, AddressOf OnWorkCompleted
        MyWorker.Jobs = MyJobs
        MyWorker.StartWork()
    End Sub

#Region "Events"
    Private Sub OnJobCompleted(ByVal Worker As IWork(Of J, T, S, O), ByVal Job As J)
        RaiseEvent Update("Worker " & Worker.Name & " completed Job " & Job.name)
    End Sub

    Private Sub OnWorkCompleted(ByVal Worker As IWork(Of J, T, S, O))
        RaiseEvent Update("Worker " & Worker.Name & " has completed all Jobs")
        RaiseEvent Completed()
    End Sub
#End Region

#Region "Properties"
    Public Property ConfigFile() As String Implements IManage(Of W, J, T, S, O).ConfigFile
        Get
            Return MyConfigFile
        End Get
        Set(ByVal value As String)
            MyConfigFile = value
        End Set
    End Property

    Public Property Jobs() As System.Collections.Generic.List(Of J) Implements IManage(Of W, J, T, S, O).Jobs
        Get
            Return MyJobs
        End Get
        Set(ByVal value As System.Collections.Generic.List(Of J))
            MyJobs = value
        End Set
    End Property

    Public Property Workers() As System.Collections.Generic.List(Of W) Implements IManage(Of W, J, T, S, O).Workers
        Get
            Dim newList As Generic.List(Of W) = New Generic.List(Of W)
            newList.Add(MyWorker)
            Return newList
        End Get
        Set(ByVal value As System.Collections.Generic.List(Of W))
            MyWorker = value.Item(0)
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

    Public Property Name() As String Implements IManage(Of W, J, T, S, O).Name
        Get
            Return Me.ToString
        End Get
        Set(ByVal value As String)

        End Set
    End Property
#End Region
End Class
