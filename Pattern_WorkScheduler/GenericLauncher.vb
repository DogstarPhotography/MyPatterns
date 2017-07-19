Public Class GenericLauncher(Of W As {IWork(Of J, T, S, O), New}, J As IJob(Of T, S, O), T As ITask(Of S, O), S As IState(Of O), O)
    Implements ILaunch(Of W, J, T, S, O)

    Private MyWorker As W
    Private MyConfigFile As String
    Private MyJobs As Generic.List(Of J)

    Public Sub New()
        MyJobs = New Generic.List(Of J)
    End Sub

    Public Sub Prepare() Implements ILaunch(Of W, J, T, S, O).Prepare
        'Fetch jobs and create worker
        MyWorker = New W
        'Fetch config file
        'TODO
        'Read config file
        'TODO
        'Build jobs and tasks from config
        'TODO
    End Sub

    Public Sub Launch() Implements ILaunch(Of W, J, T, S, O).Launch
        'Start worker
        MyWorker.StartWork()
    End Sub

    Public Sub Cancel() Implements ILaunch(Of W, J, T, S, O).Cancel
        MyWorker.StopWork()
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
End Class
