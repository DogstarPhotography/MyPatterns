Public Class GenericManager(Of W As {IWork(Of J, T, S, O), New}, J As IJob(Of T, S, O), T As ITask(Of S, O), S As IState(Of O), O)
    Implements IManage(Of W, J, T, S, O)

    Public Event Completed() Implements IManage(Of W, J, T, S, O).Completed
    Public Event Update(ByVal Message As String) Implements IManage(Of W, J, T, S, O).Update

    Private MyWorker As W
    Private MyConfigFile As String
    Private MyJobs As Generic.List(Of J)

    Public Sub New()
        MyJobs = New Generic.List(Of J)
    End Sub

    Public Sub Launch() Implements IManage(Of W, J, T, S, O).Launch
        'Fetch jobs and create worker
        MyWorker = New W
        AddHandler MyWorker.JobCompleted, AddressOf OnJobCompleted
        AddHandler MyWorker.WorkCompleted, AddressOf OnWorkCompleted
        'Fetch config file
        'TODO
        'Read config file
        'TODO
        'Build jobs and tasks from config
        'TODO
        'Start worker
        MyWorker.StartWork()
    End Sub

    Public Sub Cancel() Implements IManage(Of W, J, T, S, O).Cancel
        MyWorker.StopWork()
    End Sub

#Region "Events"
    Private Sub OnJobCompleted(ByVal Worker As IWork(Of J, T, S, O), ByVal Job As J)
        RaiseEvent Update("Worker " & Worker.Name & " completed Job " & Job.Name)
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

    Public Property Name() As String Implements IManage(Of W, J, T, S, O).Name
        Get
            Return Me.ToString
        End Get
        Set(ByVal value As String)

        End Set
    End Property
#End Region
End Class
