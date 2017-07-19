Public Interface IManage(Of W As IWork(Of J, T, S, O), J As IJob(Of T, S, O), T As ITask(Of S, O), S As IState(Of O), O)

    Event Update(ByVal Message As String)
    Event Completed()

    Property Jobs() As Generic.List(Of J)
    Property Workers() As Generic.List(Of W)

    Property ConfigFile() As String

    Property Name() As String
    'Property Description() As String

    Sub Launch()
    Sub Cancel()

End Interface
