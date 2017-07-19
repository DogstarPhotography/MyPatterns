Public Interface ITask(Of S As IState(Of O), O)

    Event Completed(ByVal Task As ITask(Of S, O))

    Property State() As S

    ReadOnly Property Complete() As Boolean

    'Property TaskID() As Guid
    'Property Name() As String
    'Property Description() As String

    Sub RunTask() 'This is the procedure that actually does something, it may be blocking or asynchronous

End Interface
