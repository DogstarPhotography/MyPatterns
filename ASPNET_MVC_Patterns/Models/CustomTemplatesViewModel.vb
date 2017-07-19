Public Class CustomTemplatesViewModel

    Property TableA As New DataListViewModel
    Property TableB As IEnumerable(Of DataItem)

End Class

Public Class DataListViewModel

    Property Title As String
    Property Items As New List(Of DataItem)

End Class

Public Class DataItem

    Property Name As String
    Property Description As String

End Class