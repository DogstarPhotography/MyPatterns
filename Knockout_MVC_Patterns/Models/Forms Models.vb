Public Class FormData
    Inherits SelfSerializer

    Property Data As New List(Of DataItem)
End Class

Public Class DataItem
    Property Name As String = ""
    Property Value As String = ""
End Class

Public Class FormModel
    Property Content As String = ""
    Property Menu As String = ""
    Property State As String = ""
    Property Script As String = ""
End Class

