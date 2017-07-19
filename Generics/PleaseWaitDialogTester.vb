Public Class PleaseWaitDialogTester

    Public Shared Sub TestPleaseWaitDialog()

        Dim newDailog As New PleaseWaitDialog
        newDailog.Setup(AddressOf PleaseWaitTest, 10000, 1000)

    End Sub

    Public Shared Function PleaseWaitTest() As Boolean
        Static Counter As Integer = 0
        Counter = Counter + 1
        If Counter > 7 Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
