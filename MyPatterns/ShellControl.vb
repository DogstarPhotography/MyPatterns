Public Class ShellControl

    Public Sub Run(ByVal CommandLine As String)

        Shell(CommandLine, AppWinStyle.NormalFocus)

    End Sub

End Class
