Public Class frmSetWindowPos

    Private Sub btnFloatAPI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFloatAPI.Click

        Dim Flags As UInteger = WindowsAPI.SWP_NOSIZE Or WindowsAPI.SWP_NOMOVE

        WindowsAPI.SetWindowPos(Me.Handle, CType(WindowsAPI.HWND_TOPMOST, IntPtr), 0, 0, 0, 0, Flags)

    End Sub

    Private Sub btnSinkAPI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSinkAPI.Click

        Dim Flags As UInteger = WindowsAPI.SWP_NOSIZE Or WindowsAPI.SWP_NOMOVE

        WindowsAPI.SetWindowPos(Me.Handle, CType(WindowsAPI.HWND_NOTOPMOST, IntPtr), 0, 0, 0, 0, Flags)

    End Sub

    Private Sub btnFloatDotNET_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFloatDotNET.Click

        Me.TopMost = True

    End Sub

    Private Sub btnSinkDotNET_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSinkDotNET.Click

        Me.TopMost = False

    End Sub
End Class