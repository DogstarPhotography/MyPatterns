Imports System.Runtime.InteropServices

Public Class WindowsAPI

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function SetWindowPos(ByVal hWnd As IntPtr, _
        ByVal hWndInsertAfter As IntPtr, _
        ByVal X As Integer, _
        ByVal Y As Integer, _
        ByVal cx As Integer, _
        ByVal cy As Integer, _
        ByVal uFlags As UInteger) As Boolean
    End Function

    Public Const SWP_NOSIZE As Integer = &H1
    Public Const SWP_NOMOVE As Integer = &H2
    Public Const SWP_NOZORDER As Integer = &H4
    Public Const SWP_NOREDRAW As Integer = &H8
    Public Const SWP_NOACTIVATE As Integer = &H10
    Public Const SWP_DRAWFRAME As Integer = &H20
    Public Const SWP_FRAMECHANGED As Integer = &H20
    Public Const SWP_SHOWWINDOW As Integer = &H40
    Public Const SWP_HIDEWINDOW As Integer = &H80
    Public Const SWP_NOCOPYBITS As Integer = &H100
    Public Const SWP_NOOWNERZORDER As Integer = &H200
    Public Const SWP_NOREPOSITION As Integer = &H200
    Public Const SWP_NOSENDCHANGING As Integer = &H400
    Public Const SWP_DEFERERASE As Integer = &H2000
    Public Const SWP_ASYNCWINDOWPOS As Integer = &H4000

    Public Const HWND_TOP As Integer = 0
    Public Const HWND_BOTTOM As Integer = 1
    Public Const HWND_TOPMOST As Integer = -1
    Public Const HWND_NOTOPMOST As Integer = -2

End Class
