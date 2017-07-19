Public Class frmKeyPreview
    Inherits Windows.Forms.Form
    'Implements IMessageFilter

#Region "Key Interceptor"
    Private Const WM_KEYDOWN As Integer = &H100
    Private Const WM_SYSKEYDOWN As Integer = &H101

    Private WithEvents MyMessageFilter As SN_MessageFilter

    'This works but requires event handling
    Private Sub frmKeyPreview_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        MyMessageFilter = New SN_MessageFilter(Me)
        Application.AddMessageFilter(MyMessageFilter)
    End Sub

    Private Sub MyMessageFilter_MessageFiltered() Handles MyMessageFilter.MessageFiltered
        Label1.Text = Label1.Text & "CHK:ENTER "
    End Sub

    ''This will intercept keys for forms
    'Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
    '    Select Case msg.Msg
    '        Case WM_KEYDOWN
    '            If keyData = Keys.Return And chkEnter.Checked = True Then
    '                Label1.Text = Label1.Text & "CHK:" & keyData & " "
    '                Return True
    '            Else
    '                Return False
    '            End If
    '        Case WM_SYSKEYDOWN
    '            If keyData = Keys.Return And chkEnter.Checked = True Then
    '                Label1.Text = Label1.Text & "CHK:" & keyData & " "
    '                Return True
    '            Else
    '                Return False
    '            End If
    '    End Select
    'End Function


    'Doesn't work for forms
    'Public Function PreFilterMessage(ByRef m As System.Windows.Forms.Message) As Boolean Implements System.Windows.Forms.IMessageFilter.PreFilterMessage
    '    Select Case m.Msg
    '        Case WM_KEYDOWN
    '            Select Case m.WParam.ToInt32
    '                Case Keys.Return
    '                    If chkEnter.Checked = True Then
    '                        Label1.Text = Label1.Text & "CHK:ENTER "
    '                        Return True
    '                    Else
    '                        Return False
    '                    End If
    '                Case Else
    '                    Return False
    '            End Select
    '        Case Else
    '            Return False
    '    End Select
    'End Function

    'This doesn't work for a form
    'Protected Overrides Function IsInputChar(ByVal charCode As Char) As Boolean
    '    If charCode = Chr(13) And chkEnter.Checked = True Then
    '        Return True
    '    Else
    '        Return False
    '    End If
    'End Function

    ''This doesn't work for a form
    'Public Overrides Function PreProcessMessage(ByRef msg As Message) As Boolean
    '    Select Case msg.Msg
    '        Case WM_KEYDOWN
    '            Select Case msg.WParam.ToInt32
    '                Case Keys.Return
    '                    If chkEnter.Checked = True Then
    '                        Return True
    '                    Else
    '                        Return False
    '                    End If
    '            End Select
    '    End Select
    'End Function

    ''This doesn't work either
    'Protected Overrides Function ProcessKeyPreview(ByRef m As Message) As Boolean
    '    Select Case m.Msg
    '        Case WM_KEYDOWN
    '            Select Case m.WParam.ToInt32
    '                Case Keys.Return
    '                    If chkEnter.Checked = True Then
    '                        Return True
    '                    Else
    '                        Return False
    '                    End If
    '            End Select
    '        Case WM_SYSKEYDOWN
    '            Select Case m.WParam.ToInt32
    '                Case Keys.Return
    '                    If chkEnter.Checked = True Then
    '                        Return True
    '                    Else
    '                        Return False
    '                    End If
    '            End Select
    '    End Select
    'End Function

    'THIS DOESN'T WORK....
    'Private Sub frmKeyPreview_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles Me.PreviewKeyDown
    '    If e.KeyCode = Keys.Return And chkEnter.Checked = True Then
    '        Label1.Text = Label1.Text & "CHK:" & e.KeyCode & " "
    '        e.IsInputKey = False
    '    End If
    'End Sub

#End Region

    Private Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub chkKeyPreview_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkKeyPreview.CheckedChanged

        Me.KeyPreview = chkKeyPreview.Checked
        Label1.Text = Label1.Text & "PVW:" & Me.KeyPreview & " "
    End Sub

    Private Sub frmKeyPreview_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Label1.Text = Label1.Text & "v:" & e.KeyCode & " "
    End Sub

    Private Sub frmKeyPreview_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Label1.Text = Label1.Text & "P:" & e.KeyChar & " "
    End Sub

    Private Sub frmKeyPreview_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        Label1.Text = Label1.Text & "^:" & e.KeyCode & " "
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Label1.Text = Label1.Text & "P:Button1_Click "
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Label1.Text = Label1.Text & "P:Button2_Click "
    End Sub
End Class

''Sub Main... 
'Public Sub Main()
'    nMainForm = New frmMain

'    Dim nMessageFilter As New SN_MessageFilter(nMainForm)
'    Application.AddMessageFilter(nMessageFilter)

'    nMainForm.ShowDialog()

'    Application.RemoveMessageFilter(nMessageFilter)
'End Sub

Public Class SN_MessageFilter
    Implements IMessageFilter

    Public Event MessageFiltered()

    Private Const WM_KEYDOWN As Integer = &H100
    Private owner As Form

    Public Sub New(ByVal ownerForm As Form)
        owner = ownerForm
    End Sub

    Public Function PreFilterMessage(ByRef m As System.Windows.Forms.Message) As Boolean Implements IMessageFilter.PreFilterMessage
        Select Case m.Msg
            Case WM_KEYDOWN
                Select Case m.WParam.ToInt32
                    Case Keys.Return
                        RaiseEvent MessageFiltered()
                        Return True
                End Select
        End Select
        Return False
    End Function

End Class
