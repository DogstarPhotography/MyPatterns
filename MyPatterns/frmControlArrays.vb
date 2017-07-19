Public Class frmControlArrays
    ''' <summary>
    ''' Load some controls
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmControlArrays_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ctrContainers As Integer
        Dim ctrControls As Integer
        Dim curContainer As Windows.Forms.Control
        Dim curControl As Windows.Forms.Control
        'Add containers and controls
        For ctrContainers = 1 To 3
            'Add a container
            curContainer = New GroupBox
            'Set some properties
            curContainer.Text = "Container " & ctrContainers.ToString
            'Position the container
            curContainer.Top = 6
            curContainer.Left = ((ctrContainers - 1) * curContainer.Width) + (ctrContainers * 12)
            For ctrControls = 1 To 4
                'Add a control
                curControl = New Label
                'Set some properties
                curControl.Text = "Control " & ctrControls.ToString
                'Position the control
                curControl.Left = 6
                curControl.Top = ((ctrControls - 1) * curControl.Height) + (ctrControls * 3) + 13
                'Add an event handler
                AddHandler curControl.Click, AddressOf ControlEventHandler
                'Add the control to the container
                curContainer.Controls.Add(curControl)
            Next
            'Set the container height
            curContainer.Height = curControl.Top + curControl.Height + 6
            'Add the container to the controls collection of the form
            Me.Controls.Add(curContainer)
        Next

    End Sub
    ''' <summary>
    ''' Event handler for controls
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ControlEventHandler(ByVal sender As Object, ByVal e As System.EventArgs)
        'CODE_HERE
    End Sub

End Class