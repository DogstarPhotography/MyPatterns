<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSetWindowPos
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.grpAPI = New System.Windows.Forms.GroupBox
        Me.btnSinkAPI = New System.Windows.Forms.Button
        Me.btnFloatAPI = New System.Windows.Forms.Button
        Me.grpDotNet = New System.Windows.Forms.GroupBox
        Me.btnSinkDotNET = New System.Windows.Forms.Button
        Me.btnFloatDotNET = New System.Windows.Forms.Button
        Me.grpAPI.SuspendLayout()
        Me.grpDotNet.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpAPI
        '
        Me.grpAPI.Controls.Add(Me.btnSinkAPI)
        Me.grpAPI.Controls.Add(Me.btnFloatAPI)
        Me.grpAPI.Location = New System.Drawing.Point(12, 12)
        Me.grpAPI.Name = "grpAPI"
        Me.grpAPI.Size = New System.Drawing.Size(92, 80)
        Me.grpAPI.TabIndex = 2
        Me.grpAPI.TabStop = False
        Me.grpAPI.Text = "API"
        '
        'btnSinkAPI
        '
        Me.btnSinkAPI.Location = New System.Drawing.Point(6, 48)
        Me.btnSinkAPI.Name = "btnSinkAPI"
        Me.btnSinkAPI.Size = New System.Drawing.Size(75, 23)
        Me.btnSinkAPI.TabIndex = 3
        Me.btnSinkAPI.Text = "Sink"
        Me.btnSinkAPI.UseVisualStyleBackColor = True
        '
        'btnFloatAPI
        '
        Me.btnFloatAPI.Location = New System.Drawing.Point(6, 19)
        Me.btnFloatAPI.Name = "btnFloatAPI"
        Me.btnFloatAPI.Size = New System.Drawing.Size(75, 23)
        Me.btnFloatAPI.TabIndex = 2
        Me.btnFloatAPI.Text = "Float"
        Me.btnFloatAPI.UseVisualStyleBackColor = True
        '
        'grpDotNet
        '
        Me.grpDotNet.Controls.Add(Me.btnSinkDotNET)
        Me.grpDotNet.Controls.Add(Me.btnFloatDotNET)
        Me.grpDotNet.Location = New System.Drawing.Point(110, 12)
        Me.grpDotNet.Name = "grpDotNet"
        Me.grpDotNet.Size = New System.Drawing.Size(92, 80)
        Me.grpDotNet.TabIndex = 3
        Me.grpDotNet.TabStop = False
        Me.grpDotNet.Text = "DotNET"
        '
        'btnSinkDotNET
        '
        Me.btnSinkDotNET.Location = New System.Drawing.Point(6, 48)
        Me.btnSinkDotNET.Name = "btnSinkDotNET"
        Me.btnSinkDotNET.Size = New System.Drawing.Size(75, 23)
        Me.btnSinkDotNET.TabIndex = 3
        Me.btnSinkDotNET.Text = "Sink"
        Me.btnSinkDotNET.UseVisualStyleBackColor = True
        '
        'btnFloatDotNET
        '
        Me.btnFloatDotNET.Location = New System.Drawing.Point(6, 19)
        Me.btnFloatDotNET.Name = "btnFloatDotNET"
        Me.btnFloatDotNET.Size = New System.Drawing.Size(75, 23)
        Me.btnFloatDotNET.TabIndex = 2
        Me.btnFloatDotNET.Text = "Float"
        Me.btnFloatDotNET.UseVisualStyleBackColor = True
        '
        'frmSetWindowPos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(211, 104)
        Me.Controls.Add(Me.grpDotNet)
        Me.Controls.Add(Me.grpAPI)
        Me.Name = "frmSetWindowPos"
        Me.Text = "SetWindowPos"
        Me.grpAPI.ResumeLayout(False)
        Me.grpDotNet.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grpAPI As System.Windows.Forms.GroupBox
    Friend WithEvents btnSinkAPI As System.Windows.Forms.Button
    Friend WithEvents btnFloatAPI As System.Windows.Forms.Button
    Friend WithEvents grpDotNet As System.Windows.Forms.GroupBox
    Friend WithEvents btnSinkDotNET As System.Windows.Forms.Button
    Friend WithEvents btnFloatDotNET As System.Windows.Forms.Button
End Class
