<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmKeyPreview
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.chkEnter = New System.Windows.Forms.CheckBox
        Me.chkKeyPreview = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Location = New System.Drawing.Point(9, 61)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(262, 241)
        Me.Label1.TabIndex = 0
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(12, 35)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 20)
        Me.TextBox1.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(118, 35)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(199, 35)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'chkEnter
        '
        Me.chkEnter.AutoSize = True
        Me.chkEnter.Checked = True
        Me.chkEnter.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkEnter.Location = New System.Drawing.Point(159, 12)
        Me.chkEnter.Name = "chkEnter"
        Me.chkEnter.Size = New System.Drawing.Size(115, 17)
        Me.chkEnter.TabIndex = 4
        Me.chkEnter.Text = "Intercept enter key"
        Me.chkEnter.UseVisualStyleBackColor = True
        '
        'chkKeyPreview
        '
        Me.chkKeyPreview.AutoSize = True
        Me.chkKeyPreview.Checked = True
        Me.chkKeyPreview.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkKeyPreview.Location = New System.Drawing.Point(12, 12)
        Me.chkKeyPreview.Name = "chkKeyPreview"
        Me.chkKeyPreview.Size = New System.Drawing.Size(108, 17)
        Me.chkKeyPreview.TabIndex = 5
        Me.chkKeyPreview.Text = "Form.KeyPreview"
        Me.chkKeyPreview.UseVisualStyleBackColor = True
        '
        'frmKeyPreview
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(281, 312)
        Me.Controls.Add(Me.chkKeyPreview)
        Me.Controls.Add(Me.chkEnter)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.KeyPreview = True
        Me.Name = "frmKeyPreview"
        Me.Text = "frmKeyPreview"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents chkEnter As System.Windows.Forms.CheckBox
    Friend WithEvents chkKeyPreview As System.Windows.Forms.CheckBox
End Class
