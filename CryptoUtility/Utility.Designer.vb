<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Utility
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtInput = New System.Windows.Forms.TextBox()
        Me.lblInput = New System.Windows.Forms.Label()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.lblSalt = New System.Windows.Forms.Label()
        Me.txtSalt = New System.Windows.Forms.TextBox()
        Me.btnPasswordFill = New System.Windows.Forms.Button()
        Me.btnSaltFill = New System.Windows.Forms.Button()
        Me.cboMethod = New System.Windows.Forms.ComboBox()
        Me.lblMethod = New System.Windows.Forms.Label()
        Me.lblOutput = New System.Windows.Forms.Label()
        Me.txtOutput = New System.Windows.Forms.TextBox()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.lblLength = New System.Windows.Forms.Label()
        Me.txtLength = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtInput
        '
        Me.txtInput.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInput.Location = New System.Drawing.Point(71, 12)
        Me.txtInput.Name = "txtInput"
        Me.txtInput.Size = New System.Drawing.Size(679, 20)
        Me.txtInput.TabIndex = 0
        '
        'lblInput
        '
        Me.lblInput.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblInput.AutoSize = True
        Me.lblInput.Location = New System.Drawing.Point(31, 15)
        Me.lblInput.Name = "lblInput"
        Me.lblInput.Size = New System.Drawing.Size(34, 13)
        Me.lblInput.TabIndex = 1
        Me.lblInput.Text = "Input:"
        Me.lblInput.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPassword
        '
        Me.lblPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPassword.AutoSize = True
        Me.lblPassword.Location = New System.Drawing.Point(9, 41)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(56, 13)
        Me.lblPassword.TabIndex = 3
        Me.lblPassword.Text = "Password:"
        Me.lblPassword.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPassword
        '
        Me.txtPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPassword.Location = New System.Drawing.Point(71, 38)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(598, 20)
        Me.txtPassword.TabIndex = 2
        '
        'lblSalt
        '
        Me.lblSalt.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSalt.AutoSize = True
        Me.lblSalt.Location = New System.Drawing.Point(37, 67)
        Me.lblSalt.Name = "lblSalt"
        Me.lblSalt.Size = New System.Drawing.Size(28, 13)
        Me.lblSalt.TabIndex = 5
        Me.lblSalt.Text = "Salt:"
        Me.lblSalt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSalt
        '
        Me.txtSalt.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSalt.Location = New System.Drawing.Point(71, 64)
        Me.txtSalt.Name = "txtSalt"
        Me.txtSalt.Size = New System.Drawing.Size(598, 20)
        Me.txtSalt.TabIndex = 4
        '
        'btnPasswordFill
        '
        Me.btnPasswordFill.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPasswordFill.Location = New System.Drawing.Point(675, 36)
        Me.btnPasswordFill.Name = "btnPasswordFill"
        Me.btnPasswordFill.Size = New System.Drawing.Size(75, 23)
        Me.btnPasswordFill.TabIndex = 6
        Me.btnPasswordFill.Text = "Fill"
        Me.btnPasswordFill.UseVisualStyleBackColor = True
        '
        'btnSaltFill
        '
        Me.btnSaltFill.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSaltFill.Location = New System.Drawing.Point(675, 61)
        Me.btnSaltFill.Name = "btnSaltFill"
        Me.btnSaltFill.Size = New System.Drawing.Size(75, 23)
        Me.btnSaltFill.TabIndex = 7
        Me.btnSaltFill.Text = "Fill"
        Me.btnSaltFill.UseVisualStyleBackColor = True
        '
        'cboMethod
        '
        Me.cboMethod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMethod.FormattingEnabled = True
        Me.cboMethod.Location = New System.Drawing.Point(71, 116)
        Me.cboMethod.Name = "cboMethod"
        Me.cboMethod.Size = New System.Drawing.Size(679, 21)
        Me.cboMethod.TabIndex = 8
        '
        'lblMethod
        '
        Me.lblMethod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblMethod.AutoSize = True
        Me.lblMethod.Location = New System.Drawing.Point(19, 119)
        Me.lblMethod.Name = "lblMethod"
        Me.lblMethod.Size = New System.Drawing.Size(46, 13)
        Me.lblMethod.TabIndex = 9
        Me.lblMethod.Text = "Method:"
        Me.lblMethod.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOutput
        '
        Me.lblOutput.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblOutput.AutoSize = True
        Me.lblOutput.Location = New System.Drawing.Point(23, 175)
        Me.lblOutput.Name = "lblOutput"
        Me.lblOutput.Size = New System.Drawing.Size(42, 13)
        Me.lblOutput.TabIndex = 11
        Me.lblOutput.Text = "Output:"
        Me.lblOutput.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOutput
        '
        Me.txtOutput.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOutput.Location = New System.Drawing.Point(71, 172)
        Me.txtOutput.Multiline = True
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.Size = New System.Drawing.Size(679, 283)
        Me.txtOutput.TabIndex = 10
        '
        'btnGenerate
        '
        Me.btnGenerate.Location = New System.Drawing.Point(71, 143)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(75, 23)
        Me.btnGenerate.TabIndex = 12
        Me.btnGenerate.Text = "Generate"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'lblLength
        '
        Me.lblLength.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblLength.AutoSize = True
        Me.lblLength.Location = New System.Drawing.Point(22, 93)
        Me.lblLength.Name = "lblLength"
        Me.lblLength.Size = New System.Drawing.Size(43, 13)
        Me.lblLength.TabIndex = 14
        Me.lblLength.Text = "Length:"
        Me.lblLength.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtLength
        '
        Me.txtLength.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLength.Location = New System.Drawing.Point(71, 90)
        Me.txtLength.MaxLength = 6
        Me.txtLength.Name = "txtLength"
        Me.txtLength.Size = New System.Drawing.Size(75, 20)
        Me.txtLength.TabIndex = 13
        Me.txtLength.Text = "256"
        '
        'Utility
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(762, 467)
        Me.Controls.Add(Me.lblLength)
        Me.Controls.Add(Me.txtLength)
        Me.Controls.Add(Me.btnGenerate)
        Me.Controls.Add(Me.lblOutput)
        Me.Controls.Add(Me.txtOutput)
        Me.Controls.Add(Me.lblMethod)
        Me.Controls.Add(Me.cboMethod)
        Me.Controls.Add(Me.btnSaltFill)
        Me.Controls.Add(Me.btnPasswordFill)
        Me.Controls.Add(Me.lblSalt)
        Me.Controls.Add(Me.txtSalt)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.lblInput)
        Me.Controls.Add(Me.txtInput)
        Me.Name = "Utility"
        Me.Text = "Crypto Utility"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtInput As TextBox
    Friend WithEvents lblInput As Label
    Friend WithEvents lblPassword As Label
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents lblSalt As Label
    Friend WithEvents txtSalt As TextBox
    Friend WithEvents btnPasswordFill As Button
    Friend WithEvents btnSaltFill As Button
    Friend WithEvents cboMethod As ComboBox
    Friend WithEvents lblMethod As Label
    Friend WithEvents lblOutput As Label
    Friend WithEvents txtOutput As TextBox
    Friend WithEvents btnGenerate As Button
    Friend WithEvents lblLength As Label
    Friend WithEvents txtLength As TextBox
End Class
