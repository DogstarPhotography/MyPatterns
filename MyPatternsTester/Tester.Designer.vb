<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Tester
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
        Me.tabMain = New System.Windows.Forms.TabControl
        Me.tabReflection = New System.Windows.Forms.TabPage
        Me.txtReflection = New System.Windows.Forms.TextBox
        Me.btnReflection = New System.Windows.Forms.Button
        Me.tabThreading = New System.Windows.Forms.TabPage
        Me.txtMessages = New System.Windows.Forms.TextBox
        Me.btnClientController = New System.Windows.Forms.Button
        Me.tabSettings = New System.Windows.Forms.TabPage
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.btnSave = New System.Windows.Forms.Button
        Me.cboXMLSelector = New System.Windows.Forms.ComboBox
        Me.txtSerialization = New System.Windows.Forms.TextBox
        Me.btnSerialize = New System.Windows.Forms.Button
        Me.dlgSaveFile = New System.Windows.Forms.SaveFileDialog
        Me.tabMain.SuspendLayout()
        Me.tabReflection.SuspendLayout()
        Me.tabThreading.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabMain
        '
        Me.tabMain.Controls.Add(Me.tabReflection)
        Me.tabMain.Controls.Add(Me.tabThreading)
        Me.tabMain.Controls.Add(Me.tabSettings)
        Me.tabMain.Controls.Add(Me.TabPage1)
        Me.tabMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabMain.Location = New System.Drawing.Point(3, 3)
        Me.tabMain.Name = "tabMain"
        Me.tabMain.SelectedIndex = 0
        Me.tabMain.Size = New System.Drawing.Size(665, 562)
        Me.tabMain.TabIndex = 0
        '
        'tabReflection
        '
        Me.tabReflection.Controls.Add(Me.txtReflection)
        Me.tabReflection.Controls.Add(Me.btnReflection)
        Me.tabReflection.Location = New System.Drawing.Point(4, 22)
        Me.tabReflection.Name = "tabReflection"
        Me.tabReflection.Size = New System.Drawing.Size(657, 536)
        Me.tabReflection.TabIndex = 2
        Me.tabReflection.Text = "Reflection"
        Me.tabReflection.UseVisualStyleBackColor = True
        '
        'txtReflection
        '
        Me.txtReflection.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReflection.Location = New System.Drawing.Point(11, 40)
        Me.txtReflection.Multiline = True
        Me.txtReflection.Name = "txtReflection"
        Me.txtReflection.Size = New System.Drawing.Size(631, 481)
        Me.txtReflection.TabIndex = 1
        '
        'btnReflection
        '
        Me.btnReflection.AutoSize = True
        Me.btnReflection.Location = New System.Drawing.Point(11, 11)
        Me.btnReflection.Name = "btnReflection"
        Me.btnReflection.Size = New System.Drawing.Size(89, 23)
        Me.btnReflection.TabIndex = 0
        Me.btnReflection.Text = "Test Reflection"
        Me.btnReflection.UseVisualStyleBackColor = True
        '
        'tabThreading
        '
        Me.tabThreading.Controls.Add(Me.txtMessages)
        Me.tabThreading.Controls.Add(Me.btnClientController)
        Me.tabThreading.Location = New System.Drawing.Point(4, 22)
        Me.tabThreading.Name = "tabThreading"
        Me.tabThreading.Padding = New System.Windows.Forms.Padding(3)
        Me.tabThreading.Size = New System.Drawing.Size(657, 536)
        Me.tabThreading.TabIndex = 1
        Me.tabThreading.Text = "Threading"
        Me.tabThreading.UseVisualStyleBackColor = True
        '
        'txtMessages
        '
        Me.txtMessages.Location = New System.Drawing.Point(11, 40)
        Me.txtMessages.Multiline = True
        Me.txtMessages.Name = "txtMessages"
        Me.txtMessages.Size = New System.Drawing.Size(323, 271)
        Me.txtMessages.TabIndex = 1
        '
        'btnClientController
        '
        Me.btnClientController.AutoSize = True
        Me.btnClientController.Location = New System.Drawing.Point(11, 11)
        Me.btnClientController.Name = "btnClientController"
        Me.btnClientController.Size = New System.Drawing.Size(114, 23)
        Me.btnClientController.TabIndex = 0
        Me.btnClientController.Text = "Test Client Controller"
        Me.btnClientController.UseVisualStyleBackColor = True
        '
        'tabSettings
        '
        Me.tabSettings.Location = New System.Drawing.Point(4, 22)
        Me.tabSettings.Name = "tabSettings"
        Me.tabSettings.Padding = New System.Windows.Forms.Padding(3)
        Me.tabSettings.Size = New System.Drawing.Size(657, 536)
        Me.tabSettings.TabIndex = 0
        Me.tabSettings.Text = "Log && Settings"
        Me.tabSettings.UseVisualStyleBackColor = True
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.btnSave)
        Me.TabPage1.Controls.Add(Me.cboXMLSelector)
        Me.TabPage1.Controls.Add(Me.txtSerialization)
        Me.TabPage1.Controls.Add(Me.btnSerialize)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(657, 536)
        Me.TabPage1.TabIndex = 3
        Me.TabPage1.Text = "XML Serialization"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(576, 6)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 3
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'cboXMLSelector
        '
        Me.cboXMLSelector.FormattingEnabled = True
        Me.cboXMLSelector.Items.AddRange(New Object() {"SerializableComplexClass", "XXXs", "YYYs", "ZZZs", "XXX", "YYY", "ZZZ"})
        Me.cboXMLSelector.Location = New System.Drawing.Point(6, 8)
        Me.cboXMLSelector.Name = "cboXMLSelector"
        Me.cboXMLSelector.Size = New System.Drawing.Size(205, 21)
        Me.cboXMLSelector.TabIndex = 2
        Me.cboXMLSelector.Text = "SerializableComplexClass"
        '
        'txtSerialization
        '
        Me.txtSerialization.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSerialization.Location = New System.Drawing.Point(6, 35)
        Me.txtSerialization.Multiline = True
        Me.txtSerialization.Name = "txtSerialization"
        Me.txtSerialization.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSerialization.Size = New System.Drawing.Size(645, 495)
        Me.txtSerialization.TabIndex = 1
        '
        'btnSerialize
        '
        Me.btnSerialize.Location = New System.Drawing.Point(217, 6)
        Me.btnSerialize.Name = "btnSerialize"
        Me.btnSerialize.Size = New System.Drawing.Size(75, 23)
        Me.btnSerialize.TabIndex = 0
        Me.btnSerialize.Text = "Serialize"
        Me.btnSerialize.UseVisualStyleBackColor = True
        '
        'Tester
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(671, 568)
        Me.Controls.Add(Me.tabMain)
        Me.Name = "Tester"
        Me.Padding = New System.Windows.Forms.Padding(3)
        Me.Text = "Form1"
        Me.tabMain.ResumeLayout(False)
        Me.tabReflection.ResumeLayout(False)
        Me.tabReflection.PerformLayout()
        Me.tabThreading.ResumeLayout(False)
        Me.tabThreading.PerformLayout()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tabMain As System.Windows.Forms.TabControl
    Friend WithEvents tabSettings As System.Windows.Forms.TabPage
    Friend WithEvents tabThreading As System.Windows.Forms.TabPage
    Friend WithEvents btnClientController As System.Windows.Forms.Button
    Friend WithEvents txtMessages As System.Windows.Forms.TextBox
    Friend WithEvents tabReflection As System.Windows.Forms.TabPage
    Friend WithEvents txtReflection As System.Windows.Forms.TextBox
    Friend WithEvents btnReflection As System.Windows.Forms.Button
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents cboXMLSelector As System.Windows.Forms.ComboBox
    Friend WithEvents txtSerialization As System.Windows.Forms.TextBox
    Friend WithEvents btnSerialize As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dlgSaveFile As System.Windows.Forms.SaveFileDialog

End Class
