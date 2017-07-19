<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgWizard
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
        Me.pnlButtons = New System.Windows.Forms.TableLayoutPanel
        Me.btnFinish = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnBack = New System.Windows.Forms.Button
        Me.btnNext = New System.Windows.Forms.Button
        Me.tabSteps = New System.Windows.Forms.TabControl
        Me.tabStep1 = New System.Windows.Forms.TabPage
        Me.tabStep2 = New System.Windows.Forms.TabPage
        Me.tabStep3 = New System.Windows.Forms.TabPage
        Me.pnlButtons.SuspendLayout()
        Me.tabSteps.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlButtons
        '
        Me.pnlButtons.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlButtons.ColumnCount = 4
        Me.pnlButtons.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.pnlButtons.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.pnlButtons.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.pnlButtons.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.pnlButtons.Controls.Add(Me.btnNext, 0, 0)
        Me.pnlButtons.Controls.Add(Me.btnBack, 0, 0)
        Me.pnlButtons.Controls.Add(Me.btnFinish, 2, 0)
        Me.pnlButtons.Controls.Add(Me.btnCancel, 3, 0)
        Me.pnlButtons.Location = New System.Drawing.Point(106, 274)
        Me.pnlButtons.Name = "pnlButtons"
        Me.pnlButtons.RowCount = 1
        Me.pnlButtons.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.pnlButtons.Size = New System.Drawing.Size(317, 29)
        Me.pnlButtons.TabIndex = 0
        '
        'btnFinish
        '
        Me.btnFinish.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnFinish.Location = New System.Drawing.Point(164, 3)
        Me.btnFinish.Name = "btnFinish"
        Me.btnFinish.Size = New System.Drawing.Size(67, 23)
        Me.btnFinish.TabIndex = 0
        Me.btnFinish.Text = "Finish"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(243, 3)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(67, 23)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        '
        'btnBack
        '
        Me.btnBack.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnBack.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnBack.Location = New System.Drawing.Point(6, 3)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(67, 23)
        Me.btnBack.TabIndex = 2
        Me.btnBack.Text = "< Back"
        '
        'btnNext
        '
        Me.btnNext.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnNext.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnNext.Location = New System.Drawing.Point(85, 3)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(67, 23)
        Me.btnNext.TabIndex = 3
        Me.btnNext.Text = "Next >"
        '
        'tabSteps
        '
        Me.tabSteps.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.tabSteps.Controls.Add(Me.tabStep1)
        Me.tabSteps.Controls.Add(Me.tabStep2)
        Me.tabSteps.Controls.Add(Me.tabStep3)
        Me.tabSteps.Location = New System.Drawing.Point(12, 12)
        Me.tabSteps.Name = "tabSteps"
        Me.tabSteps.SelectedIndex = 0
        Me.tabSteps.Size = New System.Drawing.Size(411, 256)
        Me.tabSteps.TabIndex = 1
        '
        'tabStep1
        '
        Me.tabStep1.Location = New System.Drawing.Point(4, 25)
        Me.tabStep1.Name = "tabStep1"
        Me.tabStep1.Padding = New System.Windows.Forms.Padding(3)
        Me.tabStep1.Size = New System.Drawing.Size(403, 227)
        Me.tabStep1.TabIndex = 0
        Me.tabStep1.Text = "Step 1"
        Me.tabStep1.UseVisualStyleBackColor = True
        '
        'tabStep2
        '
        Me.tabStep2.Location = New System.Drawing.Point(4, 25)
        Me.tabStep2.Name = "tabStep2"
        Me.tabStep2.Padding = New System.Windows.Forms.Padding(3)
        Me.tabStep2.Size = New System.Drawing.Size(403, 227)
        Me.tabStep2.TabIndex = 1
        Me.tabStep2.Text = "Step 2"
        Me.tabStep2.UseVisualStyleBackColor = True
        '
        'tabStep3
        '
        Me.tabStep3.Location = New System.Drawing.Point(4, 25)
        Me.tabStep3.Name = "tabStep3"
        Me.tabStep3.Padding = New System.Windows.Forms.Padding(3)
        Me.tabStep3.Size = New System.Drawing.Size(403, 227)
        Me.tabStep3.TabIndex = 2
        Me.tabStep3.Text = "Step 3"
        Me.tabStep3.UseVisualStyleBackColor = True
        '
        'dlgWizard
        '
        Me.AcceptButton = Me.btnFinish
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(435, 315)
        Me.Controls.Add(Me.tabSteps)
        Me.Controls.Add(Me.pnlButtons)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgWizard"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Wizard"
        Me.pnlButtons.ResumeLayout(False)
        Me.tabSteps.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlButtons As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents btnFinish As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents tabSteps As System.Windows.Forms.TabControl
    Friend WithEvents tabStep1 As System.Windows.Forms.TabPage
    Friend WithEvents tabStep2 As System.Windows.Forms.TabPage
    Friend WithEvents tabStep3 As System.Windows.Forms.TabPage

End Class
