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
        Me.btnSerialize = New System.Windows.Forms.Button
        Me.txtXML = New System.Windows.Forms.TextBox
        Me.btnSave = New System.Windows.Forms.Button
        Me.dlgSaveFile = New System.Windows.Forms.SaveFileDialog
        Me.SuspendLayout()
        '
        'btnSerialize
        '
        Me.btnSerialize.Location = New System.Drawing.Point(12, 12)
        Me.btnSerialize.Name = "btnSerialize"
        Me.btnSerialize.Size = New System.Drawing.Size(75, 23)
        Me.btnSerialize.TabIndex = 0
        Me.btnSerialize.Text = "Serialize!"
        Me.btnSerialize.UseVisualStyleBackColor = True
        '
        'txtXML
        '
        Me.txtXML.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtXML.Location = New System.Drawing.Point(12, 41)
        Me.txtXML.Multiline = True
        Me.txtXML.Name = "txtXML"
        Me.txtXML.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtXML.Size = New System.Drawing.Size(501, 444)
        Me.txtXML.TabIndex = 1
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(438, 12)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 3
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dlgSaveFile
        '
        Me.dlgSaveFile.DefaultExt = "XML"
        '
        'Tester
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(525, 497)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.txtXML)
        Me.Controls.Add(Me.btnSerialize)
        Me.Name = "Tester"
        Me.Text = "Serialization Tester"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSerialize As System.Windows.Forms.Button
    Friend WithEvents txtXML As System.Windows.Forms.TextBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dlgSaveFile As System.Windows.Forms.SaveFileDialog

End Class
