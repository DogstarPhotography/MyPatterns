<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Examples
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
        Me.grpExamples = New System.Windows.Forms.GroupBox()
        Me.btnHeirarchy = New System.Windows.Forms.Button()
        Me.txtOutput = New System.Windows.Forms.TextBox()
        Me.grpExamples.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpExamples
        '
        Me.grpExamples.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.grpExamples.Controls.Add(Me.btnHeirarchy)
        Me.grpExamples.Location = New System.Drawing.Point(12, 12)
        Me.grpExamples.Name = "grpExamples"
        Me.grpExamples.Size = New System.Drawing.Size(200, 726)
        Me.grpExamples.TabIndex = 0
        Me.grpExamples.TabStop = False
        Me.grpExamples.Text = "Examples"
        '
        'btnHeirarchy
        '
        Me.btnHeirarchy.Location = New System.Drawing.Point(6, 19)
        Me.btnHeirarchy.Name = "btnHeirarchy"
        Me.btnHeirarchy.Size = New System.Drawing.Size(188, 23)
        Me.btnHeirarchy.TabIndex = 0
        Me.btnHeirarchy.Text = "Heirarchy"
        Me.btnHeirarchy.UseVisualStyleBackColor = True
        '
        'txtOutput
        '
        Me.txtOutput.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOutput.Location = New System.Drawing.Point(218, 12)
        Me.txtOutput.Multiline = True
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.Size = New System.Drawing.Size(909, 726)
        Me.txtOutput.TabIndex = 1
        '
        'Examples
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1139, 750)
        Me.Controls.Add(Me.txtOutput)
        Me.Controls.Add(Me.grpExamples)
        Me.Name = "Examples"
        Me.Text = "Examples"
        Me.grpExamples.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents grpExamples As GroupBox
    Friend WithEvents btnHeirarchy As Button
    Friend WithEvents txtOutput As TextBox
End Class
