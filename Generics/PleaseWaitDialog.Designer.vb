<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PleaseWaitDialog
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PleaseWaitDialog))
        Me.picDistraction = New System.Windows.Forms.PictureBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.myTimer = New System.Windows.Forms.Timer(Me.components)
        CType(Me.picDistraction, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'picDistraction
        '
        Me.picDistraction.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.picDistraction.ErrorImage = CType(resources.GetObject("picDistraction.ErrorImage"), System.Drawing.Image)
        Me.picDistraction.Image = CType(resources.GetObject("picDistraction.Image"), System.Drawing.Image)
        Me.picDistraction.InitialImage = CType(resources.GetObject("picDistraction.InitialImage"), System.Drawing.Image)
        Me.picDistraction.Location = New System.Drawing.Point(12, 12)
        Me.picDistraction.Name = "picDistraction"
        Me.picDistraction.Size = New System.Drawing.Size(175, 100)
        Me.picDistraction.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.picDistraction.TabIndex = 0
        Me.picDistraction.TabStop = False
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(12, 118)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(175, 30)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'PleaseWaitDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(199, 160)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.picDistraction)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PleaseWaitDialog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Please Wait"
        CType(Me.picDistraction, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents picDistraction As System.Windows.Forms.PictureBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents myTimer As System.Windows.Forms.Timer
End Class
