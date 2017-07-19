<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NameValueBox
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.pnlTableLayout = New System.Windows.Forms.TableLayoutPanel
        Me.txtValue = New System.Windows.Forms.TextBox
        Me.lblName = New System.Windows.Forms.Label
        Me.pnlTableLayout.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTableLayout
        '
        Me.pnlTableLayout.BackColor = System.Drawing.SystemColors.Control
        Me.pnlTableLayout.ColumnCount = 2
        Me.pnlTableLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.pnlTableLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 66.66666!))
        Me.pnlTableLayout.Controls.Add(Me.txtValue, 1, 0)
        Me.pnlTableLayout.Controls.Add(Me.lblName, 0, 0)
        Me.pnlTableLayout.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlTableLayout.Location = New System.Drawing.Point(0, 0)
        Me.pnlTableLayout.Name = "pnlTableLayout"
        Me.pnlTableLayout.RowCount = 1
        Me.pnlTableLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.pnlTableLayout.Size = New System.Drawing.Size(195, 28)
        Me.pnlTableLayout.TabIndex = 0
        '
        'txtValue
        '
        Me.txtValue.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtValue.Location = New System.Drawing.Point(68, 4)
        Me.txtValue.Name = "txtValue"
        Me.txtValue.Size = New System.Drawing.Size(124, 20)
        Me.txtValue.TabIndex = 4
        Me.txtValue.Text = "Value"
        '
        'lblName
        '
        Me.lblName.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblName.AutoSize = True
        Me.lblName.Location = New System.Drawing.Point(3, 7)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(59, 13)
        Me.lblName.TabIndex = 5
        Me.lblName.Text = "Name"
        '
        'NameValueBox
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.pnlTableLayout)
        Me.Name = "NameValueBox"
        Me.Size = New System.Drawing.Size(195, 28)
        Me.pnlTableLayout.ResumeLayout(False)
        Me.pnlTableLayout.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlTableLayout As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents txtValue As System.Windows.Forms.TextBox
    Friend WithEvents lblName As System.Windows.Forms.Label

End Class
