<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbxEnvironment = New System.Windows.Forms.ComboBox
        Me.lblServer = New System.Windows.Forms.Label
        Me.cbxServer = New System.Windows.Forms.ComboBox
        Me.lbxTables = New System.Windows.Forms.ListBox
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.chkIncludeAttributes = New System.Windows.Forms.CheckBox
        Me.chkReadOnly = New System.Windows.Forms.CheckBox
        Me.chkIncludeNew = New System.Windows.Forms.CheckBox
        Me.pnlOptions = New System.Windows.Forms.Panel
        Me.btnShowReaderTemplate = New System.Windows.Forms.Button
        Me.btnSchema = New System.Windows.Forms.Button
        Me.pnlOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Environment"
        '
        'cbxEnvironment
        '
        Me.cbxEnvironment.FormattingEnabled = True
        Me.cbxEnvironment.Location = New System.Drawing.Point(85, 10)
        Me.cbxEnvironment.Name = "cbxEnvironment"
        Me.cbxEnvironment.Size = New System.Drawing.Size(114, 21)
        Me.cbxEnvironment.TabIndex = 1
        '
        'lblServer
        '
        Me.lblServer.AutoSize = True
        Me.lblServer.Location = New System.Drawing.Point(16, 47)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(38, 13)
        Me.lblServer.TabIndex = 2
        Me.lblServer.Text = "Server"
        '
        'cbxServer
        '
        Me.cbxServer.FormattingEnabled = True
        Me.cbxServer.Location = New System.Drawing.Point(85, 44)
        Me.cbxServer.Name = "cbxServer"
        Me.cbxServer.Size = New System.Drawing.Size(114, 21)
        Me.cbxServer.TabIndex = 3
        '
        'lbxTables
        '
        Me.lbxTables.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbxTables.FormattingEnabled = True
        Me.lbxTables.Location = New System.Drawing.Point(12, 80)
        Me.lbxTables.Name = "lbxTables"
        Me.lbxTables.Size = New System.Drawing.Size(187, 407)
        Me.lbxTables.TabIndex = 4
        '
        'txtCode
        '
        Me.txtCode.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCode.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCode.Location = New System.Drawing.Point(218, 47)
        Me.txtCode.Multiline = True
        Me.txtCode.Name = "txtCode"
        Me.txtCode.ReadOnly = True
        Me.txtCode.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtCode.Size = New System.Drawing.Size(536, 440)
        Me.txtCode.TabIndex = 5
        '
        'chkIncludeAttributes
        '
        Me.chkIncludeAttributes.AutoSize = True
        Me.chkIncludeAttributes.Checked = True
        Me.chkIncludeAttributes.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkIncludeAttributes.Location = New System.Drawing.Point(3, 5)
        Me.chkIncludeAttributes.Name = "chkIncludeAttributes"
        Me.chkIncludeAttributes.Size = New System.Drawing.Size(108, 17)
        Me.chkIncludeAttributes.TabIndex = 6
        Me.chkIncludeAttributes.Text = "Include Attributes"
        Me.chkIncludeAttributes.UseVisualStyleBackColor = True
        '
        'chkReadOnly
        '
        Me.chkReadOnly.AutoSize = True
        Me.chkReadOnly.Location = New System.Drawing.Point(114, 5)
        Me.chkReadOnly.Name = "chkReadOnly"
        Me.chkReadOnly.Size = New System.Drawing.Size(123, 17)
        Me.chkReadOnly.TabIndex = 7
        Me.chkReadOnly.Text = "ReadOnly Properties"
        Me.chkReadOnly.UseVisualStyleBackColor = True
        '
        'chkIncludeNew
        '
        Me.chkIncludeNew.AutoSize = True
        Me.chkIncludeNew.Location = New System.Drawing.Point(239, 5)
        Me.chkIncludeNew.Name = "chkIncludeNew"
        Me.chkIncludeNew.Size = New System.Drawing.Size(108, 17)
        Me.chkIncludeNew.TabIndex = 8
        Me.chkIncludeNew.Text = "Incldue New Sub"
        Me.chkIncludeNew.UseVisualStyleBackColor = True
        '
        'pnlOptions
        '
        Me.pnlOptions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlOptions.Controls.Add(Me.btnSchema)
        Me.pnlOptions.Controls.Add(Me.btnShowReaderTemplate)
        Me.pnlOptions.Controls.Add(Me.chkIncludeAttributes)
        Me.pnlOptions.Controls.Add(Me.chkIncludeNew)
        Me.pnlOptions.Controls.Add(Me.chkReadOnly)
        Me.pnlOptions.Enabled = False
        Me.pnlOptions.Location = New System.Drawing.Point(218, 10)
        Me.pnlOptions.Name = "pnlOptions"
        Me.pnlOptions.Size = New System.Drawing.Size(536, 27)
        Me.pnlOptions.TabIndex = 9
        '
        'btnShowReaderTemplate
        '
        Me.btnShowReaderTemplate.Location = New System.Drawing.Point(431, 3)
        Me.btnShowReaderTemplate.Name = "btnShowReaderTemplate"
        Me.btnShowReaderTemplate.Size = New System.Drawing.Size(102, 23)
        Me.btnShowReaderTemplate.TabIndex = 9
        Me.btnShowReaderTemplate.Text = "Reader Template"
        Me.btnShowReaderTemplate.UseVisualStyleBackColor = True
        '
        'btnSchema
        '
        Me.btnSchema.Location = New System.Drawing.Point(353, 3)
        Me.btnSchema.Name = "btnSchema"
        Me.btnSchema.Size = New System.Drawing.Size(72, 23)
        Me.btnSchema.TabIndex = 10
        Me.btnSchema.Text = "Schema"
        Me.btnSchema.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(766, 499)
        Me.Controls.Add(Me.pnlOptions)
        Me.Controls.Add(Me.txtCode)
        Me.Controls.Add(Me.lbxTables)
        Me.Controls.Add(Me.cbxServer)
        Me.Controls.Add(Me.lblServer)
        Me.Controls.Add(Me.cbxEnvironment)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "CreateTableObject"
        Me.pnlOptions.ResumeLayout(False)
        Me.pnlOptions.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbxEnvironment As System.Windows.Forms.ComboBox
    Friend WithEvents lblServer As System.Windows.Forms.Label
    Friend WithEvents cbxServer As System.Windows.Forms.ComboBox
    Friend WithEvents lbxTables As System.Windows.Forms.ListBox
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents chkIncludeAttributes As System.Windows.Forms.CheckBox
    Friend WithEvents chkReadOnly As System.Windows.Forms.CheckBox
    Friend WithEvents chkIncludeNew As System.Windows.Forms.CheckBox
    Friend WithEvents pnlOptions As System.Windows.Forms.Panel
    Friend WithEvents btnShowReaderTemplate As System.Windows.Forms.Button
    Friend WithEvents btnSchema As System.Windows.Forms.Button

End Class
