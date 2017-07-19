<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMainModelFirst
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
		Me.grpNewMain = New System.Windows.Forms.GroupBox()
		Me.btnAddNew = New System.Windows.Forms.Button()
		Me.txtName = New System.Windows.Forms.TextBox()
		Me.lblName = New System.Windows.Forms.Label()
		Me.dgvData = New System.Windows.Forms.DataGridView()
		Me.dgvSubdata = New System.Windows.Forms.DataGridView()
		Me.dgvTertiary = New System.Windows.Forms.DataGridView()
		Me.dgvAdditional = New System.Windows.Forms.DataGridView()
		Me.grpNewMain.SuspendLayout()
		CType(Me.dgvData, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.dgvSubdata, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.dgvTertiary, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.dgvAdditional, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'grpNewMain
		'
		Me.grpNewMain.Controls.Add(Me.btnAddNew)
		Me.grpNewMain.Controls.Add(Me.txtName)
		Me.grpNewMain.Controls.Add(Me.lblName)
		Me.grpNewMain.Location = New System.Drawing.Point(12, 12)
		Me.grpNewMain.Name = "grpNewMain"
		Me.grpNewMain.Size = New System.Drawing.Size(264, 53)
		Me.grpNewMain.TabIndex = 0
		Me.grpNewMain.TabStop = False
		Me.grpNewMain.Text = "Add New Item"
		'
		'btnAddNew
		'
		Me.btnAddNew.Location = New System.Drawing.Point(157, 20)
		Me.btnAddNew.Name = "btnAddNew"
		Me.btnAddNew.Size = New System.Drawing.Size(100, 23)
		Me.btnAddNew.TabIndex = 4
		Me.btnAddNew.Text = "Add New"
		Me.btnAddNew.UseVisualStyleBackColor = True
		'
		'txtName
		'
		Me.txtName.Location = New System.Drawing.Point(51, 20)
		Me.txtName.Name = "txtName"
		Me.txtName.Size = New System.Drawing.Size(100, 20)
		Me.txtName.TabIndex = 2
		'
		'lblName
		'
		Me.lblName.AutoSize = True
		Me.lblName.Location = New System.Drawing.Point(7, 20)
		Me.lblName.Name = "lblName"
		Me.lblName.Size = New System.Drawing.Size(38, 13)
		Me.lblName.TabIndex = 0
		Me.lblName.Text = "Name:"
		'
		'dgvData
		'
		Me.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgvData.Location = New System.Drawing.Point(12, 71)
		Me.dgvData.Name = "dgvData"
		Me.dgvData.Size = New System.Drawing.Size(625, 211)
		Me.dgvData.TabIndex = 1
		'
		'dgvSubdata
		'
		Me.dgvSubdata.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgvSubdata.Location = New System.Drawing.Point(12, 288)
		Me.dgvSubdata.Name = "dgvSubdata"
		Me.dgvSubdata.Size = New System.Drawing.Size(625, 211)
		Me.dgvSubdata.TabIndex = 2
		'
		'dgvTertiary
		'
		Me.dgvTertiary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgvTertiary.Location = New System.Drawing.Point(12, 505)
		Me.dgvTertiary.Name = "dgvTertiary"
		Me.dgvTertiary.Size = New System.Drawing.Size(625, 211)
		Me.dgvTertiary.TabIndex = 3
		'
		'dgvAdditional
		'
		Me.dgvAdditional.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgvAdditional.Location = New System.Drawing.Point(643, 71)
		Me.dgvAdditional.Name = "dgvAdditional"
		Me.dgvAdditional.Size = New System.Drawing.Size(408, 645)
		Me.dgvAdditional.TabIndex = 4
		'
		'frmMainModelFirst
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(1063, 726)
		Me.Controls.Add(Me.dgvAdditional)
		Me.Controls.Add(Me.dgvTertiary)
		Me.Controls.Add(Me.dgvSubdata)
		Me.Controls.Add(Me.dgvData)
		Me.Controls.Add(Me.grpNewMain)
		Me.Name = "frmMainModelFirst"
		Me.Text = "Main"
		Me.grpNewMain.ResumeLayout(False)
		Me.grpNewMain.PerformLayout()
		CType(Me.dgvData, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.dgvSubdata, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.dgvTertiary, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.dgvAdditional, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)

	End Sub
	Friend WithEvents grpNewMain As System.Windows.Forms.GroupBox
	Friend WithEvents btnAddNew As System.Windows.Forms.Button
	Friend WithEvents txtName As System.Windows.Forms.TextBox
	Friend WithEvents lblName As System.Windows.Forms.Label
	Friend WithEvents dgvData As System.Windows.Forms.DataGridView
	Friend WithEvents dgvSubdata As System.Windows.Forms.DataGridView
	Friend WithEvents dgvTertiary As System.Windows.Forms.DataGridView
	Friend WithEvents dgvAdditional As System.Windows.Forms.DataGridView

End Class
