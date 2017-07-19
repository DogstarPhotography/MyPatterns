<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMainInheritance
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
		Me.txtIdentifier = New System.Windows.Forms.TextBox()
		Me.lblIdentifier = New System.Windows.Forms.Label()
		Me.btnAddNew = New System.Windows.Forms.Button()
		Me.txtValue = New System.Windows.Forms.TextBox()
		Me.txtName = New System.Windows.Forms.TextBox()
		Me.lblValue = New System.Windows.Forms.Label()
		Me.lblName = New System.Windows.Forms.Label()
		Me.dgvData = New System.Windows.Forms.DataGridView()
		Me.grpNewTPH = New System.Windows.Forms.GroupBox()
		Me.txtTPHIdentifier = New System.Windows.Forms.TextBox()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.btnNewTPHItem = New System.Windows.Forms.Button()
		Me.txtTPHValue = New System.Windows.Forms.TextBox()
		Me.txtTPHName = New System.Windows.Forms.TextBox()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.dgvTPH = New System.Windows.Forms.DataGridView()
		Me.grpNewMain.SuspendLayout()
		CType(Me.dgvData, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.grpNewTPH.SuspendLayout()
		CType(Me.dgvTPH, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'grpNewMain
		'
		Me.grpNewMain.Controls.Add(Me.txtIdentifier)
		Me.grpNewMain.Controls.Add(Me.lblIdentifier)
		Me.grpNewMain.Controls.Add(Me.btnAddNew)
		Me.grpNewMain.Controls.Add(Me.txtValue)
		Me.grpNewMain.Controls.Add(Me.txtName)
		Me.grpNewMain.Controls.Add(Me.lblValue)
		Me.grpNewMain.Controls.Add(Me.lblName)
		Me.grpNewMain.Location = New System.Drawing.Point(12, 12)
		Me.grpNewMain.Name = "grpNewMain"
		Me.grpNewMain.Size = New System.Drawing.Size(160, 135)
		Me.grpNewMain.TabIndex = 0
		Me.grpNewMain.TabStop = False
		Me.grpNewMain.Text = "Add New TPT Primary"
		'
		'txtIdentifier
		'
		Me.txtIdentifier.Location = New System.Drawing.Point(51, 72)
		Me.txtIdentifier.Name = "txtIdentifier"
		Me.txtIdentifier.Size = New System.Drawing.Size(100, 20)
		Me.txtIdentifier.TabIndex = 6
		'
		'lblIdentifier
		'
		Me.lblIdentifier.AutoSize = True
		Me.lblIdentifier.Location = New System.Drawing.Point(6, 72)
		Me.lblIdentifier.Name = "lblIdentifier"
		Me.lblIdentifier.Size = New System.Drawing.Size(50, 13)
		Me.lblIdentifier.TabIndex = 5
		Me.lblIdentifier.Text = "Identifier:"
		'
		'btnAddNew
		'
		Me.btnAddNew.Location = New System.Drawing.Point(51, 98)
		Me.btnAddNew.Name = "btnAddNew"
		Me.btnAddNew.Size = New System.Drawing.Size(100, 23)
		Me.btnAddNew.TabIndex = 4
		Me.btnAddNew.Text = "Add New"
		Me.btnAddNew.UseVisualStyleBackColor = True
		'
		'txtValue
		'
		Me.txtValue.Location = New System.Drawing.Point(51, 46)
		Me.txtValue.Name = "txtValue"
		Me.txtValue.Size = New System.Drawing.Size(100, 20)
		Me.txtValue.TabIndex = 3
		'
		'txtName
		'
		Me.txtName.Location = New System.Drawing.Point(51, 20)
		Me.txtName.Name = "txtName"
		Me.txtName.Size = New System.Drawing.Size(100, 20)
		Me.txtName.TabIndex = 2
		'
		'lblValue
		'
		Me.lblValue.AutoSize = True
		Me.lblValue.Location = New System.Drawing.Point(6, 46)
		Me.lblValue.Name = "lblValue"
		Me.lblValue.Size = New System.Drawing.Size(37, 13)
		Me.lblValue.TabIndex = 1
		Me.lblValue.Text = "Value:"
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
		Me.dgvData.Location = New System.Drawing.Point(178, 12)
		Me.dgvData.Name = "dgvData"
		Me.dgvData.Size = New System.Drawing.Size(459, 135)
		Me.dgvData.TabIndex = 1
		'
		'grpNewTPH
		'
		Me.grpNewTPH.Controls.Add(Me.txtTPHIdentifier)
		Me.grpNewTPH.Controls.Add(Me.Label1)
		Me.grpNewTPH.Controls.Add(Me.btnNewTPHItem)
		Me.grpNewTPH.Controls.Add(Me.txtTPHValue)
		Me.grpNewTPH.Controls.Add(Me.txtTPHName)
		Me.grpNewTPH.Controls.Add(Me.Label2)
		Me.grpNewTPH.Controls.Add(Me.Label3)
		Me.grpNewTPH.Location = New System.Drawing.Point(12, 153)
		Me.grpNewTPH.Name = "grpNewTPH"
		Me.grpNewTPH.Size = New System.Drawing.Size(160, 135)
		Me.grpNewTPH.TabIndex = 2
		Me.grpNewTPH.TabStop = False
		Me.grpNewTPH.Text = "Add New TPH Primary"
		'
		'txtTPHIdentifier
		'
		Me.txtTPHIdentifier.Location = New System.Drawing.Point(51, 72)
		Me.txtTPHIdentifier.Name = "txtTPHIdentifier"
		Me.txtTPHIdentifier.Size = New System.Drawing.Size(100, 20)
		Me.txtTPHIdentifier.TabIndex = 6
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(6, 72)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(50, 13)
		Me.Label1.TabIndex = 5
		Me.Label1.Text = "Identifier:"
		'
		'btnNewTPHItem
		'
		Me.btnNewTPHItem.Location = New System.Drawing.Point(51, 98)
		Me.btnNewTPHItem.Name = "btnNewTPHItem"
		Me.btnNewTPHItem.Size = New System.Drawing.Size(100, 23)
		Me.btnNewTPHItem.TabIndex = 4
		Me.btnNewTPHItem.Text = "Add New"
		Me.btnNewTPHItem.UseVisualStyleBackColor = True
		'
		'txtTPHValue
		'
		Me.txtTPHValue.Location = New System.Drawing.Point(51, 46)
		Me.txtTPHValue.Name = "txtTPHValue"
		Me.txtTPHValue.Size = New System.Drawing.Size(100, 20)
		Me.txtTPHValue.TabIndex = 3
		'
		'txtTPHName
		'
		Me.txtTPHName.Location = New System.Drawing.Point(51, 20)
		Me.txtTPHName.Name = "txtTPHName"
		Me.txtTPHName.Size = New System.Drawing.Size(100, 20)
		Me.txtTPHName.TabIndex = 2
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Location = New System.Drawing.Point(6, 46)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(37, 13)
		Me.Label2.TabIndex = 1
		Me.Label2.Text = "Value:"
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Location = New System.Drawing.Point(7, 20)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(38, 13)
		Me.Label3.TabIndex = 0
		Me.Label3.Text = "Name:"
		'
		'dgvTPH
		'
		Me.dgvTPH.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgvTPH.Location = New System.Drawing.Point(178, 162)
		Me.dgvTPH.Name = "dgvTPH"
		Me.dgvTPH.Size = New System.Drawing.Size(459, 126)
		Me.dgvTPH.TabIndex = 3
		'
		'frmMain
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(649, 305)
		Me.Controls.Add(Me.dgvTPH)
		Me.Controls.Add(Me.grpNewTPH)
		Me.Controls.Add(Me.dgvData)
		Me.Controls.Add(Me.grpNewMain)
		Me.Name = "frmMain"
		Me.Text = "Main"
		Me.grpNewMain.ResumeLayout(False)
		Me.grpNewMain.PerformLayout()
		CType(Me.dgvData, System.ComponentModel.ISupportInitialize).EndInit()
		Me.grpNewTPH.ResumeLayout(False)
		Me.grpNewTPH.PerformLayout()
		CType(Me.dgvTPH, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)

	End Sub
	Friend WithEvents grpNewMain As System.Windows.Forms.GroupBox
	Friend WithEvents btnAddNew As System.Windows.Forms.Button
	Friend WithEvents txtValue As System.Windows.Forms.TextBox
	Friend WithEvents txtName As System.Windows.Forms.TextBox
	Friend WithEvents lblValue As System.Windows.Forms.Label
	Friend WithEvents lblName As System.Windows.Forms.Label
	Friend WithEvents dgvData As System.Windows.Forms.DataGridView
	Friend WithEvents txtIdentifier As System.Windows.Forms.TextBox
	Friend WithEvents lblIdentifier As System.Windows.Forms.Label
	Friend WithEvents grpNewTPH As System.Windows.Forms.GroupBox
	Friend WithEvents txtTPHIdentifier As System.Windows.Forms.TextBox
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents btnNewTPHItem As System.Windows.Forms.Button
	Friend WithEvents txtTPHValue As System.Windows.Forms.TextBox
	Friend WithEvents txtTPHName As System.Windows.Forms.TextBox
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents dgvTPH As System.Windows.Forms.DataGridView

End Class
