<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DAAB
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
		Me.dgvSampleData = New System.Windows.Forms.DataGridView()
		Me.GroupBox1 = New System.Windows.Forms.GroupBox()
		Me.btnAccessor = New System.Windows.Forms.Button()
		Me.btnDelete = New System.Windows.Forms.Button()
		Me.btnOutParameter = New System.Windows.Forms.Button()
		Me.btnSingleRow = New System.Windows.Forms.Button()
		Me.btnInsert = New System.Windows.Forms.Button()
		Me.btnScalar = New System.Windows.Forms.Button()
		Me.btnDataSetToGrid = New System.Windows.Forms.Button()
		Me.btnReader = New System.Windows.Forms.Button()
		Me.txtResults = New System.Windows.Forms.TextBox()
		CType(Me.dgvSampleData, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.GroupBox1.SuspendLayout()
		Me.SuspendLayout()
		'
		'dgvSampleData
		'
		Me.dgvSampleData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgvSampleData.Location = New System.Drawing.Point(358, 12)
		Me.dgvSampleData.Name = "dgvSampleData"
		Me.dgvSampleData.Size = New System.Drawing.Size(611, 279)
		Me.dgvSampleData.TabIndex = 0
		'
		'GroupBox1
		'
		Me.GroupBox1.Controls.Add(Me.btnAccessor)
		Me.GroupBox1.Controls.Add(Me.btnDelete)
		Me.GroupBox1.Controls.Add(Me.btnOutParameter)
		Me.GroupBox1.Controls.Add(Me.btnSingleRow)
		Me.GroupBox1.Controls.Add(Me.btnInsert)
		Me.GroupBox1.Controls.Add(Me.btnScalar)
		Me.GroupBox1.Controls.Add(Me.btnDataSetToGrid)
		Me.GroupBox1.Controls.Add(Me.btnReader)
		Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(340, 279)
		Me.GroupBox1.TabIndex = 1
		Me.GroupBox1.TabStop = False
		Me.GroupBox1.Text = "Examples"
		'
		'btnAccessor
		'
		Me.btnAccessor.Location = New System.Drawing.Point(6, 222)
		Me.btnAccessor.Name = "btnAccessor"
		Me.btnAccessor.Size = New System.Drawing.Size(328, 23)
		Me.btnAccessor.TabIndex = 6
		Me.btnAccessor.Text = "Accessor"
		Me.btnAccessor.UseVisualStyleBackColor = True
		'
		'btnDelete
		'
		Me.btnDelete.Location = New System.Drawing.Point(6, 193)
		Me.btnDelete.Name = "btnDelete"
		Me.btnDelete.Size = New System.Drawing.Size(328, 23)
		Me.btnDelete.TabIndex = 6
		Me.btnDelete.Text = "Delete"
		Me.btnDelete.UseVisualStyleBackColor = True
		'
		'btnOutParameter
		'
		Me.btnOutParameter.Location = New System.Drawing.Point(6, 164)
		Me.btnOutParameter.Name = "btnOutParameter"
		Me.btnOutParameter.Size = New System.Drawing.Size(328, 23)
		Me.btnOutParameter.TabIndex = 5
		Me.btnOutParameter.Text = "Get Out Parameter"
		Me.btnOutParameter.UseVisualStyleBackColor = True
		'
		'btnSingleRow
		'
		Me.btnSingleRow.Location = New System.Drawing.Point(6, 135)
		Me.btnSingleRow.Name = "btnSingleRow"
		Me.btnSingleRow.Size = New System.Drawing.Size(328, 23)
		Me.btnSingleRow.TabIndex = 4
		Me.btnSingleRow.Text = "Get Single Row"
		Me.btnSingleRow.UseVisualStyleBackColor = True
		'
		'btnInsert
		'
		Me.btnInsert.Location = New System.Drawing.Point(6, 106)
		Me.btnInsert.Name = "btnInsert"
		Me.btnInsert.Size = New System.Drawing.Size(328, 23)
		Me.btnInsert.TabIndex = 3
		Me.btnInsert.Text = "GetStoredProccommand Insert"
		Me.btnInsert.UseVisualStyleBackColor = True
		'
		'btnScalar
		'
		Me.btnScalar.Location = New System.Drawing.Point(6, 19)
		Me.btnScalar.Name = "btnScalar"
		Me.btnScalar.Size = New System.Drawing.Size(328, 23)
		Me.btnScalar.TabIndex = 2
		Me.btnScalar.Text = "ExecuteScalar"
		Me.btnScalar.UseVisualStyleBackColor = True
		'
		'btnDataSetToGrid
		'
		Me.btnDataSetToGrid.Location = New System.Drawing.Point(6, 48)
		Me.btnDataSetToGrid.Name = "btnDataSetToGrid"
		Me.btnDataSetToGrid.Size = New System.Drawing.Size(328, 23)
		Me.btnDataSetToGrid.TabIndex = 1
		Me.btnDataSetToGrid.Text = "DataSet Grid Bind"
		Me.btnDataSetToGrid.UseVisualStyleBackColor = True
		'
		'btnReader
		'
		Me.btnReader.Location = New System.Drawing.Point(6, 77)
		Me.btnReader.Name = "btnReader"
		Me.btnReader.Size = New System.Drawing.Size(328, 23)
		Me.btnReader.TabIndex = 0
		Me.btnReader.Text = "DataReader"
		Me.btnReader.UseVisualStyleBackColor = True
		'
		'txtResults
		'
		Me.txtResults.Location = New System.Drawing.Point(12, 297)
		Me.txtResults.Multiline = True
		Me.txtResults.Name = "txtResults"
		Me.txtResults.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.txtResults.Size = New System.Drawing.Size(957, 321)
		Me.txtResults.TabIndex = 3
		'
		'DAAB
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(982, 630)
		Me.Controls.Add(Me.txtResults)
		Me.Controls.Add(Me.GroupBox1)
		Me.Controls.Add(Me.dgvSampleData)
		Me.Name = "DAAB"
		Me.Text = "DAAB"
		CType(Me.dgvSampleData, System.ComponentModel.ISupportInitialize).EndInit()
		Me.GroupBox1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents dgvSampleData As System.Windows.Forms.DataGridView
	Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
	Friend WithEvents btnReader As System.Windows.Forms.Button
	Friend WithEvents txtResults As System.Windows.Forms.TextBox
	Friend WithEvents btnDataSetToGrid As System.Windows.Forms.Button
	Friend WithEvents btnScalar As System.Windows.Forms.Button
	Friend WithEvents btnInsert As System.Windows.Forms.Button
	Friend WithEvents btnSingleRow As System.Windows.Forms.Button
	Friend WithEvents btnOutParameter As System.Windows.Forms.Button
	Friend WithEvents btnDelete As System.Windows.Forms.Button
	Friend WithEvents btnAccessor As System.Windows.Forms.Button
End Class
