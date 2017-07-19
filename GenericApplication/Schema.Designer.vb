<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Schema
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
        Me.dgvSchema = New System.Windows.Forms.DataGridView
        Me.ColumnNameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TypeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.MaxCharacterLengthDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.IsNullableDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.TableColumnBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.dgvSchema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TableColumnBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvSchema
        '
        Me.dgvSchema.AllowUserToAddRows = False
        Me.dgvSchema.AllowUserToDeleteRows = False
        Me.dgvSchema.AllowUserToOrderColumns = True
        Me.dgvSchema.AllowUserToResizeRows = False
        Me.dgvSchema.AutoGenerateColumns = False
        Me.dgvSchema.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgvSchema.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSchema.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColumnNameDataGridViewTextBoxColumn, Me.TypeDataGridViewTextBoxColumn, Me.MaxCharacterLengthDataGridViewTextBoxColumn, Me.IsNullableDataGridViewCheckBoxColumn})
        Me.dgvSchema.DataSource = Me.TableColumnBindingSource
        Me.dgvSchema.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSchema.Location = New System.Drawing.Point(0, 0)
        Me.dgvSchema.MultiSelect = False
        Me.dgvSchema.Name = "dgvSchema"
        Me.dgvSchema.ReadOnly = True
        Me.dgvSchema.RowHeadersVisible = False
        Me.dgvSchema.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.dgvSchema.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvSchema.Size = New System.Drawing.Size(419, 284)
        Me.dgvSchema.TabIndex = 1
        '
        'ColumnNameDataGridViewTextBoxColumn
        '
        Me.ColumnNameDataGridViewTextBoxColumn.DataPropertyName = "ColumnName"
        Me.ColumnNameDataGridViewTextBoxColumn.HeaderText = "Column Name"
        Me.ColumnNameDataGridViewTextBoxColumn.Name = "ColumnNameDataGridViewTextBoxColumn"
        Me.ColumnNameDataGridViewTextBoxColumn.ReadOnly = True
        Me.ColumnNameDataGridViewTextBoxColumn.Width = 200
        '
        'TypeDataGridViewTextBoxColumn
        '
        Me.TypeDataGridViewTextBoxColumn.DataPropertyName = "Type"
        Me.TypeDataGridViewTextBoxColumn.HeaderText = "Type"
        Me.TypeDataGridViewTextBoxColumn.Name = "TypeDataGridViewTextBoxColumn"
        Me.TypeDataGridViewTextBoxColumn.ReadOnly = True
        Me.TypeDataGridViewTextBoxColumn.Width = 75
        '
        'MaxCharacterLengthDataGridViewTextBoxColumn
        '
        Me.MaxCharacterLengthDataGridViewTextBoxColumn.DataPropertyName = "MaxCharacterLength"
        Me.MaxCharacterLengthDataGridViewTextBoxColumn.HeaderText = "Max Chars"
        Me.MaxCharacterLengthDataGridViewTextBoxColumn.Name = "MaxCharacterLengthDataGridViewTextBoxColumn"
        Me.MaxCharacterLengthDataGridViewTextBoxColumn.ReadOnly = True
        Me.MaxCharacterLengthDataGridViewTextBoxColumn.Width = 80
        '
        'IsNullableDataGridViewCheckBoxColumn
        '
        Me.IsNullableDataGridViewCheckBoxColumn.DataPropertyName = "IsNullable"
        Me.IsNullableDataGridViewCheckBoxColumn.HeaderText = "Nullable"
        Me.IsNullableDataGridViewCheckBoxColumn.Name = "IsNullableDataGridViewCheckBoxColumn"
        Me.IsNullableDataGridViewCheckBoxColumn.ReadOnly = True
        Me.IsNullableDataGridViewCheckBoxColumn.Width = 50
        '
        'TableColumnBindingSource
        '
        Me.TableColumnBindingSource.DataSource = GetType(Sirius.TableColumn)
        '
        'Schema
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(419, 284)
        Me.Controls.Add(Me.dgvSchema)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Schema"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Schema"
        CType(Me.dgvSchema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TableColumnBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvSchema As System.Windows.Forms.DataGridView
    Friend WithEvents TableColumnBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents ColumnNameDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TypeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MaxCharacterLengthDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IsNullableDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn

End Class
