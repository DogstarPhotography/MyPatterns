<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MyDictionaryEditor
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
        Me.ToolStripContainer1 = New System.Windows.Forms.ToolStripContainer
        Me.stpStatus = New System.Windows.Forms.StatusStrip
        Me.pnlItems = New System.Windows.Forms.FlowLayoutPanel
        Me.stpMain = New System.Windows.Forms.MenuStrip
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFileNew = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFileOpen = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFileSave = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFileExit = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuEdit = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuEditAdd = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuEditDelete = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuEditChange = New System.Windows.Forms.ToolStripMenuItem
        Me.dlgFileOpen = New System.Windows.Forms.OpenFileDialog
        Me.dlgFileSave = New System.Windows.Forms.SaveFileDialog
        Me.nvbItem = New MyPatterns.NameValueBox
        Me.ToolStripContainer1.BottomToolStripPanel.SuspendLayout()
        Me.ToolStripContainer1.ContentPanel.SuspendLayout()
        Me.ToolStripContainer1.TopToolStripPanel.SuspendLayout()
        Me.ToolStripContainer1.SuspendLayout()
        Me.stpMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStripContainer1
        '
        '
        'ToolStripContainer1.BottomToolStripPanel
        '
        Me.ToolStripContainer1.BottomToolStripPanel.Controls.Add(Me.stpStatus)
        '
        'ToolStripContainer1.ContentPanel
        '
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.pnlItems)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.nvbItem)
        Me.ToolStripContainer1.ContentPanel.Size = New System.Drawing.Size(741, 318)
        Me.ToolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStripContainer1.LeftToolStripPanelVisible = False
        Me.ToolStripContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStripContainer1.Name = "ToolStripContainer1"
        Me.ToolStripContainer1.RightToolStripPanelVisible = False
        Me.ToolStripContainer1.Size = New System.Drawing.Size(741, 364)
        Me.ToolStripContainer1.TabIndex = 0
        Me.ToolStripContainer1.Text = "ToolStripContainer1"
        '
        'ToolStripContainer1.TopToolStripPanel
        '
        Me.ToolStripContainer1.TopToolStripPanel.Controls.Add(Me.stpMain)
        '
        'stpStatus
        '
        Me.stpStatus.Dock = System.Windows.Forms.DockStyle.None
        Me.stpStatus.Location = New System.Drawing.Point(0, 0)
        Me.stpStatus.Name = "stpStatus"
        Me.stpStatus.Size = New System.Drawing.Size(741, 22)
        Me.stpStatus.TabIndex = 2
        '
        'pnlItems
        '
        Me.pnlItems.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlItems.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.pnlItems.Location = New System.Drawing.Point(0, 0)
        Me.pnlItems.Name = "pnlItems"
        Me.pnlItems.Size = New System.Drawing.Size(741, 318)
        Me.pnlItems.TabIndex = 3
        '
        'stpMain
        '
        Me.stpMain.Dock = System.Windows.Forms.DockStyle.None
        Me.stpMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuEdit})
        Me.stpMain.Location = New System.Drawing.Point(0, 0)
        Me.stpMain.Name = "stpMain"
        Me.stpMain.Size = New System.Drawing.Size(741, 24)
        Me.stpMain.TabIndex = 0
        Me.stpMain.Text = "MenuStrip1"
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileNew, Me.mnuFileOpen, Me.mnuFileSave, Me.mnuFileExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(35, 20)
        Me.mnuFile.Text = "&File"
        '
        'mnuFileNew
        '
        Me.mnuFileNew.Name = "mnuFileNew"
        Me.mnuFileNew.Size = New System.Drawing.Size(152, 22)
        Me.mnuFileNew.Text = "&New"
        '
        'mnuFileOpen
        '
        Me.mnuFileOpen.Name = "mnuFileOpen"
        Me.mnuFileOpen.Size = New System.Drawing.Size(152, 22)
        Me.mnuFileOpen.Text = "&Open"
        '
        'mnuFileSave
        '
        Me.mnuFileSave.Name = "mnuFileSave"
        Me.mnuFileSave.Size = New System.Drawing.Size(152, 22)
        Me.mnuFileSave.Text = "&Save"
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Name = "mnuFileExit"
        Me.mnuFileExit.Size = New System.Drawing.Size(152, 22)
        Me.mnuFileExit.Text = "E&xit"
        '
        'mnuEdit
        '
        Me.mnuEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuEditChange, Me.mnuEditAdd, Me.mnuEditDelete})
        Me.mnuEdit.Name = "mnuEdit"
        Me.mnuEdit.Size = New System.Drawing.Size(37, 20)
        Me.mnuEdit.Text = "&Edit"
        '
        'mnuEditAdd
        '
        Me.mnuEditAdd.Name = "mnuEditAdd"
        Me.mnuEditAdd.Size = New System.Drawing.Size(152, 22)
        Me.mnuEditAdd.Text = "&Add New"
        '
        'mnuEditDelete
        '
        Me.mnuEditDelete.Name = "mnuEditDelete"
        Me.mnuEditDelete.Size = New System.Drawing.Size(152, 22)
        Me.mnuEditDelete.Text = "&Delete"
        '
        'mnuEditChange
        '
        Me.mnuEditChange.Name = "mnuEditChange"
        Me.mnuEditChange.Size = New System.Drawing.Size(152, 22)
        Me.mnuEditChange.Text = "&Change Name"
        '
        'dlgFileOpen
        '
        Me.dlgFileOpen.DefaultExt = "xml"
        Me.dlgFileOpen.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
        '
        'dlgFileSave
        '
        Me.dlgFileSave.DefaultExt = "xml"
        Me.dlgFileSave.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
        '
        'nvbItem
        '
        Me.nvbItem.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.nvbItem.Location = New System.Drawing.Point(521, 73)
        Me.nvbItem.Name = "nvbItem"
        Me.nvbItem.NameText = ""
        Me.nvbItem.Size = New System.Drawing.Size(180, 28)
        Me.nvbItem.TabIndex = 1
        Me.nvbItem.ValueText = ""
        '
        'MyDictionaryEditor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(741, 364)
        Me.Controls.Add(Me.ToolStripContainer1)
        Me.MainMenuStrip = Me.stpMain
        Me.Name = "MyDictionaryEditor"
        Me.Text = "MyDictionaryEditor"
        Me.ToolStripContainer1.BottomToolStripPanel.ResumeLayout(False)
        Me.ToolStripContainer1.BottomToolStripPanel.PerformLayout()
        Me.ToolStripContainer1.ContentPanel.ResumeLayout(False)
        Me.ToolStripContainer1.TopToolStripPanel.ResumeLayout(False)
        Me.ToolStripContainer1.TopToolStripPanel.PerformLayout()
        Me.ToolStripContainer1.ResumeLayout(False)
        Me.ToolStripContainer1.PerformLayout()
        Me.stpMain.ResumeLayout(False)
        Me.stpMain.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolStripContainer1 As System.Windows.Forms.ToolStripContainer
    Friend WithEvents stpStatus As System.Windows.Forms.StatusStrip
    Friend WithEvents stpMain As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFileOpen As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFileSave As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFileExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFileNew As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents dlgFileOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents dlgFileSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents mnuEdit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuEditAdd As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuEditDelete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents nvbItem As MyPatterns.NameValueBox
    Friend WithEvents pnlItems As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents mnuEditChange As System.Windows.Forms.ToolStripMenuItem
End Class
