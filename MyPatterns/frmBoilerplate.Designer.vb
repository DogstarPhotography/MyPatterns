<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBoilerplate
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
        Me.ctlControl = New System.Windows.Forms.Label
        Me.strTools = New System.Windows.Forms.ToolStripContainer
        Me.strStatus = New System.Windows.Forms.StatusStrip
        Me.mnuStrip = New System.Windows.Forms.MenuStrip
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFileLoad = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFileSave = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFileExit = New System.Windows.Forms.ToolStripMenuItem
        Me.dlgFileOpen = New System.Windows.Forms.OpenFileDialog
        Me.dlgFileSave = New System.Windows.Forms.SaveFileDialog
        Me.strTools.BottomToolStripPanel.SuspendLayout()
        Me.strTools.ContentPanel.SuspendLayout()
        Me.strTools.TopToolStripPanel.SuspendLayout()
        Me.strTools.SuspendLayout()
        Me.mnuStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'ctlControl
        '
        Me.ctlControl.AutoSize = True
        Me.ctlControl.Location = New System.Drawing.Point(12, 9)
        Me.ctlControl.Name = "ctlControl"
        Me.ctlControl.Size = New System.Drawing.Size(51, 13)
        Me.ctlControl.TabIndex = 0
        Me.ctlControl.Text = "ctlControl"
        '
        'strTools
        '
        '
        'strTools.BottomToolStripPanel
        '
        Me.strTools.BottomToolStripPanel.Controls.Add(Me.strStatus)
        '
        'strTools.ContentPanel
        '
        Me.strTools.ContentPanel.AutoScroll = True
        Me.strTools.ContentPanel.Controls.Add(Me.ctlControl)
        Me.strTools.ContentPanel.Size = New System.Drawing.Size(292, 220)
        Me.strTools.Dock = System.Windows.Forms.DockStyle.Fill
        Me.strTools.Location = New System.Drawing.Point(0, 0)
        Me.strTools.Name = "strTools"
        Me.strTools.Size = New System.Drawing.Size(292, 266)
        Me.strTools.TabIndex = 1
        Me.strTools.Text = "ToolStripContainer1"
        '
        'strTools.TopToolStripPanel
        '
        Me.strTools.TopToolStripPanel.Controls.Add(Me.mnuStrip)
        '
        'strStatus
        '
        Me.strStatus.Dock = System.Windows.Forms.DockStyle.None
        Me.strStatus.Location = New System.Drawing.Point(0, 0)
        Me.strStatus.Name = "strStatus"
        Me.strStatus.Size = New System.Drawing.Size(292, 22)
        Me.strStatus.TabIndex = 0
        '
        'mnuStrip
        '
        Me.mnuStrip.Dock = System.Windows.Forms.DockStyle.None
        Me.mnuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile})
        Me.mnuStrip.Location = New System.Drawing.Point(0, 0)
        Me.mnuStrip.Name = "mnuStrip"
        Me.mnuStrip.Size = New System.Drawing.Size(292, 24)
        Me.mnuStrip.TabIndex = 0
        Me.mnuStrip.Text = "MenuStrip1"
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileLoad, Me.mnuFileSave, Me.mnuFileExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(35, 20)
        Me.mnuFile.Text = "&File"
        '
        'mnuFileLoad
        '
        Me.mnuFileLoad.Name = "mnuFileLoad"
        Me.mnuFileLoad.Size = New System.Drawing.Size(152, 22)
        Me.mnuFileLoad.Text = "&Load"
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
        'dlgFileOpen
        '
        Me.dlgFileOpen.FileName = "OpenFileDialog1"
        '
        'frmBoilerplate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.strTools)
        Me.MainMenuStrip = Me.mnuStrip
        Me.Name = "frmBoilerplate"
        Me.Text = "frmBoilerplate"
        Me.strTools.BottomToolStripPanel.ResumeLayout(False)
        Me.strTools.BottomToolStripPanel.PerformLayout()
        Me.strTools.ContentPanel.ResumeLayout(False)
        Me.strTools.ContentPanel.PerformLayout()
        Me.strTools.TopToolStripPanel.ResumeLayout(False)
        Me.strTools.TopToolStripPanel.PerformLayout()
        Me.strTools.ResumeLayout(False)
        Me.strTools.PerformLayout()
        Me.mnuStrip.ResumeLayout(False)
        Me.mnuStrip.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ctlControl As System.Windows.Forms.Label
    Friend WithEvents strTools As System.Windows.Forms.ToolStripContainer
    Friend WithEvents mnuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents strStatus As System.Windows.Forms.StatusStrip
    Friend WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFileLoad As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFileSave As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFileExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents dlgFileOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents dlgFileSave As System.Windows.Forms.SaveFileDialog
End Class
