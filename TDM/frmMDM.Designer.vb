<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMDM
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("AppName")
        Me.mainMenu = New System.Windows.Forms.MenuStrip()
        Me.DateiToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem_Beenden = New System.Windows.Forms.ToolStripMenuItem()
        Me.BearbeitenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AnsichtToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExtrasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem_Konfiguration = New System.Windows.Forms.ToolStripMenuItem()
        Me.SerialToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HilfeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mainSplitContainer = New System.Windows.Forms.SplitContainer()
        Me.TreeView = New System.Windows.Forms.TreeView()
        Me.PropertyGrid1 = New System.Windows.Forms.PropertyGrid()
        Me.dateTimeTimer = New System.Windows.Forms.Timer(Me.components)
        Me.Separator_1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.ToolStripStatusSeparator = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusDateTime = New System.Windows.Forms.ToolStripStatusLabel()
        Me.mainStatusStrip = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusTempInfo = New MDM.myToolStripStatusLabel()
        Me.ToolStripStatusNetworkState = New MDM.myToolStripStatusLabel()
        Me.ToolStripStatusMachineState = New MDM.myToolStripStatusLabel()
        Me.ToolStripStatusTempProcess = New MDM.myToolStripStatusLabel()
        Me.mainMenu.SuspendLayout()
        Me.mainSplitContainer.Panel1.SuspendLayout()
        Me.mainSplitContainer.Panel2.SuspendLayout()
        Me.mainSplitContainer.SuspendLayout()
        Me.mainStatusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'mainMenu
        '
        Me.mainMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DateiToolStripMenuItem, Me.BearbeitenToolStripMenuItem, Me.AnsichtToolStripMenuItem, Me.ExtrasToolStripMenuItem, Me.HilfeToolStripMenuItem})
        Me.mainMenu.Location = New System.Drawing.Point(0, 0)
        Me.mainMenu.Name = "mainMenu"
        Me.mainMenu.Size = New System.Drawing.Size(1032, 24)
        Me.mainMenu.TabIndex = 1
        Me.mainMenu.Text = "MenuStrip1"
        '
        'DateiToolStripMenuItem
        '
        Me.DateiToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem_Beenden})
        Me.DateiToolStripMenuItem.Name = "DateiToolStripMenuItem"
        Me.DateiToolStripMenuItem.Size = New System.Drawing.Size(46, 20)
        Me.DateiToolStripMenuItem.Text = "Datei"
        '
        'ToolStripMenuItem_Beenden
        '
        Me.ToolStripMenuItem_Beenden.Name = "ToolStripMenuItem_Beenden"
        Me.ToolStripMenuItem_Beenden.Size = New System.Drawing.Size(120, 22)
        Me.ToolStripMenuItem_Beenden.Text = "Beenden"
        '
        'BearbeitenToolStripMenuItem
        '
        Me.BearbeitenToolStripMenuItem.Name = "BearbeitenToolStripMenuItem"
        Me.BearbeitenToolStripMenuItem.Size = New System.Drawing.Size(75, 20)
        Me.BearbeitenToolStripMenuItem.Text = "Bearbeiten"
        '
        'AnsichtToolStripMenuItem
        '
        Me.AnsichtToolStripMenuItem.Name = "AnsichtToolStripMenuItem"
        Me.AnsichtToolStripMenuItem.Size = New System.Drawing.Size(59, 20)
        Me.AnsichtToolStripMenuItem.Text = "Ansicht"
        '
        'ExtrasToolStripMenuItem
        '
        Me.ExtrasToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem_Konfiguration, Me.SerialToolStripMenuItem})
        Me.ExtrasToolStripMenuItem.Name = "ExtrasToolStripMenuItem"
        Me.ExtrasToolStripMenuItem.Size = New System.Drawing.Size(49, 20)
        Me.ExtrasToolStripMenuItem.Text = "Extras"
        '
        'ToolStripMenuItem_Konfiguration
        '
        Me.ToolStripMenuItem_Konfiguration.Name = "ToolStripMenuItem_Konfiguration"
        Me.ToolStripMenuItem_Konfiguration.Size = New System.Drawing.Size(147, 22)
        Me.ToolStripMenuItem_Konfiguration.Text = "Konfiguration"
        '
        'SerialToolStripMenuItem
        '
        Me.SerialToolStripMenuItem.Name = "SerialToolStripMenuItem"
        Me.SerialToolStripMenuItem.Size = New System.Drawing.Size(147, 22)
        Me.SerialToolStripMenuItem.Text = "Serial"
        '
        'HilfeToolStripMenuItem
        '
        Me.HilfeToolStripMenuItem.Name = "HilfeToolStripMenuItem"
        Me.HilfeToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HilfeToolStripMenuItem.Text = "Hilfe"
        '
        'mainSplitContainer
        '
        Me.mainSplitContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.mainSplitContainer.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.mainSplitContainer.IsSplitterFixed = True
        Me.mainSplitContainer.Location = New System.Drawing.Point(0, 24)
        Me.mainSplitContainer.Name = "mainSplitContainer"
        '
        'mainSplitContainer.Panel1
        '
        Me.mainSplitContainer.Panel1.Controls.Add(Me.TreeView)
        '
        'mainSplitContainer.Panel2
        '
        Me.mainSplitContainer.Panel2.Controls.Add(Me.PropertyGrid1)
        Me.mainSplitContainer.Size = New System.Drawing.Size(1032, 492)
        Me.mainSplitContainer.SplitterDistance = 306
        Me.mainSplitContainer.TabIndex = 3
        '
        'TreeView
        '
        Me.TreeView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeView.Location = New System.Drawing.Point(0, 0)
        Me.TreeView.Name = "TreeView"
        TreeNode1.Name = "Knoten0"
        TreeNode1.Text = "AppName"
        Me.TreeView.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode1})
        Me.TreeView.Size = New System.Drawing.Size(306, 492)
        Me.TreeView.TabIndex = 0
        '
        'PropertyGrid1
        '
        Me.PropertyGrid1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PropertyGrid1.Location = New System.Drawing.Point(0, 0)
        Me.PropertyGrid1.Name = "PropertyGrid1"
        Me.PropertyGrid1.Size = New System.Drawing.Size(722, 492)
        Me.PropertyGrid1.TabIndex = 1
        '
        'dateTimeTimer
        '
        Me.dateTimeTimer.Interval = 2500
        '
        'Separator_1
        '
        Me.Separator_1.Name = "Separator_1"
        Me.Separator_1.Size = New System.Drawing.Size(10, 17)
        Me.Separator_1.Text = "|"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(10, 17)
        Me.ToolStripStatusLabel1.Text = "|"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(10, 17)
        Me.ToolStripStatusLabel2.Text = "|"
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(250, 16)
        '
        'ToolStripStatusSeparator
        '
        Me.ToolStripStatusSeparator.Name = "ToolStripStatusSeparator"
        Me.ToolStripStatusSeparator.Size = New System.Drawing.Size(361, 17)
        Me.ToolStripStatusSeparator.Spring = True
        '
        'ToolStripStatusDateTime
        '
        Me.ToolStripStatusDateTime.Name = "ToolStripStatusDateTime"
        Me.ToolStripStatusDateTime.Size = New System.Drawing.Size(57, 17)
        Me.ToolStripStatusDateTime.Text = "dateTime"
        '
        'mainStatusStrip
        '
        Me.mainStatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusTempInfo, Me.Separator_1, Me.ToolStripStatusNetworkState, Me.ToolStripStatusLabel1, Me.ToolStripStatusMachineState, Me.ToolStripStatusLabel2, Me.ToolStripStatusTempProcess, Me.ToolStripProgressBar1, Me.ToolStripStatusSeparator, Me.ToolStripStatusDateTime})
        Me.mainStatusStrip.Location = New System.Drawing.Point(0, 516)
        Me.mainStatusStrip.Name = "mainStatusStrip"
        Me.mainStatusStrip.RenderMode = System.Windows.Forms.ToolStripRenderMode.ManagerRenderMode
        Me.mainStatusStrip.Size = New System.Drawing.Size(1032, 22)
        Me.mainStatusStrip.TabIndex = 2
        Me.mainStatusStrip.Text = "StatusStrip1"
        '
        'ToolStripStatusTempInfo
        '
        Me.ToolStripStatusTempInfo.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStripStatusTempInfo.ForeColor = System.Drawing.Color.YellowGreen
        Me.ToolStripStatusTempInfo.InfoLevel = MDM.myToolStripStatusLabel.InfoLevelValues.Warning
        Me.ToolStripStatusTempInfo.Name = "ToolStripStatusTempInfo"
        Me.ToolStripStatusTempInfo.Size = New System.Drawing.Size(56, 17)
        Me.ToolStripStatusTempInfo.Text = "tempInfo"
        '
        'ToolStripStatusNetworkState
        '
        Me.ToolStripStatusNetworkState.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStripStatusNetworkState.ForeColor = System.Drawing.Color.Green
        Me.ToolStripStatusNetworkState.Name = "ToolStripStatusNetworkState"
        Me.ToolStripStatusNetworkState.Size = New System.Drawing.Size(76, 17)
        Me.ToolStripStatusNetworkState.Text = "networkState"
        '
        'ToolStripStatusMachineState
        '
        Me.ToolStripStatusMachineState.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStripStatusMachineState.ForeColor = System.Drawing.Color.Green
        Me.ToolStripStatusMachineState.Name = "ToolStripStatusMachineState"
        Me.ToolStripStatusMachineState.Size = New System.Drawing.Size(79, 17)
        Me.ToolStripStatusMachineState.Text = "machineState"
        '
        'ToolStripStatusTempProcess
        '
        Me.ToolStripStatusTempProcess.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStripStatusTempProcess.ForeColor = System.Drawing.Color.Red
        Me.ToolStripStatusTempProcess.InfoLevel = MDM.myToolStripStatusLabel.InfoLevelValues.[Error]
        Me.ToolStripStatusTempProcess.Name = "ToolStripStatusTempProcess"
        Me.ToolStripStatusTempProcess.Size = New System.Drawing.Size(75, 17)
        Me.ToolStripStatusTempProcess.Text = "tempProcess"
        '
        'frmMDM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1032, 538)
        Me.Controls.Add(Me.mainSplitContainer)
        Me.Controls.Add(Me.mainStatusStrip)
        Me.Controls.Add(Me.mainMenu)
        Me.MainMenuStrip = Me.mainMenu
        Me.Name = "frmMDM"
        Me.Text = "Form1"
        Me.mainMenu.ResumeLayout(False)
        Me.mainMenu.PerformLayout()
        Me.mainSplitContainer.Panel1.ResumeLayout(False)
        Me.mainSplitContainer.Panel2.ResumeLayout(False)
        Me.mainSplitContainer.ResumeLayout(False)
        Me.mainStatusStrip.ResumeLayout(False)
        Me.mainStatusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents mainMenu As System.Windows.Forms.MenuStrip
    Friend WithEvents DateiToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AnsichtToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExtrasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HilfeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BearbeitenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mainSplitContainer As System.Windows.Forms.SplitContainer
    Friend WithEvents TreeView As System.Windows.Forms.TreeView
    Friend WithEvents dateTimeTimer As System.Windows.Forms.Timer
    Friend WithEvents ToolStripMenuItem_Beenden As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem_Konfiguration As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripStatusTempInfo As MDM.myToolStripStatusLabel
    Friend WithEvents Separator_1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusNetworkState As MDM.myToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusMachineState As MDM.myToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusTempProcess As MDM.myToolStripStatusLabel
    Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents ToolStripStatusSeparator As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusDateTime As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents mainStatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents PropertyGrid1 As System.Windows.Forms.PropertyGrid
    Friend WithEvents SerialToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
