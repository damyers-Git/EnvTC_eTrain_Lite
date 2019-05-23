<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MainMenuToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TestToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ModeSwitchToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SeadriftToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ROHToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.MidlandToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AECOMToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CLABToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.lstFileList = New System.Windows.Forms.ListBox()
        Me.btnFindFiles = New System.Windows.Forms.Button()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.lblImportResults = New System.Windows.Forms.Label()
        Me.txtImportResults = New System.Windows.Forms.TextBox()
        Me.btnTransLIMS = New System.Windows.Forms.Button()
        Me.btnReport = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.tsslLocation = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslTeam = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslServer = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.nudSigFig = New System.Windows.Forms.NumericUpDown()
        Me.btnSelAll = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnClearList = New System.Windows.Forms.Button()
        Me.cboImportType = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnClearSamples = New System.Windows.Forms.Button()
        Me.btnSigHelp = New System.Windows.Forms.Button()
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.nudSigFig, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ModeSwitchToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(8, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(825, 28)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem, Me.MainMenuToolStripMenuItem, Me.TestToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(44, 24)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(125, 26)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'MainMenuToolStripMenuItem
        '
        Me.MainMenuToolStripMenuItem.Name = "MainMenuToolStripMenuItem"
        Me.MainMenuToolStripMenuItem.Size = New System.Drawing.Size(125, 26)
        Me.MainMenuToolStripMenuItem.Text = "E&xit"
        '
        'TestToolStripMenuItem
        '
        Me.TestToolStripMenuItem.Enabled = False
        Me.TestToolStripMenuItem.Name = "TestToolStripMenuItem"
        Me.TestToolStripMenuItem.Size = New System.Drawing.Size(125, 26)
        Me.TestToolStripMenuItem.Text = "Test"
        '
        'ModeSwitchToolStripMenuItem
        '
        Me.ModeSwitchToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SeadriftToolStripMenuItem, Me.ROHToolStripMenuItem1, Me.MidlandToolStripMenuItem})
        Me.ModeSwitchToolStripMenuItem.Name = "ModeSwitchToolStripMenuItem"
        Me.ModeSwitchToolStripMenuItem.Size = New System.Drawing.Size(98, 24)
        Me.ModeSwitchToolStripMenuItem.Text = "&LIMS Server"
        '
        'SeadriftToolStripMenuItem
        '
        Me.SeadriftToolStripMenuItem.Name = "SeadriftToolStripMenuItem"
        Me.SeadriftToolStripMenuItem.Size = New System.Drawing.Size(139, 26)
        Me.SeadriftToolStripMenuItem.Text = "&Seadrift"
        '
        'ROHToolStripMenuItem1
        '
        Me.ROHToolStripMenuItem1.Name = "ROHToolStripMenuItem1"
        Me.ROHToolStripMenuItem1.Size = New System.Drawing.Size(139, 26)
        Me.ROHToolStripMenuItem1.Text = "&ROH"
        '
        'MidlandToolStripMenuItem
        '
        Me.MidlandToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AECOMToolStripMenuItem, Me.CLABToolStripMenuItem})
        Me.MidlandToolStripMenuItem.Name = "MidlandToolStripMenuItem"
        Me.MidlandToolStripMenuItem.Size = New System.Drawing.Size(139, 26)
        Me.MidlandToolStripMenuItem.Text = "Midland"
        '
        'AECOMToolStripMenuItem
        '
        Me.AECOMToolStripMenuItem.Name = "AECOMToolStripMenuItem"
        Me.AECOMToolStripMenuItem.Size = New System.Drawing.Size(135, 26)
        Me.AECOMToolStripMenuItem.Text = "AECOM"
        '
        'CLABToolStripMenuItem
        '
        Me.CLABToolStripMenuItem.Name = "CLABToolStripMenuItem"
        Me.CLABToolStripMenuItem.Size = New System.Drawing.Size(135, 26)
        Me.CLABToolStripMenuItem.Text = "CLAB"
        '
        'lstFileList
        '
        Me.lstFileList.FormattingEnabled = True
        Me.lstFileList.HorizontalScrollbar = True
        Me.lstFileList.ItemHeight = 16
        Me.lstFileList.Location = New System.Drawing.Point(16, 39)
        Me.lstFileList.Margin = New System.Windows.Forms.Padding(4)
        Me.lstFileList.Name = "lstFileList"
        Me.lstFileList.ScrollAlwaysVisible = True
        Me.lstFileList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstFileList.Size = New System.Drawing.Size(599, 212)
        Me.lstFileList.TabIndex = 1
        '
        'btnFindFiles
        '
        Me.btnFindFiles.Enabled = False
        Me.btnFindFiles.Location = New System.Drawing.Point(649, 139)
        Me.btnFindFiles.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindFiles.Name = "btnFindFiles"
        Me.btnFindFiles.Size = New System.Drawing.Size(156, 50)
        Me.btnFindFiles.TabIndex = 4
        Me.btnFindFiles.Text = "Find Files"
        Me.btnFindFiles.UseVisualStyleBackColor = True
        '
        'btnImport
        '
        Me.btnImport.Enabled = False
        Me.btnImport.Location = New System.Drawing.Point(649, 235)
        Me.btnImport.Margin = New System.Windows.Forms.Padding(4)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(156, 50)
        Me.btnImport.TabIndex = 5
        Me.btnImport.Text = "Import"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'lblImportResults
        '
        Me.lblImportResults.AutoSize = True
        Me.lblImportResults.Location = New System.Drawing.Point(16, 302)
        Me.lblImportResults.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblImportResults.Name = "lblImportResults"
        Me.lblImportResults.Size = New System.Drawing.Size(102, 17)
        Me.lblImportResults.TabIndex = 6
        Me.lblImportResults.Text = "Import Results:"
        '
        'txtImportResults
        '
        Me.txtImportResults.Location = New System.Drawing.Point(16, 321)
        Me.txtImportResults.Margin = New System.Windows.Forms.Padding(4)
        Me.txtImportResults.Multiline = True
        Me.txtImportResults.Name = "txtImportResults"
        Me.txtImportResults.ReadOnly = True
        Me.txtImportResults.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtImportResults.Size = New System.Drawing.Size(599, 96)
        Me.txtImportResults.TabIndex = 7
        '
        'btnTransLIMS
        '
        Me.btnTransLIMS.Enabled = False
        Me.btnTransLIMS.Location = New System.Drawing.Point(649, 497)
        Me.btnTransLIMS.Margin = New System.Windows.Forms.Padding(4)
        Me.btnTransLIMS.Name = "btnTransLIMS"
        Me.btnTransLIMS.Size = New System.Drawing.Size(156, 50)
        Me.btnTransLIMS.TabIndex = 12
        Me.btnTransLIMS.Text = "Transfer to LIMS"
        Me.btnTransLIMS.UseVisualStyleBackColor = True
        '
        'btnReport
        '
        Me.btnReport.Enabled = False
        Me.btnReport.Location = New System.Drawing.Point(485, 497)
        Me.btnReport.Margin = New System.Windows.Forms.Padding(4)
        Me.btnReport.Name = "btnReport"
        Me.btnReport.Size = New System.Drawing.Size(156, 50)
        Me.btnReport.TabIndex = 13
        Me.btnReport.Text = "Generate Report(s)"
        Me.btnReport.UseVisualStyleBackColor = True
        Me.btnReport.Visible = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsslLocation, Me.tsslTeam, Me.tsslServer})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 577)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(825, 22)
        Me.StatusStrip1.TabIndex = 16
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tsslLocation
        '
        Me.tsslLocation.Name = "tsslLocation"
        Me.tsslLocation.Size = New System.Drawing.Size(73, 20)
        Me.tsslLocation.Text = "Location: "
        Me.tsslLocation.Visible = False
        '
        'tsslTeam
        '
        Me.tsslTeam.Name = "tsslTeam"
        Me.tsslTeam.Size = New System.Drawing.Size(49, 20)
        Me.tsslTeam.Text = "Team:"
        Me.tsslTeam.Visible = False
        '
        'tsslServer
        '
        Me.tsslServer.Name = "tsslServer"
        Me.tsslServer.Size = New System.Drawing.Size(57, 20)
        Me.tsslServer.Text = "Server: "
        Me.tsslServer.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(20, 480)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(185, 17)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Additional Required Options"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(20, 514)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 17)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Significant Figures:"
        '
        'nudSigFig
        '
        Me.nudSigFig.Enabled = False
        Me.nudSigFig.Location = New System.Drawing.Point(157, 512)
        Me.nudSigFig.Margin = New System.Windows.Forms.Padding(4)
        Me.nudSigFig.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.nudSigFig.Minimum = New Decimal(New Integer() {1, 0, 0, -2147483648})
        Me.nudSigFig.Name = "nudSigFig"
        Me.nudSigFig.Size = New System.Drawing.Size(77, 22)
        Me.nudSigFig.TabIndex = 24
        Me.nudSigFig.Value = New Decimal(New Integer() {1, 0, 0, -2147483648})
        '
        'btnSelAll
        '
        Me.btnSelAll.Location = New System.Drawing.Point(576, 257)
        Me.btnSelAll.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSelAll.Name = "btnSelAll"
        Me.btnSelAll.Size = New System.Drawing.Size(40, 28)
        Me.btnSelAll.TabIndex = 25
        Me.btnSelAll.Text = "All"
        Me.btnSelAll.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(415, 263)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(148, 17)
        Me.Label6.TabIndex = 26
        Me.Label6.Text = "CTRL + Click to Select"
        '
        'btnClearList
        '
        Me.btnClearList.Location = New System.Drawing.Point(16, 257)
        Me.btnClearList.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClearList.Name = "btnClearList"
        Me.btnClearList.Size = New System.Drawing.Size(175, 28)
        Me.btnClearList.TabIndex = 27
        Me.btnClearList.Text = "Clear Sample List"
        Me.btnClearList.UseVisualStyleBackColor = True
        '
        'cboImportType
        '
        Me.cboImportType.Enabled = False
        Me.cboImportType.FormattingEnabled = True
        Me.cboImportType.Location = New System.Drawing.Point(649, 106)
        Me.cboImportType.Margin = New System.Windows.Forms.Padding(4)
        Me.cboImportType.Name = "cboImportType"
        Me.cboImportType.Size = New System.Drawing.Size(155, 24)
        Me.cboImportType.TabIndex = 28
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(645, 86)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 17)
        Me.Label1.TabIndex = 29
        Me.Label1.Text = "Import Type:"
        '
        'btnClearSamples
        '
        Me.btnClearSamples.Location = New System.Drawing.Point(11, 423)
        Me.btnClearSamples.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClearSamples.Name = "btnClearSamples"
        Me.btnClearSamples.Size = New System.Drawing.Size(175, 28)
        Me.btnClearSamples.TabIndex = 30
        Me.btnClearSamples.Text = "Clear Imported Samples"
        Me.btnClearSamples.UseVisualStyleBackColor = True
        '
        'btnSigHelp
        '
        Me.btnSigHelp.Enabled = False
        Me.btnSigHelp.Location = New System.Drawing.Point(243, 512)
        Me.btnSigHelp.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSigHelp.Name = "btnSigHelp"
        Me.btnSigHelp.Size = New System.Drawing.Size(31, 25)
        Me.btnSigHelp.TabIndex = 31
        Me.btnSigHelp.Text = "?"
        Me.btnSigHelp.UseVisualStyleBackColor = True
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(825, 599)
        Me.Controls.Add(Me.btnSigHelp)
        Me.Controls.Add(Me.btnClearSamples)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboImportType)
        Me.Controls.Add(Me.btnClearList)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btnSelAll)
        Me.Controls.Add(Me.nudSigFig)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnReport)
        Me.Controls.Add(Me.btnTransLIMS)
        Me.Controls.Add(Me.txtImportResults)
        Me.Controls.Add(Me.lblImportResults)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.btnFindFiles)
        Me.Controls.Add(Me.lstFileList)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "MainForm"
        Me.Text = "eTrain"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.nudSigFig, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MainMenuToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lstFileList As System.Windows.Forms.ListBox
    Friend WithEvents btnFindFiles As System.Windows.Forms.Button
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents lblImportResults As System.Windows.Forms.Label
    Friend WithEvents txtImportResults As System.Windows.Forms.TextBox
    Friend WithEvents btnTransLIMS As System.Windows.Forms.Button
    Friend WithEvents ModeSwitchToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SeadriftToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnReport As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tsslLocation As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tsslTeam As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tsslServer As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents nudSigFig As System.Windows.Forms.NumericUpDown
    Friend WithEvents btnSelAll As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnClearList As System.Windows.Forms.Button
    Friend WithEvents cboImportType As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnClearSamples As System.Windows.Forms.Button
    Friend WithEvents btnSigHelp As System.Windows.Forms.Button
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TestToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ROHToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents MidlandToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AECOMToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CLABToolStripMenuItem As ToolStripMenuItem
End Class
