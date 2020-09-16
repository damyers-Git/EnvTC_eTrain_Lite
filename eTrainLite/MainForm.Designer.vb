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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MainMenuToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TestToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ModeSwitchToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MidlandToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CLABToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EUROLANToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ALSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SGSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TAToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FIBERTECToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VISTAToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewSampleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TANC_NEWToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.AECOMToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SeadriftToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ROHToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
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
        Me.MenuStrip1.Size = New System.Drawing.Size(619, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem, Me.MainMenuToolStripMenuItem, Me.TestToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'MainMenuToolStripMenuItem
        '
        Me.MainMenuToolStripMenuItem.Name = "MainMenuToolStripMenuItem"
        Me.MainMenuToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.MainMenuToolStripMenuItem.Text = "E&xit"
        '
        'TestToolStripMenuItem
        '
        Me.TestToolStripMenuItem.Enabled = False
        Me.TestToolStripMenuItem.Name = "TestToolStripMenuItem"
        Me.TestToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.TestToolStripMenuItem.Text = "Test"
        '
        'ModeSwitchToolStripMenuItem
        '
        Me.ModeSwitchToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MidlandToolStripMenuItem, Me.SeadriftToolStripMenuItem, Me.ROHToolStripMenuItem1})
        Me.ModeSwitchToolStripMenuItem.Name = "ModeSwitchToolStripMenuItem"
        Me.ModeSwitchToolStripMenuItem.Size = New System.Drawing.Size(80, 20)
        Me.ModeSwitchToolStripMenuItem.Text = "&LIMS Server"
        '
        'MidlandToolStripMenuItem
        '
        Me.MidlandToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CLABToolStripMenuItem, Me.NewSampleToolStripMenuItem, Me.AECOMToolStripMenuItem})
        Me.MidlandToolStripMenuItem.Name = "MidlandToolStripMenuItem"
        Me.MidlandToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.MidlandToolStripMenuItem.Text = "Midland"
        '
        'CLABToolStripMenuItem
        '
        Me.CLABToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EUROLANToolStripMenuItem, Me.ALSToolStripMenuItem, Me.SGSToolStripMenuItem, Me.TAToolStripMenuItem, Me.FIBERTECToolStripMenuItem, Me.VISTAToolStripMenuItem})
        Me.CLABToolStripMenuItem.Name = "CLABToolStripMenuItem"
        Me.CLABToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.CLABToolStripMenuItem.Text = "CLAB"
        '
        'EUROLANToolStripMenuItem
        '
        Me.EUROLANToolStripMenuItem.Name = "EUROLANToolStripMenuItem"
        Me.EUROLANToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.EUROLANToolStripMenuItem.Text = "EUROLAN"
        '
        'ALSToolStripMenuItem
        '
        Me.ALSToolStripMenuItem.Name = "ALSToolStripMenuItem"
        Me.ALSToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.ALSToolStripMenuItem.Text = "ALS"
        '
        'SGSToolStripMenuItem
        '
        Me.SGSToolStripMenuItem.Name = "SGSToolStripMenuItem"
        Me.SGSToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.SGSToolStripMenuItem.Text = "SGS"
        '
        'TAToolStripMenuItem
        '
        Me.TAToolStripMenuItem.Name = "TAToolStripMenuItem"
        Me.TAToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.TAToolStripMenuItem.Text = "TA"
        '
        'FIBERTECToolStripMenuItem
        '
        Me.FIBERTECToolStripMenuItem.Name = "FIBERTECToolStripMenuItem"
        Me.FIBERTECToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.FIBERTECToolStripMenuItem.Text = "FIBERTEC"
        '
        'VISTAToolStripMenuItem
        '
        Me.VISTAToolStripMenuItem.Name = "VISTAToolStripMenuItem"
        Me.VISTAToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.VISTAToolStripMenuItem.Text = "VISTA"
        '
        'NewSampleToolStripMenuItem
        '
        Me.NewSampleToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TANC_NEWToolStripMenuItem1})
        Me.NewSampleToolStripMenuItem.Name = "NewSampleToolStripMenuItem"
        Me.NewSampleToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.NewSampleToolStripMenuItem.Text = "NewSample"
        '
        'TANC_NEWToolStripMenuItem1
        '
        Me.TANC_NEWToolStripMenuItem1.Name = "TANC_NEWToolStripMenuItem1"
        Me.TANC_NEWToolStripMenuItem1.Size = New System.Drawing.Size(152, 22)
        Me.TANC_NEWToolStripMenuItem1.Text = "TANC_NEW"
        '
        'AECOMToolStripMenuItem
        '
        Me.AECOMToolStripMenuItem.Name = "AECOMToolStripMenuItem"
        Me.AECOMToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.AECOMToolStripMenuItem.Text = "AECOM"
        '
        'SeadriftToolStripMenuItem
        '
        Me.SeadriftToolStripMenuItem.Name = "SeadriftToolStripMenuItem"
        Me.SeadriftToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.SeadriftToolStripMenuItem.Text = "&Seadrift"
        '
        'ROHToolStripMenuItem1
        '
        Me.ROHToolStripMenuItem1.Name = "ROHToolStripMenuItem1"
        Me.ROHToolStripMenuItem1.Size = New System.Drawing.Size(152, 22)
        Me.ROHToolStripMenuItem1.Text = "&ROH"
        '
        'lstFileList
        '
        Me.lstFileList.FormattingEnabled = True
        Me.lstFileList.HorizontalScrollbar = True
        Me.lstFileList.Location = New System.Drawing.Point(12, 32)
        Me.lstFileList.Name = "lstFileList"
        Me.lstFileList.ScrollAlwaysVisible = True
        Me.lstFileList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstFileList.Size = New System.Drawing.Size(450, 173)
        Me.lstFileList.TabIndex = 1
        '
        'btnFindFiles
        '
        Me.btnFindFiles.Enabled = False
        Me.btnFindFiles.Location = New System.Drawing.Point(487, 113)
        Me.btnFindFiles.Name = "btnFindFiles"
        Me.btnFindFiles.Size = New System.Drawing.Size(117, 41)
        Me.btnFindFiles.TabIndex = 4
        Me.btnFindFiles.Text = "Find Files"
        Me.btnFindFiles.UseVisualStyleBackColor = True
        '
        'btnImport
        '
        Me.btnImport.Enabled = False
        Me.btnImport.Location = New System.Drawing.Point(487, 191)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(117, 41)
        Me.btnImport.TabIndex = 5
        Me.btnImport.Text = "Import"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'lblImportResults
        '
        Me.lblImportResults.AutoSize = True
        Me.lblImportResults.Location = New System.Drawing.Point(12, 245)
        Me.lblImportResults.Name = "lblImportResults"
        Me.lblImportResults.Size = New System.Drawing.Size(77, 13)
        Me.lblImportResults.TabIndex = 6
        Me.lblImportResults.Text = "Import Results:"
        '
        'txtImportResults
        '
        Me.txtImportResults.Location = New System.Drawing.Point(12, 261)
        Me.txtImportResults.Multiline = True
        Me.txtImportResults.Name = "txtImportResults"
        Me.txtImportResults.ReadOnly = True
        Me.txtImportResults.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtImportResults.Size = New System.Drawing.Size(450, 149)
        Me.txtImportResults.TabIndex = 7
        '
        'btnTransLIMS
        '
        Me.btnTransLIMS.Enabled = False
        Me.btnTransLIMS.Location = New System.Drawing.Point(490, 476)
        Me.btnTransLIMS.Name = "btnTransLIMS"
        Me.btnTransLIMS.Size = New System.Drawing.Size(117, 41)
        Me.btnTransLIMS.TabIndex = 12
        Me.btnTransLIMS.Text = "Transfer to LIMS"
        Me.btnTransLIMS.UseVisualStyleBackColor = True
        '
        'btnReport
        '
        Me.btnReport.Enabled = False
        Me.btnReport.Location = New System.Drawing.Point(367, 476)
        Me.btnReport.Name = "btnReport"
        Me.btnReport.Size = New System.Drawing.Size(117, 41)
        Me.btnReport.TabIndex = 13
        Me.btnReport.Text = "Generate Report(s)"
        Me.btnReport.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsslLocation, Me.tsslTeam, Me.tsslServer})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 529)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(619, 22)
        Me.StatusStrip1.TabIndex = 16
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tsslLocation
        '
        Me.tsslLocation.Name = "tsslLocation"
        Me.tsslLocation.Size = New System.Drawing.Size(59, 17)
        Me.tsslLocation.Text = "Location: "
        Me.tsslLocation.Visible = False
        '
        'tsslTeam
        '
        Me.tsslTeam.Name = "tsslTeam"
        Me.tsslTeam.Size = New System.Drawing.Size(40, 17)
        Me.tsslTeam.Text = "Team:"
        Me.tsslTeam.Visible = False
        '
        'tsslServer
        '
        Me.tsslServer.Name = "tsslServer"
        Me.tsslServer.Size = New System.Drawing.Size(45, 17)
        Me.tsslServer.Text = "Server: "
        Me.tsslServer.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(18, 462)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(138, 13)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Additional Required Options"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(18, 490)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 13)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Significant Figures:"
        '
        'nudSigFig
        '
        Me.nudSigFig.Enabled = False
        Me.nudSigFig.Location = New System.Drawing.Point(121, 488)
        Me.nudSigFig.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.nudSigFig.Minimum = New Decimal(New Integer() {1, 0, 0, -2147483648})
        Me.nudSigFig.Name = "nudSigFig"
        Me.nudSigFig.Size = New System.Drawing.Size(58, 20)
        Me.nudSigFig.TabIndex = 24
        Me.nudSigFig.Value = New Decimal(New Integer() {1, 0, 0, -2147483648})
        '
        'btnSelAll
        '
        Me.btnSelAll.Location = New System.Drawing.Point(432, 209)
        Me.btnSelAll.Name = "btnSelAll"
        Me.btnSelAll.Size = New System.Drawing.Size(30, 23)
        Me.btnSelAll.TabIndex = 25
        Me.btnSelAll.Text = "All"
        Me.btnSelAll.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(311, 214)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(115, 13)
        Me.Label6.TabIndex = 26
        Me.Label6.Text = "CTRL + Click to Select"
        '
        'btnClearList
        '
        Me.btnClearList.Location = New System.Drawing.Point(12, 209)
        Me.btnClearList.Name = "btnClearList"
        Me.btnClearList.Size = New System.Drawing.Size(131, 23)
        Me.btnClearList.TabIndex = 27
        Me.btnClearList.Text = "Clear Sample List"
        Me.btnClearList.UseVisualStyleBackColor = True
        '
        'cboImportType
        '
        Me.cboImportType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboImportType.Enabled = False
        Me.cboImportType.FormattingEnabled = True
        Me.cboImportType.Location = New System.Drawing.Point(487, 86)
        Me.cboImportType.Name = "cboImportType"
        Me.cboImportType.Size = New System.Drawing.Size(117, 21)
        Me.cboImportType.TabIndex = 28
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(484, 70)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 29
        Me.Label1.Text = "Import Type:"
        '
        'btnClearSamples
        '
        Me.btnClearSamples.Location = New System.Drawing.Point(11, 416)
        Me.btnClearSamples.Name = "btnClearSamples"
        Me.btnClearSamples.Size = New System.Drawing.Size(131, 23)
        Me.btnClearSamples.TabIndex = 30
        Me.btnClearSamples.Text = "Clear Imported Samples"
        Me.btnClearSamples.UseVisualStyleBackColor = True
        '
        'btnSigHelp
        '
        Me.btnSigHelp.Enabled = False
        Me.btnSigHelp.Location = New System.Drawing.Point(185, 488)
        Me.btnSigHelp.Name = "btnSigHelp"
        Me.btnSigHelp.Size = New System.Drawing.Size(23, 20)
        Me.btnSigHelp.TabIndex = 31
        Me.btnSigHelp.Text = "?"
        Me.btnSigHelp.UseVisualStyleBackColor = True
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(619, 551)
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
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "MainForm"
        Me.Text = "eTrainLite"
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
    Friend WithEvents EUROLANToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ALSToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SGSToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TAToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FIBERTECToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents VISTAToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents NewSampleToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TANC_NEWToolStripMenuItem1 As ToolStripMenuItem
End Class
