<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EditMethods
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.dgCompound = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col11 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.lblCboOption1 = New System.Windows.Forms.Label()
        Me.cboOption1 = New System.Windows.Forms.ComboBox()
        Me.cboOption2 = New System.Windows.Forms.ComboBox()
        Me.lblCboOption2 = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CreateNewMethodToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CreateNewMethodToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopyFromExistingMethodToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UsingChemstationDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditCurrentMethodToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.tsslLocation = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslTeam = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslServer = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnOption2Add = New System.Windows.Forms.Button()
        Me.btnOption2Del = New System.Windows.Forms.Button()
        Me.btnAddCompound = New System.Windows.Forms.Button()
        Me.btnDelCompound = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.cboOption3 = New System.Windows.Forms.ComboBox()
        Me.lblCboOption3 = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.txtOption4 = New System.Windows.Forms.TextBox()
        Me.lblTxtOption4 = New System.Windows.Forms.Label()
        Me.txtOption5 = New System.Windows.Forms.TextBox()
        Me.lblTxtOption5 = New System.Windows.Forms.Label()
        Me.btnOption3Del = New System.Windows.Forms.Button()
        Me.btnOption3Add = New System.Windows.Forms.Button()
        Me.btnOption4 = New System.Windows.Forms.Button()
        CType(Me.dgCompound, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgCompound
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgCompound.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgCompound.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgCompound.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5, Me.Col6, Me.Col7, Me.Col8, Me.Col9, Me.Col10, Me.Col11})
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgCompound.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgCompound.Location = New System.Drawing.Point(25, 67)
        Me.dgCompound.Name = "dgCompound"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgCompound.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgCompound.RowHeadersWidth = 45
        Me.dgCompound.Size = New System.Drawing.Size(766, 260)
        Me.dgCompound.TabIndex = 0
        '
        'Col1
        '
        Me.Col1.HeaderText = ""
        Me.Col1.Name = "Col1"
        Me.Col1.Width = 150
        '
        'Col2
        '
        Me.Col2.HeaderText = ""
        Me.Col2.Name = "Col2"
        '
        'Col3
        '
        Me.Col3.HeaderText = ""
        Me.Col3.Name = "Col3"
        '
        'Col4
        '
        Me.Col4.HeaderText = ""
        Me.Col4.Name = "Col4"
        '
        'Col5
        '
        Me.Col5.HeaderText = ""
        Me.Col5.Name = "Col5"
        '
        'Col6
        '
        Me.Col6.HeaderText = ""
        Me.Col6.Name = "Col6"
        '
        'Col7
        '
        Me.Col7.FillWeight = 150.0!
        Me.Col7.HeaderText = ""
        Me.Col7.Name = "Col7"
        '
        'Col8
        '
        Me.Col8.HeaderText = ""
        Me.Col8.Name = "Col8"
        '
        'Col9
        '
        Me.Col9.HeaderText = ""
        Me.Col9.Name = "Col9"
        '
        'Col10
        '
        Me.Col10.HeaderText = ""
        Me.Col10.Name = "Col10"
        '
        'Col11
        '
        Me.Col11.HeaderText = ""
        Me.Col11.Name = "Col11"
        '
        'lblCboOption1
        '
        Me.lblCboOption1.AutoSize = True
        Me.lblCboOption1.Location = New System.Drawing.Point(22, 43)
        Me.lblCboOption1.Name = "lblCboOption1"
        Me.lblCboOption1.Size = New System.Drawing.Size(95, 13)
        Me.lblCboOption1.TabIndex = 1
        Me.lblCboOption1.Text = "Method / Analysis:"
        '
        'cboOption1
        '
        Me.cboOption1.FormattingEnabled = True
        Me.cboOption1.Location = New System.Drawing.Point(123, 40)
        Me.cboOption1.Name = "cboOption1"
        Me.cboOption1.Size = New System.Drawing.Size(171, 21)
        Me.cboOption1.TabIndex = 2
        '
        'cboOption2
        '
        Me.cboOption2.Enabled = False
        Me.cboOption2.FormattingEnabled = True
        Me.cboOption2.Location = New System.Drawing.Point(385, 40)
        Me.cboOption2.Name = "cboOption2"
        Me.cboOption2.Size = New System.Drawing.Size(153, 21)
        Me.cboOption2.TabIndex = 4
        '
        'lblCboOption2
        '
        Me.lblCboOption2.AutoSize = True
        Me.lblCboOption2.Location = New System.Drawing.Point(320, 43)
        Me.lblCboOption2.Name = "lblCboOption2"
        Me.lblCboOption2.Size = New System.Drawing.Size(59, 13)
        Me.lblCboOption2.TabIndex = 3
        Me.lblCboOption2.Text = "Instrument:"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.OptionsToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(815, 24)
        Me.MenuStrip1.TabIndex = 9
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CreateNewMethodToolStripMenuItem, Me.EditCurrentMethodToolStripMenuItem})
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.OptionsToolStripMenuItem.Text = "&Options"
        '
        'CreateNewMethodToolStripMenuItem
        '
        Me.CreateNewMethodToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CreateNewMethodToolStripMenuItem1, Me.CopyFromExistingMethodToolStripMenuItem, Me.UsingChemstationDataToolStripMenuItem})
        Me.CreateNewMethodToolStripMenuItem.Name = "CreateNewMethodToolStripMenuItem"
        Me.CreateNewMethodToolStripMenuItem.Size = New System.Drawing.Size(228, 22)
        Me.CreateNewMethodToolStripMenuItem.Text = "&Create New Method / Project"
        '
        'CreateNewMethodToolStripMenuItem1
        '
        Me.CreateNewMethodToolStripMenuItem1.Name = "CreateNewMethodToolStripMenuItem1"
        Me.CreateNewMethodToolStripMenuItem1.Size = New System.Drawing.Size(186, 22)
        Me.CreateNewMethodToolStripMenuItem1.Text = "Manually"
        '
        'CopyFromExistingMethodToolStripMenuItem
        '
        Me.CopyFromExistingMethodToolStripMenuItem.Name = "CopyFromExistingMethodToolStripMenuItem"
        Me.CopyFromExistingMethodToolStripMenuItem.Size = New System.Drawing.Size(186, 22)
        Me.CopyFromExistingMethodToolStripMenuItem.Text = "Copy From Existing"
        '
        'UsingChemstationDataToolStripMenuItem
        '
        Me.UsingChemstationDataToolStripMenuItem.Name = "UsingChemstationDataToolStripMenuItem"
        Me.UsingChemstationDataToolStripMenuItem.Size = New System.Drawing.Size(186, 22)
        Me.UsingChemstationDataToolStripMenuItem.Text = "Import from Cal Data"
        '
        'EditCurrentMethodToolStripMenuItem
        '
        Me.EditCurrentMethodToolStripMenuItem.Enabled = False
        Me.EditCurrentMethodToolStripMenuItem.Name = "EditCurrentMethodToolStripMenuItem"
        Me.EditCurrentMethodToolStripMenuItem.Size = New System.Drawing.Size(228, 22)
        Me.EditCurrentMethodToolStripMenuItem.Text = "E&dit Current"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsslLocation, Me.tsslTeam, Me.tsslServer})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 427)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(815, 22)
        Me.StatusStrip1.TabIndex = 17
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tsslLocation
        '
        Me.tsslLocation.Name = "tsslLocation"
        Me.tsslLocation.Size = New System.Drawing.Size(59, 17)
        Me.tsslLocation.Text = "Location: "
        '
        'tsslTeam
        '
        Me.tsslTeam.Name = "tsslTeam"
        Me.tsslTeam.Size = New System.Drawing.Size(40, 17)
        Me.tsslTeam.Text = "Team:"
        '
        'tsslServer
        '
        Me.tsslServer.Name = "tsslServer"
        Me.tsslServer.Size = New System.Drawing.Size(45, 17)
        Me.tsslServer.Text = "Server: "
        '
        'btnOption2Add
        '
        Me.btnOption2Add.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOption2Add.Location = New System.Drawing.Point(25, 345)
        Me.btnOption2Add.Name = "btnOption2Add"
        Me.btnOption2Add.Size = New System.Drawing.Size(87, 29)
        Me.btnOption2Add.TabIndex = 18
        Me.btnOption2Add.Text = "Option2 Add"
        Me.btnOption2Add.UseVisualStyleBackColor = True
        Me.btnOption2Add.Visible = False
        '
        'btnOption2Del
        '
        Me.btnOption2Del.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOption2Del.Location = New System.Drawing.Point(25, 380)
        Me.btnOption2Del.Name = "btnOption2Del"
        Me.btnOption2Del.Size = New System.Drawing.Size(87, 29)
        Me.btnOption2Del.TabIndex = 19
        Me.btnOption2Del.Text = "Option2 Del"
        Me.btnOption2Del.UseVisualStyleBackColor = True
        Me.btnOption2Del.Visible = False
        '
        'btnAddCompound
        '
        Me.btnAddCompound.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddCompound.Location = New System.Drawing.Point(250, 345)
        Me.btnAddCompound.Name = "btnAddCompound"
        Me.btnAddCompound.Size = New System.Drawing.Size(87, 29)
        Me.btnAddCompound.TabIndex = 20
        Me.btnAddCompound.Text = "Add Compound"
        Me.btnAddCompound.UseVisualStyleBackColor = True
        Me.btnAddCompound.Visible = False
        '
        'btnDelCompound
        '
        Me.btnDelCompound.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelCompound.Location = New System.Drawing.Point(250, 380)
        Me.btnDelCompound.Name = "btnDelCompound"
        Me.btnDelCompound.Size = New System.Drawing.Size(87, 29)
        Me.btnDelCompound.TabIndex = 21
        Me.btnDelCompound.Text = "Delete Compound"
        Me.btnDelCompound.UseVisualStyleBackColor = True
        Me.btnDelCompound.Visible = False
        '
        'btnSave
        '
        Me.btnSave.Enabled = False
        Me.btnSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(704, 380)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(87, 29)
        Me.btnSave.TabIndex = 22
        Me.btnSave.Text = "Save Method"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'cboOption3
        '
        Me.cboOption3.Enabled = False
        Me.cboOption3.FormattingEnabled = True
        Me.cboOption3.Location = New System.Drawing.Point(653, 40)
        Me.cboOption3.Name = "cboOption3"
        Me.cboOption3.Size = New System.Drawing.Size(138, 21)
        Me.cboOption3.TabIndex = 24
        '
        'lblCboOption3
        '
        Me.lblCboOption3.AutoSize = True
        Me.lblCboOption3.Location = New System.Drawing.Point(575, 43)
        Me.lblCboOption3.Name = "lblCboOption3"
        Me.lblCboOption3.Size = New System.Drawing.Size(72, 13)
        Me.lblCboOption3.TabIndex = 23
        Me.lblCboOption3.Text = "Analyte Type:"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Location = New System.Drawing.Point(701, 348)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(58, 13)
        Me.lblStatus.TabIndex = 31
        Me.lblStatus.Text = "Reviewed:"
        '
        'txtOption4
        '
        Me.txtOption4.Location = New System.Drawing.Point(578, 349)
        Me.txtOption4.Name = "txtOption4"
        Me.txtOption4.Size = New System.Drawing.Size(95, 20)
        Me.txtOption4.TabIndex = 34
        '
        'lblTxtOption4
        '
        Me.lblTxtOption4.Location = New System.Drawing.Point(474, 352)
        Me.lblTxtOption4.Name = "lblTxtOption4"
        Me.lblTxtOption4.Size = New System.Drawing.Size(98, 13)
        Me.lblTxtOption4.TabIndex = 33
        Me.lblTxtOption4.Text = "ETEQ :"
        Me.lblTxtOption4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOption5
        '
        Me.txtOption5.Location = New System.Drawing.Point(578, 384)
        Me.txtOption5.Name = "txtOption5"
        Me.txtOption5.Size = New System.Drawing.Size(95, 20)
        Me.txtOption5.TabIndex = 36
        '
        'lblTxtOption5
        '
        Me.lblTxtOption5.Location = New System.Drawing.Point(474, 387)
        Me.lblTxtOption5.Name = "lblTxtOption5"
        Me.lblTxtOption5.Size = New System.Drawing.Size(98, 13)
        Me.lblTxtOption5.TabIndex = 35
        Me.lblTxtOption5.Text = "Report Tolerance:"
        Me.lblTxtOption5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnOption3Del
        '
        Me.btnOption3Del.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOption3Del.Location = New System.Drawing.Point(137, 380)
        Me.btnOption3Del.Name = "btnOption3Del"
        Me.btnOption3Del.Size = New System.Drawing.Size(87, 29)
        Me.btnOption3Del.TabIndex = 38
        Me.btnOption3Del.Text = "Option3 Del"
        Me.btnOption3Del.UseVisualStyleBackColor = True
        Me.btnOption3Del.Visible = False
        '
        'btnOption3Add
        '
        Me.btnOption3Add.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOption3Add.Location = New System.Drawing.Point(137, 345)
        Me.btnOption3Add.Name = "btnOption3Add"
        Me.btnOption3Add.Size = New System.Drawing.Size(87, 29)
        Me.btnOption3Add.TabIndex = 37
        Me.btnOption3Add.Text = "Option3 Add"
        Me.btnOption3Add.UseVisualStyleBackColor = True
        Me.btnOption3Add.Visible = False
        '
        'btnOption4
        '
        Me.btnOption4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOption4.Location = New System.Drawing.Point(365, 345)
        Me.btnOption4.Name = "btnOption4"
        Me.btnOption4.Size = New System.Drawing.Size(87, 29)
        Me.btnOption4.TabIndex = 39
        Me.btnOption4.Text = "Standard Books"
        Me.btnOption4.UseVisualStyleBackColor = True
        Me.btnOption4.Visible = False
        '
        'EditMethods
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(815, 449)
        Me.Controls.Add(Me.btnOption4)
        Me.Controls.Add(Me.btnOption3Del)
        Me.Controls.Add(Me.btnOption3Add)
        Me.Controls.Add(Me.txtOption5)
        Me.Controls.Add(Me.lblTxtOption5)
        Me.Controls.Add(Me.txtOption4)
        Me.Controls.Add(Me.lblTxtOption4)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.cboOption3)
        Me.Controls.Add(Me.lblCboOption3)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnDelCompound)
        Me.Controls.Add(Me.btnAddCompound)
        Me.Controls.Add(Me.btnOption2Del)
        Me.Controls.Add(Me.btnOption2Add)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.cboOption2)
        Me.Controls.Add(Me.lblCboOption2)
        Me.Controls.Add(Me.cboOption1)
        Me.Controls.Add(Me.lblCboOption1)
        Me.Controls.Add(Me.dgCompound)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "EditMethods"
        Me.Text = "Edit Methods - eTrain 2.0"
        CType(Me.dgCompound, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgCompound As System.Windows.Forms.DataGridView
    Friend WithEvents lblCboOption1 As System.Windows.Forms.Label
    Friend WithEvents cboOption1 As System.Windows.Forms.ComboBox
    Friend WithEvents cboOption2 As System.Windows.Forms.ComboBox
    Friend WithEvents lblCboOption2 As System.Windows.Forms.Label
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tsslLocation As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tsslTeam As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tsslServer As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CreateNewMethodToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditCurrentMethodToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnOption2Add As System.Windows.Forms.Button
    Friend WithEvents btnOption2Del As System.Windows.Forms.Button
    Friend WithEvents btnAddCompound As System.Windows.Forms.Button
    Friend WithEvents btnDelCompound As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents cboOption3 As System.Windows.Forms.ComboBox
    Friend WithEvents lblCboOption3 As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents CreateNewMethodToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopyFromExistingMethodToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UsingChemstationDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents txtOption4 As System.Windows.Forms.TextBox
    Friend WithEvents lblTxtOption4 As System.Windows.Forms.Label
    Friend WithEvents txtOption5 As System.Windows.Forms.TextBox
    Friend WithEvents lblTxtOption5 As System.Windows.Forms.Label
    Friend WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col10 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnOption3Del As System.Windows.Forms.Button
    Friend WithEvents btnOption3Add As System.Windows.Forms.Button
    Friend WithEvents btnOption4 As System.Windows.Forms.Button
End Class
