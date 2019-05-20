<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TransferForm
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
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.cbo1 = New System.Windows.Forms.ComboBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ComponentSelectionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SampleEditorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.cbo2 = New System.Windows.Forms.ComboBox()
        Me.lbl2 = New System.Windows.Forms.Label()
        Me.cbo3 = New System.Windows.Forms.ComboBox()
        Me.lbl3 = New System.Windows.Forms.Label()
        Me.cbo4 = New System.Windows.Forms.ComboBox()
        Me.lbl4 = New System.Windows.Forms.Label()
        Me.cbo5 = New System.Windows.Forms.ComboBox()
        Me.lbl5 = New System.Windows.Forms.Label()
        Me.txt1 = New System.Windows.Forms.TextBox()
        Me.lblTxt1 = New System.Windows.Forms.Label()
        Me.btnSISBrowse = New System.Windows.Forms.Button()
        Me.btnTransfer = New System.Windows.Forms.Button()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbl1
        '
        Me.lbl1.AutoSize = True
        Me.lbl1.Location = New System.Drawing.Point(9, 32)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(50, 13)
        Me.lbl1.TabIndex = 0
        Me.lbl1.Text = "Option 1:"
        '
        'cbo1
        '
        Me.cbo1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbo1.FormattingEnabled = True
        Me.cbo1.Location = New System.Drawing.Point(12, 48)
        Me.cbo1.Name = "cbo1"
        Me.cbo1.Size = New System.Drawing.Size(179, 21)
        Me.cbo1.Sorted = True
        Me.cbo1.TabIndex = 1
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.OptionsToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(408, 24)
        Me.MenuStrip1.TabIndex = 2
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
        Me.OptionsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ComponentSelectionToolStripMenuItem, Me.SampleEditorToolStripMenuItem})
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.OptionsToolStripMenuItem.Text = "&Options"
        '
        'ComponentSelectionToolStripMenuItem
        '
        Me.ComponentSelectionToolStripMenuItem.Name = "ComponentSelectionToolStripMenuItem"
        Me.ComponentSelectionToolStripMenuItem.Size = New System.Drawing.Size(189, 22)
        Me.ComponentSelectionToolStripMenuItem.Text = "&Component Selection"
        '
        'SampleEditorToolStripMenuItem
        '
        Me.SampleEditorToolStripMenuItem.Name = "SampleEditorToolStripMenuItem"
        Me.SampleEditorToolStripMenuItem.Size = New System.Drawing.Size(189, 22)
        Me.SampleEditorToolStripMenuItem.Text = "&Sample Editor"
        '
        'cbo2
        '
        Me.cbo2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbo2.FormattingEnabled = True
        Me.cbo2.Location = New System.Drawing.Point(217, 48)
        Me.cbo2.Name = "cbo2"
        Me.cbo2.Size = New System.Drawing.Size(179, 21)
        Me.cbo2.Sorted = True
        Me.cbo2.TabIndex = 2
        '
        'lbl2
        '
        Me.lbl2.AutoSize = True
        Me.lbl2.Location = New System.Drawing.Point(214, 32)
        Me.lbl2.Name = "lbl2"
        Me.lbl2.Size = New System.Drawing.Size(50, 13)
        Me.lbl2.TabIndex = 3
        Me.lbl2.Text = "Option 2:"
        '
        'cbo3
        '
        Me.cbo3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbo3.FormattingEnabled = True
        Me.cbo3.Location = New System.Drawing.Point(12, 103)
        Me.cbo3.Name = "cbo3"
        Me.cbo3.Size = New System.Drawing.Size(179, 21)
        Me.cbo3.Sorted = True
        Me.cbo3.TabIndex = 3
        '
        'lbl3
        '
        Me.lbl3.AutoSize = True
        Me.lbl3.Location = New System.Drawing.Point(9, 87)
        Me.lbl3.Name = "lbl3"
        Me.lbl3.Size = New System.Drawing.Size(50, 13)
        Me.lbl3.TabIndex = 5
        Me.lbl3.Text = "Option 3:"
        '
        'cbo4
        '
        Me.cbo4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbo4.FormattingEnabled = True
        Me.cbo4.Location = New System.Drawing.Point(217, 103)
        Me.cbo4.Name = "cbo4"
        Me.cbo4.Size = New System.Drawing.Size(179, 21)
        Me.cbo4.Sorted = True
        Me.cbo4.TabIndex = 4
        '
        'lbl4
        '
        Me.lbl4.AutoSize = True
        Me.lbl4.Location = New System.Drawing.Point(214, 87)
        Me.lbl4.Name = "lbl4"
        Me.lbl4.Size = New System.Drawing.Size(50, 13)
        Me.lbl4.TabIndex = 7
        Me.lbl4.Text = "Option 4:"
        '
        'cbo5
        '
        Me.cbo5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbo5.FormattingEnabled = True
        Me.cbo5.Location = New System.Drawing.Point(12, 157)
        Me.cbo5.Name = "cbo5"
        Me.cbo5.Size = New System.Drawing.Size(179, 21)
        Me.cbo5.TabIndex = 5
        '
        'lbl5
        '
        Me.lbl5.AutoSize = True
        Me.lbl5.Location = New System.Drawing.Point(9, 141)
        Me.lbl5.Name = "lbl5"
        Me.lbl5.Size = New System.Drawing.Size(50, 13)
        Me.lbl5.TabIndex = 9
        Me.lbl5.Text = "Option 5:"
        '
        'txt1
        '
        Me.txt1.Location = New System.Drawing.Point(12, 270)
        Me.txt1.Name = "txt1"
        Me.txt1.Size = New System.Drawing.Size(294, 20)
        Me.txt1.TabIndex = 9
        Me.txt1.Visible = False
        '
        'lblTxt1
        '
        Me.lblTxt1.AutoSize = True
        Me.lblTxt1.Location = New System.Drawing.Point(12, 254)
        Me.lblTxt1.Name = "lblTxt1"
        Me.lblTxt1.Size = New System.Drawing.Size(126, 13)
        Me.lblTxt1.TabIndex = 12
        Me.lblTxt1.Text = "Associated SIS Location:"
        Me.lblTxt1.Visible = False
        '
        'btnSISBrowse
        '
        Me.btnSISBrowse.Location = New System.Drawing.Point(321, 266)
        Me.btnSISBrowse.Name = "btnSISBrowse"
        Me.btnSISBrowse.Size = New System.Drawing.Size(75, 26)
        Me.btnSISBrowse.TabIndex = 10
        Me.btnSISBrowse.Text = "Browse..."
        Me.btnSISBrowse.UseVisualStyleBackColor = True
        '
        'btnTransfer
        '
        Me.btnTransfer.Location = New System.Drawing.Point(321, 305)
        Me.btnTransfer.Name = "btnTransfer"
        Me.btnTransfer.Size = New System.Drawing.Size(75, 26)
        Me.btnTransfer.TabIndex = 11
        Me.btnTransfer.Text = "Transfer"
        Me.btnTransfer.UseVisualStyleBackColor = True
        '
        'TransferForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(408, 346)
        Me.Controls.Add(Me.btnTransfer)
        Me.Controls.Add(Me.btnSISBrowse)
        Me.Controls.Add(Me.lblTxt1)
        Me.Controls.Add(Me.txt1)
        Me.Controls.Add(Me.cbo5)
        Me.Controls.Add(Me.lbl5)
        Me.Controls.Add(Me.cbo4)
        Me.Controls.Add(Me.lbl4)
        Me.Controls.Add(Me.cbo3)
        Me.Controls.Add(Me.lbl3)
        Me.Controls.Add(Me.cbo2)
        Me.Controls.Add(Me.lbl2)
        Me.Controls.Add(Me.cbo1)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "TransferForm"
        Me.Text = "Transfer Data - eTrain 2.0"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents cbo1 As System.Windows.Forms.ComboBox
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cbo2 As System.Windows.Forms.ComboBox
    Friend WithEvents lbl2 As System.Windows.Forms.Label
    Friend WithEvents cbo3 As System.Windows.Forms.ComboBox
    Friend WithEvents lbl3 As System.Windows.Forms.Label
    Friend WithEvents cbo4 As System.Windows.Forms.ComboBox
    Friend WithEvents lbl4 As System.Windows.Forms.Label
    Friend WithEvents cbo5 As System.Windows.Forms.ComboBox
    Friend WithEvents lbl5 As System.Windows.Forms.Label
    Friend WithEvents txt1 As System.Windows.Forms.TextBox
    Friend WithEvents lblTxt1 As System.Windows.Forms.Label
    Friend WithEvents btnSISBrowse As System.Windows.Forms.Button
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnTransfer As System.Windows.Forms.Button
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ComponentSelectionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SampleEditorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
