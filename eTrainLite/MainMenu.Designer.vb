<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainMenu
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
        Me.btnImport = New System.Windows.Forms.Button()
        Me.btnReport = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LocationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MidlandToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FastToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HighResToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ChromToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FreeportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnImport
        '
        Me.btnImport.Enabled = False
        Me.btnImport.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnImport.Location = New System.Drawing.Point(76, 55)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(105, 44)
        Me.btnImport.TabIndex = 1
        Me.btnImport.Text = "Import Samples"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'btnReport
        '
        Me.btnReport.Enabled = False
        Me.btnReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReport.Location = New System.Drawing.Point(76, 121)
        Me.btnReport.Name = "btnReport"
        Me.btnReport.Size = New System.Drawing.Size(105, 44)
        Me.btnReport.TabIndex = 2
        Me.btnReport.Text = "Generate Reports"
        Me.btnReport.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(257, 24)
        Me.MenuStrip1.TabIndex = 3
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LocationToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'LocationToolStripMenuItem
        '
        Me.LocationToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MidlandToolStripMenuItem, Me.FreeportToolStripMenuItem})
        Me.LocationToolStripMenuItem.Name = "LocationToolStripMenuItem"
        Me.LocationToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.LocationToolStripMenuItem.Text = "&Location"
        '
        'MidlandToolStripMenuItem
        '
        Me.MidlandToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FastToolStripMenuItem, Me.HighResToolStripMenuItem, Me.ChromToolStripMenuItem})
        Me.MidlandToolStripMenuItem.Name = "MidlandToolStripMenuItem"
        Me.MidlandToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.MidlandToolStripMenuItem.Text = "&Midland"
        '
        'FastToolStripMenuItem
        '
        Me.FastToolStripMenuItem.Name = "FastToolStripMenuItem"
        Me.FastToolStripMenuItem.Size = New System.Drawing.Size(121, 22)
        Me.FastToolStripMenuItem.Text = "&Fast"
        '
        'HighResToolStripMenuItem
        '
        Me.HighResToolStripMenuItem.Name = "HighResToolStripMenuItem"
        Me.HighResToolStripMenuItem.Size = New System.Drawing.Size(121, 22)
        Me.HighResToolStripMenuItem.Text = "&High Res"
        '
        'ChromToolStripMenuItem
        '
        Me.ChromToolStripMenuItem.Name = "ChromToolStripMenuItem"
        Me.ChromToolStripMenuItem.Size = New System.Drawing.Size(121, 22)
        Me.ChromToolStripMenuItem.Text = "&Chrom"
        '
        'FreeportToolStripMenuItem
        '
        Me.FreeportToolStripMenuItem.Name = "FreeportToolStripMenuItem"
        Me.FreeportToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.FreeportToolStripMenuItem.Text = "F&reeport"
        '
        'MainMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(257, 221)
        Me.Controls.Add(Me.btnReport)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "MainMenu"
        Me.Text = "Main Menu - eTrain"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents btnReport As System.Windows.Forms.Button
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LocationToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MidlandToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FastToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HighResToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ChromToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FreeportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
