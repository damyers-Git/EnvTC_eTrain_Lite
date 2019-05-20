<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SampleMerge
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
        Me.cboSampleList = New System.Windows.Forms.ComboBox()
        Me.dgComponents = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ManualToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.btnCombine = New System.Windows.Forms.Button()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.chkParent = New System.Windows.Forms.CheckBox()
        CType(Me.dgComponents, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboSampleList
        '
        Me.cboSampleList.FormattingEnabled = True
        Me.cboSampleList.Location = New System.Drawing.Point(95, 63)
        Me.cboSampleList.Name = "cboSampleList"
        Me.cboSampleList.Size = New System.Drawing.Size(235, 24)
        Me.cboSampleList.TabIndex = 0
        '
        'dgComponents
        '
        Me.dgComponents.AllowUserToAddRows = False
        Me.dgComponents.AllowUserToDeleteRows = False
        Me.dgComponents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgComponents.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2})
        Me.dgComponents.Location = New System.Drawing.Point(33, 108)
        Me.dgComponents.Name = "dgComponents"
        Me.dgComponents.RowTemplate.Height = 24
        Me.dgComponents.Size = New System.Drawing.Size(506, 433)
        Me.dgComponents.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(30, 63)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Sample:"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(572, 28)
        Me.MenuStrip1.TabIndex = 3
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ManualToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(44, 24)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'ManualToolStripMenuItem
        '
        Me.ManualToolStripMenuItem.Enabled = False
        Me.ManualToolStripMenuItem.Name = "ManualToolStripMenuItem"
        Me.ManualToolStripMenuItem.Size = New System.Drawing.Size(191, 24)
        Me.ManualToolStripMenuItem.Text = "&Manual Combine"
        '
        'btnCombine
        '
        Me.btnCombine.Location = New System.Drawing.Point(209, 556)
        Me.btnCombine.Name = "btnCombine"
        Me.btnCombine.Size = New System.Drawing.Size(154, 48)
        Me.btnCombine.TabIndex = 4
        Me.btnCombine.Text = "Merge"
        Me.btnCombine.UseVisualStyleBackColor = True
        '
        'Col1
        '
        Me.Col1.HeaderText = "Component"
        Me.Col1.Name = "Col1"
        Me.Col1.ReadOnly = True
        Me.Col1.Width = 232
        '
        'Col2
        '
        Me.Col2.HeaderText = "Include"
        Me.Col2.Name = "Col2"
        Me.Col2.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Col2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Col2.Width = 201
        '
        'chkParent
        '
        Me.chkParent.AutoSize = True
        Me.chkParent.Location = New System.Drawing.Point(369, 62)
        Me.chkParent.Name = "chkParent"
        Me.chkParent.Size = New System.Drawing.Size(126, 21)
        Me.chkParent.TabIndex = 5
        Me.chkParent.Text = "Mark as Parent"
        Me.chkParent.UseVisualStyleBackColor = True
        '
        'SampleMerge
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(572, 616)
        Me.Controls.Add(Me.chkParent)
        Me.Controls.Add(Me.btnCombine)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgComponents)
        Me.Controls.Add(Me.cboSampleList)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "SampleMerge"
        Me.Text = "Sample Merge"
        CType(Me.dgComponents, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboSampleList As System.Windows.Forms.ComboBox
    Friend WithEvents dgComponents As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ManualToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnCombine As System.Windows.Forms.Button
    Friend WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents chkParent As System.Windows.Forms.CheckBox
End Class
