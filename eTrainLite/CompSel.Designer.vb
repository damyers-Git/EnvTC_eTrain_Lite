<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CompSel
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
        Me.dgCompSelect = New System.Windows.Forms.DataGridView()
        Me.btnSel = New System.Windows.Forms.Button()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.col2 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        CType(Me.dgCompSelect, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgCompSelect
        '
        Me.dgCompSelect.AllowUserToAddRows = False
        Me.dgCompSelect.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgCompSelect.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgCompSelect.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgCompSelect.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.col1, Me.col2})
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgCompSelect.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgCompSelect.Location = New System.Drawing.Point(28, 21)
        Me.dgCompSelect.Margin = New System.Windows.Forms.Padding(4)
        Me.dgCompSelect.Name = "dgCompSelect"
        Me.dgCompSelect.Size = New System.Drawing.Size(633, 335)
        Me.dgCompSelect.TabIndex = 0
        '
        'btnSel
        '
        Me.btnSel.Location = New System.Drawing.Point(350, 378)
        Me.btnSel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSel.Name = "btnSel"
        Me.btnSel.Size = New System.Drawing.Size(123, 43)
        Me.btnSel.TabIndex = 1
        Me.btnSel.Text = "Continue"
        Me.btnSel.UseVisualStyleBackColor = True
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(219, 378)
        Me.btnBack.Margin = New System.Windows.Forms.Padding(4)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(123, 43)
        Me.btnBack.TabIndex = 3
        Me.btnBack.Text = "Back"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'col1
        '
        Me.col1.HeaderText = "Component"
        Me.col1.Name = "col1"
        Me.col1.ReadOnly = True
        Me.col1.Width = 200
        '
        'col2
        '
        Me.col2.HeaderText = "Include"
        Me.col2.Name = "col2"
        Me.col2.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.col2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'CompSel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(692, 439)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.btnSel)
        Me.Controls.Add(Me.dgCompSelect)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "CompSel"
        Me.Text = "Component Selection"
        CType(Me.dgCompSelect, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgCompSelect As System.Windows.Forms.DataGridView
    Friend WithEvents btnSel As System.Windows.Forms.Button
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col2 As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
