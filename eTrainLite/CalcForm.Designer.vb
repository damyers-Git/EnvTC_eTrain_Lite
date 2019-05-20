<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CalcForm
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5})
        Me.DataGridView1.Location = New System.Drawing.Point(24, 38)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(544, 324)
        Me.DataGridView1.TabIndex = 0
        '
        'Col1
        '
        Me.Col1.HeaderText = "Col1"
        Me.Col1.Name = "Col1"
        Me.Col1.ReadOnly = True
        '
        'Col2
        '
        Me.Col2.HeaderText = "Col2"
        Me.Col2.Name = "Col2"
        Me.Col2.ReadOnly = True
        '
        'Col3
        '
        Me.Col3.HeaderText = "Col3"
        Me.Col3.Name = "Col3"
        Me.Col3.ReadOnly = True
        '
        'Col4
        '
        Me.Col4.HeaderText = "Col4"
        Me.Col4.Name = "Col4"
        '
        'Col5
        '
        Me.Col5.HeaderText = "Col5"
        Me.Col5.Name = "Col5"
        '
        'CalcForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(593, 457)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "CalcForm"
        Me.Text = "Calculation Details - eTrain 2.0"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col5 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
