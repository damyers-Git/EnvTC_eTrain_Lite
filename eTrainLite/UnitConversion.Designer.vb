<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UnitConversion
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboReportedUnits = New System.Windows.Forms.ComboBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.cboSampleUnits = New System.Windows.Forms.ComboBox()
        Me.btnHelp = New System.Windows.Forms.Button()
        Me.lblNotDetect = New System.Windows.Forms.Label()
        Me.lblNotDetect2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(31, 41)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(95, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Sample Units:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(19, 108)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(107, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Reported Units:"
        '
        'cboReportedUnits
        '
        Me.cboReportedUnits.FormattingEnabled = True
        Me.cboReportedUnits.Items.AddRange(New Object() {"ppm", "ppb", "ppt", "ppq", "mg/kg", "ug/kg", "ng/kg", "ug/g", "ng/g", "pg/g", "ng/mg", "pg/mg", "mg/L", "ug/L", "ng/L", "pg/L", "ug/mL", "ng/mL", "pg/mL", "ng/uL", "pg/uL"})
        Me.cboReportedUnits.Location = New System.Drawing.Point(144, 104)
        Me.cboReportedUnits.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboReportedUnits.Name = "cboReportedUnits"
        Me.cboReportedUnits.Size = New System.Drawing.Size(123, 24)
        Me.cboReportedUnits.TabIndex = 1
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(144, 150)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(123, 42)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'cboSampleUnits
        '
        Me.cboSampleUnits.FormattingEnabled = True
        Me.cboSampleUnits.Items.AddRange(New Object() {"ppm", "ppb", "ppt", "ppq", "mg/kg", "ug/kg", "ng/kg", "ug/g", "ng/g", "pg/g", "ng/mg", "pg/mg", "mg/L", "ug/L", "ng/L", "pg/L", "ug/mL", "ng/mL", "pg/mL", "ng/uL", "pg/uL"})
        Me.cboSampleUnits.Location = New System.Drawing.Point(144, 37)
        Me.cboSampleUnits.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboSampleUnits.Name = "cboSampleUnits"
        Me.cboSampleUnits.Size = New System.Drawing.Size(123, 24)
        Me.cboSampleUnits.TabIndex = 0
        '
        'btnHelp
        '
        Me.btnHelp.Location = New System.Drawing.Point(22, 164)
        Me.btnHelp.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.Size = New System.Drawing.Size(31, 28)
        Me.btnHelp.TabIndex = 7
        Me.btnHelp.TabStop = False
        Me.btnHelp.Text = "?"
        Me.btnHelp.UseVisualStyleBackColor = True
        '
        'lblNotDetect
        '
        Me.lblNotDetect.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNotDetect.ForeColor = System.Drawing.Color.Red
        Me.lblNotDetect.Location = New System.Drawing.Point(141, 9)
        Me.lblNotDetect.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblNotDetect.Name = "lblNotDetect"
        Me.lblNotDetect.Size = New System.Drawing.Size(180, 26)
        Me.lblNotDetect.TabIndex = 2
        Me.lblNotDetect.Text = "**Units Not Detected"
        Me.lblNotDetect.Visible = False
        '
        'lblNotDetect2
        '
        Me.lblNotDetect2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNotDetect2.ForeColor = System.Drawing.Color.Black
        Me.lblNotDetect2.Location = New System.Drawing.Point(141, 76)
        Me.lblNotDetect2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblNotDetect2.Name = "lblNotDetect2"
        Me.lblNotDetect2.Size = New System.Drawing.Size(192, 26)
        Me.lblNotDetect2.TabIndex = 8
        Me.lblNotDetect2.Text = "**Units Detected In LIMS"
        Me.lblNotDetect2.Visible = False
        '
        'UnitConversion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(362, 213)
        Me.Controls.Add(Me.lblNotDetect2)
        Me.Controls.Add(Me.btnHelp)
        Me.Controls.Add(Me.cboSampleUnits)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.cboReportedUnits)
        Me.Controls.Add(Me.lblNotDetect)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "UnitConversion"
        Me.Text = "Unit Conversion"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboReportedUnits As System.Windows.Forms.ComboBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents cboSampleUnits As System.Windows.Forms.ComboBox
    Friend WithEvents btnHelp As System.Windows.Forms.Button
    Friend WithEvents lblNotDetect As System.Windows.Forms.Label
    Friend WithEvents lblNotDetect2 As System.Windows.Forms.Label
End Class
