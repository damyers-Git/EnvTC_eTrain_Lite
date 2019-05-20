<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RecoveryLimitSelect
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
        Me.txtLimitPath = New System.Windows.Forms.TextBox()
        Me.btnFindSheets = New System.Windows.Forms.Button()
        Me.cboSheetName = New System.Windows.Forms.ComboBox()
        Me.btnLoadLimits = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtLimitPath
        '
        Me.txtLimitPath.Location = New System.Drawing.Point(12, 31)
        Me.txtLimitPath.Name = "txtLimitPath"
        Me.txtLimitPath.Size = New System.Drawing.Size(261, 20)
        Me.txtLimitPath.TabIndex = 0
        '
        'btnFindSheets
        '
        Me.btnFindSheets.Location = New System.Drawing.Point(308, 20)
        Me.btnFindSheets.Name = "btnFindSheets"
        Me.btnFindSheets.Size = New System.Drawing.Size(95, 40)
        Me.btnFindSheets.TabIndex = 1
        Me.btnFindSheets.Text = "Find Recovery Limits"
        Me.btnFindSheets.UseVisualStyleBackColor = True
        '
        'cboSheetName
        '
        Me.cboSheetName.Enabled = False
        Me.cboSheetName.FormattingEnabled = True
        Me.cboSheetName.Location = New System.Drawing.Point(12, 98)
        Me.cboSheetName.Name = "cboSheetName"
        Me.cboSheetName.Size = New System.Drawing.Size(260, 21)
        Me.cboSheetName.TabIndex = 2
        '
        'btnLoadLimits
        '
        Me.btnLoadLimits.Enabled = False
        Me.btnLoadLimits.Location = New System.Drawing.Point(308, 87)
        Me.btnLoadLimits.Name = "btnLoadLimits"
        Me.btnLoadLimits.Size = New System.Drawing.Size(95, 40)
        Me.btnLoadLimits.TabIndex = 3
        Me.btnLoadLimits.Text = "Load Recovery Limits"
        Me.btnLoadLimits.UseVisualStyleBackColor = True
        '
        'RecoveryLimitSelect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(415, 148)
        Me.Controls.Add(Me.btnLoadLimits)
        Me.Controls.Add(Me.cboSheetName)
        Me.Controls.Add(Me.btnFindSheets)
        Me.Controls.Add(Me.txtLimitPath)
        Me.Name = "RecoveryLimitSelect"
        Me.Text = "Recovery Limits"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtLimitPath As System.Windows.Forms.TextBox
    Friend WithEvents btnFindSheets As System.Windows.Forms.Button
    Friend WithEvents cboSheetName As System.Windows.Forms.ComboBox
    Friend WithEvents btnLoadLimits As System.Windows.Forms.Button
End Class
