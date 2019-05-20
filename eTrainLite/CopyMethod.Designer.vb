<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CopyMethod
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
        Me.cboCopyFrom = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtNMethod = New System.Windows.Forms.TextBox()
        Me.btnCopy = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(39, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Copy from:"
        '
        'cboCopyFrom
        '
        Me.cboCopyFrom.FormattingEnabled = True
        Me.cboCopyFrom.Location = New System.Drawing.Point(112, 29)
        Me.cboCopyFrom.Name = "cboCopyFrom"
        Me.cboCopyFrom.Size = New System.Drawing.Size(196, 21)
        Me.cboCopyFrom.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(18, 77)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Name for New:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtNMethod
        '
        Me.txtNMethod.Location = New System.Drawing.Point(112, 74)
        Me.txtNMethod.Name = "txtNMethod"
        Me.txtNMethod.Size = New System.Drawing.Size(196, 20)
        Me.txtNMethod.TabIndex = 3
        '
        'btnCopy
        '
        Me.btnCopy.Location = New System.Drawing.Point(126, 108)
        Me.btnCopy.Name = "btnCopy"
        Me.btnCopy.Size = New System.Drawing.Size(75, 35)
        Me.btnCopy.TabIndex = 4
        Me.btnCopy.Text = "Copy"
        Me.btnCopy.UseVisualStyleBackColor = True
        '
        'CopyMethod
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(327, 161)
        Me.Controls.Add(Me.btnCopy)
        Me.Controls.Add(Me.txtNMethod)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboCopyFrom)
        Me.Controls.Add(Me.Label1)
        Me.Name = "CopyMethod"
        Me.Text = "Copy From Existing"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboCopyFrom As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtNMethod As System.Windows.Forms.TextBox
    Friend WithEvents btnCopy As System.Windows.Forms.Button
End Class
