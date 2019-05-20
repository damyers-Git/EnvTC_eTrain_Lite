<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SpikeInfo
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
        Me.btnCalc = New System.Windows.Forms.Button()
        Me.txtVol = New System.Windows.Forms.TextBox()
        Me.txtConc = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblSampleName = New System.Windows.Forms.Label()
        Me.lblUnits = New System.Windows.Forms.Label()
        Me.lblDil = New System.Windows.Forms.Label()
        Me.btnTip = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnCalc
        '
        Me.btnCalc.Location = New System.Drawing.Point(72, 197)
        Me.btnCalc.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCalc.Name = "btnCalc"
        Me.btnCalc.Size = New System.Drawing.Size(169, 57)
        Me.btnCalc.TabIndex = 0
        Me.btnCalc.Text = "Calculate Corrected Spike Amt"
        Me.btnCalc.UseVisualStyleBackColor = True
        '
        'txtVol
        '
        Me.txtVol.Location = New System.Drawing.Point(141, 121)
        Me.txtVol.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtVol.Name = "txtVol"
        Me.txtVol.Size = New System.Drawing.Size(132, 22)
        Me.txtVol.TabIndex = 1
        '
        'txtConc
        '
        Me.txtConc.Location = New System.Drawing.Point(141, 153)
        Me.txtConc.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtConc.Name = "txtConc"
        Me.txtConc.Size = New System.Drawing.Size(132, 22)
        Me.txtConc.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(32, 124)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Spike Vol:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(32, 156)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Concentration:"
        '
        'lblSampleName
        '
        Me.lblSampleName.AutoSize = True
        Me.lblSampleName.Location = New System.Drawing.Point(32, 23)
        Me.lblSampleName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSampleName.Name = "lblSampleName"
        Me.lblSampleName.Size = New System.Drawing.Size(199, 17)
        Me.lblSampleName.TabIndex = 5
        Me.lblSampleName.Text = "Sample: Test Sample MS DUP"
        '
        'lblUnits
        '
        Me.lblUnits.AutoSize = True
        Me.lblUnits.Location = New System.Drawing.Point(32, 89)
        Me.lblUnits.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblUnits.Name = "lblUnits"
        Me.lblUnits.Size = New System.Drawing.Size(209, 17)
        Me.lblUnits.TabIndex = 6
        Me.lblUnits.Text = "Ending Conversion Units: (ug/L)"
        '
        'lblDil
        '
        Me.lblDil.AutoSize = True
        Me.lblDil.Location = New System.Drawing.Point(32, 57)
        Me.lblDil.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblDil.Name = "lblDil"
        Me.lblDil.Size = New System.Drawing.Size(115, 17)
        Me.lblDil.TabIndex = 7
        Me.lblDil.Text = "Dilution Factor: 1"
        '
        'btnTip
        '
        Me.btnTip.Location = New System.Drawing.Point(263, 235)
        Me.btnTip.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnTip.Name = "btnTip"
        Me.btnTip.Size = New System.Drawing.Size(33, 26)
        Me.btnTip.TabIndex = 4
        Me.btnTip.TabStop = False
        Me.btnTip.Text = "?"
        Me.btnTip.UseVisualStyleBackColor = True
        '
        'SpikeInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(312, 276)
        Me.Controls.Add(Me.btnTip)
        Me.Controls.Add(Me.lblDil)
        Me.Controls.Add(Me.lblUnits)
        Me.Controls.Add(Me.lblSampleName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtConc)
        Me.Controls.Add(Me.txtVol)
        Me.Controls.Add(Me.btnCalc)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "SpikeInfo"
        Me.Text = "Spike Information"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCalc As System.Windows.Forms.Button
    Friend WithEvents txtVol As System.Windows.Forms.TextBox
    Friend WithEvents txtConc As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblUnits As System.Windows.Forms.Label
    Friend WithEvents lblDil As System.Windows.Forms.Label
    Friend WithEvents lblSampleName As System.Windows.Forms.Label
    Friend WithEvents btnTip As System.Windows.Forms.Button
End Class
