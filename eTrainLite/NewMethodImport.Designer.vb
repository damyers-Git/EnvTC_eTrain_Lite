<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NewMethodImport
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
        Me.txt1 = New System.Windows.Forms.TextBox()
        Me.txt2 = New System.Windows.Forms.TextBox()
        Me.Label02 = New System.Windows.Forms.Label()
        Me.txtInstrument = New System.Windows.Forms.TextBox()
        Me.txtMethodName = New System.Windows.Forms.TextBox()
        Me.Label01 = New System.Windows.Forms.Label()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.radXls = New System.Windows.Forms.RadioButton()
        Me.radMH = New System.Windows.Forms.RadioButton()
        Me.radChem = New System.Windows.Forms.RadioButton()
        Me.btn1 = New System.Windows.Forms.Button()
        Me.btn2 = New System.Windows.Forms.Button()
        Me.lblNote = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 70)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Calrpt.txt:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Daprtmth.txt:"
        '
        'txt1
        '
        Me.txt1.Location = New System.Drawing.Point(101, 67)
        Me.txt1.Name = "txt1"
        Me.txt1.Size = New System.Drawing.Size(316, 20)
        Me.txt1.TabIndex = 2
        '
        'txt2
        '
        Me.txt2.Location = New System.Drawing.Point(101, 93)
        Me.txt2.Name = "txt2"
        Me.txt2.Size = New System.Drawing.Size(316, 20)
        Me.txt2.TabIndex = 4
        '
        'Label02
        '
        Me.Label02.AutoSize = True
        Me.Label02.Location = New System.Drawing.Point(16, 44)
        Me.Label02.Name = "Label02"
        Me.Label02.Size = New System.Drawing.Size(59, 13)
        Me.Label02.TabIndex = 4
        Me.Label02.Text = "Instrument:"
        '
        'txtInstrument
        '
        Me.txtInstrument.Location = New System.Drawing.Point(101, 41)
        Me.txtInstrument.Name = "txtInstrument"
        Me.txtInstrument.Size = New System.Drawing.Size(229, 20)
        Me.txtInstrument.TabIndex = 1
        '
        'txtMethodName
        '
        Me.txtMethodName.Location = New System.Drawing.Point(143, 15)
        Me.txtMethodName.Name = "txtMethodName"
        Me.txtMethodName.Size = New System.Drawing.Size(187, 20)
        Me.txtMethodName.TabIndex = 0
        '
        'Label01
        '
        Me.Label01.AutoSize = True
        Me.Label01.Location = New System.Drawing.Point(16, 18)
        Me.Label01.Name = "Label01"
        Me.Label01.Size = New System.Drawing.Size(121, 13)
        Me.Label01.TabIndex = 6
        Me.Label01.Text = "Method / Project Name:"
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(221, 165)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(75, 35)
        Me.btnImport.TabIndex = 9
        Me.btnImport.Text = "Import"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.radXls)
        Me.GroupBox1.Controls.Add(Me.radMH)
        Me.GroupBox1.Controls.Add(Me.radChem)
        Me.GroupBox1.Location = New System.Drawing.Point(336, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(165, 50)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Import From:"
        '
        'radXls
        '
        Me.radXls.AutoSize = True
        Me.radXls.Location = New System.Drawing.Point(113, 20)
        Me.radXls.Name = "radXls"
        Me.radXls.Size = New System.Drawing.Size(45, 17)
        Me.radXls.TabIndex = 2
        Me.radXls.Text = "XLS"
        Me.radXls.UseVisualStyleBackColor = True
        '
        'radMH
        '
        Me.radMH.AutoSize = True
        Me.radMH.Location = New System.Drawing.Point(65, 20)
        Me.radMH.Name = "radMH"
        Me.radMH.Size = New System.Drawing.Size(42, 17)
        Me.radMH.TabIndex = 1
        Me.radMH.Text = "MH"
        Me.radMH.UseVisualStyleBackColor = True
        '
        'radChem
        '
        Me.radChem.AutoSize = True
        Me.radChem.Checked = True
        Me.radChem.Location = New System.Drawing.Point(7, 20)
        Me.radChem.Name = "radChem"
        Me.radChem.Size = New System.Drawing.Size(52, 17)
        Me.radChem.TabIndex = 0
        Me.radChem.TabStop = True
        Me.radChem.Text = "Chem"
        Me.radChem.UseVisualStyleBackColor = True
        '
        'btn1
        '
        Me.btn1.Location = New System.Drawing.Point(426, 65)
        Me.btn1.Name = "btn1"
        Me.btn1.Size = New System.Drawing.Size(75, 23)
        Me.btn1.TabIndex = 3
        Me.btn1.Text = "Browse..."
        Me.btn1.UseVisualStyleBackColor = True
        '
        'btn2
        '
        Me.btn2.Location = New System.Drawing.Point(426, 91)
        Me.btn2.Name = "btn2"
        Me.btn2.Size = New System.Drawing.Size(75, 23)
        Me.btn2.TabIndex = 5
        Me.btn2.Text = "Browse..."
        Me.btn2.UseVisualStyleBackColor = True
        '
        'lblNote
        '
        Me.lblNote.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNote.ForeColor = System.Drawing.Color.Red
        Me.lblNote.Location = New System.Drawing.Point(119, 124)
        Me.lblNote.Name = "lblNote"
        Me.lblNote.Size = New System.Drawing.Size(279, 38)
        Me.lblNote.TabIndex = 21
        Me.lblNote.Text = "***Calibration amount (ng) for 13C's and Injection will need to be manually enter" & _
    "ed after import!"
        Me.lblNote.Visible = False
        '
        'NewMethodImport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(516, 212)
        Me.Controls.Add(Me.lblNote)
        Me.Controls.Add(Me.btn2)
        Me.Controls.Add(Me.btn1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.txtMethodName)
        Me.Controls.Add(Me.Label01)
        Me.Controls.Add(Me.txtInstrument)
        Me.Controls.Add(Me.Label02)
        Me.Controls.Add(Me.txt2)
        Me.Controls.Add(Me.txt1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "NewMethodImport"
        Me.Text = "Using Imported Data"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txt1 As System.Windows.Forms.TextBox
    Friend WithEvents txt2 As System.Windows.Forms.TextBox
    Friend WithEvents Label02 As System.Windows.Forms.Label
    Friend WithEvents txtInstrument As System.Windows.Forms.TextBox
    Friend WithEvents txtMethodName As System.Windows.Forms.TextBox
    Friend WithEvents Label01 As System.Windows.Forms.Label
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents radXls As System.Windows.Forms.RadioButton
    Friend WithEvents radMH As System.Windows.Forms.RadioButton
    Friend WithEvents radChem As System.Windows.Forms.RadioButton
    Friend WithEvents btn1 As System.Windows.Forms.Button
    Friend WithEvents btn2 As System.Windows.Forms.Button
    Friend WithEvents lblNote As System.Windows.Forms.Label
End Class
