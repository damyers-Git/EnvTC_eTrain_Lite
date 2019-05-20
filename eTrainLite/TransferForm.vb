Public Class TransferForm


    Private Sub cbo1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbo1.SelectedIndexChanged
        Dim aPermit As Permit
        Dim aMethod As Method
        Dim aInstrument As mInstrument

        If GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                'Clear out cbos because of change
                cbo2.Items.Clear()
                cbo3.Items.Clear()
                cbo4.Items.Clear()
                cbo3.Enabled = False
                cbo4.Enabled = False
                cbo5.Enabled = False
                cbo5.SelectedIndex = -1
                'Clear out loaded permits
                GlobalVariables.PermitList.Clear()
                If cbo1.Text = "LIMS" Then
                    GlobalVariables.Permit.LoadLimsLimit()
                    txt1.Text = "\\helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\FPT_Spike Recovery Limits.xlsx"
                ElseIf cbo1.Text = "eTrain File" Then
                    GlobalVariables.Permit.LoadPermitNames()
                    txt1.Text = "\\helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\FPT_Spike Recovery Limits.xlsx"
                ElseIf cbo1.Text = "Non Compliance" Then
                    GlobalVariables.Permit.LoadNonCompliance()
                End If
                'Load up Permit cbo
                For Each aPermit In GlobalVariables.PermitList
                    cbo2.Items.Add(aPermit.Name)
                Next
                cbo2.Enabled = True

            End If
        ElseIf GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                cbo2.Items.Clear()
                cbo2.Text = ""
                'Load method details
                GlobalVariables.Method.LoadMethod(cbo1.Text)
                For Each aMethod In GlobalVariables.MethodList
                    If cbo1.Text = aMethod.Name Then
                        For Each aInstrument In aMethod.mInstrumentList
                            cbo2.Items.Add(aInstrument.Name)
                        Next
                    End If
                Next
            End If
        End If

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub cbo2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbo2.SelectedIndexChanged
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aMethod As Method
        Dim aInstrument As mInstrument

        If GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                'Clear out cbo because of change
                cbo3.Items.Clear()
                cbo4.Items.Clear()
                cbo5.SelectedIndex = -1
                cbo4.Enabled = False
                cbo5.Enabled = False
                If cbo1.Text <> "LIMS" Then
                    GlobalVariables.Permit.LoadPermit(cbo2.Text)
                End If
                For Each aPermit In GlobalVariables.PermitList
                    If aPermit.Name = cbo2.Text Then
                        GlobalVariables.selPermit = aPermit
                        For Each aProject In aPermit.ProjectList
                            cbo3.Items.Add(aProject.Name)
                        Next
                    End If
                Next
                cbo3.Enabled = True
            End If
        ElseIf GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                'Associate 13c's
                For Each aMethod In GlobalVariables.MethodList
                    If aMethod.Name = cbo1.Text Then
                        For Each aInstrument In aMethod.mInstrumentList
                            If aInstrument.Name = cbo2.Text Then
                                GlobalVariables.Method.Associate13cs(aInstrument)
                            End If
                        Next
                    End If
                Next

            End If
        End If
    End Sub

    Private Sub btnSISBrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnSISBrowse.Click
        Dim fd As New OpenFileDialog()
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                fd.Title = "Open File Dialog"
                fd.InitialDirectory = "C:\"
                fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
                fd.FilterIndex = 2
                fd.RestoreDirectory = True
                If fd.ShowDialog() = DialogResult.OK Then
                    txt1.Text = fd.FileName
                End If
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                fd.Title = "Open File Dialog"
                fd.InitialDirectory = "C:\"
                fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
                fd.FilterIndex = 2
                fd.RestoreDirectory = True
                If fd.ShowDialog() = DialogResult.OK Then
                    txt1.Text = fd.FileName
                End If
            End If
        End If
    End Sub

    Private Sub btnTransfer_Click(sender As System.Object, e As System.EventArgs) Handles btnTransfer.Click
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                If txt1.Text <> "" Then
                    GlobalVariables.Calculations.MidlandFAST(txt1.Text)
                    GlobalVariables.Transfer.ToLIMS(InputBox("Please enter your UserID for LIMS Transfer", "eTrain 2.0"))
                Else
                    MsgBox("Please select the SIS associated with the samples to be transfered.", MsgBoxStyle.Critical, "eTrain 2.0")
                    txt1.Focus()
                End If
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                'Determine if LIMS Unit is setup
                For Each aPermit In GlobalVariables.PermitList
                    If aPermit.Name = GlobalVariables.selPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If aProject.Name = GlobalVariables.selProject Then
                                'aProject.LimsUnits = "" 'WT 10/19/2017 -> Temp commented for EDD Test
                                If aProject.LimsUnits = "" Then
                                    GlobalVariables.ContinueTransfer = False
                                    MsgBox("LIMS Units not detected, transfer will not continue!", MsgBoxStyle.Critical, "eTrain 2.0")
                                    Exit Sub
                                End If
                            End If
                        Next
                    End If
                Next
                SampleEdit.ShowDialog() 'WT -> Remove?
                If GlobalVariables.ContinueTransfer Then
                    CompSel.ShowDialog() 'WT -> Remove?
                    If GlobalVariables.ContinueTransfer Then
                        If GlobalVariables.Calculations.FreeportChrom(cbo5.Text, txt1.Text, True) Then
                            If GlobalVariables.Transfer.ToLIMS(InputBox("Please enter your UserID for LIMS Transfer", "eTrain 2.0")) Then
                                MsgBox("Data transfer complete!", MsgBoxStyle.Information, "eTrain 2.0")
                            Else
                                MsgBox("Error sending data to LIMS! Data not sent.", MsgBoxStyle.Critical, "eTrain 2.0")
                            End If
                        End If
                    End If
                End If
                'Else
                '   MsgBox("Please select the Recovery Limits file associated with the samples to be transfered.", MsgBoxStyle.Critical, "eTrain 2.0")
                '   txt1.Focus()
                'End If
            End If
        End If
    End Sub

    Private Sub cbo3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbo3.SelectedIndexChanged
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aInstrument As mInstrument
        If GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                cbo4.Items.Clear()
                cbo5.Enabled = False
                cbo5.SelectedIndex = -1
                For Each aPermit In GlobalVariables.PermitList
                    If aPermit.Name = cbo2.Text Then
                        For Each aProject In aPermit.ProjectList
                            If aProject.Name = cbo3.Text Then
                                GlobalVariables.selProject = aProject.Name
                                For Each aInstrument In aProject.mInstrumentList
                                    cbo4.Items.Add(aInstrument.Name)
                                Next
                            End If
                        Next
                    End If
                Next
                cbo4.Enabled = True

            End If
        End If
    End Sub

    Private Sub cbo4_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbo4.SelectedIndexChanged
        If GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                GlobalVariables.selInstrument = cbo4.Text
                cbo5.Enabled = True
            End If
        End If
    End Sub

    Private Sub ComponentSelectionToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ComponentSelectionToolStripMenuItem.Click
        CompSel.ShowDialog()
    End Sub

    Private Sub SampleEditorToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SampleEditorToolStripMenuItem.Click
        SampleEdit.ShowDialog()
    End Sub
    Private Sub Me_FormClosing(sender As Object, e As FormClosingEventArgs) _
     Handles Me.FormClosing

        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            cbo1.Items.Clear()
            cbo2.Items.Clear()
            cbo3.Items.Clear()
            cbo4.Items.Clear()
            cbo5.Items.Clear()
            txt1.Text = ""
            Me.Hide()
        End If

    End Sub

    Private Sub TransferForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                cbo1.Items.Add("eTrain File")
                cbo1.Items.Add("Non Compliance")
                cbo1.Items.Add("LIMS")
                cbo5.Items.Add("N/A")
                cbo5.Items.Add("MDL")
                cbo5.Items.Add("PQL")
                cbo5.Items.Add("RL")
            End If
        End If
    End Sub

    Private Sub cbo5_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbo5.SelectedIndexChanged

    End Sub
End Class