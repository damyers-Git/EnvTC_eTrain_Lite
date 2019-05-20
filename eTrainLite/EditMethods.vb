Imports System.IO
Public Class EditMethods

    Private strSelInstrument As String
    Private strSelProject As String
    Private blnEdit As Boolean

    Private Sub cboOption1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboOption1.SelectedIndexChanged
        Dim aMethod As Method
        Dim aPermit As Permit
        Dim aProject As Project

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                EditCurrentMethodToolStripMenuItem.Enabled = False
                cboOption2.Items.Clear()
                btnOption4.Visible = False
                dgCompound.Rows.Clear()
                cboOption2.Text = ""
                cboOption3.Text = ""
                
                txtOption4.Text = ""
                txtOption5.Text = ""
                cboOption2.Enabled = True
                cboOption3.Enabled = False
                'Load method details
                GlobalVariables.Method.LoadMethod(cboOption1.Text)
                For Each aMethod In GlobalVariables.MethodList
                    If cboOption1.Text = aMethod.Name Then
                        GlobalVariables.selMethod = aMethod
                        For Each aInstrument In aMethod.mInstrumentList
                            cboOption2.Items.Add(aInstrument.Name)
                        Next
                        
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                EditCurrentMethodToolStripMenuItem.Enabled = False
                cboOption2.Items.Clear()
                btnOption4.Visible = False
                dgCompound.Rows.Clear()
                cboOption2.Text = ""
                cboOption3.Text = ""

                txtOption4.Text = ""
                txtOption5.Text = ""
                cboOption2.Enabled = True
                cboOption3.Enabled = False
                'Load method details
                GlobalVariables.Method.LoadMethod(cboOption1.Text)
                For Each aMethod In GlobalVariables.MethodList
                    If cboOption1.Text = aMethod.Name Then
                        GlobalVariables.selMethod = aMethod
                        For Each aInstrument In aMethod.mInstrumentList
                            cboOption2.Items.Add(aInstrument.Name)
                        Next

                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                cboOption2.Items.Clear()
                cboOption3.Items.Clear()
                dgCompound.Rows.Clear()
                strSelProject = ""
                strSelInstrument = ""
                cboOption2.Text = ""
                cboOption3.Text = ""
                
                cboOption2.Enabled = True
                cboOption3.Enabled = False
                'Load permit details
                GlobalVariables.Permit.LoadPermit(cboOption1.Text)
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            cboOption2.Items.Add(aProject.Name)
                        Next
                    End If
                Next
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                cboOption2.Items.Clear()
                cboOption3.Items.Clear()
                dgCompound.Rows.Clear()
                strSelProject = ""
                strSelInstrument = ""
                cboOption2.Text = ""
                cboOption3.Text = ""
               
                cboOption2.Enabled = True
                cboOption3.Enabled = False
                'Load permit details
                GlobalVariables.Permit.LoadPermit(cboOption1.Text)
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            cboOption2.Items.Add(aProject.Name)
                        Next
                    End If
                Next
            End If
        End If


    End Sub

    Private Sub cboOption2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboOption2.SelectedIndexChanged
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aInstrument As mInstrument

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                cboOption3.Enabled = True
                btnOption4.Visible = False
                EditCurrentMethodToolStripMenuItem.Enabled = False
                'Check if editing or not
                If blnEdit Then
                    'Clear compounds and load in new table
                    SaveDetailsToObject()
                End If
                strSelInstrument = cboOption2.Text
                cboOption3.Text = ""
                dgCompound.Rows.Clear()
                If blnEdit = True Then
                    cboOption3.Text = "Standard"
                    'Fills datagrid
                    FillDataGrid()
                End If

            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                cboOption3.Enabled = True
                btnOption4.Visible = False
                EditCurrentMethodToolStripMenuItem.Enabled = False
                'Check if editing or not
                If blnEdit Then
                    'Clear compounds and load in new table
                    SaveDetailsToObject()
                End If
                strSelInstrument = cboOption2.Text
                cboOption3.Text = ""
                dgCompound.Rows.Clear()
                If blnEdit = True Then
                    cboOption3.Text = "Standard"
                    'Fills datagrid
                    FillDataGrid()
                End If
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then

                If blnEdit Then
                    'Clear compounds and load in new table
                    SaveDetailsToObject()
                End If
                strSelProject = cboOption2.Text
                cboOption3.Enabled = True
                cboOption3.Items.Clear()
                cboOption3.Text = ""
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If cboOption2.Text = aProject.Name Then
                                For Each aInstrument In aProject.mInstrumentList
                                    cboOption3.Items.Add(aInstrument.Name)
                                Next
                            End If
                        Next
                    End If
                Next

                FillDataGrid()
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                If blnEdit Then
                    'Clear compounds and load in new table
                    SaveDetailsToObject()
                End If
                strSelProject = cboOption2.Text
                cboOption3.Enabled = True
                cboOption3.Items.Clear()
                cboOption3.Text = ""
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If cboOption2.Text = aProject.Name Then
                                For Each aInstrument In aProject.mInstrumentList
                                    cboOption3.Items.Add(aInstrument.Name)
                                Next
                            End If
                        Next
                    End If
                Next

                FillDataGrid()
            End If
        End If


    End Sub

    Private Sub EditCurrentMethodToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EditCurrentMethodToolStripMenuItem.Click
        EditForm()
    End Sub

    Sub EditForm()
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                If EditCurrentMethodToolStripMenuItem.Checked Then
                    btnSave.Enabled = False
                    btnAddCompound.Visible = False
                    btnOption2Add.Visible = False
                    btnOption3Add.Visible = False
                    btnDelCompound.Visible = False
                    btnOption2Del.Visible = False
                    btnOption3Del.Visible = False
                    cboOption1.Enabled = True
                    cboOption2.Enabled = True
                    If cboOption3.Text = "Standard" Then
                        btnOption3Add.Text = "Add Standard"
                        btnOption3Del.Text = "Delete Standard"
                    Else
                        btnOption3Add.Text = "Add Compound"
                        btnOption3Del.Text = "Delete Compound"
                    End If
                    blnEdit = False
                    dgCompound.ScrollBars = ScrollBars.None
                    dgCompound.ScrollBars = ScrollBars.Both
                    EditCurrentMethodToolStripMenuItem.Checked = False
                Else
                    btnSave.Enabled = True
                    btnAddCompound.Visible = False
                    btnOption2Add.Visible = True
                    btnOption3Add.Visible = True
                    btnDelCompound.Visible = False
                    btnOption2Del.Visible = True
                    btnOption3Del.Visible = True
                    cboOption1.Enabled = False
                    cboOption2.Enabled = False
                    btnOption3Add.Visible = True
                    If cboOption3.Text = "Standard" Then
                        btnOption3Add.Text = "Add Standard"
                        btnOption3Del.Text = "Delete Standard"
                    Else
                        btnOption3Add.Text = "Add Compound"
                        btnOption3Del.Text = "Delete Compound"
                    End If
                    blnEdit = True
                    dgCompound.ScrollBars = ScrollBars.None
                    dgCompound.ScrollBars = ScrollBars.Both
                    EditCurrentMethodToolStripMenuItem.Checked = True
                End If
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                If EditCurrentMethodToolStripMenuItem.Checked Then
                    btnSave.Enabled = False
                    btnAddCompound.Visible = False
                    btnOption2Add.Visible = False
                    btnOption3Add.Visible = False
                    btnDelCompound.Visible = False
                    btnOption2Del.Visible = False
                    btnOption3Del.Visible = False
                    cboOption1.Enabled = True
                    cboOption2.Enabled = True
                    If cboOption3.Text = "Standard" Then
                        btnOption3Add.Text = "Add Standard"
                        btnOption3Del.Text = "Delete Standard"
                    Else
                        btnOption3Add.Text = "Add Compound"
                        btnOption3Del.Text = "Delete Compound"
                    End If
                    blnEdit = False
                    dgCompound.ScrollBars = ScrollBars.None
                    dgCompound.ScrollBars = ScrollBars.Both
                    EditCurrentMethodToolStripMenuItem.Checked = False
                Else
                    btnSave.Enabled = True
                    btnAddCompound.Visible = False
                    btnOption2Add.Visible = True
                    btnOption3Add.Visible = True
                    btnDelCompound.Visible = False
                    btnOption2Del.Visible = True
                    btnOption3Del.Visible = True
                    cboOption1.Enabled = False
                    cboOption2.Enabled = False
                    btnOption3Add.Visible = True
                    If cboOption3.Text = "Standard" Then
                        btnOption3Add.Text = "Add Standard"
                        btnOption3Del.Text = "Delete Standard"
                    Else
                        btnOption3Add.Text = "Add Compound"
                        btnOption3Del.Text = "Delete Compound"
                    End If
                    blnEdit = True
                    dgCompound.ScrollBars = ScrollBars.None
                    dgCompound.ScrollBars = ScrollBars.Both
                    EditCurrentMethodToolStripMenuItem.Checked = True
                End If
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                btnSave.Enabled = True
                btnAddCompound.Visible = True
                btnOption2Add.Visible = True
                btnDelCompound.Visible = True
                btnOption2Del.Visible = True
                btnOption3Add.Visible = True
                btnOption3Del.Visible = True
                cboOption1.Enabled = False
                blnEdit = True
                dgCompound.ScrollBars = ScrollBars.None
                dgCompound.ScrollBars = ScrollBars.Both
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                btnSave.Enabled = True
                btnAddCompound.Visible = True
                btnOption2Add.Visible = True
                btnDelCompound.Visible = True
                btnOption2Del.Visible = True
                btnOption3Add.Visible = True
                btnOption3Del.Visible = True
                cboOption1.Enabled = False
                blnEdit = True
                dgCompound.ScrollBars = ScrollBars.None
                dgCompound.ScrollBars = ScrollBars.Both
            End If
        End If
    End Sub

    Private Sub btnOption2Add_Click(sender As System.Object, e As System.EventArgs) Handles btnOption2Add.Click
        Dim aInstrument As mInstrument
        Dim aExistInstrument As mInstrument
        Dim aMethod As Method
        Dim aPermit As Permit
        Dim aProject As Project
        Dim strCalPath As String
        Dim strDapPath As String

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                'add in new instrument
                aInstrument = New mInstrument
                aInstrument.Name = InputBox("Please enter the Instrument name:", "Edit Method - eTrain 2.0")
                aInstrument.Reviewed = False
                aInstrument.ReviewedDate = CDate("1/1/1970")
                If aInstrument.Name <> "" Then
                    For Each aMethod In GlobalVariables.MethodList
                        If aMethod.Name = cboOption1.Text Then
                            'Copy Everything std/comp wise except cal information if another instrument exists
                            If aMethod.mInstrumentList.Count > 0 Then
                                For Each aExistInstrument In aMethod.mInstrumentList
                                    aInstrument.CopyMethodInfo(aExistInstrument)
                                    Exit For
                                Next
                            End If
                            aMethod.mInstrumentList.Add(aInstrument)
                            If MsgBox("Do you want to import Calibration data for this instrument?", MsgBoxStyle.YesNo, "eTrain 2.0") Then
                                strCalPath = InputBox("Please enter the Path to the CalRpt.txt you wish to import from.", "eTrain 2.0")
                                strDapPath = InputBox("Please enter the Path to the Daprtmth.txt you wish to import from.", "eTrain 2.0")
                                If File.Exists(strCalPath) Then
                                    NewMethodImport.ChemCalrptImport(aMethod, aInstrument.Name, strCalPath)
                                Else
                                    MsgBox("Calrpt.txt not found.", MsgBoxStyle.Exclamation, "eTrain 2.0")
                                End If
                                If File.Exists(strCalPath) Then
                                    NewMethodImport.ChemDaprtmthImport(aMethod, aInstrument.Name, strDapPath)
                                Else
                                    MsgBox("Daprtmth.txt not found.", MsgBoxStyle.Exclamation, "eTrain 2.0")
                                End If
                            End If
                            blnEdit = False
                            'Update instrument combo box
                            cboOption2.Items.Add(aInstrument.Name)
                            cboOption2.Text = aInstrument.Name
                            lblStatus.Text = "Reviewed: " & aInstrument.Reviewed
                            cboOption3.Text = "Standard"
                            btnOption3Add.Text = "Add Standard"
                            btnOption3Del.Text = "Delete Standard"
                            btnOption3Add.Visible = True
                            btnOption3Del.Visible = True
                            If dgCompound.Rows.Count <= 1 Then
                                FillDataGrid()
                            End If
                            blnEdit = True
                        End If
                    Next
                End If
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                'add in new instrument
                aInstrument = New mInstrument
                aInstrument.Name = InputBox("Please enter the Instrument name:", "Edit Method - eTrain 2.0")
                aInstrument.Reviewed = False
                aInstrument.ReviewedDate = CDate("1/1/1970")
                If aInstrument.Name <> "" Then
                    For Each aMethod In GlobalVariables.MethodList
                        If aMethod.Name = cboOption1.Text Then
                            'Copy Everything std/comp wise except cal information if another instrument exists
                            If aMethod.mInstrumentList.Count > 0 Then
                                For Each aExistInstrument In aMethod.mInstrumentList
                                    aInstrument.CopyMethodInfo(aExistInstrument)
                                    Exit For
                                Next
                            End If
                            aMethod.mInstrumentList.Add(aInstrument)
                            blnEdit = False
                            'Update instrument combo box
                            cboOption2.Items.Add(aInstrument.Name)
                            cboOption2.Text = aInstrument.Name
                            lblStatus.Text = "Reviewed: " & aInstrument.Reviewed
                            cboOption3.Text = "Standard"
                            btnOption3Add.Text = "Add Standard"
                            btnOption3Del.Text = "Delete Standard"
                            btnOption3Add.Visible = True
                            btnOption3Del.Visible = True
                        If dgCompound.Rows.Count <= 1 Then
                            FillDataGrid()
                        End If
                        blnEdit = True
                        End If
                    Next
                End If
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                'New project
               
                aProject = New Project
                aProject.Name = InputBox("Please enter the Analysis name:", "eTrain 2.0")
                strSelProject = aProject.Name
                aProject.Reviewed = False
                aProject.ReviewedDate = CDate("1/1/1970")
                If aProject.Name <> "" Then
                    For Each aPermit In GlobalVariables.PermitList
                        If cboOption1.Text = aPermit.Name Then
                            aPermit.ProjectList.Add(aProject)
                            'Update cbooption2
                            cboOption2.Items.Add(aProject.Name)
                            cboOption2.Text = aProject.Name
                            lblStatus.Text = "Reviewed: " & aProject.Reviewed
                            btnOption3Add.Text = "Add Instrument"
                            btnOption3Del.Text = "Delete Instrument"
                            btnOption3Add.Visible = True
                            btnOption3Del.Visible = True
                        End If
                    Next
                End If
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                'New project
                
                aProject = New Project
                aProject.Name = InputBox("Please enter the Analysis name:", "eTrain 2.0")
                strSelProject = aProject.Name
                aProject.Reviewed = False
                aProject.ReviewedDate = CDate("1/1/1970")
                If aProject.Name <> "" Then
                    For Each aPermit In GlobalVariables.PermitList
                        If cboOption1.Text = aPermit.Name Then
                            aPermit.ProjectList.Add(aProject)
                            'Update cbooption2
                            cboOption2.Items.Add(aProject.Name)
                            cboOption2.Text = aProject.Name
                            lblStatus.Text = "Reviewed: " & aProject.Reviewed
                            btnOption3Add.Text = "Add Instrument"
                            btnOption3Del.Text = "Delete Instrument"
                            btnOption3Add.Visible = True
                            btnOption3Del.Visible = True
                        End If
                    Next
                End If
            End If
        End If

    End Sub

    Private Sub btnAddCompound_Click(sender As System.Object, e As System.EventArgs) Handles btnAddCompound.Click
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                'Create new row
                If cboOption3.Text = "Standard" Then
                    dgCompound.Rows.Add("", "", "", "", "", "", "", "", "", "")
                ElseIf cboOption3.Text = "Compound" Then
                    dgCompound.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                End If
                btnSave.Enabled = True
                blnEdit = True
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                'Create new row
                If cboOption3.Text = "Standard" Then
                    dgCompound.Rows.Add("", "", "", "", "", "", "", "", "", "")
                ElseIf cboOption3.Text = "Compound" Then
                    dgCompound.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                End If
                btnSave.Enabled = True
                blnEdit = True
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                dgCompound.Rows.Add("", "", "", "", "", "", "")
                blnEdit = True
                btnSave.Enabled = True
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                dgCompound.Rows.Add("", "", "", "", "", "", "")
                blnEdit = True
                btnSave.Enabled = True
            End If
        End If

    End Sub

    Private Sub btnDelCompound_Click(sender As System.Object, e As System.EventArgs) Handles btnDelCompound.Click
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                For Each r In dgCompound.SelectedRows
                    Try
                        dgCompound.Rows.RemoveAt(r.Index)
                    Catch ex As Exception
                        MsgBox("Error: No need to delete starter row, it will not show up in Method.", MsgBoxStyle.Exclamation)
                    End Try
                Next
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                For Each r In dgCompound.SelectedRows
                    Try
                        dgCompound.Rows.RemoveAt(r.Index)
                    Catch ex As Exception
                        MsgBox("Error: No need to delete starter row, it will not show up in Method.", MsgBoxStyle.Exclamation)
                    End Try
                Next
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                For Each r In dgCompound.SelectedRows
                    Try
                        dgCompound.Rows.RemoveAt(r.Index)
                    Catch ex As Exception
                        MsgBox("Error: No need to delete starter row, it will not show up in Permit.", MsgBoxStyle.Exclamation)
                    End Try
                Next
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                For Each r In dgCompound.SelectedRows
                    Try
                        dgCompound.Rows.RemoveAt(r.Index)
                    Catch ex As Exception
                        MsgBox("Error: No need to delete starter row, it will not show up in Permit.", MsgBoxStyle.Exclamation)
                    End Try
                Next
            End If
        End If

    End Sub

    Private Sub btnOption2Del_Click(sender As System.Object, e As System.EventArgs) Handles btnOption2Del.Click
        Dim aPermit As Permit
        Dim aProject As Project

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then

                'Delete instrument from new method
                For Each aMethod In GlobalVariables.MethodList
                    If aMethod.Name = cboOption1.Text Then
                        For Each aInstrument In aMethod.mInstrumentList
                            If aInstrument.Name = cboOption2.Text Then
                                aMethod.mInstrumentList.Remove(aInstrument)
                            End If
                        Next
                    End If
                Next

                'Reset form
                dgCompound.Rows.Clear()
                cboOption2.Items.Clear()
                lblStatus.Text = "Reviewed: "
                cboOption2.Text = ""
                For Each aMethod In GlobalVariables.MethodList
                    If cboOption1.Text = aMethod.Name Then
                        For Each aInstrument In aMethod.mInstrumentList
                            cboOption2.Items.Add(aInstrument.Name)
                        Next
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "HR" Then

                'Delete instrument from new method
                For Each aMethod In GlobalVariables.MethodList
                    If aMethod.Name = cboOption1.Text Then
                        For Each aInstrument In aMethod.mInstrumentList
                            If aInstrument.Name = cboOption2.Text Then
                                aMethod.mInstrumentList.Remove(aInstrument)
                            End If
                        Next
                    End If
                Next

                'Reset form
                dgCompound.Rows.Clear()
                cboOption2.Items.Clear()
                lblStatus.Text = "Reviewed: "
                cboOption2.Text = ""
                For Each aMethod In GlobalVariables.MethodList
                    If cboOption1.Text = aMethod.Name Then
                        For Each aInstrument In aMethod.mInstrumentList
                            cboOption2.Items.Add(aInstrument.Name)
                        Next
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If cboOption2.Text = aProject.Name Then
                                aPermit.ProjectList.Remove(aProject)
                            End If
                        Next
                    End If
                Next
                'Reset form
                dgCompound.Rows.Clear()
                cboOption2.Items.Clear()
                lblStatus.Text = "Reviewed: "
                cboOption2.Text = ""
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            cboOption2.Items.Add(aProject.Name)
                        Next
                    End If
                Next
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If cboOption2.Text = aProject.Name Then
                                aPermit.ProjectList.Remove(aProject)
                            End If
                        Next
                    End If
                Next
                'Reset form
                dgCompound.Rows.Clear()
                cboOption2.Items.Clear()
                lblStatus.Text = "Reviewed: "
                cboOption2.Text = ""
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            cboOption2.Items.Add(aProject.Name)
                        Next
                    End If
                Next
            End If
        End If

    End Sub

    Private Sub Me_FormClosing(sender As Object, e As FormClosingEventArgs) _
     Handles Me.FormClosing

        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            dgCompound.Rows.Clear()
            cboOption1.Items.Clear()
            cboOption2.Items.Clear()
            cboOption3.Items.Clear()
            txtOption4.Text = ""
            txtOption5.Text = ""
            Me.Hide()
        End If

    End Sub

    Private Sub EditMethods_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim dtAssoc13cs As New DataTable()
        Dim comboBoxColumn = New DataGridViewComboBoxColumn()

        blnEdit = False
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                If GlobalVariables.Associated13Cs.Count = 0 Then
                    GlobalVariables.Method.LoadAssoc13cFile()

                End If
                comboBoxColumn.Name = "Col12"
                comboBoxColumn.Width = 150
                For Each c In GlobalVariables.Associated13Cs
                    comboBoxColumn.Items.Add(c)
                Next
                comboBoxColumn.Visible = False
                dgCompound.Columns.Add(comboBoxColumn)
                strSelInstrument = ""
                cboOption3.Items.Add("Standard")
                cboOption3.Items.Add("Compound")
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                If GlobalVariables.Associated13Cs.Count = 0 Then
                    GlobalVariables.Method.LoadAssoc13cFile()

                End If
                comboBoxColumn.Name = "Col12"
                comboBoxColumn.Width = 150
                For Each c In GlobalVariables.Associated13Cs
                    comboBoxColumn.Items.Add(c)
                Next
                comboBoxColumn.Visible = False
                dgCompound.Columns.Add(comboBoxColumn)
                strSelInstrument = ""
                cboOption3.Items.Add("Standard")
                cboOption3.Items.Add("Compound")

            End If
        End If

    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        Dim aMethod As Method
        Dim aInstrument As mInstrument
        Dim aPermit As Permit
        Dim aProject As Project
        Dim blnSave As Boolean

        blnSave = False

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                'Clear compounds and load in new table
                SaveDetailsToObject()

                For Each aMethod In GlobalVariables.MethodList
                    If cboOption1.Text = aMethod.Name Then
                        'Check to see if aMethod, Instrment and some compounds exist first
                        If aMethod.ETEQ <> "" And aMethod.RptTolerance <> "" Then
                            If aMethod.mInstrumentList.Count > 0 Then
                                For Each aInstrument In aMethod.mInstrumentList
                                    If aInstrument.mStandardList.Count > 0 Then
                                        If aInstrument.mCompoundList.Count > 0 Then
                                            blnSave = True
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If GlobalVariables.Method.SaveMethod(aMethod) And blnSave Then
                            EditForm()
                            'blnEdit = False
                            'btnSave.Enabled = False
                            'cboOption1.Enabled = True
                            MsgBox("Saved successfully!", MsgBoxStyle.Information)
                            GlobalVariables.Method.LoadMethodNames()
                            Exit For
                        Else
                            MsgBox("Unable to Save. Method not filled out completely.", MsgBoxStyle.Information)
                        End If
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                'Clear compounds and load in new table
                SaveDetailsToObject()

                For Each aMethod In GlobalVariables.MethodList
                    If cboOption1.Text = aMethod.Name Then
                        'Check to see if aMethod, Instrment and some compounds exist first
                        If aMethod.ETEQ <> "" And aMethod.RptTolerance <> "" Then
                            If aMethod.mInstrumentList.Count > 0 Then
                                For Each aInstrument In aMethod.mInstrumentList
                                    If aInstrument.mStandardList.Count > 0 Then
                                        If aInstrument.mCompoundList.Count > 0 Then
                                            blnSave = True
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If GlobalVariables.Method.SaveMethod(aMethod) And blnSave Then
                            EditForm()
                            'blnEdit = False
                            'btnSave.Enabled = False
                            'cboOption1.Enabled = True
                            MsgBox("Saved successfully!", MsgBoxStyle.Information)
                            GlobalVariables.Method.LoadMethodNames()
                            Exit For
                        Else
                            MsgBox("Unable to Save. Method not filled out completely.", MsgBoxStyle.Information)
                        End If
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                'Clear compounds and load in new table
                SaveDetailsToObject()
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        'See if we can save it
                        If aPermit.ProjectList.Count > 0 Then
                            For Each aProject In aPermit.ProjectList
                                If aProject.mInstrumentList.Count > 0 Then
                                    For Each aInstrument In aProject.mInstrumentList
                                        If aInstrument.mCompoundList.Count > 0 Then
                                            blnSave = True
                                        End If
                                    Next
                                End If
                            Next
                        End If
                        If GlobalVariables.Permit.SavePermit(aPermit) And blnSave Then
                            EditForm()
                            'blnEdit = False
                            'btnSave.Enabled = False
                            'cboOption1.Enabled = True
                            MsgBox("Saved successfully!", MsgBoxStyle.Information)
                            GlobalVariables.Permit.LoadPermitNames()
                            Exit For
                        Else
                            MsgBox("Unable to Save. Source not filled out completely.", MsgBoxStyle.Information)
                        End If
                    End If
                Next
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                'Clear compounds and load in new table
                SaveDetailsToObject()
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        'See if we can save it
                        If aPermit.ProjectList.Count > 0 Then
                            For Each aProject In aPermit.ProjectList
                                If aProject.mInstrumentList.Count > 0 Then
                                    For Each aInstrument In aProject.mInstrumentList
                                        If aInstrument.mCompoundList.Count > 0 Then
                                            blnSave = True
                                        End If
                                    Next
                                End If
                            Next
                        End If
                        If GlobalVariables.Permit.SavePermit(aPermit) Then
                            EditForm()
                            'blnEdit = False
                            'btnSave.Enabled = False
                            'cboOption1.Enabled = True
                            MsgBox("Saved successfully!", MsgBoxStyle.Information)
                            GlobalVariables.Permit.LoadPermitNames()
                            Exit For
                        Else
                            MsgBox("Unable to Save. Source not filled out completely.", MsgBoxStyle.Information)
                        End If
                    End If
                Next
            End If
        End If


    End Sub

    Private Sub cboOption3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboOption3.SelectedIndexChanged
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                EditCurrentMethodToolStripMenuItem.Enabled = True
                'Check if editing or not
                If blnEdit = True Then
                    SaveDetailsToObject()

                    If cboOption3.Text = "Compound" Then
                        btnOption3Add.Text = "Add Compound"
                        btnOption3Del.Text = "Delete Compound"

                    ElseIf cboOption3.Text = "Standard" Then
                        btnOption3Add.Text = "Add Standard"
                        btnOption3Del.Text = "Delete Standard"

                    End If
                End If
                btnOption4.Visible = True
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                EditCurrentMethodToolStripMenuItem.Enabled = True
                'Check if editing or not
                If blnEdit = True Then
                    SaveDetailsToObject()

                    If cboOption3.Text = "Compound" Then
                        btnOption3Add.Text = "Add Compound"
                        btnOption3Del.Text = "Delete Compound"

                    ElseIf cboOption3.Text = "Standard" Then
                        btnOption3Add.Text = "Add Standard"
                        btnOption3Del.Text = "Delete Standard"

                    End If
                End If
                btnOption4.Visible = True
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                EditCurrentMethodToolStripMenuItem.Enabled = True
                If blnEdit = True And dgCompound.Rows.Count > 1 Then
                    SaveDetailsToObject()
                    btnAddCompound.Text = "Add Analyte"
                    btnDelCompound.Text = "Delete Analyte"
                End If
                strSelInstrument = cboOption3.Text
                strSelProject = cboOption2.Text

            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                EditCurrentMethodToolStripMenuItem.Enabled = True
                If blnEdit = True And dgCompound.Rows.Count > 1 Then
                    SaveDetailsToObject()
                    btnAddCompound.Text = "Add Analyte"
                    btnDelCompound.Text = "Delete Analyte"
                End If
                strSelInstrument = cboOption3.Text
                strSelProject = cboOption2.Text
            End If
        End If
        'Fills datagrid
        FillDataGrid()

    End Sub

    Private Sub CopyFromExistingMethodToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CopyFromExistingMethodToolStripMenuItem.Click
        'EditForm()
        CopyMethod.ShowDialog()

    End Sub

    Private Sub CreateNewMethodToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles CreateNewMethodToolStripMenuItem1.Click
        Dim curDate As Date
        Dim aMethod As Method
        Dim aPermit As Permit
        EditForm()
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                curDate = DateTime.Now
                aMethod = New Method
                'Get new method name
                aMethod.Name = InputBox("Please enter the new Method name:", "eTrain 2.0")
                aMethod.CreatedDate = curDate.Month & "/" & curDate.Day & "/" & curDate.Year
                If aMethod.Name <> "" Then
                    cboOption1.Enabled = False
                    cboOption1.Text = aMethod.Name
                    'Add new method to list
                    GlobalVariables.MethodList.Add(aMethod)
                    cboOption2.Items.Clear()
                    cboOption2.Text = ""
                    lblStatus.Text = "Reviewed: "
                    dgCompound.Rows.Clear()
                    
                    btnOption2Add.Visible = True
                    btnOption2Del.Visible = True
                End If
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                curDate = DateTime.Now
                btnOption3Add.Visible = False
                btnOption3Del.Visible = False
                aMethod = New Method
                'Get new method name
                aMethod.Name = InputBox("Please enter the new Method name:", "eTrain 2.0")
                aMethod.CreatedDate = curDate.Month & "/" & curDate.Day & "/" & curDate.Year
                If aMethod.Name <> "" Then
                    cboOption1.Enabled = False
                    cboOption1.Text = aMethod.Name
                    'Add new method to list
                    GlobalVariables.MethodList.Add(aMethod)
                    cboOption2.Items.Clear()
                    cboOption2.Text = ""
                    lblStatus.Text = "Reviewed: "
                    dgCompound.Rows.Clear()

                    btnOption2Add.Visible = True
                    btnOption2Del.Visible = True
                End If
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                curDate = DateTime.Now
                aPermit = New Permit
                'Get new permit name
                aPermit.Name = InputBox("Please enter the new Source name:", "eTrain 2.0")
                aPermit.CreatedDate = curDate.Month & "/" & curDate.Day & "/" & curDate.Year
                If aPermit.Name <> "" Then
                    cboOption1.Enabled = False
                    cboOption1.Text = aPermit.Name
                    'Add new method to list
                    GlobalVariables.PermitList.Add(aPermit)
                    cboOption2.Items.Clear()
                    cboOption2.Text = ""
                    cboOption3.Items.Clear()
                    cboOption3.Text = ""
                    lblStatus.Text = "Reviewed: "
                    dgCompound.Rows.Clear()
                    
                    btnOption2Add.Visible = True
                    btnOption2Del.Visible = True
                End If
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                curDate = DateTime.Now
                aPermit = New Permit
                'Get new permit name
                aPermit.Name = InputBox("Please enter the new Source name:", "eTrain 2.0")
                aPermit.CreatedDate = curDate.Month & "/" & curDate.Day & "/" & curDate.Year
                If aPermit.Name <> "" Then
                    cboOption1.Enabled = False
                    cboOption1.Text = aPermit.Name
                    'Add new method to list
                    GlobalVariables.PermitList.Add(aPermit)
                    cboOption2.Items.Clear()
                    cboOption2.Text = ""
                    cboOption3.Items.Clear()
                    cboOption3.Text = ""
                    lblStatus.Text = "Reviewed: "
                    dgCompound.Rows.Clear()
                   
                    btnOption2Add.Visible = True
                    btnOption2Del.Visible = True
                End If
            End If
        End If

    End Sub

    Private Sub UsingChemstationDataToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles UsingChemstationDataToolStripMenuItem.Click
        Dim aMethod As Method
        EditForm()
        NewMethodImport.ShowDialog()
        'Reload Form
        GlobalVariables.Method.LoadMethodNames()
        cboOption1.Items.Clear()
        For Each aMethod In GlobalVariables.MethodList
            cboOption1.Items.Add(aMethod.Name)
        Next

    End Sub

    Sub SaveDetailsToObject()
        Dim aCompound As mCompound
        Dim aStandard As mStandard
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aInstrument As mInstrument
        Dim aMethod As Method

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then

                For Each aMethod In GlobalVariables.MethodList
                    If aMethod.Name = cboOption1.Text Then
                        aMethod.ETEQ = txtOption4.Text
                        aMethod.RptTolerance = txtOption5.Text
                        For Each aInstrument In aMethod.mInstrumentList
                            If aInstrument.Name = strSelInstrument Then
                                'Save new information, look at columns to see what the type was
                                If dgCompound.Columns(1).HeaderText = "Avg Area" Then
                                    aInstrument.mStandardList.Clear()
                                    For Each r In dgCompound.Rows
                                        If r.Cells.Item(0).Value <> "" Then
                                            aStandard = New mStandard
                                            aStandard.Name = r.Cells.Item(0).Value
                                            aStandard.AvgArea = r.Cells.Item(1).Value
                                            aStandard.CalAmt = r.Cells.Item(2).Value
                                            aStandard.Conc = r.Cells.Item(3).Value
                                            aStandard.RecLowLim = r.Cells.Item(4).Value
                                            aStandard.RecUpLim = r.Cells.Item(5).Value
                                            aStandard.IonTarget = r.Cells.Item(6).Value
                                            aStandard.IonQual = r.Cells.Item(7).Value
                                            aStandard.AbundTarget = r.Cells.Item(8).Value
                                            aStandard.AbundQual = r.Cells.Item(9).Value
                                            aInstrument.mStandardList.Add(aStandard)
                                        End If
                                    Next
                                ElseIf Not dgCompound.Columns(1).HeaderText = "Avg Area" Then
                                    aInstrument.mCompoundList.Clear()
                                    For Each r In dgCompound.Rows
                                        If r.Cells.Item(0).Value <> "" Then
                                            aCompound = New mCompound
                                            aCompound.Name = r.Cells.Item(0).Value
                                            aCompound.RRF = r.Cells.Item(1).Value
                                            aCompound.RSD = r.Cells.Item(2).Value
                                            aCompound.MaxPeakArea = r.Cells.Item(3).Value
                                            aCompound.Conc = r.Cells.Item(4).Value
                                            aCompound.CS3Amt = r.Cells.Item(5).Value
                                            aCompound.TEF = r.Cells.Item(6).Value
                                            aCompound.Ion = r.Cells.Item(7).Value
                                            aCompound.Abundance = r.Cells.Item(8).Value
                                            aCompound.LCSLLim = r.Cells.Item(9).Value
                                            aCompound.LCSULim = r.Cells.Item(10).Value
                                            aCompound.Assoc13C = r.Cells.Item(11).Value
                                            aInstrument.mCompoundList.Add(aCompound)
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "HR" Then

                For Each aMethod In GlobalVariables.MethodList
                    If aMethod.Name = cboOption1.Text Then
                        aMethod.ETEQ = txtOption4.Text
                        aMethod.RptTolerance = txtOption5.Text
                        For Each aInstrument In aMethod.mInstrumentList
                            If aInstrument.Name = strSelInstrument Then
                                'Save new information, look at columns to see what the type was
                                If dgCompound.Columns(1).HeaderText = "Avg Area" Then
                                    aInstrument.mStandardList.Clear()
                                    For Each r In dgCompound.Rows
                                        If r.Cells.Item(0).Value <> "" Then
                                            aStandard = New mStandard
                                            aStandard.Name = r.Cells.Item(0).Value
                                            aStandard.AvgArea = r.Cells.Item(1).Value
                                            aStandard.CalAmt = r.Cells.Item(2).Value
                                            aStandard.Conc = r.Cells.Item(3).Value
                                            aStandard.RecLowLim = r.Cells.Item(4).Value
                                            aStandard.RecUpLim = r.Cells.Item(5).Value
                                            aStandard.IonTarget = r.Cells.Item(6).Value
                                            aStandard.IonQual = r.Cells.Item(7).Value
                                            aStandard.AbundTarget = r.Cells.Item(8).Value
                                            aStandard.AbundQual = r.Cells.Item(9).Value
                                            aInstrument.mStandardList.Add(aStandard)
                                        End If
                                    Next
                                ElseIf Not dgCompound.Columns(1).HeaderText = "Avg Area" Then
                                    aInstrument.mCompoundList.Clear()
                                    For Each r In dgCompound.Rows
                                        If r.Cells.Item(0).Value <> "" Then
                                            aCompound = New mCompound
                                            aCompound.Name = r.Cells.Item(0).Value
                                            aCompound.RRF = r.Cells.Item(1).Value
                                            aCompound.RSD = r.Cells.Item(2).Value
                                            aCompound.MaxPeakArea = r.Cells.Item(3).Value
                                            aCompound.Conc = r.Cells.Item(4).Value
                                            aCompound.CalAmt = r.Cells.Item(5).Value
                                            aCompound.TEF = r.Cells.Item(6).Value
                                            aCompound.Ion = r.Cells.Item(7).Value
                                            aCompound.Abundance = r.Cells.Item(8).Value
                                            aCompound.LCSLLim = r.Cells.Item(9).Value
                                            aCompound.LCSULim = r.Cells.Item(10).Value
                                            aCompound.Assoc13C = r.Cells.Item(11).Value
                                            aInstrument.mCompoundList.Add(aCompound)
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If strSelProject = aProject.Name Then
                                For Each aInstrument In aProject.mInstrumentList
                                    If strSelInstrument = aInstrument.Name Then
                                        aInstrument.mCompoundList.Clear()
                                        For Each r In dgCompound.Rows
                                            If r.Cells.Item(0).Value <> "" Then
                                                aCompound = New mCompound
                                                aCompound.Name = r.Cells.Item(0).Value
                                                aCompound.CAS = r.Cells.Item(1).Value
                                                aCompound.RL = r.Cells.Item(2).Value
                                                aCompound.MDL = r.Cells.Item(3).Value
                                                aCompound.PQL = r.Cells.Item(4).Value
                                                aCompound.RecLLim = r.Cells.Item(5).Value
                                                aCompound.RecULim = r.Cells.Item(6).Value
                                                aInstrument.mCompoundList.Add(aCompound)
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If strSelProject = aProject.Name Then

                                For Each aInstrument In aProject.mInstrumentList
                                    If strSelInstrument = aInstrument.Name Then
                                        aInstrument.mCompoundList.Clear()
                                        For Each r In dgCompound.Rows
                                            If r.Cells.Item(0).Value <> "" Then
                                                aCompound = New mCompound
                                                aCompound.Name = r.Cells.Item(0).Value
                                                aCompound.CAS = r.Cells.Item(1).Value
                                                aCompound.RL = r.Cells.Item(2).Value
                                                aCompound.MDL = r.Cells.Item(3).Value
                                                aCompound.PQL = r.Cells.Item(4).Value
                                                aCompound.RecLLim = r.Cells.Item(5).Value
                                                aCompound.RecULim = r.Cells.Item(6).Value
                                                aInstrument.mCompoundList.Add(aCompound)
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        End If

    End Sub

    Sub FillDataGrid()
        Dim aMethod As Method
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aInstrument As mInstrument
        Dim aStandard As mStandard
        Dim aCompound As mCompound
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                dgCompound.Rows.Clear()
                For Each aMethod In GlobalVariables.MethodList
                    If cboOption1.Text = aMethod.Name Then
                        txtOption4.Text = aMethod.ETEQ
                        txtOption5.Text = aMethod.RptTolerance
                        For Each aInstrument In aMethod.mInstrumentList
                            If strSelInstrument = aInstrument.Name Then
                                lblStatus.Text = "Reviewed: " & aInstrument.Reviewed
                                If cboOption3.Text = "Standard" Then
                                    dgCompound.Columns(0).HeaderText = "Name"
                                    dgCompound.Columns(1).HeaderText = "Avg Area"
                                    dgCompound.Columns(2).HeaderText = "Cal Amt (ng)"
                                    dgCompound.Columns(3).HeaderText = "Conc (ng/uL)"
                                    dgCompound.Columns(4).HeaderText = "Recovery Lower Limit"
                                    dgCompound.Columns(5).HeaderText = "Recovery Upper Limit"
                                    dgCompound.Columns(6).HeaderText = "Ion Target"
                                    dgCompound.Columns(7).HeaderText = "Ion Qual"
                                    dgCompound.Columns(8).HeaderText = "Abundance Target"
                                    dgCompound.Columns(9).HeaderText = "Abundance Qual"
                                    dgCompound.Columns(10).Visible = False
                                    dgCompound.Columns(11).Visible = False
                                    For Each aStandard In aInstrument.mStandardList
                                        dgCompound.Rows.Add(aStandard.Name, aStandard.AvgArea, aStandard.CalAmt, aStandard.Conc, aStandard.RecLowLim, aStandard.RecUpLim, aStandard.IonTarget, aStandard.IonQual, aStandard.AbundTarget, aStandard.AbundQual, "", "")
                                    Next
                                ElseIf cboOption3.Text = "Compound" Then
                                    dgCompound.Columns(0).HeaderText = "Name"
                                    dgCompound.Columns(1).HeaderText = "RRF"
                                    dgCompound.Columns(2).HeaderText = "% RSD"
                                    dgCompound.Columns(3).HeaderText = "Max Peak Area"
                                    dgCompound.Columns(4).HeaderText = "Conc (ng/uL)"
                                    dgCompound.Columns(5).HeaderText = "CS3 Amt (ng)"
                                    dgCompound.Columns(6).HeaderText = "TEF"
                                    dgCompound.Columns(7).HeaderText = "Ion"
                                    dgCompound.Columns(8).HeaderText = "Abundance"
                                    dgCompound.Columns(9).HeaderText = "LCS Lower Limit"
                                    dgCompound.Columns(10).HeaderText = "LCS Upper Limit"
                                    dgCompound.Columns(11).HeaderText = "Associated 13C"
                                    'Change last column to drop box
                                    dgCompound.Columns(10).Visible = True
                                    dgCompound.Columns(11).Visible = True
                                    For Each aCompound In aInstrument.mCompoundList
                                        dgCompound.Rows.Add(aCompound.Name, aCompound.RRF, aCompound.RSD, aCompound.MaxPeakArea, aCompound.Conc, aCompound.CS3Amt, aCompound.TEF, aCompound.Ion, aCompound.Abundance, aCompound.LCSLLim, aCompound.LCSULim, aCompound.Assoc13C)
                                    Next
                                End If
                            End If
                        Next
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                dgCompound.Rows.Clear()
                For Each aMethod In GlobalVariables.MethodList
                    If cboOption1.Text = aMethod.Name Then
                        txtOption4.Text = aMethod.ETEQ
                        txtOption5.Text = aMethod.RptTolerance
                        For Each aInstrument In aMethod.mInstrumentList
                            If strSelInstrument = aInstrument.Name Then
                                lblStatus.Text = "Reviewed: " & aInstrument.Reviewed
                                If cboOption3.Text = "Standard" Then
                                    dgCompound.Columns(0).HeaderText = "Name"
                                    dgCompound.Columns(1).HeaderText = "Avg Area"
                                    dgCompound.Columns(2).HeaderText = "Cal Amt (ng)"
                                    dgCompound.Columns(3).HeaderText = "Conc (ng/uL)"
                                    dgCompound.Columns(4).HeaderText = "Recovery Lower Limit"
                                    dgCompound.Columns(5).HeaderText = "Recovery Upper Limit"
                                    dgCompound.Columns(6).HeaderText = "Ion Target"
                                    dgCompound.Columns(7).HeaderText = "Ion Qual"
                                    dgCompound.Columns(8).HeaderText = "Abundance Target"
                                    dgCompound.Columns(9).HeaderText = "Abundance Qual"
                                    dgCompound.Columns(10).Visible = False
                                    dgCompound.Columns(11).Visible = False
                                    For Each aStandard In aInstrument.mStandardList
                                        dgCompound.Rows.Add(aStandard.Name, aStandard.AvgArea, aStandard.CalAmt, aStandard.Conc, aStandard.RecLowLim, aStandard.RecUpLim, aStandard.IonTarget, aStandard.IonQual, aStandard.AbundTarget, aStandard.AbundQual, "", "")
                                    Next
                                ElseIf cboOption3.Text = "Compound" Then
                                    dgCompound.Columns(0).HeaderText = "Name"
                                    dgCompound.Columns(1).HeaderText = "RRF"
                                    dgCompound.Columns(2).HeaderText = "% RSD"
                                    dgCompound.Columns(3).HeaderText = "Max Peak Area"
                                    dgCompound.Columns(4).HeaderText = "Conc (ng/uL)"
                                    dgCompound.Columns(5).HeaderText = "Cal Amt (ng)"
                                    dgCompound.Columns(6).HeaderText = "TEF"
                                    dgCompound.Columns(7).HeaderText = "Ion"
                                    dgCompound.Columns(8).HeaderText = "Abundance"
                                    dgCompound.Columns(9).HeaderText = "LCS Lower Limit"
                                    dgCompound.Columns(10).HeaderText = "LCS Upper Limit"
                                    dgCompound.Columns(11).HeaderText = "Associated 13C"
                                    'Change last column to drop box
                                    dgCompound.Columns(10).Visible = True
                                    dgCompound.Columns(11).Visible = True
                                    For Each aCompound In aInstrument.mCompoundList
                                        dgCompound.Rows.Add(aCompound.Name, aCompound.RRF, aCompound.RSD, aCompound.MaxPeakArea, aCompound.Conc, aCompound.CalAmt, aCompound.TEF, aCompound.Ion, aCompound.Abundance, aCompound.LCSLLim, aCompound.LCSULim, aCompound.Assoc13C)
                                    Next
                                End If
                            End If
                        Next
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                dgCompound.Rows.Clear()
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If cboOption2.Text = aProject.Name Then
                                For Each aInstrument In aProject.mInstrumentList
                                    If cboOption3.Text = aInstrument.Name Then
                                        dgCompound.Columns(0).HeaderText = "Name"
                                        dgCompound.Columns(1).HeaderText = "CAS"
                                        dgCompound.Columns(2).HeaderText = "RL"
                                        dgCompound.Columns(3).HeaderText = "MDL"
                                        dgCompound.Columns(4).HeaderText = "PQL"
                                        dgCompound.Columns(5).HeaderText = "Recovery Lower Limit"
                                        dgCompound.Columns(6).HeaderText = "Recovery Upper Limit"
                                        dgCompound.Columns(7).Visible = False
                                        dgCompound.Columns(8).Visible = False
                                        dgCompound.Columns(9).Visible = False
                                        dgCompound.Columns(10).Visible = False
                                        For Each aCompound In aInstrument.mCompoundList
                                            dgCompound.Rows.Add(aCompound.Name, aCompound.CAS, aCompound.RL, aCompound.MDL, aCompound.PQL, aCompound.RecLLim, aCompound.RecULim, "", "", "", "")
                                        Next
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                dgCompound.Rows.Clear()
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If cboOption2.Text = aProject.Name Then
                            
                                For Each aInstrument In aProject.mInstrumentList
                                    If cboOption3.Text = aInstrument.Name Then
                                        dgCompound.Columns(0).HeaderText = "Name"
                                        dgCompound.Columns(1).HeaderText = "CAS"
                                        dgCompound.Columns(2).HeaderText = "RL"
                                        dgCompound.Columns(3).HeaderText = "MDL"
                                        dgCompound.Columns(4).HeaderText = "PQL"
                                        dgCompound.Columns(5).HeaderText = "Recovery Lower Limit"
                                        dgCompound.Columns(6).HeaderText = "Recovery Upper Limit"
                                        dgCompound.Columns(7).Visible = False
                                        dgCompound.Columns(8).Visible = False
                                        dgCompound.Columns(9).Visible = False
                                        dgCompound.Columns(10).Visible = False
                                        For Each aCompound In aInstrument.mCompoundList
                                            dgCompound.Rows.Add(aCompound.Name, aCompound.CAS, aCompound.RL, aCompound.MDL, aCompound.PQL, aCompound.RecLLim, aCompound.RecULim, "", "", "", "")
                                        Next
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub btnOption3Add_Click(sender As System.Object, e As System.EventArgs) Handles btnOption3Add.Click
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aInstrument As mInstrument

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                'Create new row
                If cboOption3.Text = "Standard" Then
                    dgCompound.Rows.Add("", "", "", "", "", "", "", "", "", "")
                ElseIf cboOption3.Text = "Compound" Then
                    dgCompound.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                End If
                btnSave.Enabled = True
                blnEdit = True
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                'Create new row
                If cboOption3.Text = "Standard" Then
                    dgCompound.Rows.Add("", "", "", "", "", "", "", "", "", "")
                ElseIf cboOption3.Text = "Compound" Then
                    dgCompound.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                End If
                btnSave.Enabled = True
                blnEdit = True


            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                aInstrument = New mInstrument
                aInstrument.Name = InputBox("Please enter the Instrument name:", "eTrain 2.0")
                strSelInstrument = aInstrument.Name
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If cboOption2.Text = aProject.Name Then
                                aProject.mInstrumentList.Add(aInstrument)
                                'Update cbooption2
                                cboOption3.Items.Add(aInstrument.Name)
                                cboOption3.Text = aInstrument.Name
                                btnAddCompound.Text = "Add Analyte"
                                btnDelCompound.Text = "Delete Analyte"
                                btnAddCompound.Visible = True
                                btnDelCompound.Visible = True
                            End If
                        Next
                    End If
                Next
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                aInstrument = New mInstrument
                aInstrument.Name = InputBox("Please enter the Instrument name:", "eTrain 2.0")
                strSelInstrument = aInstrument.Name
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If cboOption2.Text = aProject.Name Then
                                aProject.mInstrumentList.Add(aInstrument)
                                'Update cbooption2
                                cboOption3.Items.Add(aInstrument.Name)
                                cboOption3.Text = aInstrument.Name
                                btnAddCompound.Text = "Add Analyte"
                                btnDelCompound.Text = "Delete Analyte"
                                btnAddCompound.Visible = True
                                btnDelCompound.Visible = True
                            End If
                        Next
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub btnOption3Del_Click(sender As System.Object, e As System.EventArgs) Handles btnOption3Del.Click
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aInstrument As mInstrument

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If cboOption2.Text = aProject.Name Then
                                For Each aInstrument In aProject.mInstrumentList
                                    If cboOption3.Text = aInstrument.Name Then
                                        aProject.mInstrumentList.Remove(aInstrument)
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
                'Reset form
                dgCompound.Rows.Clear()
                cboOption3.Items.Clear()
                lblStatus.Text = "Reviewed: "
                cboOption3.Text = ""
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If cboOption2.Text = aProject.Name Then
                                For Each aInstrument In aProject.mInstrumentList
                                    cboOption3.Items.Add(aInstrument)
                                Next
                            End If
                        Next
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "FAST" Then
                For Each r In dgCompound.SelectedRows
                    Try
                        dgCompound.Rows.RemoveAt(r.Index)
                    Catch ex As Exception
                        MsgBox("Error: No need to delete starter row, it will not show up in Method.", MsgBoxStyle.Exclamation)
                    End Try
                Next
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                For Each r In dgCompound.SelectedRows
                    Try
                        dgCompound.Rows.RemoveAt(r.Index)
                    Catch ex As Exception
                        MsgBox("Error: No need to delete starter row, it will not show up in Method.", MsgBoxStyle.Exclamation)
                    End Try
                Next
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If cboOption2.Text = aProject.Name Then
                                For Each aInstrument In aProject.mInstrumentList
                                    If cboOption3.Text = aInstrument.Name Then
                                        aProject.mInstrumentList.Remove(aInstrument)
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
                'Reset form
                dgCompound.Rows.Clear()
                cboOption3.Items.Clear()
                lblStatus.Text = "Reviewed: "
                cboOption3.Text = ""
                For Each aPermit In GlobalVariables.PermitList
                    If cboOption1.Text = aPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If cboOption2.Text = aProject.Name Then
                                For Each aInstrument In aProject.mInstrumentList
                                    cboOption3.Items.Add(aInstrument)
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        End If

    End Sub

    Private Sub btnOption4_Click(sender As System.Object, e As System.EventArgs) Handles btnOption4.Click
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                RefBookEdit.ShowDialog()
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                RefBookEdit.ShowDialog()
            End If
        End If

    End Sub
End Class