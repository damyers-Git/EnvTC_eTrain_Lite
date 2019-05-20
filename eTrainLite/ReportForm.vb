Imports System.IO
Imports Syncfusion.XlsIO

Public Class ReportForm

    Private Sub btnBrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnReportSaveBrowse.Click
        Dim fDialog As New FolderBrowserDialog
        If fDialog.ShowDialog() = DialogResult.OK Then
            txtReportSaveLoc.Text = fDialog.SelectedPath
        End If

    End Sub

    'Generate Report
    Private Sub btnGen_Click(sender As System.Object, e As System.EventArgs) Handles btnGen.Click
        Dim exEngine As New ExcelEngine
        Dim exApp As IApplication
        Dim aSample As Sample
        Dim blnFlg As Boolean

        blnFlg = False
        GlobalVariables.ReportSamList.Clear()
        'Check for entry
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                If cboType.Text <> "" Then
                    If txt1.Text <> "" And File.Exists(txt1.Text) Then
                        If txtReportSaveLoc.Text <> "" And Directory.Exists(txtReportSaveLoc.Text) Then
                            If txtRName.Text <> "" Then
                                'Set report values
                                For Each aSample In GlobalVariables.SampleList
                                    GlobalVariables.ReportSamList.Add(aSample)
                                Next
                                GlobalVariables.Report.SavLoc = txtReportSaveLoc.Text
                                GlobalVariables.Report.RName = txtRName.Text
                                'Do Calculations
                                If GlobalVariables.NeedsCalculation Then
                                    If Not GlobalVariables.Calculations.MidlandFAST(txt1.Text) Then
                                        Exit Sub
                                    End If
                                End If
                                'Types of Reports
                                Select Case cboType.Text
                                    Case "Sample Report"
                                        GlobalVariables.Report.MidlandFASTSampleReport(cbo1.Text)
                                        txtRName.Text = ""
                                    Case "CS3 Check Report"
                                        GlobalVariables.Report.MidlandFASTCS3Report()
                                        txtRName.Text = ""
                                    Case "LCS Check Report"
                                        GlobalVariables.Report.MidlandFASTLCSReport()
                                        txtRName.Text = ""
                                    Case "Final Data"
                                        If MsgBox("Do you want to create supporting sample reports as well?", MsgBoxStyle.YesNo, "eTrain 2.0") = vbYes Then
                                            GlobalVariables.Report.MidlandFASTSampleReport(cbo1.Text)
                                        End If
                                        GlobalVariables.Report.MidlandFASTFinalDataReport(txt1.Text)
                                        txtRName.Text = ""
                                End Select
                                MsgBox("Report Generation Completed!", MsgBoxStyle.Information)
                            Else
                                MsgBox("Please enter a Valid Report Name before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                                txtRName.Focus()
                            End If
                        Else
                            MsgBox("Please enter a Valid Directory before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                            txtReportSaveLoc.Focus()
                        End If
                    Else
                        MsgBox("Please enter a Valid Path to a SIS 2.0 Spreadsheet before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                        txt1.Focus()
                    End If
                Else
                    MsgBox("Please select a Valid Report Type before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                    cboType.Focus()
                End If
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                If cboType.Text <> "" Then
                    If txt1.Text <> "" And File.Exists(txt1.Text) Then
                        If txtReportSaveLoc.Text <> "" And Directory.Exists(txtReportSaveLoc.Text) Then
                            If txtRName.Text <> "" Then
                                'Set report values
                                GlobalVariables.Report.SavLoc = txtReportSaveLoc.Text
                                GlobalVariables.Report.RName = txtRName.Text
                                'Do Calculations
                                If GlobalVariables.NeedsCalculation Then
                                    If Not GlobalVariables.Calculations.MidlandHR(txt1.Text) Then
                                        Exit Sub
                                    End If
                                End If
                                'Types of Reports
                                Select Case cboType.Text
                                    Case "Sample Report"
                                        ' GlobalVariables.Report.MidlandFASTSampleReport()
                                        txtRName.Text = ""
                                    Case "CS3 Check Report"
                                        GlobalVariables.Report.MidlandFASTCS3Report()
                                        txtRName.Text = ""
                                    Case "LCS Check Report"
                                        GlobalVariables.Report.MidlandFASTLCSReport()
                                        txtRName.Text = ""
                                    Case "Final Data"
                                        If MsgBox("Do you want to create supporting sample reports as well?", MsgBoxStyle.YesNo, "eTrain 2.0") = vbYes Then
                                            '    GlobalVariables.Report.MidlandFASTSampleReport()
                                        End If
                                        GlobalVariables.Report.MidlandFASTFinalDataReport(txt1.Text)
                                        txtRName.Text = ""
                                End Select
                                MsgBox("Report Generation Completed!", MsgBoxStyle.Information)
                            Else
                                MsgBox("Please enter a Valid Report Name before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                                txtRName.Focus()
                            End If
                        Else
                            MsgBox("Please enter a Valid Directory before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                            txtReportSaveLoc.Focus()
                        End If
                    Else
                        MsgBox("Please enter a Valid Path to a SIS 2.0 Spreadsheet before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                        txt1.Focus()
                    End If
                Else
                    MsgBox("Please select a Valid Report Type before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                    cboType.Focus()
                End If
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                If cboType.Text <> "" Or cboType.Text <> "Custom Report" Then
                    If txtReportSaveLoc.Text <> "" And Directory.Exists(txtReportSaveLoc.Text) Then
                        If txtRName.Text <> "" Then
                            'Set report values
                            GlobalVariables.Report.SavLoc = txtReportSaveLoc.Text
                            GlobalVariables.Report.RName = txtRName.Text
                            GlobalVariables.selLimit = cbo5.Text
                            'Move samples to report sample list for manipulation
                            For Each aSample In GlobalVariables.SampleList
                                GlobalVariables.ReportSamList.Add(aSample)
                            Next
                            SampleEdit.ShowDialog()
                            If GlobalVariables.ContinueReport = True Then
                                CompSel.ShowDialog()
                                If GlobalVariables.ContinueReport = True Then

                                    If MsgBox("Do you want to sort components alphabetically?", MsgBoxStyle.YesNo) = vbYes Then
                                        aSample = Nothing
                                        'Order components
                                        For Each aSample In GlobalVariables.ReportSamList
                                            aSample.SortStandards()
                                            aSample.SortSSurrogates()
                                            aSample.SortCompounds()
                                        Next
                                    Else
                                        aSample = Nothing
                                        If GlobalVariables.Import.ElutionImport("\\mdrnd\AS-Global\Special_Access\EAC\Data\eTrain\DataFiles\Midland\Chrom\TemplateInfo\ElutionDictionary.txt") Then
                                            For Each aSample In GlobalVariables.ReportSamList
                                                aSample.ESortStandards()
                                                aSample.ESortSurrogates()
                                                aSample.ESortCompounds()
                                            Next
                                        End If
                                    End If

                                    'Do Calculations
                                    If GlobalVariables.NeedsCalculation Then
                                        If Not GlobalVariables.Calculations.MidlandChrom(GlobalVariables.selLimit, cbo6.Text, txt1.Text, False) Then
                                            Exit Sub
                                        End If
                                    End If
                                    Try
                                        'Reset Reported Values
                                        For Each aSample In GlobalVariables.ReportSamList
                                            aSample.Reported = False
                                        Next
                                        exApp = exEngine.Excel
                                        GlobalVariables.workbook = exApp.Workbooks.Create(1)
                                        'Types of Reports
                                        Select Case cboType.Text
                                            Case "Summary Report"
                                                If GlobalVariables.Report.MidlandChromSummaryReport(GlobalVariables.selLimit) Then
                                                    txtRName.Text = ""
                                                Else
                                                    MsgBox("Report not created!", MsgBoxStyle.Information)
                                                    GlobalVariables.workbook.Close()
                                                    exEngine.Dispose()
                                                    Exit Sub
                                                End If
                                            Case "Spike Recovery Report"
                                                For Each aSample In GlobalVariables.ReportSamList
                                                    If aSample.Type = "MS" Then
                                                        blnFlg = True
                                                    End If
                                                Next
                                                If blnFlg Then
                                                    If GlobalVariables.Report.MidlandChromMSReport(GlobalVariables.selLimit) Then
                                                        txtRName.Text = ""
                                                    Else
                                                        MsgBox("Report not created!", MsgBoxStyle.Information)
                                                        GlobalVariables.workbook.Close()
                                                        exEngine.Dispose()
                                                        Exit Sub
                                                    End If
                                                Else
                                                    MsgBox("Report not created, sample of required type not found to generate report!", MsgBoxStyle.Information)
                                                    GlobalVariables.workbook.Close()
                                                    exEngine.Dispose()
                                                    Exit Sub
                                                End If

                                            Case "LCS Report"
                                                For Each aSample In GlobalVariables.ReportSamList
                                                    If aSample.Type = "LCS" Then
                                                        blnFlg = True
                                                    End If
                                                Next
                                                If blnFlg Then
                                                    If GlobalVariables.Report.MidlandChromLCSReport() Then
                                                        txtRName.Text = ""
                                                    Else
                                                        MsgBox("Report not created!", MsgBoxStyle.Information)
                                                        GlobalVariables.workbook.Close()
                                                        exEngine.Dispose()
                                                        Exit Sub
                                                    End If
                                                Else
                                                    MsgBox("Report not created, sample of required type not found to generate report!", MsgBoxStyle.Information)
                                                    GlobalVariables.workbook.Close()
                                                    exEngine.Dispose()
                                                    Exit Sub
                                                End If
                                            Case "Duplicate Report"
                                                For Each aSample In GlobalVariables.ReportSamList
                                                    If aSample.Type = "DUP" Then
                                                        blnFlg = True
                                                    End If
                                                Next
                                                If blnFlg Then
                                                    If GlobalVariables.Report.MidlandChromDUPReport(GlobalVariables.selLimit) Then
                                                        txtRName.Text = ""
                                                    Else
                                                        MsgBox("Report not created!", MsgBoxStyle.Information)
                                                        GlobalVariables.workbook.Close()
                                                        exEngine.Dispose()
                                                        Exit Sub
                                                    End If
                                                Else
                                                    MsgBox("Report not created, sample of required type not found to generate report!", MsgBoxStyle.Information)
                                                    GlobalVariables.workbook.Close()
                                                    exEngine.Dispose()
                                                    Exit Sub
                                                End If

                                            Case "Method Blank Report"
                                                For Each aSample In GlobalVariables.ReportSamList
                                                    If aSample.Type = "MB" Then
                                                        blnFlg = True
                                                    End If
                                                Next
                                                'If blnFlg Then
                                                '    If GlobalVariables.Report.MidlandChromMBReport(GlobalVariables.selLimitPath) Then
                                                '        txtRName.Text = ""
                                                '    Else
                                                '        MsgBox("Report not created!", MsgBoxStyle.Information)
                                                '        GlobalVariables.workbook.Close()
                                                '        exEngine.Dispose()
                                                '        Exit Sub
                                                '    End If
                                                'Else
                                                '    MsgBox("Report not created, sample of required type not found to generate report!", MsgBoxStyle.Information)
                                                '    GlobalVariables.workbook.Close()
                                                '    exEngine.Dispose()
                                                '    Exit Sub
                                                'End If

                                            Case "Custom Report"
                                                GlobalVariables.CustomReportError = False
                                                CustomReport.ShowDialog()
                                        End Select
                                        'Start setting up save
                                        If GlobalVariables.CustomReportError Then
                                            GlobalVariables.workbook.Worksheets.Remove("Sheet1")
                                            GlobalVariables.workbook.Close()
                                            exEngine.Dispose()
                                            Exit Sub
                                        Else
                                            GlobalVariables.workbook.Version = ExcelVersion.Excel2010
                                            GlobalVariables.Report.RName = "\" & GlobalVariables.Report.RName & ".xlsx"
                                            GlobalVariables.workbook.Worksheets.Remove("Sheet1")
                                            GlobalVariables.workbook.SaveAs(GlobalVariables.Report.SavLoc & GlobalVariables.Report.RName)
                                            GlobalVariables.workbook.Close()
                                            MsgBox("Report Generation Completed!", MsgBoxStyle.Information)
                                            exEngine.Dispose()
                                            blnFlg = False
                                            txtRName.Text = ""
                                        End If
                                        UnitConversion.Close()
                                        SpikeInfo.Close()
                                        CustomReport.Close()
                                        GlobalVariables.ReportSamList.Clear()
                                    Catch ex As Exception
                                        MsgBox("Error generating report!" & vbCrLf & _
                                                    "Sub Procedure: btnGen_Click_ReportForm - Midland()" & vbCrLf & _
                                                    "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                                    End Try

                                End If
                            End If
                        Else
                            MsgBox("Please enter a Valid Report Name before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                            txtRName.Focus()
                        End If
                    Else
                        MsgBox("Please enter a Valid Directory before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                        txtReportSaveLoc.Focus()
                    End If
                Else
                    MsgBox("Please select a Valid Report Type before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                    cboType.Focus()
                End If
                End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                If cboType.Text <> "" Or cboType.Text <> "Custom Report" Then
                    If txt1.Text <> "" Then
                        If txtReportSaveLoc.Text <> "" And Directory.Exists(txtReportSaveLoc.Text) Then
                            If txtRName.Text <> "" Then
                                'Set report values
                                GlobalVariables.Report.SavLoc = txtReportSaveLoc.Text
                                GlobalVariables.Report.RName = txtRName.Text
                                GlobalVariables.selLimit = cbo5.Text
                                GlobalVariables.selLimitPath = txt1.Text
                                'Move samples to report sample list for manipulation
                                For Each aSample In GlobalVariables.SampleList
                                    GlobalVariables.ReportSamList.Add(aSample)
                                Next
                                SampleEdit.ShowDialog()
                                If GlobalVariables.ContinueReport = True Then
                                    CompSel.ShowDialog()
                                    If GlobalVariables.ContinueReport = True Then
                                        'Do Calculations
                                        If GlobalVariables.NeedsCalculation Then
                                            If Not GlobalVariables.Calculations.FreeportChrom(GlobalVariables.selLimit, GlobalVariables.selLimitPath, False) Then
                                                Exit Sub
                                            End If
                                        End If
                                        Try
                                            'Reset Reported Values
                                            For Each aSample In GlobalVariables.ReportSamList
                                                aSample.Reported = False
                                            Next
                                            exApp = exEngine.Excel
                                            GlobalVariables.workbook = exApp.Workbooks.Create(1)
                                            'Types of Reports
                                            Select Case cboType.Text
                                                Case "Summary Report"
                                                    If GlobalVariables.Report.FreeportChromSummaryReport(GlobalVariables.selLimit) Then
                                                        txtRName.Text = ""
                                                    Else
                                                        MsgBox("Report not created!", MsgBoxStyle.Information)
                                                        GlobalVariables.workbook.Close()
                                                        exEngine.Dispose()
                                                        Exit Sub
                                                    End If
                                                Case "Spike Recovery Report"
                                                    For Each aSample In GlobalVariables.ReportSamList
                                                        If aSample.Type = "MS" Then
                                                            blnFlg = True
                                                        End If
                                                    Next
                                                    If blnFlg Then
                                                        If GlobalVariables.Report.FreeportChromMSReport(GlobalVariables.selLimit) Then
                                                            txtRName.Text = ""
                                                        Else
                                                            MsgBox("Report not created!", MsgBoxStyle.Information)
                                                            GlobalVariables.workbook.Close()
                                                            exEngine.Dispose()
                                                            Exit Sub
                                                        End If
                                                    Else
                                                        MsgBox("Report not created, sample of required type not found to generate report!", MsgBoxStyle.Information)
                                                        GlobalVariables.workbook.Close()
                                                        exEngine.Dispose()
                                                        Exit Sub
                                                    End If

                                                Case "LCS Report"
                                                    For Each aSample In GlobalVariables.ReportSamList
                                                        If aSample.Type = "LCS" Then
                                                            blnFlg = True
                                                        End If
                                                    Next
                                                    If blnFlg Then
                                                        If GlobalVariables.Report.FreeportChromLCSReport() Then
                                                            txtRName.Text = ""
                                                        Else
                                                            MsgBox("Report not created!", MsgBoxStyle.Information)
                                                            GlobalVariables.workbook.Close()
                                                            exEngine.Dispose()
                                                            Exit Sub
                                                        End If
                                                    Else
                                                        MsgBox("Report not created, sample of required type not found to generate report!", MsgBoxStyle.Information)
                                                        GlobalVariables.workbook.Close()
                                                        exEngine.Dispose()
                                                        Exit Sub
                                                    End If

                                                Case "ICV Report"
                                                    For Each aSample In GlobalVariables.ReportSamList
                                                        If aSample.Type = "ICV" Then
                                                            blnFlg = True
                                                        End If
                                                    Next
                                                    If blnFlg Then
                                                        If GlobalVariables.Report.FreeportChromICVReport() Then
                                                            txtRName.Text = ""
                                                        Else
                                                            MsgBox("Report not created!", MsgBoxStyle.Information)
                                                            GlobalVariables.workbook.Close()
                                                            exEngine.Dispose()
                                                            Exit Sub
                                                        End If
                                                    Else
                                                        MsgBox("Report not created, sample of required type not found to generate report!", MsgBoxStyle.Information)
                                                        GlobalVariables.workbook.Close()
                                                        exEngine.Dispose()
                                                        Exit Sub
                                                    End If

                                                Case "CVS Report"
                                                    For Each aSample In GlobalVariables.ReportSamList
                                                        If aSample.Type = "CVS" Then
                                                            blnFlg = True
                                                        End If
                                                    Next
                                                    If blnFlg Then
                                                        If GlobalVariables.Report.FreeportChromCVSReport() Then
                                                            txtRName.Text = ""
                                                        Else
                                                            MsgBox("Report not created!", MsgBoxStyle.Information)
                                                            GlobalVariables.workbook.Close()
                                                            exEngine.Dispose()
                                                            Exit Sub
                                                        End If
                                                    Else
                                                        MsgBox("Report not created, sample of required type not found to generate report!", MsgBoxStyle.Information)
                                                        GlobalVariables.workbook.Close()
                                                        exEngine.Dispose()
                                                        Exit Sub
                                                    End If

                                                Case "Duplicate Report"
                                                    For Each aSample In GlobalVariables.ReportSamList
                                                        If aSample.Type = "DUP" Then
                                                            blnFlg = True
                                                        End If
                                                    Next
                                                    If blnFlg Then
                                                        If GlobalVariables.Report.FreeportChromDUPReport(GlobalVariables.selLimit) Then
                                                            txtRName.Text = ""
                                                        Else
                                                            MsgBox("Report not created!", MsgBoxStyle.Information)
                                                            GlobalVariables.workbook.Close()
                                                            exEngine.Dispose()
                                                            Exit Sub
                                                        End If
                                                    Else
                                                        MsgBox("Report not created, sample of required type not found to generate report!", MsgBoxStyle.Information)
                                                        GlobalVariables.workbook.Close()
                                                        exEngine.Dispose()
                                                        Exit Sub
                                                    End If

                                                Case "Method Blank Report"
                                                    For Each aSample In GlobalVariables.ReportSamList
                                                        If aSample.Type = "MB" Then
                                                            blnFlg = True
                                                        End If
                                                    Next
                                                    If blnFlg Then
                                                        If GlobalVariables.Report.FreeportChromMBReport(GlobalVariables.selLimitPath, GlobalVariables.selLimit) Then
                                                            txtRName.Text = ""
                                                        Else
                                                            MsgBox("Report not created!", MsgBoxStyle.Information)
                                                            GlobalVariables.workbook.Close()
                                                            exEngine.Dispose()
                                                            Exit Sub
                                                        End If
                                                    Else
                                                        MsgBox("Report not created, sample of required type not found to generate report!", MsgBoxStyle.Information)
                                                        GlobalVariables.workbook.Close()
                                                        exEngine.Dispose()
                                                        Exit Sub
                                                    End If

                                                Case "Custom Report"
                                                    GlobalVariables.CustomReportError = False
                                                    CustomReport.ShowDialog()
                                            End Select
                                            'Start setting up save
                                            If GlobalVariables.CustomReportError Then
                                                GlobalVariables.workbook.Worksheets.Remove("Sheet1")
                                                GlobalVariables.workbook.Close()
                                                exEngine.Dispose()
                                                Exit Sub
                                            Else
                                                GlobalVariables.workbook.Version = ExcelVersion.Excel2010
                                                GlobalVariables.Report.RName = "\" & GlobalVariables.Report.RName & ".xlsx"
                                                GlobalVariables.workbook.Worksheets.Remove("Sheet1")
                                                GlobalVariables.workbook.SaveAs(GlobalVariables.Report.SavLoc & GlobalVariables.Report.RName)
                                                GlobalVariables.workbook.Close()
                                                MsgBox("Report Generation Completed!", MsgBoxStyle.Information)
                                                exEngine.Dispose()
                                                blnFlg = False
                                                txtRName.Text = ""
                                            End If
                                            UnitConversion.Close()
                                            SpikeInfo.Close()
                                            CustomReport.Close()
                                            GlobalVariables.ReportSamList.Clear()
                                        Catch ex As Exception
                                            MsgBox("Error generating report!" & vbCrLf & _
                                                      "Sub Procedure: btnGen_Click_ReportForm - Freeport()" & vbCrLf & _
                                                      "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                                        End Try
                                    End If
                                End If
                            Else
                                MsgBox("Please enter a Valid Report Name before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                                txtRName.Focus()
                            End If
                        Else
                            MsgBox("Please enter a Valid Directory before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                            txtReportSaveLoc.Focus()
                        End If
                    Else
                        MsgBox("Please enter a Valid Path to the Recovery Limits Spreadsheet before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                        txt1.Focus()
                    End If
                Else
                    MsgBox("Please select a Valid Report Type before trying to Generate a Report.", MsgBoxStyle.Exclamation)
                    cboType.Focus()
                End If
            End If
        End If

    End Sub

    Private Sub Me_FormClosing(sender As Object, e As FormClosingEventArgs) _
     Handles Me.FormClosing

        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            cboType.Items.Clear()
            cbo1.Items.Clear()
            cbo2.Items.Clear()
            cbo3.Items.Clear()
            cbo4.Items.Clear()
            cbo5.Items.Clear()
            cbo6.Items.Clear()
            txt1.Text = ""
            txtReportSaveLoc.Text = ""
            txtRName.Text = ""
            Me.Hide()
        End If

    End Sub

    Private Sub ReportForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                cboType.Items.Add("Sample Report")
                cboType.Items.Add("CS3 Check Report")
                cboType.Items.Add("LCS Check Report")
                cboType.Items.Add("Final Data")
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                cboType.Items.Add("Sample Report")
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                'cboType.Items.Add("CVS Report")
                cboType.Items.Add("Duplicate Report")
                'cboType.Items.Add("ICV Report")
                cboType.Items.Add("LCS Report")
                'cboType.Items.Add("Method Blank Report")
                cboType.Items.Add("Spike Recovery Report")
                cboType.Items.Add("Summary Report")
                cboType.Items.Add("Custom Report")
                cbo1.Items.Add("eTrain File")
                cbo1.Items.Add("Non Compliance")
                cbo1.Items.Add("LIMS")
                cbo5.Items.Add("N/A")
                cbo5.Items.Add("MDL")
                cbo5.Items.Add("PQL")
                cbo5.Items.Add("RL")
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                cboType.Items.Add("CVS Report")
                cboType.Items.Add("Duplicate Report")
                cboType.Items.Add("ICV Report")
                cboType.Items.Add("LCS Report")
                cboType.Items.Add("Method Blank Report")
                cboType.Items.Add("Spike Recovery Report")
                cboType.Items.Add("Summary Report")
                cboType.Items.Add("Custom Report")
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

    Private Sub SISBrowse_Click(sender As System.Object, e As System.EventArgs) Handles SISBrowse.Click
        Dim fd As New OpenFileDialog()
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                fd.Title = "Open File Dialog"
                fd.InitialDirectory = "\\mdrnd\AS-Global\Special_Access\EAC\Trace\Data\FA-Analysis"
                fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
                fd.FilterIndex = 2
                fd.RestoreDirectory = True
                If fd.ShowDialog() = DialogResult.OK Then
                    txt1.Text = fd.FileName
                End If
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                fd.Title = "Open File Dialog"
                fd.InitialDirectory = "\\mdrnd\AS-Global\Special_Access\EAC\Trace\Data\FA-Analysis"
                fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
                fd.FilterIndex = 2
                fd.RestoreDirectory = True
                If fd.ShowDialog() = DialogResult.OK Then
                    txt1.Text = fd.FileName
                End If
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                fd.Title = "Open File Dialog"
                fd.InitialDirectory = "\\mdrnd\AS-Global\Special_Access\EAC\Chrom\"
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

    Private Sub cboType_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboType.SelectedIndexChanged
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                If cboType.Text = "Sample Report" Then
                    txtRName.Text = "SampleRpt"
                ElseIf cboType.Text = "Final Data" Then
                    txtRName.Text = "FinalData"
                ElseIf cboType.Text = "CS3 Check Report" Then
                    txtRName.Text = "CS3 Check Report"
                ElseIf cboType.Text = "LCS Check Report" Then
                    txtRName.Text = "LCS Check Report"
                End If
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                If cboType.Text = "Sample Report" Then

                End If
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                If cboType.Text = "Summary Report" Then
                    txtRName.Text = "Summary Report"
                    GlobalVariables.Report.Type = ""
                ElseIf cboType.Text = "LCS Report" Then
                    txtRName.Text = "LCS Report"
                    GlobalVariables.Report.Type = "LCS"
                ElseIf cboType.Text = "Spike Recovery Report" Then
                    txtRName.Text = "MS Report"
                    GlobalVariables.Report.Type = "MS"
                ElseIf cboType.Text = "Duplicate Report" Then
                    txtRName.Text = "Duplicate Report"
                    GlobalVariables.Report.Type = "DUP"
                ElseIf cboType.Text = "Method Blank Report" Then
                    txtRName.Text = "MB Report"
                    GlobalVariables.Report.Type = "MB"
                Else
                    GlobalVariables.Report.Type = ""
                End If
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                If cboType.Text = "Summary Report" Then
                    txtRName.Text = "Summary Report"
                End If
            End If
        End If
    End Sub

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
                cbo6.Items.Clear()
                cbo3.Enabled = False
                cbo4.Enabled = False
                cbo5.Enabled = False
                cbo6.Enabled = False
                cbo5.SelectedIndex = -1
                'Clear out loaded permits
                GlobalVariables.PermitList.Clear()
                txt1.Visible = True
                SISBrowse.Visible = True
                lblTxt1.Visible = True
                If cbo1.Text = "LIMS" Then
                    GlobalVariables.Permit.LoadLimsLimit()
                    If cboType.Text = "CVS Report" Then
                        GlobalVariables.Report.Type = "CVS"
                        txt1.Text = "\\helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\FPT_Spike Recovery Limits.xlsx"
                    ElseIf cboType.Text = "Duplicate Report" Then
                        GlobalVariables.Report.Type = "DUP"
                        txt1.Text = "\\helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\FPT_Spike Recovery Limits.xlsx"
                    ElseIf cboType.Text = "ICV Report" Then
                        GlobalVariables.Report.Type = "ICV"
                        txt1.Text = "\\helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\FPT_Spike Recovery Limits.xlsx"
                    ElseIf cboType.Text = "Method Blank Report" Then
                        GlobalVariables.Report.Type = "MB"
                    ElseIf cboType.Text = "Spike Recovery Report" Then
                        GlobalVariables.Report.Type = "MS"
                        txt1.Text = "\\helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\FPT_Spike Recovery Limits.xlsx"
                    ElseIf cboType.Text = "LCS Report" Then
                        GlobalVariables.Report.Type = "LCS"
                        txt1.Text = "\\helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\FPT_Spike Recovery Limits.xlsx"
                    ElseIf cboType.Text = "Summary Report" Then
                        GlobalVariables.Report.Type = "SUM"
                        txt1.Text = "\\helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\FPT_Spike Recovery Limits.xlsx"
                    ElseIf cboType.Text = "Custom Report" Then
                        GlobalVariables.Report.Type = "CUS"
                        txt1.Text = "\\helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\FPT_Spike Recovery Limits.xlsx"
                    End If
                ElseIf cbo1.Text = "eTrain File" Then
                    GlobalVariables.Permit.LoadPermitNames()
                ElseIf cbo1.Text = "Non Compliance" Then
                    txt1.Text = "NonCompliance"
                    GlobalVariables.Permit.LoadNonCompliance()
                    txt1.Visible = False
                    SISBrowse.Visible = False
                    lblTxt1.Visible = False
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
                cbo2.Enabled = True
                'Load method details
                GlobalVariables.Method.LoadMethod(cbo1.Text)
                For Each aMethod In GlobalVariables.MethodList
                    If cbo1.Text = aMethod.Name Then
                        GlobalVariables.selMethod = aMethod
                        For Each aInstrument In aMethod.mInstrumentList
                            cbo2.Items.Add(aInstrument.Name)
                        Next
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                cbo2.Items.Clear()
                cbo2.Text = ""
                cbo2.Enabled = True
                'Load method details
                GlobalVariables.Method.LoadMethod(cbo1.Text)
                For Each aMethod In GlobalVariables.MethodList
                    If cbo1.Text = aMethod.Name Then
                        GlobalVariables.selMethod = aMethod
                        For Each aInstrument In aMethod.mInstrumentList
                            cbo2.Items.Add(aInstrument.Name)
                        Next
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                'Clear out cbos because of change
                lblTxt1.Visible = True
                txt1.Visible = True
                SISBrowse.Visible = True
                cbo2.Items.Clear()
                cbo3.Items.Clear()
                cbo4.Items.Clear()
                cbo6.Items.Clear()
                cbo6.Enabled = False
                cbo3.Enabled = False
                cbo4.Enabled = False
                cbo5.Enabled = False
                cbo6.Enabled = False
                cbo5.SelectedIndex = -1
                GlobalVariables.MidlandChromRLimitNames.Clear()
                'Clear out loaded permits
                GlobalVariables.PermitList.Clear()
                If cbo1.Text = "LIMS" Then
                    If GlobalVariables.Permit.LoadLimsLimit() Then
                        lblTxt1.Visible = True
                        txt1.Visible = True
                        cbo6.Visible = True
                        lbl6.Visible = True
                        SISBrowse.Visible = True
                        'Load up Permit cbo
                        For Each aPermit In GlobalVariables.PermitList
                            cbo2.Items.Add(aPermit.Name)
                        Next
                        cbo2.Enabled = True
                    End If
                    'Load Limits sheet names into 6th cbo
                    If GlobalVariables.Import.MidlandChromRecLimitsNames() Then
                        For Each itm In GlobalVariables.MidlandChromRLimitNames
                            cbo6.Items.Add(itm)
                        Next
                    End If
                ElseIf cbo1.Text = "eTrain File" Then
                    If GlobalVariables.Permit.LoadPermitNames() Then
                        lblTxt1.Visible = True
                        txt1.Visible = True
                        cbo6.Visible = True
                        lbl6.Visible = True
                        SISBrowse.Visible = True
                        'Load up Permit cbo
                        For Each aPermit In GlobalVariables.PermitList
                            cbo2.Items.Add(aPermit.Name)
                        Next
                        cbo2.Enabled = True
                    End If
                    'Load Limits sheet names into 6th cbo
                    If GlobalVariables.Import.MidlandChromRecLimitsNames() Then
                        For Each itm In GlobalVariables.MidlandChromRLimitNames
                            cbo6.Items.Add(itm)
                        Next
                    End If
                ElseIf cbo1.Text = "Non Compliance" Then
                    If GlobalVariables.Permit.LoadNonCompliance() Then
                        lblTxt1.Visible = False
                        txt1.Visible = False
                        cbo6.Visible = False
                        lbl6.Visible = False
                        SISBrowse.Visible = False
                        cbo6.Text = "N/A"
                        'Load up Permit cbo
                        For Each aPermit In GlobalVariables.PermitList
                            cbo2.Items.Add(aPermit.Name)
                        Next
                        cbo2.Enabled = True
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub cbo2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbo2.SelectedIndexChanged
        Dim aPermit As Permit
        Dim aProject As Project
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
                                GlobalVariables.selInstrument = aInstrument.name
                                GlobalVariables.Report.Instrument = aInstrument.name
                                GlobalVariables.Method.Associate13cs(aInstrument)
                            End If
                        Next
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                'Associate 13c's
                For Each aMethod In GlobalVariables.MethodList
                    If aMethod.Name = cbo1.Text Then
                        For Each aInstrument In aMethod.mInstrumentList
                            If aInstrument.Name = cbo2.Text Then
                                GlobalVariables.selInstrument = aInstrument.name
                                GlobalVariables.Report.Instrument = aInstrument.name
                                GlobalVariables.Method.Associate13cs(aInstrument)
                            End If
                        Next
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
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
                                If cboType.Text = "Method Blank Report" And GlobalVariables.selProject = "TPH_DUP" Then
                                    txt1.Text = "\\Helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\MBPCB.txt"
                                    GlobalVariables.strFreeportAnalysis = "PCB"
                                ElseIf cboType.Text = "Method Blank Report" And GlobalVariables.selProject = "M624H_DUP" Then
                                    txt1.Text = "\\Helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\MB624.txt"
                                    GlobalVariables.strFreeportAnalysis = "M624"
                                ElseIf cboType.Text = "Method Blank Report" And GlobalVariables.selProject = "HS_FID_DUP" Then
                                    txt1.Text = "\\Helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\MBBEV.txt"
                                    GlobalVariables.strFreeportAnalysis = "BevCan"
                                ElseIf cboType.Text = "Method Blank Report" Then
                                    txt1.Text = "\\Helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\MBETRAIN.txt"
                                    GlobalVariables.strFreeportAnalysis = "eTrainFile"
                                End If
                                For Each aInstrument In aProject.mInstrumentList
                                    cbo4.Items.Add(aInstrument.Name)
                                Next
                            End If
                        Next
                    End If
                Next
                cbo4.Enabled = True

            End If
        ElseIf GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                cbo4.Items.Clear()
                cbo5.Enabled = False
                cbo5.SelectedIndex = -1
                For Each aPermit In GlobalVariables.PermitList
                    If aPermit.Name = cbo2.Text Then
                        For Each aProject In aPermit.ProjectList
                            If aProject.Name = cbo3.Text Then
                                GlobalVariables.selProject = aProject.Name
                                If cboType.Text = "Method Blank Report" And GlobalVariables.selPermit.Name = "NonCompliance" Then
                                    GlobalVariables.selLimitPath = "\\Helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\TemplateInfo\MBETRAIN.txt"
                                End If

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
        ElseIf GlobalVariables.eTrain.Location = "MIDLAND" Then
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

    Private Sub ExitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.close
    End Sub

    Private Sub cbo5_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbo5.SelectedIndexChanged
        If GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                GlobalVariables.selLimit = cbo5.Text
            End If
        ElseIf GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                GlobalVariables.selLimit = cbo5.Text
                cbo6.Enabled = True
            End If
        End If
    End Sub

    Private Sub UnitConversionToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles UnitConversionToolStripMenuItem.Click
        'UnitConversion.ShowDialog()
    End Sub


    Private Sub cbo6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbo6.SelectedIndexChanged
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                GlobalVariables.Import.MidlandChromBuildRecLimits(cbo6.Text)
            End If
        End If
    End Sub
End Class