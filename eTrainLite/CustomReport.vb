Imports System.IO
Imports Syncfusion.XlsIO
Public Class CustomReport

    Private Sub CustomReport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Dim comboBoxColumn = New DataGridViewComboBoxColumn
        Dim arrCol1() As String

        dgReportTypes.Rows.Clear()
        If GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                If dgReportTypes.ColumnCount < 1 Then
                    comboBoxColumn.Name = "Col1"
                    comboBoxColumn.HeaderText = "Report Type"
                    arrCol1 = {"CVS Report", "Duplicate Report", "ICV Report", "LCS Report", "Method Blank Report", "Spike Recovery Report", "Summary Report"}
                    comboBoxColumn.Width = 200
                    For Each c In arrCol1
                        comboBoxColumn.Items.Add(c)
                    Next
                    comboBoxColumn.Visible = True
                    dgReportTypes.Columns.Add(comboBoxColumn)
                End If
            End If
        ElseIf GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                If dgReportTypes.ColumnCount < 1 Then
                    comboBoxColumn.Name = "Col1"
                    comboBoxColumn.HeaderText = "Report Type"
                    arrCol1 = {"Cover Page, Narrative, Flags", "Duplicate Report", "LCS Report", "Method Blank Report", "Spike Recovery Report", "Summary Report"}
                    comboBoxColumn.Width = 200
                    For Each c In arrCol1
                        comboBoxColumn.Items.Add(c)
                    Next
                    comboBoxColumn.Visible = True
                    dgReportTypes.Columns.Add(comboBoxColumn)
                End If
            End If
        End If
    End Sub

    Private Sub btnSelect_Click(sender As System.Object, e As System.EventArgs) Handles btnGenerate.Click
        Dim blnFlg As Boolean
        Dim aSample As Sample

        If GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                Try
                    For Each r In dgReportTypes.Rows
                        blnFlg = False
                        For Each aSample In GlobalVariables.SampleList
                            aSample.Reported = False
                        Next
                        If r.Cells.Item(0).Value = "Summary Report" Then
                            If Not GlobalVariables.Report.FreeportChromSummaryReport(GlobalVariables.selLimit) Then
                                MsgBox("Summary Report not created!", MsgBoxStyle.Information)
                            End If
                        ElseIf r.Cells.Item(0).Value = "Spike Recovery Report" Then
                            For Each aSample In GlobalVariables.SampleList
                                If aSample.Type = "MS" Then
                                    blnFlg = True
                                End If
                            Next
                            If blnFlg Then
                                If Not GlobalVariables.Report.FreeportChromMSReport(GlobalVariables.selLimit) Then
                                    MsgBox("Spike Recovery Report not created!", MsgBoxStyle.Information)
                                End If
                            Else
                                MsgBox("Spike Recovery Report not created, sample of required type (MS) not found to generate report!", MsgBoxStyle.Information)
                            End If
                        ElseIf r.Cells.Item(0).Value = "Duplicate Report" Then
                            For Each aSample In GlobalVariables.SampleList
                                If aSample.Type = "DUP" Then
                                    blnFlg = True
                                End If
                            Next
                            If blnFlg Then
                                If Not GlobalVariables.Report.FreeportChromDUPReport(GlobalVariables.selLimit) Then
                                    MsgBox("Duplicate Report not created!", MsgBoxStyle.Information)
                                End If
                            Else
                                MsgBox("Duplicate Report not created, sample of required type (DUP) not found to generate report!", MsgBoxStyle.Information)
                            End If
                        ElseIf r.Cells.Item(0).Value = "LCS Report" Then
                            For Each aSample In GlobalVariables.SampleList
                                If aSample.Type = "LCS" Then
                                    blnFlg = True
                                End If
                            Next
                            If blnFlg Then
                                If Not GlobalVariables.Report.FreeportChromLCSReport() Then
                                    MsgBox("LCS Report not created!", MsgBoxStyle.Information)
                                End If
                            Else
                                MsgBox("LCS Report not created, sample of required type (LCS) not found to generate report!", MsgBoxStyle.Information)
                            End If
                        ElseIf r.Cells.Item(0).Value = "Method Blank Report" Then
                            For Each aSample In GlobalVariables.SampleList
                                If aSample.Type = "MB" Then
                                    blnFlg = True
                                End If
                            Next
                            If blnFlg Then
                                If Not GlobalVariables.Report.FreeportChromMBReport(GlobalVariables.selLimitPath, GlobalVariables.selLimit) Then
                                    MsgBox("Method Blank Report not created!", MsgBoxStyle.Information)
                                End If
                            Else
                                MsgBox("Method Blank Report not created, sample of required type (MB) not found to generate report!", MsgBoxStyle.Information)
                            End If
                        ElseIf r.Cells.Item(0).Value = "CVS Report" Then
                            For Each aSample In GlobalVariables.SampleList
                                If aSample.Type = "CVS" Then
                                    blnFlg = True
                                End If
                            Next
                            If blnFlg Then
                                If Not GlobalVariables.Report.FreeportChromCVSReport() Then
                                    MsgBox("CVS Report not created!", MsgBoxStyle.Information)
                                End If
                            Else
                                MsgBox("CVS Report not created, sample of required type (CVS) not found to generate report!", MsgBoxStyle.Information)
                            End If
                        ElseIf r.Cells.Item(0).Value = "ICV Report" Then
                            For Each aSample In GlobalVariables.SampleList
                                If aSample.Type = "ICV" Then
                                    blnFlg = True
                                End If
                            Next
                            If blnFlg Then
                                If Not GlobalVariables.Report.FreeportChromICVReport() Then
                                    MsgBox("ICV Report not created!", MsgBoxStyle.Information)
                                End If
                            Else
                                MsgBox("ICV Report not created, sample of required type (ICV) not found to generate report!", MsgBoxStyle.Information)
                            End If
                        End If
                    Next
                    Me.Hide()
                Catch ex As Exception
                    MsgBox("Error generating report!" & vbCrLf & _
                        "Sub Procedure: btnGenerate_Click_CustomReport()" & vbCrLf & _
                        "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                    GlobalVariables.CustomReportError = True
                    Me.Hide()
                End Try
            End If
        ElseIf GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                Try
                    For Each r In dgReportTypes.Rows
                        blnFlg = False
                        For Each aSample In GlobalVariables.SampleList
                            aSample.Reported = False
                        Next
                        If r.Cells.Item(0).Value = "Summary Report" Then
                            If Not GlobalVariables.Report.MidlandChromSummaryReport(GlobalVariables.selLimit) Then
                                MsgBox("Summary Report not created!", MsgBoxStyle.Information)
                            End If
                        ElseIf r.Cells.Item(0).Value = "Spike Recovery Report" Then
                            For Each aSample In GlobalVariables.SampleList
                                If aSample.Type = "MS" Then
                                    blnFlg = True
                                End If
                            Next
                            If blnFlg Then
                                If Not GlobalVariables.Report.MidlandChromMSReport(GlobalVariables.selLimit) Then
                                    MsgBox("Spike Recovery Report not created!", MsgBoxStyle.Information)
                                End If
                            Else
                                MsgBox("Spike Recovery Report not created, sample of required type (MS) not found to generate report!", MsgBoxStyle.Information)
                            End If
                        ElseIf r.Cells.Item(0).Value = "Duplicate Report" Then
                            For Each aSample In GlobalVariables.SampleList
                                If aSample.Type = "DUP" Then
                                    blnFlg = True
                                End If
                            Next
                            If blnFlg Then
                                If Not GlobalVariables.Report.MidlandChromDUPReport(GlobalVariables.selLimit) Then
                                    MsgBox("Duplicate Report not created!", MsgBoxStyle.Information)
                                End If
                            Else
                                MsgBox("Duplicate Report not created, sample of required type (DUP) not found to generate report!", MsgBoxStyle.Information)
                            End If
                        ElseIf r.Cells.Item(0).Value = "LCS Report" Then
                            For Each aSample In GlobalVariables.SampleList
                                If aSample.Type = "LCS" Then
                                    blnFlg = True
                                End If
                            Next
                            If blnFlg Then
                                If Not GlobalVariables.Report.MidlandChromLCSReport() Then
                                    MsgBox("LCS Report not created!", MsgBoxStyle.Information)
                                End If
                            Else
                                MsgBox("LCS Report not created, sample of required type (LCS) not found to generate report!", MsgBoxStyle.Information)
                            End If
                        ElseIf r.Cells.Item(0).Value = "Method Blank Report" Then
                            For Each aSample In GlobalVariables.SampleList
                                If aSample.Type = "MB" Then
                                    blnFlg = True
                                End If
                            Next
                            'If blnFlg Then
                            '    If Not GlobalVariables.Report.MidlandChromMBReport(GlobalVariables.selLimitPath) Then
                            '        MsgBox("Method Blank Report not created!", MsgBoxStyle.Information)
                            '    End If
                            'Else
                            '    MsgBox("Method Blank Report not created, sample of required type (MB) not found to generate report!", MsgBoxStyle.Information)
                            'End If
                        ElseIf r.Cells.Item(0).Value = "Cover Page, Narrative, Flags" Then
                            If Not GlobalVariables.Report.MidlandChromCustomerReport() Then
                                MsgBox("Cover Page, Narrative, Flags not created!", MsgBoxStyle.Information)
                            End If
                        End If
                    Next
                    Me.Hide()
                Catch ex As Exception
                    MsgBox("Error generating report!" & vbCrLf & _
                        "Sub Procedure: btnGenerate_Click_CustomReport()" & vbCrLf & _
                        "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                    GlobalVariables.CustomReportError = True
                    Me.Hide()
                End Try
            End If
        End If
    End Sub

    Private Sub btnAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd.Click
        dgReportTypes.Rows.Add("", "")
    End Sub

    Private Sub btnRemove_Click(sender As System.Object, e As System.EventArgs) Handles btnRemove.Click
        For Each r In dgReportTypes.SelectedRows
            Try
                dgReportTypes.Rows.RemoveAt(r.Index)
            Catch ex As Exception
                MsgBox("Error: No need to delete starter row.", MsgBoxStyle.Exclamation)
            End Try
        Next
    End Sub
    Private Sub Me_FormClosing(sender As Object, e As FormClosingEventArgs) _
    Handles Me.FormClosing

        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
        End If

    End Sub
End Class