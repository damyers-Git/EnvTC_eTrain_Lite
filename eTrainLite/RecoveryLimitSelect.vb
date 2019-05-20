Imports System.IO
Imports Syncfusion.XlsIO
Imports System.Text.RegularExpressions
Public Class RecoveryLimitSelect

    Private Sub btnFindSheets_Click(sender As System.Object, e As System.EventArgs) Handles btnFindSheets.Click
        Dim exEngine As New ExcelEngine
        Dim exApp As IApplication
        Dim workbook As IWorkbook
        Dim worksheet As IWorksheet

        Try
            exApp = exEngine.Excel
            workbook = exApp.Workbooks.Open(txtLimitPath.Text)
            For Each worksheet In workbook.Worksheets
                cboSheetName.Items.Add(worksheet.Name)
            Next
            workbook.Close()
            exEngine.Dispose()
            cboSheetName.Enabled = True
            btnLoadLimits.Enabled = True

        Catch ex As Exception
            MsgBox("Error reading Recovery Limits File: " & txtLimitPath.Text & vbCrLf & _
                   "Sub Procedure: btnFindSheets_Click()" & vbCrLf & _
                    "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnLoadLimits_Click(sender As System.Object, e As System.EventArgs) Handles btnLoadLimits.Click
        GlobalVariables.Import.FreeportChromBuildRecLimits(txtLimitPath.Text)
        Me.Close()
    End Sub
End Class