Imports Syncfusion.XlsIO

Public Class MHImport

    'Midland Chrom MassHunter Import
    Function MidlandChromImport() As Boolean
        Dim exEngine As New ExcelEngine
        Dim exApp As IApplication
        Dim aSample As Sample
        Dim aCompound As Compound
        Dim aSurrogate As Surrogate
        Dim intSC As Integer
        Dim intColC As Integer
        Dim workbook As IWorkbook
        Dim worksheet As IWorksheet
        Dim arrSplText() As String

        Try
            'Import file with analytes and limits for method blank report
            exApp = exEngine.Excel
            workbook = exApp.Workbooks.Open(GlobalVariables.Import.FilePath)
            worksheet = workbook.Worksheets(0)
            intSC = 0
            intColC = 3
            Do Until worksheet.Range((3 + intSC), intColC).Value = ""
                aSample = New Sample
                aSample.Name = worksheet.Range((3 + intSC), intColC).Value
                'Sample type assignment
                If InStr(aSample.Name, "Method Blank") Then
                    aSample.Type = "MB"
                ElseIf InStr(aSample.Name, "Lab Control Spike DUP") Or InStr(aSample.Name, "LCSD", CompareMethod.Binary) Then
                    aSample.Type = "LCSD"
                ElseIf InStr(aSample.Name, "Lab Control Spike") Or InStr(aSample.Name, "LCS", CompareMethod.Binary) Then
                    aSample.Type = "LCS"
                ElseIf InStr(aSample.Name, "MSD", CompareMethod.Binary) Then
                    aSample.Type = "MSD"
                ElseIf InStr(aSample.Name, "MS", CompareMethod.Binary) Then
                    aSample.Type = "MS"
                ElseIf InStr(aSample.Name, "Lab Blank") Then
                    aSample.Type = "LB"
                ElseIf InStr(aSample.Name, "CVS", CompareMethod.Binary) Then
                    aSample.Type = "CVS"
                ElseIf InStr(aSample.Name, "ICV", CompareMethod.Binary) Then
                    aSample.Type = "ICV"
                ElseIf InStr(aSample.Name, "DUP", CompareMethod.Binary) Then
                    aSample.Type = "DUP"
                ElseIf InStr(aSample.Name, "CS", CompareMethod.Binary) Then
                    aSample.Type = "CHECK"
                ElseIf InStr(aSample.Name, "Standard") Then
                    aSample.Type = "STD"
                Else
                    aSample.Type = "SAMPLE"
                End If
                aSample.DataFile = worksheet.Range((3 + intSC), intColC + 1).Value
                aSample.Misc = worksheet.Range((3 + intSC), intColC + 3).Value
                If aSample.Misc <> "" And InStr(aSample.Misc, ",") Then
                    arrSplText = aSample.Misc.Split(",")
                    aSample.LimsID = arrSplText(0)
                    aSample.SampleDate = CDate(arrSplText(1))
                    aSample.DilutionFactor = arrSplText(2)
                    aSample.DetectLimitType = arrSplText(3)
                    aSample.Matrix = arrSplText(4)
                End If
                aSample.QMethFile = worksheet.Range((3 + intSC), intColC + 4).Value
                aSample.QuantTime = CDate(worksheet.Range((3 + intSC), intColC + 5).Value)
                Do Until worksheet.Range((3 + intSC), intColC + 6).Value = ""
                    If worksheet.Range((3 + intSC), intColC + 6).Value = "Surrogate" Then
                        aSurrogate = New Surrogate
                        aSurrogate.Units = worksheet.Range((3 + intSC), intColC + 7).Value
                        aSample.Units = aSurrogate.Units
                        aSurrogate.Name = worksheet.Range((3 + intSC), intColC + 8).Value
                        aSurrogate.Response = worksheet.Range((3 + intSC), intColC + 9).Value
                        aSurrogate.Conc = worksheet.Range((3 + intSC), intColC + 10).Value
                        aSurrogate.MI = CBool(worksheet.Range((3 + intSC), intColC + 11).Value)
                        intColC = intColC + 6
                        aSample.SurrogateList.Add(aSurrogate)
                    ElseIf worksheet.Range((3 + intSC), intColC + 6).Value = "Target" Then
                        aCompound = New Compound
                        aCompound.Units = worksheet.Range((3 + intSC), intColC + 7).Value
                        aSample.Units = aCompound.Units
                        aCompound.Name = worksheet.Range((3 + intSC), intColC + 8).Value
                        aCompound.Response = worksheet.Range((3 + intSC), intColC + 9).Value
                        aCompound.Conc = worksheet.Range((3 + intSC), intColC + 10).Value
                        aCompound.MI = CBool(worksheet.Range((3 + intSC), intColC + 11).Value)
                        intColC = intColC + 6
                        aSample.CompoundList.Add(aCompound)
                    End If
                Loop
                intSC = intSC + 1
                intColC = 3
                GlobalVariables.SampleList.Add(aSample)
            Loop
            workbook.Close()
            exEngine.Dispose()
            Return True
        Catch ex As Exception
            MsgBox("Error reading MassHunter file" & vbCrLf & _
                   "Sub Procedure: MidlandChromImport()" & vbCrLf & _
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            If Not IsNothing(workbook) Then
                workbook.Close()
                exEngine.Dispose()
            End If
            Return False
        End Try
    End Function


End Class
