Imports Syncfusion.XlsIO

Public Class TQ3Import

    'Midland High Res TQ3 Import
    Function MidlandHRImport() As Boolean
        Dim exEngine As New ExcelEngine
        Dim exApp As IApplication
        Dim aSample As Sample
        Dim aCompound As Compound
        Dim aStandard As Standard
        Dim workbook As IWorkbook
        Dim worksheet As IWorksheet
        Dim i As Integer

        Try
            'Import file with analytes and limits for method blank report
            aSample = New Sample
            exApp = exEngine.Excel
            workbook = exApp.Workbooks.Open(GlobalVariables.Import.FilePath)
            worksheet = workbook.Worksheets(0)

            'Load TQ3 details
            aSample.TQ3QuanFile = worksheet.Range("B1").Value & worksheet.Range("C1").Value & worksheet.Range("D1").Value
            aSample.TQ3DataFile = worksheet.Range("B2").Value & worksheet.Range("C2").Value
            aSample.TQ3ResponseFile = worksheet.Range("B3").Value & worksheet.Range("C3").Value & worksheet.Range("D3").Value
            aSample.TQ3Entries = worksheet.Range("B4").Value
            aSample.AcqDate = worksheet.Range("B5").Value
            aSample.Misc = worksheet.Range("B6").Value
            aSample.Name = worksheet.Range("B7").Value
            'Set sample type
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
            ElseIf InStr(aSample.Name, "TEF Cal", CompareMethod.Binary) Then
                aSample.Type = "TEFCAL"
            ElseIf InStr(aSample.Name, "ICV", CompareMethod.Binary) Then
                aSample.Type = "ICV"
            ElseIf InStr(aSample.Name, "DUP", CompareMethod.Binary) Then
                aSample.Type = "DUP"
            Else
                aSample.Type = "SAMPLE"
            End If

            aSample.Vial = worksheet.Range("B8").Value
            aSample.TQ3SampleID = worksheet.Range("B9").Value
            aSample.TQ3Study = worksheet.Range("B10").Value
            aSample.TQ3Client = worksheet.Range("B11").Value
            aSample.TQ3Laboratory = worksheet.Range("B12").Value
            aSample.TQ3Operator = worksheet.Range("B13").Value
            aSample.TQ3Phone = worksheet.Range("B14").Value
            aSample.TQ3Barcode = worksheet.Range("B15").Value
            aSample.TQ3QUALCompatMode = worksheet.Range("B16").Value
            aSample.TQ3InjectionVol = worksheet.Range("B17").Value
            aSample.TQ3SampleVol = worksheet.Range("B18").Value
            aSample.TQ3SampleWeight = worksheet.Range("B19").Value
            aSample.TQ3DilutionFactor = worksheet.Range("B20").Value
            aSample.TQ3DetLimitFactor = worksheet.Range("B21").Value
            aSample.TQ3DisplayQuantStatusArea = worksheet.Range("B22").Value
            aSample.TQ3DisplayQuantStatusHeight = worksheet.Range("B23").Value
            aSample.TQ3SumQMRM1 = worksheet.Range("B24").Value
            aSample.TQ3SumQMRM2 = worksheet.Range("B25").Value
            aSample.TQ3SinglePointRF = worksheet.Range("B26").Value
            aSample.TQ3AvgRF = worksheet.Range("B27").Value
            aSample.TQ3RFvsArea = worksheet.Range("B28").Value
            aSample.TQ3AreaRatiovsConc = worksheet.Range("B29").Value
            aSample.TQ3LinearFit = worksheet.Range("B30").Value
            aSample.TQ3SquareFit = worksheet.Range("B31").Value
            aSample.TQ3NonWeightedRegress = worksheet.Range("B32").Value
            aSample.TQ3RegressWeighted1Amt = worksheet.Range("B33").Value
            aSample.TQ3RegressWeighted1Resp = worksheet.Range("B34").Value
            aSample.TQ3WeightedRegressFactor = worksheet.Range("B35").Value

            'Compound/Standard import
            i = 37
            Do Until (worksheet.Range(i, 2).Value = "")
                If InStr(worksheet.Range(i, 2).Value, "13C") Then
                    aStandard = New Standard
                    aStandard.Name = worksheet.Range(i, 2).Value
                    aStandard.TQ3QuanMass = worksheet.Range(i, 3).Value
                    aStandard.TQ3QMHeight = worksheet.Range(i, 4).Value
                    aStandard.TQ3QMArea = worksheet.Range(i, 5).Value
                    aStandard.TQ3QMSigNoi = worksheet.Range(i, 6).Value
                    aStandard.TQ3QMNoise = worksheet.Range(i, 7).Value
                    aStandard.TQ3RatioMass = worksheet.Range(i, 8).Value
                    aStandard.TQ3RM1Height = worksheet.Range(i, 9).Value
                    aStandard.TQ3RM1Area = worksheet.Range(i, 10).Value
                    aStandard.TQ3RM1SigNoi = worksheet.Range(i, 11).Value
                    aStandard.TQ3RM1Noise = worksheet.Range(i, 12).Value
                    aStandard.TQ3R1R2 = worksheet.Range(i, 13).Value
                    aStandard.TQ3SpecAmt = worksheet.Range(i, 14).Value
                    aStandard.TQ3MI1 = worksheet.Range(i, 15).Value
                    aStandard.TQ3MI2 = worksheet.Range(i, 16).Value
                    aSample.InternalStdList.Add(aStandard)
                Else
                    aCompound = New Compound
                    aCompound.Name = worksheet.Range(i, 2).Value
                    aCompound.TQ3QuanMass = worksheet.Range(i, 3).Value
                    aCompound.TQ3QMHeight = worksheet.Range(i, 4).Value
                    aCompound.TQ3QMArea = worksheet.Range(i, 5).Value
                    aCompound.TQ3QMSigNoi = worksheet.Range(i, 6).Value
                    aCompound.TQ3QMNoise = worksheet.Range(i, 7).Value
                    aCompound.TQ3RatioMass = worksheet.Range(i, 8).Value
                    aCompound.TQ3RM1Height = worksheet.Range(i, 9).Value
                    aCompound.TQ3RM1Area = worksheet.Range(i, 10).Value
                    aCompound.TQ3RM1SigNoi = worksheet.Range(i, 11).Value
                    aCompound.TQ3RM1Noise = worksheet.Range(i, 12).Value
                    aCompound.TQ3R1R2 = worksheet.Range(i, 13).Value
                    aCompound.TQ3SpecAmt = worksheet.Range(i, 14).Value
                    aCompound.TQ3MI1 = worksheet.Range(i, 15).Value
                    aCompound.TQ3MI2 = worksheet.Range(i, 16).Value
                    aSample.CompoundList.Add(aCompound)
                End If


                i = i + 1
            Loop
            workbook.Close()
            exEngine.Dispose()
            GlobalVariables.SampleList.Add(aSample)
            Return True
        Catch ex As Exception
            MsgBox("Error reading TQ3 file" & vbCrLf & _
                   "Sub Procedure: MidlandHRImport()" & vbCrLf & _
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try


    End Function

End Class
