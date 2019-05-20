Imports System.Data
Imports System.IO
Public Class Calculations

    'Freeport Chrom reported Amount calculations
    Function FreeportChrom(ByVal strLimitType As String, ByVal strRecLimitPath As String, ByVal blnTransfer As Boolean) As Boolean
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aInstrument As mInstrument
        Dim aSample As Sample
        Dim aSample2 As Sample
        Dim aDupSample As Sample
        Dim aQCSample As Sample
        Dim aOGSample As Sample
        Dim aCompound As Compound
        Dim aCompound2 As Compound
        Dim aSurrogate As Surrogate
        Dim amSurrogate As mSurrogate
        Dim amCompound As mCompound
        Dim blnMatch As Boolean
        Dim blnIndirect As Boolean

        Try
            UnitConversion.ShowDialog()
            blnIndirect = False
            For Each aPermit In GlobalVariables.PermitList
                If aPermit.Name = GlobalVariables.selPermit.Name Then
                    For Each aProject In aPermit.ProjectList
                        If aProject.Name = GlobalVariables.selProject Then
                            For Each aInstrument In aProject.mInstrumentList
                                If aInstrument.Name = GlobalVariables.selInstrument Then
                                    'Load in recovery limits if not noncompliance
                                    If strRecLimitPath <> "NonCompliance" Then
                                        If Not GlobalVariables.Import.FreeportChromBuildRecLimits(strRecLimitPath) Then
                                            Return False
                                        End If
                                    Else
                                        'Assign recovery limits if noncomplance
                                        For Each aSample In GlobalVariables.ReportSamList
                                            For Each aSurrogate In aSample.SurrogateList
                                                For Each amSurrogate In aInstrument.mSurrogateList
                                                    If aSurrogate.Name = amSurrogate.Name Then
                                                        aSurrogate.ChromLowContLim = amSurrogate.RecLLim
                                                        aSurrogate.ChromUpContLim = amSurrogate.RecULim
                                                        aSurrogate.ChromLowLCSLim = amSurrogate.LCSLLim
                                                        aSurrogate.ChromUpLCSLim = amSurrogate.LCSULim
                                                        aSurrogate.ChromLowMSLim = amSurrogate.MSLLim
                                                        aSurrogate.ChromUpMSLim = amSurrogate.MSULim
                                                    End If
                                                Next
                                            Next
                                            For Each aCompound In aSample.CompoundList
                                                For Each amCompound In aInstrument.mCompoundList
                                                    If aCompound.Name = amCompound.Name Then
                                                        aCompound.ChromLowContLim = amCompound.RecLLim
                                                        aCompound.ChromUpContLim = amCompound.RecULim
                                                        aCompound.ChromLowLCSLim = amCompound.LCSLLim
                                                        aCompound.ChromUpLCSLim = amCompound.LCSULim
                                                        aCompound.ChromLowMSLim = amCompound.MSLLim
                                                        aCompound.ChromUpMSLim = amCompound.MSULim
                                                    End If
                                                Next
                                            Next
                                        Next
                                    End If
                                    For Each aSample In GlobalVariables.ReportSamList
                                        'Compounds
                                        If Not aSample.Calculated Then
                                            For Each aCompound In aSample.CompoundList
                                                'Determine compound and mcompound relationship
                                                blnMatch = False

                                                For Each amCompound In aInstrument.mCompoundList
                                                    If aCompound.Name = amCompound.Name Then
                                                        blnMatch = True
                                                    ElseIf UCase(aCompound.Name) = UCase(amCompound.Name) Then
                                                        If blnTransfer Then
                                                            If blnIndirect Then
                                                                blnMatch = True
                                                            Else
                                                                If vbYes = MsgBox("Compound: " & aCompound.Name & " is an indirect match for LIMS Component: " & amCompound.Name & ". Do you want to continue?", MsgBoxStyle.YesNo, "eTrain 2.0") Then
                                                                    If vbYes = MsgBox("Would you like to continue transfer including all indirect matches?", MsgBoxStyle.YesNo, "eTrain 2.0") Then
                                                                        blnIndirect = True
                                                                    End If
                                                                    blnMatch = True
                                                                Else
                                                                    MsgBox("Please update Method file so Compound: " & aCompound.Name & " matches LIMS Component: " & amCompound.Name & " exactly and try transfer again.", MsgBoxStyle.Exclamation, "eTrain 2.0")
                                                                    Return False
                                                                End If
                                                            End If
                                                        Else
                                                            blnMatch = True
                                                        End If
                                                    ElseIf amCompound.AliasList.Count > 0 Then
                                                        If amCompound.DetermineMatch(aCompound) Then
                                                            blnMatch = True
                                                        End If
                                                    End If
                                                    If blnMatch Then
                                                        If strLimitType = "MDL" Then
                                                            aCompound.ChromReportLimit = amCompound.MDL
                                                            aCompound.ChromAdjustedLimit = CDbl(amCompound.MDL) * CDbl(aSample.DilutionFactor)
                                                            If IsNumeric(aCompound.Conc) Then
                                                                aCompound.ChromAdjustConc = CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)
                                                                If CDbl(aCompound.ChromAdjustConc) > CDbl(aCompound.ChromAdjustedLimit) Then
                                                                    aCompound.ReportedAmt = aCompound.ChromAdjustConc
                                                                Else
                                                                    aCompound.ReportedAmt = "ND"
                                                                End If
                                                            Else
                                                                aCompound.ReportedAmt = "ND"
                                                            End If
                                                            'Format
                                                            ' aCompound.ReportedAmt = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
                                                        ElseIf strLimitType = "PQL" Then
                                                            aCompound.ChromReportLimit = amCompound.PQL
                                                            aCompound.ChromAdjustedLimit = CDbl(amCompound.PQL) * CDbl(aSample.DilutionFactor)
                                                            If IsNumeric(aCompound.Conc) Then
                                                                aCompound.ChromAdjustConc = CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)
                                                                If CDbl(aCompound.ChromAdjustConc) > CDbl(aCompound.ChromAdjustedLimit) Then
                                                                    aCompound.ReportedAmt = aCompound.ChromAdjustConc
                                                                Else
                                                                    aCompound.ReportedAmt = "ND"
                                                                End If
                                                            Else
                                                                aCompound.ReportedAmt = "ND"
                                                            End If
                                                            ' aCompound.ReportedAmt = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
                                                        ElseIf strLimitType = "RL" Then
                                                            If IsNumeric(aCompound.Conc) Then
                                                                aCompound.ChromAdjustConc = CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)
                                                                If CDbl(aCompound.ChromAdjustConc) > CDbl(amCompound.RL) Then
                                                                    aCompound.ReportedAmt = aCompound.ChromAdjustConc
                                                                Else
                                                                    aCompound.ReportedAmt = "ND"
                                                                End If
                                                            Else
                                                                aCompound.ReportedAmt = "ND"
                                                            End If
                                                            ' aCompound.ReportedAmt = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
                                                        ElseIf strLimitType = "N/A" Then
                                                            aCompound.ReportedAmt = aCompound.Conc
                                                            ' aCompound.ReportedAmt = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
                                                        End If
                                                        blnMatch = False
                                                    End If
                                                Next
                                            Next
                                        End If
                                    Next
                                    'Only Run this section if not transfer
                                    If Not blnTransfer Then
                                        'Calculate MS Recoveries & RPD
                                        For Each aQCSample In GlobalVariables.ReportSamList
                                            If Not aQCSample.Calculated Then
                                                If aQCSample.Type = "MS" And GlobalVariables.Report.Type = "MS" Then
                                                    'Find regular sample
                                                    For Each aOGSample In GlobalVariables.ReportSamList
                                                        If InStr(aQCSample.Name, aOGSample.Name) And aOGSample.Type = "SAMPLE" And Not aQCSample.SpikeCalculated Then
                                                            'MS Recovery
                                                            With SpikeInfo
                                                                .lblSampleName.Text = "Sample: " & aQCSample.Name
                                                                .lblDil.Text = "Dilution Factor: " & aQCSample.DilutionFactor
                                                                .lblUnits.Text = "Ending Conversion Units: " & aQCSample.Units
                                                                .txtConc.Clear()
                                                                .txtVol.Clear()
                                                            End With
                                                            SpikeInfo.ShowDialog()
                                                            For Each aCompound In aQCSample.CompoundList
                                                                If IsNumeric(aCompound.Conc) Then
                                                                    aCompound.ChromCorrectedSpike = aQCSample.ChromSpikeAmt
                                                                    aCompound.ChromSpikeRecovery = CStr((CDbl(aCompound.Conc) / CDbl(aCompound.ChromCorrectedSpike)) * 100)
                                                                    If IsNumeric(aCompound.ChromLowContLim) And IsNumeric(aCompound.ChromUpContLim) Then
                                                                        If CDbl(aCompound.ChromSpikeRecovery) >= CDbl(aCompound.ChromLowMSLim) And CDbl(aCompound.ChromSpikeRecovery) <= CDbl(aCompound.ChromUpMSLim) Then
                                                                            aCompound.ChromSpikePass = True
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next
                                                            'See if MSD sample
                                                            For Each aDupSample In GlobalVariables.ReportSamList
                                                                If aDupSample.Type = "MSD" And InStr(aDupSample.Name, aOGSample.Name) Then
                                                                    'MS MSD RPD
                                                                    For Each aCompound In aQCSample.CompoundList
                                                                        For Each aCompound2 In aDupSample.CompoundList
                                                                            If aCompound.Name = aCompound2.Name Then
                                                                                If IsNumeric(aCompound.Conc) And IsNumeric(aCompound2.Conc) Then
                                                                                    aCompound.ChromRPD = Math.Abs(CDbl(aCompound.Conc) - CDbl(aCompound2.Conc)) / (CDbl(aCompound.Conc) + CDbl(aCompound2.Conc) / 2) * 100
                                                                                    aCompound2.ChromRPD = aCompound.ChromRPD
                                                                                    aCompound.ChromRPDLimit = "30"
                                                                                    aCompound2.ChromRPDLimit = "30"
                                                                                Else
                                                                                    aCompound.ChromRPD = "N/A"
                                                                                    aCompound2.ChromRPD = aCompound.ChromRPD
                                                                                    aCompound.ChromRPDLimit = "30"
                                                                                    aCompound2.ChromRPDLimit = "30"
                                                                                End If
                                                                            End If
                                                                        Next
                                                                    Next
                                                                    'MSD Recovery
                                                                    With SpikeInfo
                                                                        .lblSampleName.Text = "Sample: " & aDupSample.Name
                                                                        .lblDil.Text = "Dilution Factor: " & aDupSample.DilutionFactor
                                                                        .lblUnits.Text = "Ending Conversion Units: " & aDupSample.Units
                                                                        .txtConc.Clear()
                                                                        .txtVol.Clear()
                                                                    End With
                                                                    SpikeInfo.ShowDialog()
                                                                    For Each aCompound In aDupSample.CompoundList
                                                                        If IsNumeric(aCompound.Conc) Then
                                                                            aCompound.ChromCorrectedSpike = aDupSample.ChromSpikeAmt
                                                                            aCompound.ChromSpikeRecovery = CStr((CDbl(aCompound.Conc) / CDbl(aCompound.ChromCorrectedSpike)) * 100)
                                                                            If IsNumeric(aCompound.ChromLowContLim) And IsNumeric(aCompound.ChromUpContLim) Then
                                                                                If CDbl(aCompound.ChromSpikeRecovery) >= CDbl(aCompound.ChromLowMSLim) And CDbl(aCompound.ChromSpikeRecovery) <= CDbl(aCompound.ChromUpMSLim) Then
                                                                                    aCompound.ChromSpikePass = True
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Next
                                                                End If
                                                            Next
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Next
                                        'Calculate LCS Recoveries
                                        For Each aSample In GlobalVariables.ReportSamList
                                            If Not aSample.Calculated Then
                                                If aSample.Type = "LCS" Or aSample.Type = "LCSD" And GlobalVariables.Report.Type = "LCS" And Not aSample.SpikeCalculated Then
                                                    With SpikeInfo
                                                        .lblSampleName.Text = "Sample: " & aSample.Name
                                                        .lblDil.Text = "Dilution Factor: " & aSample.DilutionFactor
                                                        .lblUnits.Text = "Ending Conversion Units: " & aSample.Units
                                                        .txtConc.Clear()
                                                        .txtVol.Clear()
                                                    End With
                                                    SpikeInfo.ShowDialog()
                                                    For Each aCompound In aSample.CompoundList
                                                        If IsNumeric(aCompound.Conc) Then
                                                            aCompound.ChromCorrectedSpike = aSample.ChromSpikeAmt
                                                            aCompound.ChromSpikeRecovery = CStr((CDbl(aCompound.Conc) / CDbl(aCompound.ChromCorrectedSpike)) * 100)
                                                            If IsNumeric(aCompound.ChromLowContLim) And IsNumeric(aCompound.ChromUpContLim) Then
                                                                If CDbl(aCompound.ChromSpikeRecovery) >= CDbl(aCompound.ChromLowLCSLim) And CDbl(aCompound.ChromSpikeRecovery) <= CDbl(aCompound.ChromUpLCSLim) Then
                                                                    aCompound.ChromSpikePass = True
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Next
                                        'ICV/CVS
                                        For Each aSample In GlobalVariables.ReportSamList
                                            If Not aSample.Calculated Then
                                                If aSample.Type = "ICV" Then
                                                    For Each aCompound In aSample.CompoundList
                                                        If aCompound.ChromLowICVLim = "" Then
                                                            If IsNumeric(aCompound.Conc) Then
                                                                If CDbl(aCompound.Conc) >= CDbl(aCompound.ChromLowContLim) And CDbl(aCompound.Conc) <= CDbl(aCompound.ChromUpContLim) Then
                                                                    aCompound.ChromICVPass = True
                                                                End If
                                                            Else
                                                                aCompound.ChromICVPass = False
                                                            End If
                                                        Else
                                                            If IsNumeric(aCompound.Conc) Then
                                                                If CDbl(aCompound.Conc) >= CDbl(aCompound.ChromLowICVLim) And CDbl(aCompound.Conc) <= CDbl(aCompound.ChromUpICVLim) Then
                                                                    aCompound.ChromICVPass = True
                                                                End If
                                                            Else
                                                                aCompound.ChromICVPass = False
                                                            End If
                                                        End If
                                                    Next
                                                ElseIf aSample.Type = "CVS" Then
                                                    For Each aCompound In aSample.CompoundList
                                                        If aCompound.ChromLowCVSLim = "" Then
                                                            If IsNumeric(aCompound.Conc) Then
                                                                If CDbl(aCompound.Conc) >= CDbl(aCompound.ChromLowContLim) And CDbl(aCompound.Conc) <= CDbl(aCompound.ChromUpContLim) Then
                                                                    aCompound.ChromCVSPass = True
                                                                End If
                                                            Else
                                                                aCompound.ChromCVSPass = False
                                                            End If
                                                        Else
                                                            If IsNumeric(aCompound.Conc) Then
                                                                If CDbl(aCompound.Conc) >= CDbl(aCompound.ChromLowCVSLim) And CDbl(aCompound.Conc) <= CDbl(aCompound.ChromUpCVSLim) Then
                                                                    aCompound.ChromCVSPass = True
                                                                End If
                                                            Else
                                                                aCompound.ChromCVSPass = False
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Next
                                        'RPD
                                        For Each aSample In GlobalVariables.ReportSamList
                                            If Not aSample.Calculated Then
                                                For Each aSample2 In GlobalVariables.ReportSamList
                                                    If InStr(aSample2.Name, "DUP") And aSample.Name = Trim(aSample2.Name.Substring(0, aSample2.Name.Length - 3)) And aSample.Name <> aSample2.Name Then
                                                        For Each aCompound In aSample.CompoundList
                                                            For Each aCompound2 In aSample2.CompoundList
                                                                If aCompound.Name = aCompound2.Name Then
                                                                    If IsNumeric(aCompound.Conc) And IsNumeric(aCompound2.Conc) Then
                                                                        aCompound.ChromRPD = Math.Abs(CDbl(aCompound.Conc) - CDbl(aCompound2.Conc)) / (CDbl(aCompound.Conc) + CDbl(aCompound2.Conc) / 2) * 100
                                                                        aCompound2.ChromRPD = aCompound.ChromRPD
                                                                    Else
                                                                        aCompound.ChromRPD = "N/A"
                                                                        aCompound2.ChromRPD = aCompound.ChromRPD
                                                                    End If
                                                                End If
                                                            Next
                                                        Next
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox("Error occured during calculation!" & vbCrLf & _
                  "Sub Procedure: FreeportChrom()" & vbCrLf & _
                  "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try

        'Set samples to calculated
        If Not blnTransfer Then
            For Each aSample In GlobalVariables.SampleList
                aSample.Calculated = True
            Next

            GlobalVariables.NeedsCalculation = False
        End If

        Return True

    End Function

    'Midland Chrom reported Amount calculations
    Function MidlandChrom(ByVal strLimitType As String, ByVal strSelectLim As String, ByVal strSISLoc As String, ByVal blnTransfer As Boolean) As Boolean
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aInstrument As mInstrument
        Dim aSample As Sample
        Dim aSample2 As Sample
        Dim aDupSample As Sample
        Dim aQCSample As Sample
        Dim aOGSample As Sample
        Dim aCompound As Compound
        Dim aCompound2 As Compound
        Dim amCompound As mCompound
        Dim blnMatch As Boolean
        Dim blnIndirect As Boolean
        Dim aSurrogate As Surrogate
        Dim amSurrogate As mSurrogate
        Dim flgSIS As Boolean

        Try

            'Import SIS
            If strSISLoc <> "" Then
                flgSIS = False
                For Each aSIS In GlobalVariables.SISList
                    If InStr(strSISLoc, aSIS.ProjNum) Then
                        flgSIS = True
                    End If
                Next
                If Not flgSIS Then
                    If Not GlobalVariables.Import.SISImport(strSISLoc) Then
                        Return False
                    End If
                End If
            End If
            'UnitConversion.ShowDialog()
            blnIndirect = False
            For Each aPermit In GlobalVariables.PermitList
                If aPermit.Name = GlobalVariables.selPermit.Name Then
                    For Each aProject In aPermit.ProjectList
                        If aProject.Name = GlobalVariables.selProject Then
                            For Each aInstrument In aProject.mInstrumentList
                                If aInstrument.Name = GlobalVariables.selInstrument Then
                                    'Load limits if not 
                                    If strSelectLim <> "" Then
                                        If Not GlobalVariables.Import.MidlandChromBuildRecLimits(strSelectLim) Then
                                            If Not MsgBox("Error occured during Recovery Limit Import! Continue processing without limits?" & vbCrLf & "Sub Procedure: MidlandChrom()" & MsgBoxStyle.YesNo) Then
                                                Return False
                                            End If
                                        End If
                                    Else
                                        'Assign recovery limits if noncomplance
                                        For Each aSample In GlobalVariables.ReportSamList
                                            For Each aSurrogate In aSample.SurrogateList
                                                For Each amSurrogate In aInstrument.mSurrogateList
                                                    If aSurrogate.Name = amSurrogate.Name Then
                                                        aSurrogate.ChromLowContLim = amSurrogate.RecLLim
                                                        aSurrogate.ChromUpContLim = amSurrogate.RecULim
                                                        aSurrogate.ChromLowLCSLim = amSurrogate.LCSLLim
                                                        aSurrogate.ChromUpLCSLim = amSurrogate.LCSULim
                                                        aSurrogate.ChromLowMSLim = amSurrogate.MSLLim
                                                        aSurrogate.ChromUpMSLim = amSurrogate.MSULim
                                                    End If
                                                Next
                                            Next
                                            For Each aCompound In aSample.CompoundList
                                                For Each amCompound In aInstrument.mCompoundList
                                                    If aCompound.Name = amCompound.Name Then
                                                        aCompound.ChromLowContLim = amCompound.RecLLim
                                                        aCompound.ChromUpContLim = amCompound.RecULim
                                                        aCompound.ChromLowLCSLim = amCompound.LCSLLim
                                                        aCompound.ChromUpLCSLim = amCompound.LCSULim
                                                        aCompound.ChromLowMSLim = amCompound.MSLLim
                                                        aCompound.ChromUpMSLim = amCompound.MSULim
                                                    End If
                                                Next
                                            Next
                                        Next
                                    End If


                                    For Each aSample In GlobalVariables.ReportSamList
                                        'Deal with units
                                        aSample.ReportedUnits = aSample.Units
                                        'Compounds
                                        If Not aSample.Calculated Then
                                            'Calculate Surrogate Recovery
                                            For Each aSurrogate In aSample.SurrogateList
                                                If aSurrogate.Recovery = "NA" Then
                                                    For Each amSurrogate In aInstrument.mSurrogateList
                                                        If aSurrogate.Name = amSurrogate.Name Then
                                                            aSurrogate.Recovery = CStr((CDbl(aSurrogate.Conc) / CDbl(amSurrogate.Conc)) * 100)
                                                            aSurrogate.ChromLowContLim = amSurrogate.RecLLim
                                                            aSurrogate.ChromUpContLim = amSurrogate.RecULim
                                                        End If
                                                    Next
                                                End If
                                            Next
                                            For Each aCompound In aSample.CompoundList
                                                'Determine compound and mcompound relationship
                                                blnMatch = False

                                                For Each amCompound In aInstrument.mCompoundList
                                                    If aCompound.Name = amCompound.Name Then
                                                        blnMatch = True
                                                    ElseIf UCase(aCompound.Name) = UCase(amCompound.Name) Then
                                                        If blnTransfer Then
                                                            If blnIndirect Then
                                                                blnMatch = True
                                                            Else
                                                                If vbYes = MsgBox("Compound: " & aCompound.Name & " is an indirect match for LIMS Component: " & amCompound.Name & ". Do you want to continue?", MsgBoxStyle.YesNo, "eTrain 2.0") Then
                                                                    If vbYes = MsgBox("Would you like to continue transfer including all indirect matches?", MsgBoxStyle.YesNo, "eTrain 2.0") Then
                                                                        blnIndirect = True
                                                                    End If
                                                                    blnMatch = True
                                                                Else
                                                                    MsgBox("Please update Method file so Compound: " & aCompound.Name & " matches LIMS Component: " & amCompound.Name & " exactly and try transfer again.", MsgBoxStyle.Exclamation, "eTrain 2.0")
                                                                    Return False
                                                                End If
                                                            End If
                                                        Else
                                                            blnMatch = True
                                                        End If
                                                    ElseIf amCompound.AliasList.Count > 0 Then
                                                        If amCompound.DetermineMatch(aCompound) Then
                                                            blnMatch = True
                                                        End If
                                                    End If
                                                    If blnMatch Then
                                                        If strLimitType = "MDL" Then
                                                            aCompound.ChromReportLimit = amCompound.MDL
                                                            aCompound.ChromAdjustedLimit = CDbl(amCompound.MDL) * CDbl(aSample.DilutionFactor)
                                                            If IsNumeric(aCompound.Conc) Then
                                                                aCompound.ChromAdjustConc = CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)
                                                                If CDbl(aCompound.ChromAdjustConc) > CDbl(aCompound.ChromAdjustedLimit) Then
                                                                    aCompound.ReportedAmt = aCompound.ChromAdjustConc
                                                                Else
                                                                    aCompound.ReportedAmt = "ND"
                                                                End If
                                                            Else
                                                                aCompound.ReportedAmt = "ND"
                                                            End If
                                                            'Format
                                                            'aCompound.ReportedAmt = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
                                                        ElseIf strLimitType = "PQL" Then
                                                            aCompound.ChromReportLimit = amCompound.PQL
                                                            aCompound.ChromAdjustedLimit = CDbl(amCompound.PQL) * CDbl(aSample.DilutionFactor)

                                                            If IsNumeric(aCompound.Conc) Then
                                                                aCompound.ChromAdjustConc = CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)
                                                                If CDbl(aCompound.ChromAdjustConc) > CDbl(aCompound.ChromAdjustedLimit) Then
                                                                    aCompound.ReportedAmt = aCompound.ChromAdjustConc
                                                                Else
                                                                    aCompound.ReportedAmt = "ND"
                                                                End If
                                                            Else
                                                                aCompound.ReportedAmt = "ND"
                                                            End If
                                                            'aCompound.ReportedAmt = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
                                                        ElseIf strLimitType = "RL" Then
                                                            If IsNumeric(aCompound.Conc) Then
                                                                aCompound.ChromAdjustConc = CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)
                                                                If CDbl(aCompound.ChromAdjustConc) > CDbl(amCompound.RL) Then
                                                                    aCompound.ReportedAmt = aCompound.ChromAdjustConc
                                                                Else
                                                                    aCompound.ReportedAmt = "ND"
                                                                End If
                                                            Else
                                                                aCompound.ReportedAmt = "ND"
                                                            End If
                                                            'aCompound.ReportedAmt = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
                                                        ElseIf strLimitType = "N/A" Then
                                                            aCompound.ChromAdjustConc = CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)
                                                            aCompound.ReportedAmt = aCompound.ChromAdjustConc
                                                            'aCompound.ReportedAmt = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
                                                        End If
                                                        blnMatch = False
                                                    End If
                                                Next
                                            Next

                                        End If
                                    Next

                                    'Only Run this section if not transfer
                                    If Not blnTransfer Then
                                        'Calculate MS Recoveries & RPD
                                        For Each aQCSample In GlobalVariables.ReportSamList
                                            If Not aQCSample.Calculated And aQCSample.Include Then
                                                If aQCSample.Type = "MS" And GlobalVariables.Report.Type = "MS" Then
                                                    'Find regular sample
                                                    For Each aOGSample In GlobalVariables.ReportSamList
                                                        If InStr(aQCSample.Name, aOGSample.Name) And aOGSample.Type = "SAMPLE" And Not aQCSample.SpikeCalculated Then
                                                            'MS Recovery
                                                            With SpikeInfo
                                                                .lblSampleName.Text = "Sample: " & aQCSample.Name
                                                                .lblDil.Text = "Dilution Factor: " & aQCSample.DilutionFactor
                                                                .lblUnits.Text = "Ending Conversion Units: " & aQCSample.Units
                                                                .txtConc.Clear()
                                                                .txtVol.Clear()
                                                            End With
                                                            SpikeInfo.ShowDialog()
                                                            For Each aCompound In aQCSample.CompoundList
                                                                If IsNumeric(aCompound.Conc) Then
                                                                    aCompound.ChromCorrectedSpike = aQCSample.ChromSpikeAmt
                                                                    aCompound.ChromSpikeRecovery = CStr((CDbl(aCompound.Conc) / CDbl(aCompound.ChromCorrectedSpike)) * 100)
                                                                    If IsNumeric(aCompound.ChromLowContLim) And IsNumeric(aCompound.ChromUpContLim) Then
                                                                        If CDbl(aCompound.ChromSpikeRecovery) >= CDbl(aCompound.ChromLowMSLim) And CDbl(aCompound.ChromSpikeRecovery) <= CDbl(aCompound.ChromUpMSLim) Then
                                                                            aCompound.ChromSpikePass = True
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next
                                                            'See if MSD sample
                                                            For Each aDupSample In GlobalVariables.ReportSamList
                                                                If aDupSample.Type = "MSD" And InStr(aDupSample.Name, aOGSample.Name) Then
                                                                    'MS MSD RPD
                                                                    For Each aCompound In aQCSample.CompoundList
                                                                        For Each aCompound2 In aDupSample.CompoundList
                                                                            If aCompound.Name = aCompound2.Name Then
                                                                                If IsNumeric(aCompound.Conc) And IsNumeric(aCompound2.Conc) Then
                                                                                    aCompound.ChromRPD = Math.Abs(CDbl(aCompound.Conc) - CDbl(aCompound2.Conc)) / (CDbl(aCompound.Conc) + CDbl(aCompound2.Conc) / 2) * 100
                                                                                    aCompound2.ChromRPD = aCompound.ChromRPD
                                                                                    aCompound.ChromRPDLimit = "30"
                                                                                    aCompound2.ChromRPDLimit = "30"
                                                                                Else
                                                                                    aCompound.ChromRPD = "N/A"
                                                                                    aCompound2.ChromRPD = aCompound.ChromRPD
                                                                                    aCompound.ChromRPDLimit = "30"
                                                                                    aCompound2.ChromRPDLimit = "30"
                                                                                End If
                                                                            End If
                                                                        Next
                                                                    Next
                                                                    'MSD Recovery
                                                                    With SpikeInfo
                                                                        .lblSampleName.Text = "Sample: " & aDupSample.Name
                                                                        .lblDil.Text = "Dilution Factor: " & aDupSample.DilutionFactor
                                                                        .lblUnits.Text = "Ending Conversion Units: " & aDupSample.Units
                                                                        .txtConc.Clear()
                                                                        .txtVol.Clear()
                                                                    End With
                                                                    SpikeInfo.ShowDialog()
                                                                    For Each aCompound In aDupSample.CompoundList
                                                                        If IsNumeric(aCompound.Conc) Then
                                                                            aCompound.ChromCorrectedSpike = aDupSample.ChromSpikeAmt
                                                                            aCompound.ChromSpikeRecovery = CStr((CDbl(aCompound.Conc) / CDbl(aCompound.ChromCorrectedSpike)) * 100)
                                                                            If IsNumeric(aCompound.ChromLowContLim) And IsNumeric(aCompound.ChromUpContLim) Then
                                                                                If CDbl(aCompound.ChromSpikeRecovery) >= CDbl(aCompound.ChromLowMSLim) And CDbl(aCompound.ChromSpikeRecovery) <= CDbl(aCompound.ChromUpMSLim) Then
                                                                                    aCompound.ChromSpikePass = True
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Next
                                                                End If
                                                            Next
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Next
                                        'Calculate LCS Recoveries
                                        For Each aSample In GlobalVariables.ReportSamList
                                            If Not aSample.Calculated And aSample.Include Then
                                                If aSample.Type = "LCS" Or aSample.Type = "LCSD" And GlobalVariables.Report.Type = "LCS" And Not aSample.SpikeCalculated Then
                                                    With SpikeInfo
                                                        .lblSampleName.Text = "Sample: " & aSample.Name
                                                        .lblDil.Text = "Dilution Factor: " & aSample.DilutionFactor
                                                        .lblUnits.Text = "Ending Conversion Units: " & aSample.Units
                                                        .txtConc.Clear()
                                                        .txtVol.Clear()
                                                    End With
                                                    SpikeInfo.ShowDialog()
                                                    For Each aCompound In aSample.CompoundList
                                                        If IsNumeric(aCompound.Conc) Then
                                                            aCompound.ChromCorrectedSpike = aSample.ChromSpikeAmt
                                                            aCompound.ChromSpikeRecovery = CStr((CDbl(aCompound.Conc) / CDbl(aCompound.ChromCorrectedSpike)) * 100)
                                                            If IsNumeric(aCompound.ChromLowLCSLim) And IsNumeric(aCompound.ChromUpLCSLim) Then
                                                                If CDbl(aCompound.ChromSpikeRecovery) >= CDbl(aCompound.ChromLowLCSLim) And CDbl(aCompound.ChromSpikeRecovery) <= CDbl(aCompound.ChromUpLCSLim) Then
                                                                    aCompound.ChromSpikePass = True
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Next
                                        'ICV/CVS
                                        For Each aSample In GlobalVariables.ReportSamList
                                            If Not aSample.Calculated And aSample.Include Then
                                                If aSample.Type = "ICV" Then
                                                    For Each aCompound In aSample.CompoundList
                                                        If aCompound.ChromLowICVLim = "" Then
                                                            If IsNumeric(aCompound.Conc) Then
                                                                If CDbl(aCompound.Conc) >= CDbl(aCompound.ChromLowContLim) And CDbl(aCompound.Conc) <= CDbl(aCompound.ChromUpContLim) Then
                                                                    aCompound.ChromICVPass = True
                                                                End If
                                                            Else
                                                                aCompound.ChromICVPass = False
                                                            End If
                                                        Else
                                                            If IsNumeric(aCompound.Conc) Then
                                                                If CDbl(aCompound.Conc) >= CDbl(aCompound.ChromLowICVLim) And CDbl(aCompound.Conc) <= CDbl(aCompound.ChromUpICVLim) Then
                                                                    aCompound.ChromICVPass = True
                                                                End If
                                                            Else
                                                                aCompound.ChromICVPass = False
                                                            End If
                                                        End If
                                                    Next
                                                ElseIf aSample.Type = "CVS" Then
                                                    For Each aCompound In aSample.CompoundList
                                                        If aCompound.ChromLowCVSLim = "" Then
                                                            If IsNumeric(aCompound.Conc) Then
                                                                If CDbl(aCompound.Conc) >= CDbl(aCompound.ChromLowContLim) And CDbl(aCompound.Conc) <= CDbl(aCompound.ChromUpContLim) Then
                                                                    aCompound.ChromCVSPass = True
                                                                End If
                                                            Else
                                                                aCompound.ChromCVSPass = False
                                                            End If
                                                        Else
                                                            If IsNumeric(aCompound.Conc) Then
                                                                If CDbl(aCompound.Conc) >= CDbl(aCompound.ChromLowCVSLim) And CDbl(aCompound.Conc) <= CDbl(aCompound.ChromUpCVSLim) Then
                                                                    aCompound.ChromCVSPass = True
                                                                End If
                                                            Else
                                                                aCompound.ChromCVSPass = False
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Next
                                        'RPD
                                        For Each aSample In GlobalVariables.ReportSamList
                                            If Not aSample.Calculated And aSample.Include Then
                                                For Each aSample2 In GlobalVariables.SampleList
                                                    If InStr(UCase(aSample2.Name), "DUP") Then
                                                        If aSample.Name = Trim(aSample2.Name.Substring(0, aSample2.Name.Length - 3)) And aSample.Name <> aSample2.Name Then
                                                            For Each aCompound In aSample.CompoundList
                                                                For Each aCompound2 In aSample2.CompoundList
                                                                    If aCompound.Name = aCompound2.Name Then
                                                                        If IsNumeric(aCompound.Conc) And IsNumeric(aCompound2.Conc) Then
                                                                            aCompound.ChromRPD = Math.Abs(CDbl(aCompound.Conc) - CDbl(aCompound2.Conc)) / ((CDbl(aCompound.Conc) + CDbl(aCompound2.Conc)) / 2) * 100
                                                                            aCompound2.ChromRPD = aCompound.ChromRPD
                                                                            aCompound.ChromRPDLimit = "30"
                                                                            aCompound2.ChromRPDLimit = "30"
                                                                        Else
                                                                            aCompound.ChromRPD = "N/A"
                                                                            aCompound2.ChromRPD = aCompound.ChromRPD
                                                                            aCompound.ChromRPDLimit = "30"
                                                                            aCompound2.ChromRPDLimit = "30"
                                                                        End If
                                                                    End If
                                                                Next
                                                            Next
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If

                                End If
                            Next
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox("Error occured during calculation!" & vbCrLf & _
                  "Sub Procedure: MidlandChrom()" & vbCrLf & _
                  "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        'Reset Spike Calc
        For Each aSample In GlobalVariables.ReportSamList
            aSample.SpikeCalculated = False
        Next
        

        Return True

    End Function

    Function MidlandHR(ByVal strSISLoc As String) As Boolean
        Dim aSample As Sample
        Dim aSISSample As Sample
        Dim aSIS As SIS
        Dim aCS1Sample As Sample
        Dim aLCSSample As Sample
        Dim intCalSampleCount As Integer
        'Dim aStandard As Standard
        'Dim aSwapCompound As Compound
        'Dim aInjCompound As Compound
        'Dim aCompound1 As Compound
        'Dim aCompound2 As Compound
        Dim aCompound As Compound
        'Dim aCS1Compound1 As Compound
        'Dim aCS1Compound2 As Compound
        'Dim aTCompound As Compound
        'Dim aInstrument As mInstrument
        'Dim amCompound1 As mCompound
        Dim amCompound As mCompound
        'Dim amCompound2 As mCompound
        'Dim amStandard As mStandard
        'Dim amInjStandard As mStandard
        Dim flg As Boolean
        Dim flgSIS As Boolean
        'Dim strFlags As String
        Dim strSpikes() As String

        'Import SIS
        flgSIS = False
        For Each aSIS In GlobalVariables.SISList
            If InStr(strSISLoc, aSIS.ProjNum) Then
                flgSIS = True
            End If
        Next
        If Not flgSIS Then
            If Not GlobalVariables.Import.SISImport(strSISLoc) Then
                Return False
            End If
        End If

        Try
            'Set Aliquot and spike amounts for each sample from SIS
            For Each aSample In GlobalVariables.SampleList
                aSIS = GlobalVariables.SISList(0)
                For Each aSISSample In aSIS.SampleList
                    If aSample.LimsID = aSISSample.SISLabNum Then
                        If aSISSample.SISDefaultAliquot <> "" Then
                            aSample.Aliquot = aSISSample.SISDefaultAliquot
                        End If
                        If aSISSample.SISSpikeMult <> "" Then
                            strSpikes = aSISSample.SISSpikeMult.Split(",")
                            aSample.StdSpikeAmt = Trim(strSpikes(0))
                            aSample.InjSpikeAmt = Trim(strSpikes(1))
                            aSample.LCSSpikeAmt = Trim(strSpikes(2))
                        End If
                    End If
                Next
            Next

        Catch ex As Exception
            MsgBox("Error assigning Sample values from SIS!" & vbCrLf & _
            "Sub Procedure: Calculations.MidlandHR()" & vbCrLf & _
            "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        'Get CS1
        aCS1Sample = Nothing
        For Each aSample In GlobalVariables.SampleList
            If InStr(aSample.DataFile, "CS1") Then
                aCS1Sample = aSample
            End If
        Next
        'No CS1 found?
        If IsNothing(aCS1Sample) Then
            MsgBox("No CS1 Found or Selected!" & vbCrLf & _
                    "Sub Procedure: Calculations.MidlandHR()", MsgBoxStyle.Critical)
            Return False
        End If
        aSample = Nothing
        'Get LCS Sample
        aLCSSample = Nothing
        For Each aSample In GlobalVariables.SampleList
            If InStr(aSample.DataFile, "P") Then
                aLCSSample = aSample
            End If
        Next
        'No LCS found?
        If IsNothing(aLCSSample) Then
            If MsgBox("No LCS Found! Continue processing?" & vbCrLf & _
                    "Sub Procedure: Calculations.MidlandHR()", MsgBoxStyle.YesNo) = vbNo Then
                Return False
            End If
        End If
        aSample = Nothing

        'Calc Average QM Area
        intCalSampleCount = 0
        For Each aSample In GlobalVariables.SampleList
            'Get List created and begin averaging area
            If aSample.Type = "TEFCAL" Then
                intCalSampleCount = intCalSampleCount + 1
                For Each aCompound In aSample.CompoundList
                    flg = False
                    If GlobalVariables.MidlandHRAvgAreaCompList.Count = 0 Then
                        amCompound = New mCompound
                        amCompound.Name = aCompound.Name
                        amCompound.AvgArea = aCompound.TQ3QMArea
                        GlobalVariables.MidlandHRAvgAreaCompList.Add(amCompound)
                    Else
                        For Each amCompound In GlobalVariables.MidlandHRAvgAreaCompList
                            If amCompound.Name = aCompound.Name Then
                                flg = True
                                amCompound.AvgArea = CStr(CDbl(amCompound.AvgArea) + CDbl(aCompound.TQ3QMArea))
                            End If
                        Next
                        If Not flg Then
                            amCompound = New mCompound
                            amCompound.Name = aCompound.Name
                            amCompound.AvgArea = aCompound.TQ3QMArea
                            GlobalVariables.MidlandHRAvgAreaCompList.Add(amCompound)
                        End If
                    End If
                Next
            End If
        Next
        For Each aSample In GlobalVariables.SampleList
            If aSample.Type = "TEFCAL" Then
                For Each aCompound In aSample.CompoundList
                    For Each amCompound In GlobalVariables.MidlandHRAvgAreaCompList
                        If aCompound.Name = amCompound.Name Then
                            aCompound.TQ3QMAreaAvg = CStr(CDbl(amCompound.AvgArea) / intCalSampleCount)
                            Exit For
                        End If
                    Next
                Next

            End If
        Next

    End Function


    'Midland FAST Calculations happen here
    Function MidlandFAST(ByVal strSISLoc As String) As Boolean
        Dim aSample As Sample
        Dim aSISSample As Sample
        Dim aSIS As SIS
        Dim aCS1Sample As Sample
        Dim aLCSSample As Sample
        Dim aCS3Sample As Sample
        Dim aStandard As Standard
        Dim aSwapCompound As Compound
        Dim aInjCompound As Compound
        Dim aCompound1 As Compound
        Dim aCompound2 As Compound
        Dim aCompound As Compound
        Dim aCS1Compound1 As Compound
        Dim aCS1Compound2 As Compound
        Dim aTCompound As Compound
        Dim aInstrument As mInstrument
        Dim amCompound1 As mCompound
        Dim amCompound As mCompound
        Dim amCompound2 As mCompound
        Dim amStandard As mStandard
        Dim amInjStandard As mStandard
        Dim flg As Integer
        Dim flgSIS As Boolean
        Dim strFlags As String
        Dim strSpikes() As String


        'Import SIS
        flgSIS = False
        For Each aSIS In GlobalVariables.SISList
            If InStr(strSISLoc, aSIS.ProjNum) Then
                flgSIS = True
            End If
        Next
        If Not flgSIS Then
            If Not GlobalVariables.Import.SISImport(strSISLoc) Then
                Return False
            End If
        End If

        Try
            'Set Aliquot and spike amounts for each sample from SIS
            For Each aSample In GlobalVariables.ReportSamList
                aSIS = GlobalVariables.SISList(0)
                For Each aSISSample In aSIS.SampleList
                    If aSample.LimsID = aSISSample.SISLabNum Then
                        If aSISSample.SISDefaultAliquot <> "" Then
                            aSample.Aliquot = aSISSample.SISDefaultAliquot
                        End If
                        If aSISSample.SISSpikeMult <> "" Then
                            strSpikes = aSISSample.SISSpikeMult.Split(",")
                            aSample.StdSpikeAmt = Trim(strSpikes(0))
                            aSample.InjSpikeAmt = Trim(strSpikes(1))
                            aSample.LCSSpikeAmt = Trim(strSpikes(2))
                        End If
                    End If
                Next
            Next

        Catch ex As Exception
            MsgBox("Error assigning Sample values from SIS!" & vbCrLf & _
            "Sub Procedure: MidlandFAST()" & vbCrLf & _
            "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try


        'Generate Compound list of names for each sample
        For Each aSample In GlobalVariables.ReportSamList
            For Each aCompound1 In aSample.CompoundList
                flg = 0
                For Each aCompound2 In aSample.CompoundNameList
                    If Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) = aCompound2.Name Then
                        flg = 1
                    End If
                Next
                If flg = 0 Then
                    aCompound = New Compound
                    aCompound.Name = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1))
                    aSample.CompoundNameList.Add(aCompound)
                    aCompound = Nothing
                End If
            Next
        Next

        'Get CS1(LOQ)
        aCS1Sample = Nothing
        For Each aSample In GlobalVariables.ReportSamList
            If InStr(aSample.DataFile, "CS1") Then
                aCS1Sample = aSample
            End If
        Next
        'No CS1 found?
        If IsNothing(aCS1Sample) Then
            MsgBox("No CS1 Found or Selected!" & vbCrLf & _
                    "Sub Procedure: Calculations.MidlandFAST()", MsgBoxStyle.Critical)
            Return False
        End If
        aSample = Nothing
        'Get LCS Sample
        aLCSSample = Nothing
        For Each aSample In GlobalVariables.ReportSamList
            If InStr(aSample.DataFile, "P") Then
                aLCSSample = aSample
            End If
        Next
        'No LCS found?
        If IsNothing(aLCSSample) Then
            If MsgBox("No LCS Found! Continue processing?" & vbCrLf & _
                    "Sub Procedure: Calculations.MidlandFAST()", MsgBoxStyle.YesNo) = vbNo Then
                Return False
            End If
        End If
        aSample = Nothing
        'Get CS3
        aCS3Sample = Nothing
        For Each aSample In GlobalVariables.ReportSamList
            If InStr(aSample.DataFile, "CS3") Then
                aCS3Sample = aSample
            End If
        Next
        'No CS3 found?
        If IsNothing(aCS3Sample) Then
            If MsgBox("No CS3 Found! Continue processing?" & vbCrLf & _
                    "Sub Procedure: Calculations.MidlandFAST()", MsgBoxStyle.YesNo) = vbNo Then
                Return False
            End If
        End If
        aSample = Nothing

        Try

            'Determine what is considered (1) and (2) for compounds
            For Each aSample In GlobalVariables.ReportSamList
                For Each aCompound1 In aSample.CompoundList
                    For Each aCompound2 In aSample.CompoundList
                        If Trim(aCompound2.Name.Substring(0, InStrRev(aCompound2.Name, "(") - 1)) = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) And aCompound1.Name <> aCompound2.Name And aCompound1.MidFIsTarget = False And aCompound2.MidFIsTarget = False Then
                            For Each aInstrument In GlobalVariables.selMethod.mInstrumentList
                                For Each amCompound1 In aInstrument.mCompoundList
                                    If amCompound1.Name = aCompound1.Name Then
                                        If aCompound1.QIon = amCompound1.Ion Then
                                            aCompound1.MidFIsTarget = True
                                            aCompound2.MidFIsQual = True
                                        ElseIf aCompound2.QIon = amCompound1.Ion Then
                                            aCompound1.MidFIsQual = True
                                            aCompound2.MidFIsTarget = True
                                        End If
                                        Exit For
                                    End If
                                Next
                            Next
                        End If
                    Next
                Next
            Next

            'Build Theoretical limits
            If Not GlobalVariables.Calculations.BuildTheoComps() Then
                Return False
            End If

            'Calculate Ion Ratios
            For Each aSample In GlobalVariables.ReportSamList
                'Standards
                If Not aSample.Calculated Then
                    For Each aStandard In aSample.InternalStdList
                        For Each aTCompound In GlobalVariables.TheoComps
                            If aStandard.Name = aTCompound.Name Then
                                aStandard.MidFIonRatio = (CDbl(aStandard.Q1Resp / aStandard.Response) * 100)
                                'Determine if Ion Ratio is between 85 and 115%
                                If aStandard.MidFIonRatio >= CDbl(aTCompound.MidFLLim) And aStandard.MidFIonRatio <= CDbl(aTCompound.MidFULim) Then
                                    aStandard.MidFIonRatioInLim = True
                                    aStandard.MidFQC6 = False
                                Else
                                    aStandard.MidFIonRatioInLim = False
                                    aStandard.MidFQC6 = True
                                End If
                                Exit For
                            End If
                        Next
                        'Handle QC Flag 7
                        If aStandard.MI Then
                            aStandard.MidFQC7 = True
                        Else
                            aStandard.MidFQC7 = False
                        End If
                    Next
                    'Injection and Compounds
                    For Each aCompound1 In aSample.CompoundList
                        If InStr(aCompound1.Name, "(INJ)") Then
                            For Each aTCompound In GlobalVariables.TheoComps
                                If aTCompound.Name = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) Then
                                    aCompound1.MidFIonRatio = (CDbl(aCompound1.Q1Resp / aCompound1.Response) * 100)
                                    'Determine if Ion Ratio is between 85 and 115%
                                    If aCompound1.MidFIonRatio >= CDbl(aTCompound.MidFLLim) And aCompound1.MidFIonRatio <= CDbl(aTCompound.MidFULim) Then
                                        aCompound1.MidFIonRatioInLim = True
                                        aCompound1.MidFQC1 = True
                                    Else
                                        aCompound1.MidFIonRatioInLim = False
                                        aCompound1.MidFQC2 = True
                                    End If
                                    Exit For
                                End If
                            Next
                        End If
                        For Each aCompound2 In aSample.CompoundList
                            If Trim(aCompound2.Name.Substring(0, InStrRev(aCompound2.Name, "(") - 1)) = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) And aCompound1.Name <> aCompound2.Name And aCompound1.MidFIonRatio = -1.0 And aCompound2.MidFIonRatio = -1.0 Then
                                For Each aTCompound In GlobalVariables.TheoComps
                                    If aTCompound.Name = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) Then
                                        If Math.Round(CDbl(aCompound1.TSignal)) = CDbl(aTCompound.TSignal) And aCompound1.MidFIonRatio = -1 Then
                                            aCompound1.MidFIsTarget = True
                                            aCompound2.MidFIsQual = True
                                            If CDbl(aCompound1.Response) = 0 Or CDbl(aCompound2.Response) = 0 Then
                                                aCompound1.MidFIonRatio = 0
                                                aCompound2.MidFIonRatio = aCompound1.MidFIonRatio
                                                aCompound1.MidFNonDetect = True
                                                aCompound2.MidFNonDetect = True
                                            Else
                                                aCompound1.MidFIonRatio = (CDbl(aCompound2.Response / aCompound1.Response) * 100)
                                                aCompound2.MidFIonRatio = aCompound1.MidFIonRatio
                                            End If
                                        Else
                                            aCompound1.MidFIsQual = True
                                            aCompound2.MidFIsTarget = True
                                            If CDbl(aCompound1.Response) = 0 Or CDbl(aCompound2.Response) = 0 Then
                                                aCompound1.MidFIonRatio = 0
                                                aCompound2.MidFIonRatio = aCompound1.MidFIonRatio
                                                aCompound1.MidFNonDetect = True
                                                aCompound2.MidFNonDetect = True
                                            Else
                                                aCompound1.MidFIonRatio = (CDbl(aCompound1.Response / aCompound2.Response) * 100)
                                                aCompound2.MidFIonRatio = aCompound1.MidFIonRatio
                                            End If
                                        End If
                                        'Determine if Ion Ratio is between 85 and 115%
                                        If aCompound1.MidFIonRatio >= CDbl(aTCompound.MidFLLim) And aCompound1.MidFIonRatio <= CDbl(aTCompound.MidFULim) Then
                                            aCompound1.MidFIonRatioInLim = True
                                            aCompound2.MidFIonRatioInLim = True
                                            aCompound1.MidFQC1 = True
                                            aCompound2.MidFQC1 = True
                                        Else
                                            aCompound1.MidFIonRatioInLim = False
                                            aCompound2.MidFIonRatioInLim = False
                                            aCompound1.MidFQC2 = True
                                            aCompound2.MidFQC2 = True
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox("Error calculating Ion Ratios!" & vbCrLf & _
                  "Sub Procedure: MidlandFAST()" & vbCrLf & _
                  "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        Try
            'Calculate LOQ Amt
            For Each aSample In GlobalVariables.ReportSamList
                'Ignore CS1 , CS3, S1
                If InStr(aSample.DataFile.Substring(0, 3), "CS1") Or InStr(aSample.DataFile.Substring(0, 3), "CS3") Or InStr(aSample.DataFile.Substring(0, 2), "S1") Then
                    'LOQ for CS1 is just the concentration amt from instrument
                    For Each aCompound In aSample.CompoundList
                        If InStr(aCompound.Name, "(INJ)") Then
                            aCompound.MidFLoq = 0
                        Else
                            aCompound.MidFLoq = aCompound.Conc
                        End If
                    Next
                    For Each aCompound1 In aSample.CompoundList
                        'Get second ion
                        For Each aCompound2 In aSample.CompoundList
                            'If match, find CS1 ions
                            If Trim(aCompound2.Name.Substring(0, InStrRev(aCompound2.Name, "(") - 1)) = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) And _
                                aCompound1.Name <> aCompound2.Name Then
                                If aCompound1.MidFIsQual Then
                                    aSwapCompound = aCompound1
                                    aCompound1 = aCompound2
                                    aCompound2 = aSwapCompound
                                End If
                                'While I am here, deal with some QC Flags (4,5,8,9)
                                If aCompound1.MI Then
                                    aCompound1.MidFQC8 = True
                                End If
                                If aCompound2.MI Then
                                    aCompound2.MidFQC9 = True
                                End If
                            End If
                        Next
                    Next
                Else
                    'Get compound
                    If Not aSample.Calculated Then
                        For Each aCompound1 In aSample.CompoundList
                            'Get second ion
                            For Each aCompound2 In aSample.CompoundList
                                'If match, find CS1 ions
                                If Trim(aCompound2.Name.Substring(0, InStrRev(aCompound2.Name, "(") - 1)) = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) And _
                                    aCompound1.Name <> aCompound2.Name And aCompound1.MidFLoq = -1 And aCompound2.MidFLoq = -1 Then
                                    'Make sure aCompound1 is target ion, and second is qual ion
                                    If aCompound1.MidFIsQual Then
                                        aSwapCompound = aCompound1
                                        aCompound1 = aCompound2
                                        aCompound2 = aSwapCompound
                                    End If
                                    For Each aCS1Compound1 In aCS1Sample.CompoundList
                                        If aCS1Compound1.Name = aCompound1.Name Then
                                            For Each aCS1Compound2 In aCS1Sample.CompoundList
                                                If aCS1Compound2.Name = aCompound2.Name Then
                                                    'CS1 ions found, get standard
                                                    For Each aStandard In aSample.InternalStdList
                                                        'Search for corresponding standard and 13c
                                                        If aStandard.Name = aCompound1.MidF13CAssoc Then
                                                            'Now get method information
                                                            For Each aInstrument In GlobalVariables.selMethod.mInstrumentList
                                                                'Get instrument
                                                                If GlobalVariables.Report.Instrument = aInstrument.Name Then
                                                                    For Each amCompound1 In aInstrument.mCompoundList
                                                                        'Get compound
                                                                        If amCompound1.Name = aCompound1.Name Then
                                                                            For Each amCompound2 In aInstrument.mCompoundList
                                                                                If amCompound2.Name = aCompound2.Name Then
                                                                                    'Calculate amounts for each ion
                                                                                    aCompound1.MidFLoq = (CDbl(aCS1Compound1.Response) / CDbl(aStandard.Response)) * (CDbl(aStandard.Conc) / CDbl(amCompound1.RRF))
                                                                                    aCompound2.MidFLoq = (CDbl(aCS1Compound2.Response) / CDbl(aStandard.Response)) * (CDbl(aStandard.Conc) / CDbl(amCompound2.RRF))
                                                                                    Exit For
                                                                                End If
                                                                            Next
                                                                        End If
                                                                    Next
                                                                End If
                                                            Next
                                                        End If
                                                    Next
                                                    'While I am here, deal with some QC Flags (4,5,8,9)
                                                    If CDbl(aCompound1.Response) < CDbl(aCS1Compound1.Response) Then
                                                        aCompound1.MidFQC4 = True
                                                    End If
                                                    If CDbl(aCompound2.Response) < CDbl(aCS1Compound2.Response) Then
                                                        aCompound2.MidFQC5 = True
                                                    End If
                                                    If aCompound1.MI Then
                                                        aCompound1.MidFQC8 = True
                                                    End If
                                                    If aCompound2.MI Then
                                                        aCompound2.MidFQC9 = True
                                                    End If
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            Next
                        Next
                    End If
                End If

            Next
        Catch ex As Exception
            MsgBox("Error calculating LOQ Amounts!" & vbCrLf & _
                  "Sub Procedure: MidlandFAST()" & vbCrLf & _
                  "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        Try
            'Calulate 13c amt
            For Each aSample In GlobalVariables.ReportSamList
                If aSample.InjSpikeAmt <> 0 Then
                    For Each aInjCompound In aSample.CompoundList
                        'Get inj compound
                        If InStr(aInjCompound.Name, "(INJ)", CompareMethod.Binary) Then
                            'Go through every standard
                            For Each aStandard In aSample.InternalStdList
                                'Get instrument
                                For Each aInstrument In GlobalVariables.selMethod.mInstrumentList
                                    If aInstrument.Name = GlobalVariables.Report.Instrument Then
                                        'Get Cal information for inj compound
                                        For Each amInjStandard In aInstrument.mStandardList
                                            If amInjStandard.Name = aInjCompound.Name Then
                                                'Get Cal information for standard
                                                For Each amStandard In aInstrument.mStandardList
                                                    If aStandard.Name = amStandard.Name Then
                                                        aStandard.MidF13CAmt = (CDbl(aStandard.Response) / CDbl(aInjCompound.Response)) * ((aSample.InjSpikeAmt * CDbl(amInjStandard.Conc)) / ((CDbl(amStandard.AvgArea) / CDbl(amInjStandard.AvgArea)) * (CDbl(amInjStandard.CalAmt) / CDbl(amStandard.CalAmt))))
                                                    End If

                                                Next
                                            End If
                                        Next
                                    End If
                                Next
                            Next
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox("Error calculating 13C Amounts!" & vbCrLf & _
                  "Sub Procedure: MidlandFAST()" & vbCrLf & _
                  "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        Try
            'Calculate 13c recovery
            For Each aSample In GlobalVariables.ReportSamList
                If aSample.InjSpikeAmt <> 0 Then
                    If Not aSample.Calculated Then
                        For Each aStandard In aSample.InternalStdList
                            For Each aInstrument In GlobalVariables.selMethod.mInstrumentList
                                If aInstrument.Name = GlobalVariables.Report.Instrument Then
                                    For Each amStandard In aInstrument.mStandardList
                                        If amStandard.Name = aStandard.Name Then
                                            aStandard.MidF13CRecovery = (aStandard.MidF13CAmt / (CDbl(aSample.StdSpikeAmt) * CDbl(amStandard.Conc))) * CDbl(aSample.Aliquot) * 100
                                            If aStandard.MidF13CRecovery >= amStandard.RecLowLim And aStandard.MidF13CRecovery <= amStandard.RecUpLim Then
                                                aStandard.MidF13CRecoveryInLim = True
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("Error calculating 13C Recoveries!" & vbCrLf & _
                    "Sub Procedure: MidlandFAST()" & vbCrLf & _
                    "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        Try
            'Calculate final reported amounts/LOQ
            For Each aSample In GlobalVariables.ReportSamList
                If Not aSample.Calculated Then
                    For Each aCompound1 In aSample.CompoundList
                        For Each aCompound2 In aSample.CompoundList
                            'If match, find CS1 ions
                            If Trim(aCompound2.Name.Substring(0, InStrRev(aCompound2.Name, "(") - 1)) = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) And _
                                aCompound1.Name <> aCompound2.Name And aCompound1.MidFReportedAmt = "-1" And aCompound2.MidFReportedAmt = "-1" And aCompound1.MidFReportedLOQAmt = "-1" And aCompound2.MidFReportedLOQAmt = "-1" Then
                                If aCompound1.MidFQC1 Then
                                    If aCompound1.MidFNonDetect Then
                                        'Final LOQ amount
                                        aCompound1.MidFReportedLOQAmt = CStr((CDbl(aCompound1.MidFLoq) + CDbl(aCompound2.MidFLoq)) / 2)
                                        aCompound2.MidFReportedLOQAmt = aCompound1.MidFReportedLOQAmt
                                        'Final Compound amount
                                        If InStr(aCompound1.Name, "TCDD") Then
                                            aCompound1.MidFReportedAmt = aCompound1.MidFReportedLOQAmt
                                            aCompound2.MidFReportedAmt = aCompound1.MidFReportedAmt
                                        Else
                                            aCompound1.MidFReportedAmt = CStr(CDbl(aCompound1.MidFReportedLOQAmt) / 3)
                                            aCompound2.MidFReportedAmt = aCompound1.MidFReportedAmt
                                        End If
                                        'Deal with some QC Flags (3)
                                        If CDbl(aCompound1.MidFReportedAmt) < CDbl(aCompound1.MidFReportedLOQAmt) Then
                                            aCompound1.MidFQC3 = True
                                            aCompound2.MidFQC3 = True
                                        End If
                                        aCompound1.MidFQC12 = True
                                        aCompound2.MidFQC12 = True
                                    Else
                                        'Final LOQ amount
                                        aCompound1.MidFReportedLOQAmt = CStr((CDbl(aCompound1.MidFLoq) + CDbl(aCompound2.MidFLoq)) / 2)
                                        aCompound2.MidFReportedLOQAmt = aCompound1.MidFReportedLOQAmt
                                        'Final Compound amount
                                        aCompound1.MidFReportedAmt = CStr((CDbl(aCompound1.Conc) + CDbl(aCompound2.Conc)) / 2)
                                        aCompound2.MidFReportedAmt = aCompound1.MidFReportedAmt
                                        'Deal with some QC Flags (3)
                                        If CDbl(aCompound1.MidFReportedAmt) < CDbl(aCompound1.MidFReportedLOQAmt) Then
                                            aCompound1.MidFQC3 = True
                                            aCompound2.MidFQC3 = True
                                        End If
                                    End If
                                ElseIf aCompound1.MidFQC2 Then
                                    'Final LOQ Amount
                                    aCompound1.MidFReportedLOQAmt = CStr((CDbl(aCompound1.MidFLoq) + CDbl(aCompound2.MidFLoq)) / 2)
                                    aCompound2.MidFReportedLOQAmt = aCompound1.MidFReportedLOQAmt
                                    If aCompound1.MidFNonDetect Then
                                        If InStr(aCompound1.Name, "TCDD") Then
                                            aCompound1.MidFReportedAmt = aCompound1.MidFReportedLOQAmt
                                            aCompound2.MidFReportedAmt = aCompound1.MidFReportedAmt
                                        Else
                                            aCompound1.MidFReportedAmt = CStr(CDbl(aCompound1.MidFReportedLOQAmt) / 3)
                                            aCompound2.MidFReportedAmt = aCompound1.MidFReportedAmt
                                        End If
                                        aCompound1.MidFQC12 = True
                                        aCompound2.MidFQC12 = True
                                    Else
                                        'Final Compound amount
                                        If CDbl(aCompound1.Conc) > CDbl(aCompound2.Conc) Then
                                            aCompound1.MidFReportedAmt = aCompound2.Conc
                                            aCompound2.MidFReportedAmt = aCompound1.MidFReportedAmt
                                        ElseIf CDbl(aCompound1.Conc) < CDbl(aCompound2.Conc) Then
                                            aCompound1.MidFReportedAmt = aCompound1.Conc
                                            aCompound2.MidFReportedAmt = aCompound1.MidFReportedAmt
                                        Else
                                            aCompound1.MidFReportedAmt = aCompound1.Conc
                                            aCompound2.MidFReportedAmt = aCompound1.MidFReportedAmt
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox("Error calculating Final Reported Amounts!" & vbCrLf & _
                  "Sub Procedure: MidlandFAST()" & vbCrLf & _
                  "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        'More QC setting
        For Each aSample In GlobalVariables.ReportSamList
            For Each aCompound In aSample.CompoundList
                For Each aStandard In aSample.InternalStdList
                    If aCompound.MidF13CAssoc = aStandard.Name Then
                        If aStandard.MI Then
                            aCompound.MidFQC7 = True
                        End If
                        Exit For
                    End If
                Next
            Next
        Next
        For Each aSample In GlobalVariables.ReportSamList
            For Each aCompound In aSample.CompoundList
                'Flags
                strFlags = ""
                If aCompound.MidFQC2 Then
                    If strFlags = "" Then
                        strFlags = "Y"
                    Else
                        strFlags = strFlags & ",Y"
                    End If
                End If
                If aCompound.MidFQC4 Or aCompound.MidFQC5 Then
                    If strFlags = "" Then
                        strFlags = "J"
                    Else
                        strFlags = strFlags & ",J"
                    End If
                End If
                aCompound.MidFFlags = strFlags
            Next
            For Each aStandard In aSample.InternalStdList
                'Flags
                strFlags = ""
                If Not aStandard.MidF13CRecoveryInLim Then
                    If strFlags = "" Then
                        strFlags = "A"
                    Else
                        strFlags = strFlags & ",A"
                    End If
                End If
                If Not aStandard.MidFIonRatioInLim Then
                    If strFlags = "" Then
                        strFlags = "W"
                    Else
                        strFlags = strFlags & ",W"
                    End If
                End If
                aStandard.MidFFlags = strFlags
            Next
        Next

        Try
            'LCS Calculations
            If Not IsNothing(aLCSSample) Then
                For Each aCompound1 In aLCSSample.CompoundList
                    For Each aInstrument In GlobalVariables.selMethod.mInstrumentList
                        If aInstrument.Name = GlobalVariables.Report.Instrument Then
                            For Each amCompound In aInstrument.mCompoundList
                                If amCompound.Name = aCompound1.Name Then
                                    For Each aStandard In aLCSSample.InternalStdList
                                        If aStandard.Name = aCompound1.MidF13CAssoc Then
                                            For Each amStandard In aInstrument.mStandardList
                                                If aStandard.Name = amStandard.Name Then
                                                    aCompound1.MidFLCSAmtAdded = CStr(((CDbl(aLCSSample.LCSSpikeAmt) * CDbl(amCompound.Conc)) / (CDbl(aLCSSample.StdSpikeAmt) * CDbl(amStandard.Conc))) * CDbl(aStandard.Conc))
                                                    aCompound1.MidFLCSAmtRecovered = aCompound1.Conc
                                                    aCompound1.MidFLCSPercRecovered = CStr((CDbl(aCompound1.MidFLCSAmtRecovered) / CDbl(aCompound1.MidFLCSAmtAdded)) * 100)
                                                    aCompound1.MidFLCSLowLim = amCompound.LCSLLim
                                                    aCompound1.MidFLCSHighLim = amCompound.LCSULim
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox("Error calculating LCS Values!" & vbCrLf & _
                  "Sub Procedure: MidlandFAST()" & vbCrLf & _
                  "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        Try
            'CS3 Calculations
            If Not IsNothing(aCS3Sample) Then
                For Each aCompound1 In aCS3Sample.CompoundList
                    For Each aInstrument In GlobalVariables.selMethod.mInstrumentList
                        If aInstrument.Name = GlobalVariables.Report.Instrument Then
                            For Each amCompound In aInstrument.mCompoundList
                                If amCompound.Name = aCompound1.Name Then
                                    aCompound1.MidFCS3TotalAmt = amCompound.CS3Amt
                                    aCompound1.MidFCS3AmtRecovered = aCompound1.Conc
                                    aCompound1.MidFCS3LowLim = CStr(CDbl(amCompound.CS3Amt) * ((100 - GlobalVariables.selMethod.RptTolerance) / 100))
                                    aCompound1.MidFCS3HighLim = CStr(CDbl(amCompound.CS3Amt) * ((100 + GlobalVariables.selMethod.RptTolerance) / 100))
                                End If
                            Next
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox("Error calculating CS3 Values!" & vbCrLf & _
                  "Sub Procedure: MidlandFAST()" & vbCrLf & _
                  "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        Try
            'Final Data Calulations
            For Each aSample In GlobalVariables.ReportSamList
                aSIS = GlobalVariables.SISList(0)
                For Each aSISSample In aSIS.SampleList
                    If aSample.LimsID = aSISSample.SISLabNum Then
                        For Each aCompound1 In aSample.CompoundList
                            'Get compound information
                            For Each aInstrument In GlobalVariables.selMethod.mInstrumentList
                                If aInstrument.Name = GlobalVariables.Report.Instrument Then
                                    For Each amStandard In aInstrument.mStandardList
                                        If aCompound1.MidF13CAssoc = amStandard.Name Then
                                            'Final weight
                                            aCompound1.MidFFinalWeight = CStr(((CDbl(aCompound1.MidFReportedAmt) * ((aSample.StdSpikeAmt * amStandard.Conc) / CDbl(amStandard.CalAmt))) / CDbl(aSISSample.SISFinalWeight)) * 1000)
                                            For Each amCompound1 In aInstrument.mCompoundList
                                                If amCompound1.Name = aCompound1.Name Then
                                                    'TEQ amt w/ TEF factor
                                                    aCompound1.MidFTEQFinalWeight = CStr(CDbl(aCompound1.MidFFinalWeight) * CDbl(amCompound1.TEF))
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            Next
                        Next
                    End If
                Next
            Next

            'ETEQ Calculations for Final Data
            'ETEQ0
            For Each aSample In GlobalVariables.ReportSamList
                For Each aCompound1 In aSample.CompoundList
                    For Each aCompound2 In aSample.CompoundList
                        If Trim(aCompound2.Name.Substring(0, InStrRev(aCompound2.Name, "(") - 1)) = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) And aCompound1.Name <> aCompound2.Name And aCompound1.MidFETEQ0 = False And aCompound2.MidFETEQ0 = False Then
                            If Not aCompound1.MidFNonDetect Then
                                aSample.MidFETEQ0 = CStr(CDbl(aSample.MidFETEQ0) + CDbl(aCompound1.MidFTEQFinalWeight))
                                aCompound1.MidFETEQ0 = True
                                aCompound2.MidFETEQ0 = True
                            End If
                        End If
                    Next
                Next
            Next
            'ETEQ05
            For Each aSample In GlobalVariables.ReportSamList
                For Each aCompound1 In aSample.CompoundList
                    For Each aCompound2 In aSample.CompoundList
                        If Trim(aCompound2.Name.Substring(0, InStrRev(aCompound2.Name, "(") - 1)) = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) And aCompound1.Name <> aCompound2.Name And aCompound1.MidFETEQ05 = False And aCompound2.MidFETEQ05 = False Then
                            If Not aCompound1.MidFNonDetect Then
                                aSample.MidFETEQ05 = CStr(CDbl(aSample.MidFETEQ05) + CDbl(aCompound1.MidFTEQFinalWeight))
                                aCompound1.MidFETEQ05 = True
                                aCompound2.MidFETEQ05 = True
                            Else
                                aSample.MidFETEQ05 = CStr(CDbl(aSample.MidFETEQ05) + (CDbl(aCompound1.MidFTEQFinalWeight) * 0.5))
                                aCompound1.MidFETEQ05 = True
                                aCompound2.MidFETEQ05 = True
                            End If
                        End If
                    Next
                Next
            Next
            'ETEQLOD
            For Each aSample In GlobalVariables.ReportSamList
                For Each aCompound1 In aSample.CompoundList
                    For Each aCompound2 In aSample.CompoundList
                        If Trim(aCompound2.Name.Substring(0, InStrRev(aCompound2.Name, "(") - 1)) = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) And aCompound1.Name <> aCompound2.Name And aCompound1.MidFETEQLOD = False And aCompound2.MidFETEQLOD = False Then
                            aSample.MidFETEQLOD = CStr(CDbl(aSample.MidFETEQLOD) + CDbl(aCompound1.MidFTEQFinalWeight))
                            aCompound1.MidFETEQLOD = True
                            aCompound2.MidFETEQLOD = True
                        End If
                    Next
                Next
            Next
            For Each aSample In GlobalVariables.ReportSamList
                aSample.MidFETEQ0 = CStr(CDbl(aSample.MidFETEQ0) * CDbl(GlobalVariables.selMethod.ETEQ))
                aSample.MidFETEQ05 = CStr(CDbl(aSample.MidFETEQ05) * CDbl(GlobalVariables.selMethod.ETEQ))
                aSample.MidFETEQLOD = CStr(CDbl(aSample.MidFETEQLOD) * CDbl(GlobalVariables.selMethod.ETEQ))
            Next

            'Get needed data into CompoundNameList samples
            For Each aSample In GlobalVariables.ReportSamList
                'reset comparison flag
                For Each aCompound1 In aSample.CompoundList
                    aCompound1.Written = False
                Next
                For Each aCompound1 In aSample.CompoundList
                    If Not aCompound1.Written Then
                        For Each aCompound2 In aSample.CompoundList
                            If Trim(aCompound2.Name.Substring(0, InStrRev(aCompound2.Name, "(") - 1)) = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) And aCompound1.Name <> aCompound2.Name Then
                                For Each aCompound In aSample.CompoundNameList
                                    If aCompound.Name = aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1) Then
                                        aCompound.CopyDetails(aCompound1, aCompound2)
                                        aCompound1.Written = True
                                        aCompound2.Written = True
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            Next

        Catch ex As Exception
            MsgBox("Error occured during calculation!" & vbCrLf & _
                  "Sub Procedure: MidlandFAST()" & vbCrLf & _
                  "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        'Final QC flags

        'Set samples to calculated
        For Each aSample In GlobalVariables.ReportSamList
            aSample.Calculated = True
        Next

        'Check for Unit Conversion needing to happen
        'GlobalVariables.NeedsUnitConversion = MsgBox("Would you like to change/confirm the units to report in?", MsgBoxStyle.YesNo)
        'If GlobalVariables.NeedsUnitConversion Then
        '    UnitConversion.ShowDialog()
        'End If

        GlobalVariables.NeedsCalculation = False

        Return True

    End Function

    'This builds the theoretical compounds and their ion's for Midland FAST Calculation
    Function BuildTheoComps() As Boolean
        Dim amInstrument As mInstrument
        Dim amStandard As mStandard
        Dim amCompound1 As mCompound
        Dim amCompound2 As mCompound

        Try
            For Each amInstrument In GlobalVariables.selMethod.mInstrumentList
                If amInstrument.Name = GlobalVariables.selInstrument Then
                    For Each amStandard In amInstrument.mStandardList
                        If InStr(amStandard.Name, "(INJ)") = 0 Then
                            GlobalVariables.TheoComps.Add(LoadTComps(amStandard.Name, amStandard.IonTarget, amStandard.IonQual, Math.Round(CDbl((CDbl(amStandard.AbundQual) / CDbl(amStandard.AbundTarget)) * 100), 1)))
                        End If
                    Next
                    For Each amCompound1 In amInstrument.mCompoundList
                        For Each amCompound2 In amInstrument.mCompoundList
                            If Trim(amCompound2.Name.Substring(0, InStrRev(amCompound2.Name, "(") - 1)) = Trim(amCompound1.Name.Substring(0, InStrRev(amCompound1.Name, "(") - 1)) And _
                                    amCompound1.Name <> amCompound2.Name And InStr(amCompound1.Name, "(1)") And InStr(amCompound2.Name, "(2)") And amCompound1.Calculated = False And amCompound2.Calculated = False Then
                                GlobalVariables.TheoComps.Add(LoadTComps(Trim(amCompound1.Name.Substring(0, InStrRev(amCompound1.Name, "(") - 1)), amCompound1.Ion, amCompound2.Ion, Math.Round(CDbl((CDbl(amCompound2.Abundance) / CDbl(amCompound1.Abundance)) * 100), 1)))
                            End If
                        Next
                    Next
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox("Error building Theoretical Ion Ratios!" & vbCrLf & _
                  "Sub Procedure: BuildTheoComps()" & vbCrLf & _
                  "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try
    End Function
    Function LoadTComps(ByVal n As String, ByVal ts As String, ByVal q1s As String, ByVal i As Double) As Compound
        Dim aCompound As New Compound

        aCompound.Name = n                  'Compound Name
        aCompound.TSignal = ts              'First signal 
        aCompound.Q1Signal = q1s            'Second signal
        aCompound.MidFIonRatio = i              'Ion ratio
        aCompound.MidFLLim = i * 0.85           'Ion ratio lower lim
        aCompound.MidFULim = i * 1.15           'Ion ratio upper lim
        Return aCompound

    End Function

    'SigFigs code
    Function FormatSF(ByVal n As String) As String
        Dim blnNeg As Boolean
        Dim num As Double
        Dim newNum As Decimal
        Dim blnDec As Boolean
        Dim intC As Integer
        Dim newNumS As String

        Try
            'Check if -1 sig figs for no filter
            If GlobalVariables.eTrain.SigFig = -1 Then
                Return n
            End If

            'Numeric check, if not numeric just return
            If Not IsNumeric(n) Then
                Return n
            End If

            '0 check
            If n = "0" Then
                Return n
            End If

            'NaN check
            If n = "NaN" Then
                Return n
            End If

            'negative check
            If CDbl(n) < 0 Then
                blnNeg = True
                n = CStr(Math.Abs(CDbl(n)))
            End If

            num = CDbl(n)
            newNum = Math.Pow(10, Math.Floor(Math.Log10(Math.Abs(num))) + 1)
            newNum = newNum * Math.Round(num / newNum, CInt(GlobalVariables.eTrain.SigFig))

            newNumS = newNum.ToString
            'Deal with potential 0 in decimal missing due to scalar round

            intC = (newNumS).Length
            'See if decimal already
            If InStr(newNumS, "0.") Then
                'intC = intC - 2
                blnDec = True
            ElseIf InStr(newNumS, ".") Then
                intC = intC - 1
                blnDec = True
            Else
                blnDec = False
            End If

            If blnDec Then
                Do Until intC >= GlobalVariables.eTrain.SigFig
                    newNumS = newNumS & "0"
                    intC = intC + 1
                Loop
            Else
                'If intC < GlobalVariables.eTrain.SigFig Then
                Do Until intC >= GlobalVariables.eTrain.SigFig
                    If InStr(newNumS, ".") Then
                        newNumS = newNumS & "0"
                    Else
                        newNumS = newNumS & ".0"
                    End If
                    intC = intC + 1
                Loop
                'End If

            End If

            'Zero fix - testing
            If intC > GlobalVariables.eTrain.SigFig And InStr(newNumS, ".") And InStr(newNumS.Length, newNumS, "0") Then
                newNumS = newNumS.Substring(0, newNumS.Length - 1)
            End If

            If blnNeg Then
                Return "-" & CStr(newNumS)
            Else
                Return CStr(newNumS)
            End If

        Catch ex As Exception
            MsgBox("Error formatting SigFig value!" & vbCrLf & _
                  "Sub Procedure: FormatSF()" & vbCrLf & _
                  "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
