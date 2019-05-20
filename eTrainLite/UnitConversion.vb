Public Class UnitConversion
    Private Sub Me_FormClosing(sender As Object, e As FormClosingEventArgs) _
    Handles Me.FormClosing

        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
        End If

    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        Dim aSample As Sample
        Dim aCompound As Compound
        Dim aSurrogate As Surrogate
        Dim aStandard As Standard
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aInstrument As mInstrument
        Dim amStandard As mStandard
        Dim amCompound As mCompound
        Dim amSurrogate As mSurrogate
        Dim dblPPMConv As Double
        Dim dblSampFactor As Double
        Dim dblRepFactor As Double
        Dim strLimsUnit As String

        'Convert everything to PPM then go to whatever the select unit is

        Try

       

        dblSampFactor = 0
        'Factors
        Select Case cboSampleUnits.Text
            Case "ppm"
                dblSampFactor = 1
            Case "ppb"
                dblSampFactor = 1000
            Case "ppt"
                dblSampFactor = 1000000
            Case "ppq"
                dblSampFactor = 1000000000
            Case "mg/kg"
                dblSampFactor = 1
            Case "ug/kg"
                dblSampFactor = 1000
            Case "ng/kg"
                dblSampFactor = 1000000
            Case "ug/g"
                dblSampFactor = 1
            Case "ng/g"
                dblSampFactor = 1000
            Case "pg/g"
                dblSampFactor = 1000000
            Case "ng/mg"
                dblSampFactor = 1
            Case "pg/mg"
                dblSampFactor = 1000
            Case "mg/L"
                dblSampFactor = 1
            Case "ug/L"
                dblSampFactor = 1000
            Case "ng/L"
                dblSampFactor = 1000000
            Case "pg/L"
                dblSampFactor = 1000000000
            Case "ug/mL"
                dblSampFactor = 1
            Case "ng/mL"
                dblSampFactor = 1000
            Case "pg/mL"
                dblSampFactor = 1000000
            Case "ng/uL"
                dblSampFactor = 1
            Case "pg/uL"
                dblSampFactor = 1000
            Case "%"
                dblSampFactor = 0.0001
        End Select

        dblRepFactor = 0
        'Factors
        Select Case cboReportedUnits.Text
            Case "ppm"
                dblRepFactor = 1
            Case "ppb"
                dblRepFactor = 1000
            Case "ppt"
                dblRepFactor = 1000000
            Case "ppq"
                dblRepFactor = 1000000000
            Case "mg/kg"
                dblRepFactor = 1
            Case "ug/kg"
                dblRepFactor = 1000
            Case "ng/kg"
                dblRepFactor = 1000000
            Case "ug/g"
                dblRepFactor = 1
            Case "ng/g"
                dblRepFactor = 1000
            Case "pg/g"
                dblRepFactor = 1000000
            Case "ng/mg"
                dblRepFactor = 1
            Case "pg/mg"
                dblRepFactor = 1000
            Case "mg/L"
                dblRepFactor = 1
            Case "ug/L"
                dblRepFactor = 1000
            Case "ng/L"
                dblRepFactor = 1000000
            Case "pg/L"
                dblRepFactor = 1000000000
            Case "ug/mL"
                dblRepFactor = 1
            Case "ng/mL"
                dblRepFactor = 1000
            Case "pg/mL"
                dblRepFactor = 1000000
            Case "ng/uL"
                dblRepFactor = 1
            Case "pg/uL"
                dblRepFactor = 1000
            Case "%"
                dblRepFactor = 0.0001
        End Select

        If dblSampFactor <> 0 And dblRepFactor <> 0 Then
            For Each aSample In GlobalVariables.SampleList
                For Each aStandard In aSample.InternalStdList
                    'Convert to PPM First
                    If IsNumeric(aStandard.Conc) Then
                        dblPPMConv = CDbl(aStandard.Conc) / dblSampFactor
                        'Then convert to selected unit
                        aStandard.Conc = CStr(dblPPMConv * dblRepFactor)
                        aStandard.ReportedUnits = cboReportedUnits.Text
                    End If
                Next
                For Each aSurrogate In aSample.SurrogateList
                    'Convert to PPM First
                    If IsNumeric(aSurrogate.Conc) Then
                        dblPPMConv = CDbl(aSurrogate.Conc) / dblSampFactor
                        'Then convert to selected unit
                        aSurrogate.Conc = CStr(dblPPMConv * dblRepFactor)
                        aSurrogate.ReportedUnits = cboReportedUnits.Text
                    End If
                Next
                For Each aCompound In aSample.CompoundList
                    'Convert to PPM First
                    If IsNumeric(aCompound.Conc) Then
                        dblPPMConv = CDbl(aCompound.Conc) / dblSampFactor
                        'Then convert to selected unit
                        aCompound.Conc = CStr(dblPPMConv * dblRepFactor)
                        aCompound.ReportedUnits = cboReportedUnits.Text
                    End If
                Next
                aSample.ReportedUnits = cboReportedUnits.Text
            Next
        End If

        'Reporting limit conversion
        If GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                For Each aPermit In GlobalVariables.PermitList
                    If aPermit.Name = GlobalVariables.selPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If aProject.Name = GlobalVariables.selProject Then
                                strLimsUnit = aProject.LimsUnits
                            End If
                        Next
                    End If
                Next


                dblSampFactor = 0
                'Factors
                Select Case strLimsUnit
                    Case "ppm"
                        dblSampFactor = 1
                    Case "ppb"
                        dblSampFactor = 1000
                    Case "ppt"
                        dblSampFactor = 1000000
                    Case "ppq"
                        dblSampFactor = 1000000000
                    Case "mg/kg"
                        dblSampFactor = 1
                    Case "ug/kg"
                        dblSampFactor = 1000
                    Case "ng/kg"
                        dblSampFactor = 1000000
                    Case "ug/g"
                        dblSampFactor = 1
                    Case "ng/g"
                        dblSampFactor = 1000
                    Case "pg/g"
                        dblSampFactor = 1000000
                    Case "ng/mg"
                        dblSampFactor = 1
                    Case "pg/mg"
                        dblSampFactor = 1000
                    Case "mg/L"
                        dblSampFactor = 1
                    Case "ug/L"
                        dblSampFactor = 1000
                    Case "ng/L"
                        dblSampFactor = 1000000
                    Case "pg/L"
                        dblSampFactor = 1000000000
                    Case "ug/mL"
                        dblSampFactor = 1
                    Case "ng/mL"
                        dblSampFactor = 1000
                    Case "pg/mL"
                        dblSampFactor = 1000000
                    Case "ng/uL"
                        dblSampFactor = 1
                    Case "pg/uL"
                        dblSampFactor = 1000
                    Case "%"
                        dblSampFactor = 0.0001
                End Select

                dblRepFactor = 0
                'Factors
                Select Case cboReportedUnits.Text
                    Case "ppm"
                        dblRepFactor = 1
                    Case "ppb"
                        dblRepFactor = 1000
                    Case "ppt"
                        dblRepFactor = 1000000
                    Case "ppq"
                        dblRepFactor = 1000000000
                    Case "mg/kg"
                        dblRepFactor = 1
                    Case "ug/kg"
                        dblRepFactor = 1000
                    Case "ng/kg"
                        dblRepFactor = 1000000
                    Case "ug/g"
                        dblRepFactor = 1
                    Case "ng/g"
                        dblRepFactor = 1000
                    Case "pg/g"
                        dblRepFactor = 1000000
                    Case "ng/mg"
                        dblRepFactor = 1
                    Case "pg/mg"
                        dblRepFactor = 1000
                    Case "mg/L"
                        dblRepFactor = 1
                    Case "ug/L"
                        dblRepFactor = 1000
                    Case "ng/L"
                        dblRepFactor = 1000000
                    Case "pg/L"
                        dblRepFactor = 1000000000
                    Case "ug/mL"
                        dblRepFactor = 1
                    Case "ng/mL"
                        dblRepFactor = 1000
                    Case "pg/mL"
                        dblRepFactor = 1000000
                    Case "ng/uL"
                        dblRepFactor = 1
                    Case "pg/uL"
                        dblRepFactor = 1000
                    Case "%"
                        dblRepFactor = 0.0001
                End Select

                If dblSampFactor <> 0 And dblRepFactor <> 0 Then
                    For Each aPermit In GlobalVariables.PermitList
                        If aPermit.Name = GlobalVariables.selPermit.Name Then
                            For Each aProject In aPermit.ProjectList
                                If aProject.Name = GlobalVariables.selProject Then
                                    For Each aInstrument In aProject.mInstrumentList
                                        For Each amStandard In aInstrument.mStandardList
                                            'Convert to PPM First
                                            If IsNumeric(amStandard.MDL) Then
                                                dblPPMConv = CDbl(amStandard.MDL) / dblSampFactor
                                                'Then convert to selected unit
                                                amStandard.MDL = CStr(dblPPMConv * dblRepFactor)
                                            End If

                                            'Convert to PPM First
                                            If IsNumeric(amStandard.PQL) Then
                                                dblPPMConv = CDbl(amStandard.PQL) / dblSampFactor
                                                'Then convert to selected unit
                                                amStandard.PQL = CStr(dblPPMConv * dblRepFactor)
                                            End If

                                            'Convert to PPM First
                                            If IsNumeric(amStandard.RL) Then
                                                dblPPMConv = CDbl(amStandard.RL) / dblSampFactor
                                                'Then convert to selected unit
                                                amStandard.RL = CStr(dblPPMConv * dblRepFactor)
                                            End If
                                        Next
                                        For Each amSurrogate In aInstrument.mSurrogateList
                                            'Convert to PPM First
                                            If IsNumeric(amSurrogate.MDL) Then
                                                dblPPMConv = CDbl(amSurrogate.MDL) / dblSampFactor
                                                'Then convert to selected unit
                                                amSurrogate.MDL = CStr(dblPPMConv * dblRepFactor)
                                            End If

                                            If IsNumeric(amSurrogate.PQL) Then
                                                'Convert to PPM First
                                                dblPPMConv = CDbl(amSurrogate.PQL) / dblSampFactor
                                                'Then convert to selected unit
                                                amSurrogate.PQL = CStr(dblPPMConv * dblRepFactor)
                                            End If

                                            If IsNumeric(amSurrogate.RL) Then
                                                'Convert to PPM First
                                                dblPPMConv = CDbl(amSurrogate.RL) / dblSampFactor
                                                'Then convert to selected unit
                                                amSurrogate.RL = CStr(dblPPMConv * dblRepFactor)
                                            End If

                                        Next
                                        For Each amCompound In aInstrument.mCompoundList
                                            If IsNumeric(amCompound.MDL) Then
                                                'Convert to PPM First
                                                dblPPMConv = CDbl(amCompound.MDL) / dblSampFactor
                                                'Then convert to selected unit
                                                amCompound.MDL = CStr(dblPPMConv * dblRepFactor)
                                            End If

                                            If IsNumeric(amCompound.PQL) Then
                                                'Convert to PPM First
                                                dblPPMConv = CDbl(amCompound.PQL) / dblSampFactor
                                                'Then convert to selected unit
                                                amCompound.PQL = CStr(dblPPMConv * dblRepFactor)
                                            End If

                                            If IsNumeric(amCompound.RL) Then
                                                'Convert to PPM First
                                                dblPPMConv = CDbl(amCompound.RL) / dblSampFactor
                                                'Then convert to selected unit
                                                amCompound.RL = CStr(dblPPMConv * dblRepFactor)
                                            End If

                                        Next
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        End If
        Catch ex As Exception
            MsgBox("Error occured during unit conversion!" & vbCrLf & _
                "Sub Procedure: btnSave_Click()" & vbCrLf & _
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
        lblNotDetect.Visible = False
        Me.Hide()
    End Sub

    Private Sub btnHelp_Click(sender As System.Object, e As System.EventArgs) Handles btnHelp.Click
        MsgBox("Change reported units to whatever units you would like eTrain to report in.")
    End Sub

    Private Sub UnitConversion_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim aSample As Sample
        Dim aPermit As Permit
        Dim aProject As Project
        Dim blnFound As Boolean
        Dim blnFound2 As Boolean

        blnFound = False
        aSample = GlobalVariables.SampleList.Item(0)
        For i = 0 To cboSampleUnits.Items.Count - 1
            If LCase(aSample.Units) = LCase(cboSampleUnits.Items(i)) Then
                cboSampleUnits.SelectedIndex = i
                blnFound = True
            End If
        Next
        If GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                For Each aPermit In GlobalVariables.PermitList
                    If aPermit.Name = GlobalVariables.selPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If aProject.Name = GlobalVariables.selProject Then
                                For i = 0 To cboReportedUnits.Items.Count - 1
                                    If LCase(aProject.LimsUnits) = LCase(cboReportedUnits.Items(i)) Then
                                        cboReportedUnits.SelectedIndex = i
                                        blnFound2 = True
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
                
            End If
        End If
        If Not blnFound Then
            lblNotDetect.Visible = True
        End If
        If blnFound2 Then
            lblNotDetect2.Visible = True
        End If
        aSample = Nothing
    End Sub
End Class