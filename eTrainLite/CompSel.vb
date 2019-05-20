Public Class CompSel

    Private Sub ReportDetails_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim aSample As Sample
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aInstrument As mInstrument
        Dim amStandard As mStandard
        Dim amSurrogate As mSurrogate
        Dim amCompound As mCompound
        Dim aStandard As Standard
        Dim aSurrogate As Surrogate
        Dim aCompound As Compound
        Dim blnF As Boolean

        'Grab one 
        If GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                For Each aPermit In GlobalVariables.PermitList
                    If aPermit.Name = GlobalVariables.selPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If aProject.Name = GlobalVariables.selProject Then
                                For Each aInstrument In aProject.mInstrumentList
                                    If aInstrument.Name = GlobalVariables.selInstrument Then
                                        For Each aSample In GlobalVariables.ReportSamList 'WT 10/13/2017 -> changed to GV.SampleList to test (Original = GV.ReportSamList)
                                            For Each aStandard In aSample.InternalStdList
                                                blnF = False
                                                For Each amStandard In aInstrument.mStandardList
                                                    If amStandard.Name = aStandard.Name Then
                                                        blnF = True
                                                        Exit For
                                                    End If
                                                Next
                                                If blnF Then
                                                    dgCompSelect.Rows.Add(aStandard.Name, True)
                                                Else
                                                    dgCompSelect.Rows.Add(aStandard.Name, False)
                                                End If
                                            Next
                                            For Each aSurrogate In aSample.SurrogateList
                                                blnF = False
                                                For Each amSurrogate In aInstrument.mSurrogateList
                                                    If aSurrogate.Name = amSurrogate.Name Then
                                                        blnF = True
                                                        Exit For
                                                    End If
                                                Next
                                                If blnF Then
                                                    dgCompSelect.Rows.Add(aSurrogate.Name, True)
                                                Else
                                                    dgCompSelect.Rows.Add(aSurrogate.Name, False)
                                                End If
                                            Next
                                            For Each aCompound In aSample.CompoundList
                                                For Each amCompound In aInstrument.mCompoundList
                                                    If amCompound.Name = aCompound.Name Then
                                                        blnF = True
                                                        Exit For
                                                    End If
                                                Next
                                                If blnF Then
                                                    dgCompSelect.Rows.Add(aCompound.Name, True)
                                                Else
                                                    dgCompSelect.Rows.Add(aCompound.Name, False)
                                                End If
                                            Next
                                            Exit For
                                        Next
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        ElseIf GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                For Each aPermit In GlobalVariables.PermitList
                    If aPermit.Name = GlobalVariables.selPermit.Name Then
                        For Each aProject In aPermit.ProjectList
                            If aProject.Name = GlobalVariables.selProject Then
                                For Each aInstrument In aProject.mInstrumentList
                                    If aInstrument.Name = GlobalVariables.selInstrument Then
                                    For Each aSample In GlobalVariables.ReportSamList
                                        For Each aStandard In aSample.InternalStdList
                                            blnF = False
                                            For Each amStandard In aInstrument.mStandardList
                                                If amStandard.Name = aStandard.Name Then
                                                    blnF = True
                                                    Exit For
                                                End If
                                            Next
                                            If blnF Then
                                                dgCompSelect.Rows.Add(aStandard.Name, True)
                                            Else
                                                dgCompSelect.Rows.Add(aStandard.Name, False)
                                            End If
                                        Next
                                        For Each aSurrogate In aSample.SurrogateList
                                            blnF = False
                                            For Each amSurrogate In aInstrument.mSurrogateList
                                                If aSurrogate.Name = amSurrogate.Name Then
                                                    blnF = True
                                                    Exit For
                                                End If
                                            Next
                                            If blnF Then
                                                dgCompSelect.Rows.Add(aSurrogate.Name, True)
                                            Else
                                                dgCompSelect.Rows.Add(aSurrogate.Name, False)
                                            End If
                                        Next
                                        For Each aCompound In aSample.CompoundList
                                            For Each amCompound In aInstrument.mCompoundList
                                                If amCompound.Name = aCompound.Name Then
                                                    blnF = True
                                                    Exit For
                                                End If
                                            Next
                                            If blnF Then
                                                dgCompSelect.Rows.Add(aCompound.Name, True)
                                            Else
                                                dgCompSelect.Rows.Add(aCompound.Name, False)
                                            End If
                                        Next
                                        Exit For
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
    Private Sub Me_FormClosing(sender As Object, e As FormClosingEventArgs) _
     Handles Me.FormClosing

        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            dgCompSelect.Rows.Clear()
            Me.Hide()
        End If

    End Sub

    Private Sub btnSel_Click(sender As System.Object, e As System.EventArgs) Handles btnSel.Click
        Dim aSample As Sample
        Dim aStandard As Standard
        Dim aSurrogate As Surrogate
        Dim aCompound As Compound


        For Each aSample In GlobalVariables.ReportSamList
            For Each aStandard In aSample.InternalStdList
                For Each r In dgCompSelect.Rows
                    If r.Cells.Item(0).Value <> "" Then
                        If aStandard.Name = r.Cells.Item(0).Value Then
                            aStandard.WriteToReport = r.Cells.Item(1).Value
                            Exit For
                        End If
                    End If
                Next
            Next
            For Each aSurrogate In aSample.SurrogateList
                For Each r In dgCompSelect.Rows
                    If r.Cells.Item(0).Value <> "" Then
                        If aSurrogate.Name = r.Cells.Item(0).Value Then
                            aSurrogate.WriteToReport = r.Cells.Item(1).Value
                            Exit For
                        End If
                    End If
                Next
            Next
            For Each aCompound In aSample.CompoundList
                For Each r In dgCompSelect.Rows
                    If r.Cells.Item(0).Value <> "" Then
                        If aCompound.Name = r.Cells.Item(0).Value Then
                            aCompound.WriteToReport = r.Cells.Item(1).Value
                            Exit For
                        End If
                    End If
                Next
            Next
        Next

        ''Assign Flags
        'For Each r In dgCompSelect.Rows
        '    If r.Cells.Item(0).Value <> "" Then
        '        blnMatch = False
        '        For Each aSample In GlobalVariables.SampleList
        '            If Not blnMatch Then
        '                For Each aStandard In aSample.InternalStdList
        '                    If aStandard.Name = r.Cells.Item(0).Value Then
        '                        aStandard.WriteToReport = r.Cells.Item(1).Value
        '                        blnMatch = True
        '                        Exit For
        '                    End If
        '                Next
        '            End If
        '            If Not blnMatch Then
        '                For Each aSurrogate In aSample.SurrogateList
        '                    If aSurrogate.Name = r.Cells.Item(0).Value Then
        '                        aSurrogate.WriteToReport = r.Cells.Item(1).Value
        '                        blnMatch = True
        '                        Exit For
        '                    End If
        '                Next
        '            End If
        '            If Not blnMatch Then
        '                For Each aCompound In aSample.CompoundList
        '                    If aCompound.Name = r.Cells.Item(0).Value Then
        '                        aCompound.WriteToReport = r.Cells.Item(1).Value
        '                        blnMatch = True
        '                        Exit For
        '                    End If
        '                Next
        '            End If
        '        Next
        '    End If
        'Next
        GlobalVariables.ContinueTransfer = True
        GlobalVariables.ContinueReport = True
        Me.Close()
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        GlobalVariables.ContinueTransfer = False
        GlobalVariables.ContinueReport = False
        dgCompSelect.Rows.Clear()
        Me.Close()
    End Sub
End Class