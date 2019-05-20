Public Class SampleMerge
    Dim strSelection As String

    Private Sub SampleMerge_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboSampleList.SelectedIndex = 0
    End Sub

    Private Sub btnCombine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCombine.Click
        Dim aNewSample As New Sample
        Dim aSample As Sample
        Dim aSelSample As Sample
        Dim aStandard1 As Standard
        Dim aSurrogate1 As Surrogate
        Dim aCompound1 As Compound

        'Save final selection in case changes made
        For Each aSample In GlobalVariables.ReportSamList
            If aSample.Name = strSelection Then
                aSelSample = aSample
                If chkParent.Checked Then
                    aSelSample.Parent = True
                Else
                    aSelSample.Parent = False
                End If
                For Each r In dgComponents.Rows
                    For Each aStandard1 In aSelSample.InternalStdList
                        If r.Cells.Item(0).Value = aStandard1.Name Then
                            If r.Cells.Item(1).Value Then
                                aStandard1.Keep = True
                            Else
                                aStandard1.Keep = False
                            End If
                        End If
                    Next
                    For Each aSurrogate1 In aSelSample.SurrogateList
                        If r.Cells.Item(0).Value = aSurrogate1.Name Then
                            If r.Cells.Item(1).Value Then
                                aSurrogate1.Keep = True
                            Else
                                aSurrogate1.Keep = False
                            End If
                        End If
                    Next
                    For Each aCompound1 In aSelSample.CompoundList
                        If r.Cells.Item(0).Value = aCompound1.Name Then
                            If r.Cells.Item(1).Value Then
                                aCompound1.Keep = True
                            Else
                                aCompound1.Keep = False
                            End If
                        End If
                    Next
                Next
                Exit For
            End If
        Next

        'Get selected sample and save selection
        For Each aSample In GlobalVariables.ReportSamList
            If aSample.Name = strSelection Then
                aSelSample = aSample
                If chkParent.Checked Then
                    aSelSample.Parent = True
                Else
                    aSelSample.Parent = False
                End If
                For Each r In dgComponents.Rows
                    For Each aStandard1 In aSelSample.InternalStdList
                        If r.Cells.Item(0).Value = aStandard1.Name Then
                            If r.Cells.Item(1).Value Then
                                aStandard1.Keep = True
                            Else
                                aStandard1.Keep = False
                            End If
                        End If
                    Next
                    For Each aSurrogate1 In aSelSample.SurrogateList
                        If r.Cells.Item(0).Value = aSurrogate1.Name Then
                            If r.Cells.Item(1).Value Then
                                aSurrogate1.Keep = True
                            Else
                                aSurrogate1.Keep = False
                            End If
                        End If
                    Next
                    For Each aCompound1 In aSelSample.CompoundList
                        If r.Cells.Item(0).Value = aCompound1.Name Then
                            If r.Cells.Item(1).Value Then
                                aCompound1.Keep = True
                            Else
                                aCompound1.Keep = False
                            End If
                        End If
                    Next
                Next
                Exit For
            End If
        Next

        'Cross check all items in remaining sample in cbo to ensure checks are on or off for each component 

        For Each r In dgComponents.Rows
            For Each aSample In GlobalVariables.ReportSamList
                For Each itm In cboSampleList.Items
                    If aSample.Name = itm And Not aSample.Name = aSelSample.Name Then
                        If aSelSample.Parent Then
                            aSample.Parent = False
                        End If
                        For Each aStandard1 In aSample.InternalStdList
                            For Each aStandard2 In aSelSample.InternalStdList
                                If aStandard1.Name = aStandard2.Name Then
                                    If aStandard2.Keep Then
                                        aStandard1.Keep = False
                                    End If
                                End If
                            Next
                        Next
                        For Each aSurrogate1 In aSample.SurrogateList
                            For Each aSurrogate2 In aSelSample.SurrogateList
                                If aSurrogate1.Name = aSurrogate2.Name Then
                                    If aSurrogate2.Keep Then
                                        aSurrogate1.Keep = False
                                    End If
                                End If
                            Next
                        Next
                        For Each aCompound1 In aSample.CompoundList
                            For Each aCompound2 In aSelSample.CompoundList
                                If aCompound1.Name = aCompound2.Name Then
                                    If aCompound2.Keep Then
                                        aCompound1.Keep = False
                                    End If
                                End If
                            Next
                        Next
                    End If
                Next
            Next
        Next

        'clear dg
        dgComponents.Rows.Clear()
        chkParent.Checked = False
        aSelSample = Nothing


        'Create new sample
        For Each aSample In GlobalVariables.ReportSamList
            For Each itm In cboSampleList.Items
                If aSample.Name = itm Then
                    'Copy parent information
                    If aSample.Parent Then
                        aSample.CopyMerge(aNewSample)
                    End If
                    For Each aStandard1 In aSample.InternalStdList
                        If aStandard1.Keep Then
                            aNewSample.InternalStdList.Add(aStandard1)
                        End If
                    Next
                    For Each aSurrogate1 In aSample.SurrogateList
                        If aSurrogate1.Keep Then
                            aNewSample.SurrogateList.Add(aSurrogate1)
                        End If
                    Next
                    For Each aCompound1 In aSample.CompoundList
                        If aCompound1.Keep Then
                            aNewSample.CompoundList.Add(aCompound1)
                        End If
                    Next
                End If
            Next
        Next

        'add sample to list
        GlobalVariables.ReportSamList.Add(aNewSample)
        Me.Close()

    End Sub

    Private Sub Me_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles Me.FormClosing

        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            cboSampleList.Items.Clear()
            Me.Hide()
        End If

    End Sub

    Private Sub cboSampleList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSampleList.SelectedIndexChanged
        Dim aSample As Sample
        Dim aSelSample As Sample
        Dim aStandard1 As Standard
        Dim aSurrogate1 As Surrogate
        Dim aCompound1 As Compound
        Dim aStandard2 As Standard
        Dim aSurrogate2 As Surrogate
        Dim aCompound2 As Compound

        'Get selected sample and save selection
        For Each aSample In GlobalVariables.ReportSamList
            If aSample.Name = strSelection Then
                aSelSample = aSample
                If chkParent.Checked Then
                    aSelSample.Parent = True
                Else
                    aSelSample.Parent = False
                End If
                For Each r In dgComponents.Rows
                    For Each aStandard1 In aSelSample.InternalStdList
                        If r.Cells.Item(0).Value = aStandard1.Name Then
                            If r.Cells.Item(1).Value Then
                                aStandard1.Keep = True
                            Else
                                aStandard1.Keep = False
                            End If
                        End If
                    Next
                    For Each aSurrogate1 In aSelSample.SurrogateList
                        If r.Cells.Item(0).Value = aSurrogate1.Name Then
                            If r.Cells.Item(1).Value Then
                                aSurrogate1.Keep = True
                            Else
                                aSurrogate1.Keep = False
                            End If
                        End If
                    Next
                    For Each aCompound1 In aSelSample.CompoundList
                        If r.Cells.Item(0).Value = aCompound1.Name Then
                            If r.Cells.Item(1).Value Then
                                aCompound1.Keep = True
                            Else
                                aCompound1.Keep = False
                            End If
                        End If
                    Next
                Next
                Exit For
            End If
        Next



        'Newly selected sample
        strSelection = cboSampleList.Text
        'Cross check all items in remaining sample in cbo to ensure checks are on or off for each component 

        For Each r In dgComponents.Rows
            For Each aSample In GlobalVariables.ReportSamList
                For Each itm In cboSampleList.Items
                    If aSample.Name = itm And Not aSample.Name = aSelSample.Name Then
                        If aSelSample.Parent Then
                            aSample.Parent = False
                        End If
                        For Each aStandard1 In aSample.InternalStdList
                            For Each aStandard2 In aSelSample.InternalStdList
                                If aStandard1.Name = aStandard2.Name Then
                                    If aStandard2.Keep Then
                                        aStandard1.Keep = False
                                    End If
                                End If
                            Next
                        Next
                        For Each aSurrogate1 In aSample.SurrogateList
                            For Each aSurrogate2 In aSelSample.SurrogateList
                                If aSurrogate1.Name = aSurrogate2.Name Then
                                    If aSurrogate2.Keep Then
                                        aSurrogate1.Keep = False
                                    End If
                                End If
                            Next
                        Next
                        For Each aCompound1 In aSample.CompoundList
                            For Each aCompound2 In aSelSample.CompoundList
                                If aCompound1.Name = aCompound2.Name Then
                                    If aCompound2.Keep Then
                                        aCompound1.Keep = False
                                    End If
                                End If
                            Next
                        Next
                    End If
                Next
            Next
        Next

        'clear dg
        dgComponents.Rows.Clear()
        chkParent.Checked = False
        aSelSample = Nothing

        'Get newly selected sample
        For Each aSample In GlobalVariables.ReportSamList
            If aSample.Name = strSelection Then
                aSelSample = aSample
                If aSelSample.Parent Then
                    chkParent.Checked = True
                Else
                    chkParent.Checked = False
                End If
                Exit For
            End If
        Next

        'Fill datagrid
        For Each aStandard1 In aSelSample.InternalStdList
            dgComponents.Rows.Add(aStandard1.Name, aStandard1.Keep)
        Next
        For Each aCompound1 In aSelSample.CompoundList
            dgComponents.Rows.Add(aCompound1.Name, aCompound1.Keep)
        Next
        For Each aSurrogate1 In aSelSample.SurrogateList
            dgComponents.Rows.Add(aSurrogate1.Name, aSurrogate1.Keep)
        Next
    End Sub
End Class