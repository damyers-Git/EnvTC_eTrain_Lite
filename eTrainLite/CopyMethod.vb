Public Class CopyMethod

    Private Sub CopyMethod_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim aMethod As Method

        'Load in methods
        For Each aMethod In GlobalVariables.MethodList
            cboCopyFrom.Items.Add(aMethod.Name)
        Next

    End Sub

    Private Sub btnCopy_Click(sender As System.Object, e As System.EventArgs) Handles btnCopy.Click
        Dim aNewMethod As Method
        Dim aInstrument As mInstrument
        Dim aRefBook As RefBook
        Dim aMethod As Method
        Dim curDate As Date

        'Check to make sure method doesn't already exist
        For Each aMethod In GlobalVariables.MethodList
            If UCase(txtNMethod.Text) = UCase(aMethod.Name) Then
                MsgBox("New Method cannot have same Name as existing Method!", MsgBoxStyle.Exclamation, "eTrain")
                Exit Sub
            End If
        Next

        'Check to make sure names are not the same
        If cboCopyFrom.Text = txtNMethod.Text Then
            MsgBox("New Method cannot have same Name as Copy From Method!", MsgBoxStyle.Exclamation, "eTrain")
        Else
            curDate = DateTime.Now
            For Each aMethod In GlobalVariables.MethodList
                If cboCopyFrom.Text = aMethod.Name Then
                    If Not aMethod.Loaded Then
                        GlobalVariables.Method.LoadMethod(aMethod.Name)
                    End If
                    aNewMethod = New Method
                    aNewMethod.Name = txtNMethod.Text
                    'Always set to false
                    aNewMethod.CreatedDate = CDate(curDate.Month & "/" & curDate.Day & "/" & curDate.Year)
                    For Each aRefBook In aMethod.RefBookList
                        aNewMethod.RefBookList.Add(aRefBook)
                    Next
                    For Each aInstrument In aMethod.mInstrumentList
                        aNewMethod.mInstrumentList.Add(aInstrument)
                    Next
                    aNewMethod.Loaded = True
                    GlobalVariables.MethodList.Add(aNewMethod)
                    If GlobalVariables.Method.SaveMethod(aNewMethod) Then
                        MsgBox("Copy/Creation successful!", MsgBoxStyle.Information, "eTrain")
                    End If
                    Exit For
                End If
            Next

            'Update editMethods form
            With EditMethods
                .cboOption1.Items.Clear()
                For Each aMethod In GlobalVariables.MethodList
                    .cboOption1.Items.Add(aMethod.Name)
                Next
                .cboOption2.Items.Clear()
                .dgCompound.Rows.Clear()
                .cboOption2.Text = ""
                .cboOption3.Text = ""
               
                .txtOption4.Text = ""
                .txtOption5.Text = ""
                .cboOption2.Enabled = True
                .cboOption3.Enabled = False
            End With

        End If
    End Sub
End Class