Public Class CalcDetails
    Dim sampCount As Integer

    Private Sub btnContinue_Click(sender As System.Object, e As System.EventArgs) Handles btnContinue.Click
        Save()
        DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CalcDetails_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        sampCount = 1

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                UpdateForm()
            End If
        End If

    End Sub

    Private Sub btnNext_Click(sender As System.Object, e As System.EventArgs) Handles btnNext.Click
        If Save() Then
            sampCount = sampCount + 1
            UpdateForm()
        End If

    End Sub

    Private Sub btnBack_Click(sender As System.Object, e As System.EventArgs) Handles btnBack.Click
        If Save() Then
            sampCount = sampCount - 1
            UpdateForm()
        End If

    End Sub

    Private Function Save() As Boolean
        Dim aSample As Sample
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                aSample = GlobalVariables.SampleList(sampCount - 1)
                'Save new compound information
                If IsNumeric(txtInjAmt.Text) Then
                    aSample.MidFInjAmt = CDbl(txtInjAmt.Text)
                    Return True
                Else
                    MsgBox("Value must be numeric, please re-enter Injection Amt!", MsgBoxStyle.Critical)
                    Return False
                End If
            End If
        End If
        Return False
    End Function

    Private Sub UpdateForm()
        Dim aSample As Sample

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                aSample = GlobalVariables.SampleList(sampCount - 1)
                'Fill out form
                Label2.Text = aSample.Name
                txtInjAmt.Text = CStr(aSample.MidFInjAmt)
                'Update buttons
                If sampCount + 1 <= GlobalVariables.SampleList.Count Then
                    btnNext.Visible = True
                Else
                    btnNext.Visible = False
                End If
                If sampCount - 1 <> 0 Then
                    btnBack.Visible = True
                Else
                    btnBack.Visible = False
                End If
            End If
        End If
    End Sub
End Class