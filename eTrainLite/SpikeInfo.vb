Public Class SpikeInfo

    Private Sub btnCalc_Click(sender As System.Object, e As System.EventArgs) Handles btnCalc.Click
        Dim aSample As Sample
        For Each aSample In GlobalVariables.ReportSamList
            If aSample.Name = lblSampleName.Text.Substring(8, lblSampleName.Text.Length - 8) And aSample.SpikeCalculated = False Then
                aSample.ChromSpikeAmt = CStr(CDbl(txtVol.Text) * CDbl(txtConc.Text) * CDbl(aSample.DilutionFactor))
                aSample.SpikeCalculated = True
                txtConc.Text = ""
                txtVol.Text = ""
                Me.Hide()
                Exit For
            End If
        Next
    End Sub

    Private Sub btnTip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTip.Click
        For Each aSample In GlobalVariables.ReportSamList
            If aSample.Name = lblSampleName.Text.Substring(8, lblSampleName.Text.Length - 8) And aSample.SpikeCalculated = False Then
                MsgBox("Calculation is Spike Vol (" & txtVol.Text & ") * Concentration (" & txtConc.Text & ") * Dilution Factor (" & aSample.DilutionFactor & ") = Corrected Spike Amount in Sample (" & CStr(CDbl(txtVol.Text) * CDbl(txtConc.Text) * CDbl(aSample.DilutionFactor)) & ")")
            End If
        Next

        'MsgBox("Calculation is Spike Vol * Concentration * Dilution Factor = Corrected Spike Amount in Sample")
    End Sub

    Private Sub SpikeInfo_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                txtVol.Text = ".050"
                txtConc.Text = "1000"
            End If
        End If
    End Sub
    Private Sub Me_FormClosing(sender As Object, e As FormClosingEventArgs) _
     Handles Me.FormClosing

        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            
        End If

    End Sub
End Class