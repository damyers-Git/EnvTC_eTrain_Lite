Public Class SampleEdit

    Private Sub SampleEdit_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim aSample As Sample
        GlobalVariables.ContinueTransfer = False
        GlobalVariables.ContinueReport = False
        For Each aSample In GlobalVariables.SampleList 'WT 10/13/2017 -> Changed list type to GV.SampleList (Original = ReportSamList)
            dgSamples.Rows.Add(aSample.UniqueID, aSample.LimsID, aSample.Name, aSample.Type, aSample.Include)
        Next


    End Sub
    Private Sub Me_FormClosing(sender As Object, e As FormClosingEventArgs) _
     Handles Me.FormClosing

        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            dgSamples.Rows.Clear()
            Me.Hide()
        End If

    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        Dim aSample As Sample

        For Each r In dgSamples.Rows
            For Each aSample In GlobalVariables.ReportSamList
                If aSample.UniqueID = r.Cells.Item(0).Value Then
                    aSample.Name = r.Cells.Item(2).Value
                    aSample.Type = r.Cells.Item(3).Value
                    aSample.Include = r.Cells.Item(4).Value
                End If
            Next
        Next
        GlobalVariables.ContinueTransfer = True
        GlobalVariables.ContinueReport = True
        Me.Close()

    End Sub

    Private Sub btnBack_Click(sender As System.Object, e As System.EventArgs) Handles btnBack.Click
        GlobalVariables.ContinueTransfer = False
        GlobalVariables.ContinueReport = False
        dgSamples.Rows.Clear()
        Me.Close()
    End Sub

    Private Sub btnMerge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMerge.Click
        Dim aSample As Sample
        For Each r In dgSamples.SelectedRows
            For Each aSample In GlobalVariables.ReportSamList
                If aSample.UniqueID = r.Cells.Item(0).Value Then
                    aSample.Name = r.Cells.Item(2).Value
                    aSample.Type = r.Cells.Item(3).Value
                    aSample.Include = r.Cells.Item(4).Value
                    With SampleMerge
                        .cboSampleList.Items.Insert(0, aSample.Name)
                    End With
                End If
            Next
        Next
        With SampleMerge
            .cboSampleList.SelectedIndex = 0
        End With

        SampleMerge.ShowDialog()

        'Update sample edit window
        dgSamples.Rows.Clear()
        For Each aSample In GlobalVariables.ReportSamList
            dgSamples.Rows.Add(aSample.UniqueID, aSample.LimsID, aSample.Name, aSample.Type, aSample.Include)
        Next
    End Sub
End Class