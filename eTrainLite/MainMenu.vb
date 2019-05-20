Public Class MainMenu

    Private Sub btnImport_Click(sender As System.Object, e As System.EventArgs) Handles btnImport.Click
        Dim impForm As New MainForm
        Me.Hide()
        impForm.Show()

    End Sub

    Private Sub FastToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles FastToolStripMenuItem.Click
        GlobalVariables.eTrain.Location = "MIDLAND"
        GlobalVariables.eTrain.Team = "FAST"
        btnImport.Enabled = True
    End Sub

    Private Sub HighResToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles HighResToolStripMenuItem.Click
        GlobalVariables.eTrain.Location = "MIDLAND"
        GlobalVariables.eTrain.Team = "HR"
        btnImport.Enabled = True
    End Sub

    Private Sub ChromToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ChromToolStripMenuItem.Click
        GlobalVariables.eTrain.Location = "MIDLAND"
        GlobalVariables.eTrain.Team = "CHROM"
        btnImport.Enabled = True
    End Sub

    Private Sub MainMenu_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnReport_Click(sender As System.Object, e As System.EventArgs) Handles btnReport.Click

    End Sub
End Class