Public Class RefBookEdit
    Dim intIndex As Integer
    Dim intTotalBooks As Integer

    Private Function NoEntry() As Boolean

        If txtName.Text = "" Then
            Return True
        End If
        If txtExp.Text = "" Then
            Return True
        End If
        If txtNotes.Text = "" Then
            txtNotes.Text = "NONE"
        End If

        Return False

    End Function

    Private Sub Me_FormClosing(sender As Object, e As FormClosingEventArgs) _
     Handles Me.FormClosing

        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            cboType.Items.Clear()
            txtExp.Text = ""
            txtName.Text = ""
            txtNotes.Text = ""
            Me.Hide()
        End If

    End Sub


    Private Sub RefBookEdit_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim aRefBook As RefBook
        Dim aMethod As Method

        'Load in Reference books for selected method
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                If Not EditMethods.EditCurrentMethodToolStripMenuItem.Checked Then
                    txtName.Enabled = False
                    cboType.Enabled = False
                    txtExp.Enabled = False
                    txtNotes.Enabled = False
                    btnSave.Enabled = False
                    NewToolStripMenuItem.Enabled = False
                End If
                For Each aMethod In GlobalVariables.MethodList
                    If aMethod.Name = GlobalVariables.selMethod.Name Then
                        If aMethod.RefBookList.Count <= 1 Then
                            btnNext.Enabled = False
                        End If
                        If aMethod.RefBookList.Count > 0 Then
                            aRefBook = aMethod.RefBookList.Item(0)
                            intIndex = 0
                            intTotalBooks = aMethod.RefBookList.Count
                            txtName.Text = aRefBook.Name
                            If aRefBook.Type = "13C" Then
                                cboType.SelectedIndex = 0
                            ElseIf aRefBook.Type = "Injection" Then
                                cboType.SelectedIndex = 1
                            ElseIf aRefBook.Type = "LCS" Then
                                cboType.SelectedIndex = 2
                            End If
                            txtExp.Text = aRefBook.Expiration
                            txtNotes.Text = aRefBook.Note
                        Else
                            aRefBook = New RefBook
                            aRefBook.Type = "13C"
                            aMethod.RefBookList.Add(aRefBook)
                            intIndex = 0
                            intTotalBooks = aMethod.RefBookList.Count
                            txtName.Text = aRefBook.Name
                            cboType.SelectedIndex = 0
                            txtExp.Text = aRefBook.Expiration
                            txtNotes.Text = aRefBook.Note
                            btnNext.Enabled = False
                        End If
                    End If
                Next
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                If Not EditMethods.EditCurrentMethodToolStripMenuItem.Checked Then
                    txtName.Enabled = False
                    cboType.Enabled = False
                    txtExp.Enabled = False
                    txtNotes.Enabled = False
                    btnSave.Enabled = False
                    NewToolStripMenuItem.Enabled = False
                End If
                For Each aMethod In GlobalVariables.MethodList
                    If aMethod.Name = GlobalVariables.selMethod.Name Then
                        If aMethod.RefBookList.Count <= 1 Then
                            btnNext.Enabled = False
                        End If
                        If aMethod.RefBookList.Count > 0 Then
                            aRefBook = aMethod.RefBookList.Item(0)
                            intIndex = 0
                            intTotalBooks = aMethod.RefBookList.Count
                            txtName.Text = aRefBook.Name
                            If aRefBook.Type = "13C" Then
                                cboType.SelectedIndex = 0
                            ElseIf aRefBook.Type = "Injection" Then
                                cboType.SelectedIndex = 1
                            ElseIf aRefBook.Type = "LCS" Then
                                cboType.SelectedIndex = 2
                            End If
                            txtExp.Text = aRefBook.Expiration
                            txtNotes.Text = aRefBook.Note
                        Else
                            aRefBook = New RefBook
                            aRefBook.Type = "13C"
                            aMethod.RefBookList.Add(aRefBook)
                            intIndex = 0
                            intTotalBooks = aMethod.RefBookList.Count
                            txtName.Text = aRefBook.Name
                            cboType.SelectedIndex = 0
                            txtExp.Text = aRefBook.Expiration
                            txtNotes.Text = aRefBook.Note
                            btnNext.Enabled = False
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub btnPrev_Click(sender As System.Object, e As System.EventArgs) Handles btnPrev.Click
        Dim aRefBook As RefBook
        Dim aMethod As Method

        'Go back
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                'Check for error or lack of entry
                If NoEntry Then
                    MsgBox("Make sure everything is filled out before trying clicking off this Standard Book.", MsgBoxStyle.Exclamation, "eTrain 2.0")
                    Exit Sub
                Else
                    For Each aMethod In GlobalVariables.MethodList
                        If aMethod.Name = GlobalVariables.selMethod.Name Then
                            aRefBook = aMethod.RefBookList.Item(intIndex)
                            aRefBook.Name = txtName.Text
                            aRefBook.Type = cboType.Text
                            aRefBook.Note = txtNotes.Text
                            aRefBook.Expiration = CDate(txtExp.Text)
                            intIndex = intIndex - 1
                            aRefBook = aMethod.RefBookList.Item(intIndex)
                            txtName.Text = aRefBook.Name
                            If aRefBook.Type = "13C" Then
                                cboType.SelectedIndex = 0
                            ElseIf aRefBook.Type = "Injection" Then
                                cboType.SelectedIndex = 1
                            ElseIf aRefBook.Type = "LCS" Then
                                cboType.SelectedIndex = 2
                            End If
                            txtExp.Text = aRefBook.Expiration
                            txtNotes.Text = aRefBook.Note
                        End If
                    Next
                    If intIndex = 0 Then
                        btnPrev.Enabled = False
                    End If
                    btnNext.Enabled = True
                End If
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                'Check for error or lack of entry
                If NoEntry() Then
                    MsgBox("Make sure everything is filled out before trying clicking off this Standard Book.", MsgBoxStyle.Exclamation, "eTrain 2.0")
                    Exit Sub
                Else
                    For Each aMethod In GlobalVariables.MethodList
                        If aMethod.Name = GlobalVariables.selMethod.Name Then
                            aRefBook = aMethod.RefBookList.Item(intIndex)
                            aRefBook.Name = txtName.Text
                            aRefBook.Type = cboType.Text
                            aRefBook.Note = txtNotes.Text
                            aRefBook.Expiration = CDate(txtExp.Text)
                            intIndex = intIndex - 1
                            aRefBook = aMethod.RefBookList.Item(intIndex)
                            txtName.Text = aRefBook.Name
                            If aRefBook.Type = "13C" Then
                                cboType.SelectedIndex = 0
                            ElseIf aRefBook.Type = "Injection" Then
                                cboType.SelectedIndex = 1
                            ElseIf aRefBook.Type = "LCS" Then
                                cboType.SelectedIndex = 2
                            End If
                            txtExp.Text = aRefBook.Expiration
                            txtNotes.Text = aRefBook.Note
                        End If
                    Next
                    If intIndex = 0 Then
                        btnPrev.Enabled = False
                    End If
                    btnNext.Enabled = True
                End If
            End If
        End If

    End Sub

    Private Sub btnNext_Click(sender As System.Object, e As System.EventArgs) Handles btnNext.Click
        Dim aRefBook As RefBook
        Dim aMethod As Method

        'Go forward
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                If NoEntry() Then
                    MsgBox("Make sure everything is filled out before trying clicking off this Standard Book.", MsgBoxStyle.Exclamation, "eTrain 2.0")
                    Exit Sub
                Else
                    For Each aMethod In GlobalVariables.MethodList
                        If aMethod.Name = GlobalVariables.selMethod.Name Then
                            aRefBook = aMethod.RefBookList.Item(intIndex)
                            aRefBook.Name = txtName.Text
                            aRefBook.Type = cboType.Text
                            aRefBook.Note = txtNotes.Text
                            aRefBook.Expiration = CDate(txtExp.Text)
                            intIndex = intIndex + 1
                            aRefBook = aMethod.RefBookList.Item(intIndex)
                            txtName.Text = aRefBook.Name
                            If aRefBook.Type = "13C" Then
                                cboType.SelectedIndex = 0
                            ElseIf aRefBook.Type = "Injection" Then
                                cboType.SelectedIndex = 1
                            ElseIf aRefBook.Type = "LCS" Then
                                cboType.SelectedIndex = 2
                            End If
                            txtExp.Text = aRefBook.Expiration
                            txtNotes.Text = aRefBook.Note
                        End If
                    Next
                    If intIndex >= (intTotalBooks - 1) Then
                        btnNext.Enabled = False
                    End If
                    btnPrev.Enabled = True
                End If
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                If NoEntry() Then
                    MsgBox("Make sure everything is filled out before trying clicking off this Standard Book.", MsgBoxStyle.Exclamation, "eTrain 2.0")
                    Exit Sub
                Else
                    For Each aMethod In GlobalVariables.MethodList
                        If aMethod.Name = GlobalVariables.selMethod.Name Then
                            aRefBook = aMethod.RefBookList.Item(intIndex)
                            aRefBook.Name = txtName.Text
                            aRefBook.Type = cboType.Text
                            aRefBook.Note = txtNotes.Text
                            aRefBook.Expiration = CDate(txtExp.Text)
                            intIndex = intIndex + 1
                            aRefBook = aMethod.RefBookList.Item(intIndex)
                            txtName.Text = aRefBook.Name
                            If aRefBook.Type = "13C" Then
                                cboType.SelectedIndex = 0
                            ElseIf aRefBook.Type = "Injection" Then
                                cboType.SelectedIndex = 1
                            ElseIf aRefBook.Type = "LCS" Then
                                cboType.SelectedIndex = 2
                            End If
                            txtExp.Text = aRefBook.Expiration
                            txtNotes.Text = aRefBook.Note
                        End If
                    Next
                    If intIndex >= (intTotalBooks - 1) Then
                        btnNext.Enabled = False
                    End If
                    btnPrev.Enabled = True
                End If
            End If
        End If
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        Dim aRefBook As RefBook
        Dim aMethod As Method

        'Save currently showed reference book and go back to method window
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                If NoEntry() Then
                    MsgBox("Make sure everything is filled out before trying clicking off this Standard Book.", MsgBoxStyle.Exclamation, "eTrain 2.0")
                    Exit Sub
                Else
                    For Each aMethod In GlobalVariables.MethodList
                        If aMethod.Name = GlobalVariables.selMethod.Name Then
                            aRefBook = aMethod.RefBookList.Item(intIndex)
                            aRefBook.Name = txtName.Text
                            aRefBook.Type = cboType.Text
                            aRefBook.Note = txtNotes.Text
                            aRefBook.Expiration = CDate(txtExp.Text)
                            txtExp.Text = aRefBook.Expiration
                        End If
                    Next
                End If
                'Load book details back to Edit Methods
               
               
                Me.Close()
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                If NoEntry() Then
                    MsgBox("Make sure everything is filled out before trying clicking off this Standard Book.", MsgBoxStyle.Exclamation, "eTrain 2.0")
                    Exit Sub
                Else
                    For Each aMethod In GlobalVariables.MethodList
                        If aMethod.Name = GlobalVariables.selMethod.Name Then
                            aRefBook = aMethod.RefBookList.Item(intIndex)
                            aRefBook.Name = txtName.Text
                            aRefBook.Type = cboType.Text
                            aRefBook.Note = txtNotes.Text
                            aRefBook.Expiration = CDate(txtExp.Text)
                            txtExp.Text = aRefBook.Expiration
                        End If
                    Next
                End If
                'Load book details back to Edit Methods


                Me.Close()
            End If
        End If
    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles NewToolStripMenuItem.Click
        Dim aRefBook As RefBook
        Dim aMethod As Method

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                'Create new book
                aRefBook = New RefBook
                For Each aMethod In GlobalVariables.MethodList
                    If aMethod.Name = GlobalVariables.selMethod.Name Then
                        aMethod.RefBookList.Add(aRefBook)
                        intTotalBooks = aMethod.RefBookList.Count
                        intIndex = intTotalBooks - 1
                        aRefBook = aMethod.RefBookList.Item(intIndex)
                        txtName.Text = aRefBook.Name
                        cboType.SelectedIndex = -1
                        txtExp.Text = aRefBook.Expiration
                        txtNotes.Text = aRefBook.Note
                        Exit For
                    End If
                Next
                If intIndex >= (intTotalBooks - 1) Then
                    btnNext.Enabled = False
                End If
                btnPrev.Enabled = True
            End If
        End If

    End Sub
End Class