Imports System.IO
Imports System.Text.RegularExpressions
Imports Syncfusion.XlsIO

Public Class NewMethodImport

    Dim aMethod As Method
    Dim aInstrument As mInstrument
    Dim strStdPrepLoc As String
    Dim strCalIntStdLoc As String

    Private Sub btnImport_Click(sender As System.Object, e As System.EventArgs) Handles btnImport.Click
        Dim curDate As Date
        Dim aMethod As Method
        Dim aInstrument As mInstrument


        curDate = DateTime.Now
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                If txtMethodName.Text <> "" Then
                    If txtInstrument.Text <> "" Then
                        If radChem.Checked Then
                            'Chemstation import
                            If File.Exists(txt1.Text) Then
                                If File.Exists(txt2.Text) Then
                                    'Create method and instrument based off form
                                    aMethod = New Method
                                    aInstrument = New mInstrument
                                    aMethod.Name = txtMethodName.Text
                                    aMethod.CreatedDate = curDate.Month & "/" & curDate.Day & "/" & curDate.Year
                                    aInstrument.Name = txtInstrument.Text
                                    aInstrument.Reviewed = False
                                    aInstrument.ReviewedDate = "1/1/1970"
                                    aMethod.mInstrumentList.Add(aInstrument)
                                    'Import Calrpt.txt
                                    ChemCalrptImport(aMethod, txtInstrument.Text, txt1.Text)
                                    'Import daprtmth.txt
                                    ChemDaprtmthImport(aMethod, txtInstrument.Text, txt2.Text)
                                    'Save Method
                                    If GlobalVariables.Method.SaveMethod(aMethod) Then
                                        'Save successful
                                        MsgBox("Import Successful", MsgBoxStyle.Information, "eTrain 2.0")
                                    End If
                                Else
                                    MsgBox("Daprtmth.txt cannot be found", MsgBoxStyle.Exclamation, "eTrain 2.0")
                                    txt2.Focus()
                                End If
                            Else
                                MsgBox("CalRpt.txt cannot be found", MsgBoxStyle.Exclamation, "eTrain 2.0")
                                txt1.Focus()
                            End If
                        ElseIf radMH.Checked Then
                            MsgBox("This import is not yet available.", MsgBoxStyle.Exclamation, "eTrain 2.0")
                        ElseIf radXls.Checked Then
                            'Excel Import
                            If File.Exists(txt1.Text) Then
                                'Create method and instrument based off form
                                aMethod = New Method
                                aInstrument = New mInstrument
                                aMethod.Name = txtMethodName.Text
                                aMethod.CreatedDate = curDate.Month & "/" & curDate.Day & "/" & curDate.Year
                                aInstrument.Name = txtInstrument.Text
                                aInstrument.Reviewed = False
                                aInstrument.ReviewedDate = "1/1/1970"
                                aMethod.mInstrumentList.Add(aInstrument)
                                strCalIntStdLoc = txt1.Text
                                XlsCalIntStdSumImport(aMethod)
                                'Save Method
                                If GlobalVariables.Method.SaveMethod(aMethod) Then
                                    'Save successful
                                    MsgBox("Import Successful", MsgBoxStyle.Information, "eTrain 2.0")
                                    Me.Close()
                                End If
                            Else
                                MsgBox("CalIntStdSum.xls cannot be found", MsgBoxStyle.Exclamation, "eTrain 2.0")
                                txt1.Focus()
                            End If
                        End If
                    Else
                        MsgBox("Please enter an Instrument Name", MsgBoxStyle.Exclamation, "eTrain 2.0")
                        txtInstrument.Focus()
                    End If
                Else
                    MsgBox("Please enter a Method Name", MsgBoxStyle.Exclamation, "eTrain 2.0")
                    txtMethodName.Focus()
                End If
            End If
        End If


    End Sub

    Private Sub XlsCalIntStdSumImport(ByRef aMethod As Method)
        Dim exEngine As New ExcelEngine
        Dim exApp As IApplication
        Dim workbook As IWorkbook
        Dim worksheet As IWorksheet
        Dim aCompound As mCompound
        Dim aStandard As mStandard
        Dim aInstrument As mInstrument
        Dim c As Integer

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                Try
                    For Each aInstrument In aMethod.mInstrumentList
                        If aInstrument.Name = txtInstrument.Text Then
                            exApp = exEngine.Excel
                            workbook = exApp.Workbooks.Open(strCalIntStdLoc)
                            worksheet = workbook.Worksheets(0)
                            'Get first standard
                            c = 0
                            Do Until worksheet.Range("A" & c + 4).Value = ""
                                aStandard = New mStandard
                                If InStr(worksheet.Range("A" & c + 4).Value, "(INJ)", CompareMethod.Text) Then
                                    aStandard.Name = Trim(worksheet.Range("A" & c + 4).Value)
                                    aStandard.Type = "Inj"
                                    If worksheet.Range("B" & c + 4).HasFormulaNumberValue Then
                                        aStandard.AvgArea = worksheet.Range("B" & c + 4).FormulaNumberValue
                                    Else
                                        aStandard.AvgArea = worksheet.Range("B" & c + 4).Value
                                    End If
                                Else
                                    aStandard.Name = Trim(worksheet.Range("A" & c + 4).Value)
                                    aStandard.Type = "13C"
                                    If worksheet.Range("B" & c + 4).HasFormulaNumberValue Then
                                        aStandard.AvgArea = worksheet.Range("B" & c + 4).FormulaNumberValue
                                    Else
                                        aStandard.AvgArea = worksheet.Range("B" & c + 4).Value
                                    End If
                                End If
                                aInstrument.mStandardList.Add(aStandard)
                                c += 1
                            Loop
                            'loop until compounds
                            Do Until worksheet.Range("A" & c + 4).Value <> ""
                                c += 1
                            Loop
                            'Account for heading and get next cell for first compound
                            c += 1
                            Do Until worksheet.Range("A" & c + 4).Value = ""
                                aCompound = New mCompound
                                aCompound.Name = Trim(worksheet.Range("A" & c + 4).Value)
                                aCompound.RRF = worksheet.Range("B" & c + 4).Value
                                aCompound.RSD = worksheet.Range("C" & c + 4).Value
                                aCompound.MaxPeakArea = worksheet.Range("D" & c + 4).Value
                                aInstrument.mCompoundList.Add(aCompound)
                                c += 1
                            Loop
                        End If
                    Next
                Catch ex As Exception
                    MsgBox("Error reading CalIntStdSum.xls" & vbCrLf & _
                     "Line: " & vbCrLf & _
                     "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                End Try
            End If
        End If
    End Sub

    Public Sub ChemCalrptImport(ByRef aMethod As Method, ByVal strInstrument As String, ByVal strCalRptLoc As String)
        Dim sr As StreamReader
        Dim aCompound As mCompound
        Dim aInstrument As mInstrument
        Dim line As String
        Dim strAssoc13C As String
        Dim arrSplText() As String


        line = ""
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                'Read calrpt until you hit the first line of "-----"
                Try
                    strAssoc13C = ""
                    For Each aInstrument In aMethod.mInstrumentList
                        If aInstrument.Name = strInstrument Then
                            sr = New StreamReader(strCalRptLoc)
                            line = sr.ReadLine
                            Do Until InStr(line, "------", CompareMethod.Binary)
                                line = Trim(sr.ReadLine)
                            Loop
                            line = sr.ReadLine
                            Do Until sr.EndOfStream
                                arrSplText = Regex.Split(line, "\s+")
                                If UBound(arrSplText) > 0 Then
                                    If InStr(line, "(INJ)", CompareMethod.Text) = 0 And InStr(line, "ISTD", CompareMethod.Text) = 0 Then
                                        aCompound = New mCompound
                                        'Assign associated 13c
                                        aCompound.Assoc13C = strAssoc13C
                                        'see which spot name is in
                                        If arrSplText(1).Length > 1 Then
                                            aCompound.Name = arrSplText(1)
                                        Else
                                            aCompound.Name = arrSplText(2)
                                        End If
                                        aCompound.RSD = arrSplText(UBound(arrSplText))
                                        'see if there is a 2 character note between rsd and rrf
                                        If arrSplText(UBound(arrSplText) - 1).Length = 2 Then
                                            aCompound.RRF = arrSplText(UBound(arrSplText) - 2)
                                        Else
                                            aCompound.RRF = arrSplText(UBound(arrSplText) - 1)
                                        End If
                                        aInstrument.mCompoundList.Add(aCompound)
                                    ElseIf InStr(line, "ISTD", CompareMethod.Text) Then
                                        'see which spot name is in
                                        If arrSplText(1).Length > 1 Then
                                            strAssoc13C = arrSplText(1)
                                        Else
                                            strAssoc13C = arrSplText(2)
                                        End If
                                    End If
                                End If
                                line = Trim(sr.ReadLine)
                            Loop
                        End If
                    Next
                Catch ex As Exception
                    MsgBox("Error reading Calrpt.txt" & vbCrLf & _
                      "Line: " & line & vbCrLf & _
                      "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                End Try
            End If
        End If
    End Sub

    Public Sub ChemDaprtmthImport(ByRef aMethod As Method, ByVal strInstrument As String, ByVal strDaprtmthLoc As String)
        Dim sr As StreamReader
        Dim aCompound As mCompound
        Dim aStandard As mStandard
        Dim line As String
        Dim dblArea As Double
        Dim strAmt As String
        Dim c As Integer

        Dim aInstrument As mInstrument
        Dim arrSplText() As String

        line = ""
        strAmt = ""
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                Try
                    For Each aInstrument In aMethod.mInstrumentList
                        If aInstrument.Name = strInstrument Then
                            sr = New StreamReader(strDaprtmthLoc)
                            line = sr.ReadLine
                            Do Until InStr(line, "Compound Information", CompareMethod.Text)
                                line = sr.ReadLine
                            Loop
                            line = sr.ReadLine
                            line = sr.ReadLine
                            line = sr.ReadLine
                            Do Until sr.EndOfStream
                                dblArea = 0
                                If line.Length > 2 Then
                                    If line.Substring(2, 1) = ")" Then
                                        If InStr(line, "(INJ)", CompareMethod.Text) Then
                                            aStandard = New mStandard
                                            aStandard.Name = Trim(line.Substring(5, 40))
                                            aStandard.Type = "Inj"
                                            Do Until InStr(line, "Signal", CompareMethod.Text)
                                                line = sr.ReadLine
                                            Loop
                                            line = sr.ReadLine
                                            If InStr(line, "Tgt") Then
                                                aStandard.IonTarget = line.Substring(5, 6)
                                                aStandard.AbundTarget = "100"
                                            End If
                                            line = sr.ReadLine
                                            If InStr(line, "Q1") Then
                                                aStandard.IonQual = line.Substring(5, 6)
                                                aStandard.AbundQual = line.Substring(15, 6)
                                            End If
                                            Do Until InStr(line, "Lvl ID", CompareMethod.Text)
                                                line = sr.ReadLine
                                            Loop
                                            line = sr.ReadLine
                                            c = 0
                                            Do Until line = "" Or InStr(line, "not used for this compound", CompareMethod.Text)
                                                arrSplText = Regex.Split(line, "\s+")
                                                If UBound(arrSplText) = 2 Then
                                                    strAmt = arrSplText(1)
                                                    dblArea += CDbl(arrSplText(2))
                                                    c += 1
                                                End If
                                                line = sr.ReadLine
                                            Loop
                                            aStandard.AvgArea = CStr(Math.Round((dblArea / c), 3))
                                            aStandard.CalAmt = strAmt
                                            aInstrument.mStandardList.Add(aStandard)
                                        ElseIf line.Substring(5, 3) = "13C" Then
                                            aStandard = New mStandard
                                            aStandard.Name = Trim(line.Substring(5, 40))
                                            aStandard.Type = "13C"
                                            Do Until InStr(line, "Signal", CompareMethod.Text)
                                                line = sr.ReadLine
                                            Loop
                                            line = sr.ReadLine
                                            If InStr(line, "Tgt") Then
                                                aStandard.IonTarget = line.Substring(5, 6)
                                                If line.Substring(15, 6) = "      " Then
                                                    aStandard.AbundTarget = "100"
                                                Else
                                                    aStandard.AbundTarget = line.Substring(15, 6)
                                                End If

                                            End If
                                            line = sr.ReadLine
                                            If InStr(line, "Q1") Then
                                                aStandard.IonQual = line.Substring(5, 6)
                                                aStandard.AbundQual = line.Substring(15, 6)
                                            End If
                                            Do Until InStr(line, "Lvl ID", CompareMethod.Text)
                                                line = sr.ReadLine
                                            Loop
                                            line = sr.ReadLine
                                            c = 0
                                            Do Until line = "" Or InStr(line, "not used for this compound", CompareMethod.Text)
                                                arrSplText = Regex.Split(line, "\s+")
                                                If UBound(arrSplText) = 2 Then
                                                    strAmt = arrSplText(1)
                                                    dblArea += CDbl(arrSplText(2))
                                                    c += 1
                                                End If
                                                line = sr.ReadLine
                                            Loop
                                            aStandard.CalAmt = strAmt
                                            aStandard.AvgArea = CStr(Math.Round((dblArea / c), 3))
                                            aInstrument.mStandardList.Add(aStandard)
                                        Else
                                            For Each aCompound In aInstrument.mCompoundList
                                                If aCompound.Name = Trim(line.Substring(5, 40)) Then
                                                    Do Until InStr(line, "Signal", CompareMethod.Text)
                                                        line = sr.ReadLine
                                                    Loop
                                                    line = sr.ReadLine
                                                    If InStr(line, "Tgt") Then
                                                        aCompound.Ion = line.Substring(5, 6)
                                                        aCompound.Abundance = line.Substring(15, 6)
                                                    End If
                                                    line = sr.ReadLine
                                                    Do Until InStr(line, "Lvl ID", CompareMethod.Text)
                                                        line = sr.ReadLine
                                                    Loop
                                                    line = sr.ReadLine
                                                    Do Until line = "" Or InStr(line, "not used for this compound", CompareMethod.Text)
                                                        arrSplText = Regex.Split(line, "\s+")
                                                        If UBound(arrSplText) = 2 Then
                                                            dblArea = CDbl(arrSplText(2))
                                                        End If
                                                        line = sr.ReadLine
                                                    Loop
                                                    aCompound.MaxPeakArea = CStr(dblArea)
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                End If
                                line = sr.ReadLine
                            Loop
                        End If
                        'End If
                    Next
                Catch ex As Exception
                    MsgBox("Error reading Daprtmth.txt" & vbCrLf & _
                      "Line: " & line & vbCrLf & _
                      "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                End Try

            End If
        End If
    End Sub

    Private Sub txtStdBookInfo_TextChanged(sender As System.Object, e As System.EventArgs)
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                Me.Text = ""
            End If
        End If

    End Sub

    Private Sub txtInjBookInfo_TextChanged(sender As System.Object, e As System.EventArgs)
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                Me.Text = ""
            End If
        End If
    End Sub

    Private Sub txtLCSBookInfo_TextChanged(sender As System.Object, e As System.EventArgs)
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                Me.Text = ""
            End If
        End If
    End Sub

    Private Sub radXls_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles radXls.CheckedChanged
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                If radXls.Checked Then
                    Label1.Text = "CalIntStdSum.xls:"
                    Label2.Text = ""
                    txt1.Visible = True
                    txt2.Visible = False
                    btn2.Visible = False
                    lblNote.Visible = True
                End If
            End If
        End If

    End Sub

    Private Sub radChem_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles radChem.CheckedChanged
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                If radChem.Checked Then
                    Label1.Text = "Calrpt.txt:"
                    Label2.Text = "Daprtmth.txt:"
                    txt1.Visible = True
                    txt2.Visible = True
                    btn2.Visible = True
                    lblNote.Visible = False
                End If
            End If
        End If
    End Sub

    Private Sub btn1_Click(sender As System.Object, e As System.EventArgs) Handles btn1.Click
        Dim fd As New OpenFileDialog()
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                fd.Title = "Open File Dialog"
                fd.InitialDirectory = "C:\Users\u411882\Desktop\eTrain20 TestFolder\"
                fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
                fd.FilterIndex = 2
                fd.RestoreDirectory = True
                If fd.ShowDialog() = DialogResult.OK Then
                    txt1.Text = fd.FileName
                End If
            End If
        End If
    End Sub

    Private Sub btn2_Click(sender As System.Object, e As System.EventArgs) Handles btn2.Click
        Dim fd As New OpenFileDialog()
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                fd.Title = "Open File Dialog"
                fd.InitialDirectory = "C:\Users\u411882\Desktop\eTrain20 TestFolder\"
                fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
                fd.FilterIndex = 2
                fd.RestoreDirectory = True
                If fd.ShowDialog() = DialogResult.OK Then
                    txt2.Text = fd.FileName
                End If
            End If
        End If
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim aMethod As Method

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                With EditMethods
                    EditMethods.EditForm()
                    GlobalVariables.Method.LoadMethodNames()
                    .cboOption1.Items.Clear()
                    For Each aMethod In GlobalVariables.MethodList
                        .cboOption1.Items.Add(aMethod.Name)
                    Next
                End With
            End If
        End If

    End Sub

    Private Sub Me_FormClosing(sender As Object, e As FormClosingEventArgs) _
     Handles Me.FormClosing

        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            txt1.Text = ""
            txt2.Text = ""
            txtInstrument.Text = ""
            txtMethodName.Text = ""
            Me.Hide()
        End If

    End Sub
End Class