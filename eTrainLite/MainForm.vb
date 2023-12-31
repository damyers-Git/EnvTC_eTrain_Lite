﻿Imports Syncfusion.XlsIO
Imports System.Reflection

Public Class MainForm   'FIX NEXT TIME... SSR NOT BEING SET AS CORRECT TYPE!!!!

    Private Sub btnFindFiles_Click(sender As System.Object, e As System.EventArgs) Handles btnFindFiles.Click
        Dim strImportLoc As String
        Dim arrSpl() As String

        'Set Import Type, in case it was changed
        If GlobalVariables.eTrain.Team = "CLAB" Then
            If GlobalVariables.eTrain.AnalysisLab = "EUROLAN" Then
                If cboImportType.Text = "EUROLAN" Then
                    GlobalVariables.Import.Type = "Full_List"
                ElseIf cboImportType.Text = "GRABS" Then
                    GlobalVariables.Import.Type = "GRABS"
                ElseIf cboImportType.Text = "031A" Then
                    GlobalVariables.Import.Type = "031A"
                ElseIf cboImportType.Text = "031B" Then
                    GlobalVariables.Import.Type = "031B"
                ElseIf cboImportType.Text = "031C" Then
                    GlobalVariables.Import.Type = "031C"
                ElseIf cboImportType.Text = "PESTICIDES" Then
                    GlobalVariables.Import.Type = "PESTICIDES"
                End If
            ElseIf GlobalVariables.eTrain.AnalysisLab = "ALS" Then
                If cboImportType.Text = "ALS" Then
                    GlobalVariables.Import.Type = "Full_List"
                ElseIf cboImportType.Text = "BOD-PHOS" Then
                    GlobalVariables.Import.Type = "BOD-PHOS"
                ElseIf cboImportType.Text = "QUALA_ALS" Then
                    GlobalVariables.Import.Type = "QUALA_ALS"
                ElseIf cboImportType.Text = "DPS_ALS" Then
                    GlobalVariables.Import.Type = "DPS_ALS"
                ElseIf cboImportType.Text = "XALT_ALS" Then
                    GlobalVariables.Import.Type = "XALT_ALS"
                End If
            ElseIf GlobalVariables.eTrain.AnalysisLab = "SGS" Then
                If cboImportType.Text = "SGS" Then
                    GlobalVariables.Import.Type = "Full_List"
                ElseIf cboImportType.Text = "METALS" Then
                    GlobalVariables.Import.Type = "METALS"
                ElseIf cboImportType.Text = "DPS_SGS" Then
                    GlobalVariables.Import.Type = "DPS_SGS"
                End If
            ElseIf GlobalVariables.eTrain.AnalysisLab = "TA" Then
                If cboImportType.Text = "TA" Then
                    GlobalVariables.Import.Type = "Full_List"
                ElseIf cboImportType.Text = "IONS" Then
                    GlobalVariables.Import.Type = "IONS"
                ElseIf cboImportType.Text = "VOA_MONTHLY" Then
                    GlobalVariables.Import.Type = "VOA_MONTHLY"
                ElseIf cboImportType.Text = "SFE_FLUORIDE" Then
                    GlobalVariables.Import.Type = "SFE_FLUORIDE"
                ElseIf cboImportType.Text = "1005_PS" Then
                    GlobalVariables.Import.Type = "1005_PS"
                ElseIf cboImportType.Text = "SVOA_SEMIANNUAL" Then
                    GlobalVariables.Import.Type = "SVOA_SEMIANNUAL"
                ElseIf cboImportType.Text = "PCB_TA" Then
                    GlobalVariables.Import.Type = "PCB_TA"
                ElseIf cboImportType.Text = "CYANIDE" Then
                    GlobalVariables.Import.Type = "CYANIDE"
                ElseIf cboImportType.Text = "XALT_TA" Then
                    GlobalVariables.Import.Type = "XALT_TA"
                ElseIf cboImportType.Text = "CABOT" Then
                    GlobalVariables.Import.Type = "CABOT"
                ElseIf cboImportType.Text = "KENAN_TA" Then
                    GlobalVariables.Import.Type = "KENAN_TA"
                ElseIf cboImportType.Text = "QUALA_TA" Then
                    GlobalVariables.Import.Type = "QUALA_TA"
                ElseIf cboImportType.Text = "SEWER" Then
                    GlobalVariables.Import.Type = "SEWER"
                End If
            ElseIf GlobalVariables.eTrain.AnalysisLab = "FIBERTEC" Then
                If cboImportType.Text = "FIBERTEC" Then
                    GlobalVariables.Import.Type = "Full_List"
                ElseIf cboImportType.Text = "KENAN_FIBERTEC" Then
                    GlobalVariables.Import.Type = "KENAN_FIBERTEC"
                End If
            ElseIf GlobalVariables.eTrain.AnalysisLab = "VISTA" Then
                If cboImportType.Text = "VISTA" Then
                    GlobalVariables.Import.Type = "Full_List"
                ElseIf cboImportType.Text = "DF" Then
                    GlobalVariables.Import.Type = "DF"
                End If
            End If
        ElseIf GlobalVariables.eTrain.Team = "NewSample" Then
            If GlobalVariables.eTrain.AnalysisLab = "TANC_NEW" Then
                If cboImportType.Text = "SEWER" Then
                    GlobalVariables.Import.Type = "SEWER"
                End If
            ElseIf GlobalVariables.eTrain.AnalysisLab = "DeerPark" Then
                If cboImportType.Text = "DeerPark" Then
                    GlobalVariables.Import.Type = "DeerPark"
                End If
            End If
            Else
            MsgBox("Please select an Import Type first", MsgBoxStyle.Exclamation, "eTrain 2.0")
            Exit Sub
        End If

        strImportLoc = "NULL"
        'Dialog window to selection import location
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                strImportLoc = GlobalVariables.eTrain.ChooseFolder("C:\", "Choose the location where the files reside")
            ElseIf GlobalVariables.eTrain.Team = "FAST" Then
                strImportLoc = GlobalVariables.eTrain.ChooseFolder("C:\", "Choose the location where the files reside")
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                strImportLoc = GlobalVariables.eTrain.ChooseFolder("C:\", "Choose the location where the files reside")
            ElseIf GlobalVariables.eTrain.Team = "CLAB" Then
                strImportLoc = GlobalVariables.eTrain.ChooseFolder("C:\", "Choose the location where the files reside")
            ElseIf GlobalVariables.eTrain.Team = "NewSample" Then
                strImportLoc = GlobalVariables.eTrain.ChooseFolder("C:\", "Choose the location where the files reside")
            ElseIf GlobalVariables.eTrain.Team = "AECOM" Then
                strImportLoc = GlobalVariables.eTrain.ChooseFolder("C:\", "Choose the location where the files reside")
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then

            If GlobalVariables.eTrain.Team = "CHROM" Then
                strImportLoc = GlobalVariables.eTrain.ChooseFolder("C:\", "Choose the location where the files reside")
            End If
        ElseIf GlobalVariables.eTrain.Server = "SEADRIFT" Then
            strImportLoc = GlobalVariables.eTrain.ChooseFolder("C:\", "Choose the location where the files reside")
        ElseIf GlobalVariables.eTrain.Server = "ROH" Then
            strImportLoc = GlobalVariables.eTrain.ChooseFolder("C:\", "Choose the location where the files reside")
        End If

        'User never selected a folder
        If strImportLoc = "NULL" Then

            Exit Sub
        End If
        'Search for files
        If GlobalVariables.Import.Type = "CHEM" Then
            GlobalVariables.Import.arrFileList.Clear()
            GlobalVariables.Import.FileSearch(strImportLoc, "epatemp.txt")
            For Each file In GlobalVariables.Import.arrFileList
                arrSpl = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1).Split("\")
                If UBound(arrSpl) <= 2 Then
                    lstFileList.Items.Add(file.ToString)
                Else
                    lstFileList.Items.Add("...\" & arrSpl(UBound(arrSpl) - 2) & "\" & arrSpl(UBound(arrSpl) - 1) & "\" & arrSpl(UBound(arrSpl)))
                End If
            Next
        ElseIf GlobalVariables.Import.Type = "CHEMBEVCAN" Then
            GlobalVariables.Import.arrFileList.Clear()
            GlobalVariables.Import.FileSearch(strImportLoc, "epatemp.txt")
            For Each file In GlobalVariables.Import.arrFileList
                arrSpl = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1).Split("\")
                If UBound(arrSpl) <= 2 Then
                    lstFileList.Items.Add(file.ToString)
                Else
                    lstFileList.Items.Add("...\" & arrSpl(UBound(arrSpl) - 2) & "\" & arrSpl(UBound(arrSpl) - 1) & "\" & arrSpl(UBound(arrSpl)))
                End If
            Next
        ElseIf GlobalVariables.Import.Type = "TOC" Then
            GlobalVariables.Import.arrFileList.Clear()
            GlobalVariables.Import.FileSearch(strImportLoc, "*.txt")
            For Each file In GlobalVariables.Import.arrFileList
                arrSpl = file.ToString.Split("\")
                If UBound(arrSpl) <= 2 Then
                    lstFileList.Items.Add(file.ToString)
                Else
                    lstFileList.Items.Add("...\" & arrSpl(UBound(arrSpl) - 2) & "\" & arrSpl(UBound(arrSpl) - 1) & "\" & arrSpl(UBound(arrSpl)))
                End If
            Next
        ElseIf GlobalVariables.Import.Type = "MASS" Then
            GlobalVariables.Import.arrFileList.Clear()
            GlobalVariables.Import.FileSearch(strImportLoc, "*.xlsx")
            For Each file In GlobalVariables.Import.arrFileList
                arrSpl = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1).Split("\")
                lstFileList.Items.Add(file.ToString)
            Next
        ElseIf GlobalVariables.Import.Type = "TQIII" Then
            GlobalVariables.Import.arrFileList.Clear()
            GlobalVariables.Import.FileSearch(strImportLoc, "*.xls")
            For Each file In GlobalVariables.Import.arrFileList
                arrSpl = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1).Split("\")
                lstFileList.Items.Add(file.ToString)
            Next
        ElseIf GlobalVariables.Import.Type = "EDD" Then
            GlobalVariables.Import.arrFileList.Clear()
            GlobalVariables.Import.FileSearch(strImportLoc, "*.txt")
            GlobalVariables.Import.FileSearch(strImportLoc, "*.DAT")

            GlobalVariables.Import.FileSearch(strImportLoc, "*.xls") '<- added by wmtowne for testing sewer data 1/25/2019

            For Each file In GlobalVariables.Import.arrFileList
                arrSpl = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1).Split("\")
                lstFileList.Items.Add(file.ToString) '"...\" & arrSpl(UBound(arrSpl) - 2) & "\" & arrSpl(UBound(arrSpl) - 1) & "\" & arrSpl(UBound(arrSpl)))
            Next

        ElseIf GlobalVariables.Import.Type = "SSR" Then
            GlobalVariables.Import.arrFileList.Clear()
            GlobalVariables.Import.FileSearch(strImportLoc, "*.xls*") '<- added by wmtowne for testing sewer data 1/25/2019

            For Each file In GlobalVariables.Import.arrFileList
                arrSpl = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1).Split("\")
                lstFileList.Items.Add(file.ToString) '"...\" & arrSpl(UBound(arrSpl) - 2) & "\" & arrSpl(UBound(arrSpl) - 1) & "\" & arrSpl(UBound(arrSpl)))
            Next
        ElseIf GlobalVariables.eTrain.AnalysisLab = "VISTA" Then

            GlobalVariables.Import.arrFileList.Clear()
            GlobalVariables.Import.FileSearch(strImportLoc, "*.xls*")
            GlobalVariables.Import.FileSearch(strImportLoc, "*.xlsx")
            'GlobalVariables.Import.FileSearch(strImportLoc, "*.txt*")
            'GlobalVariables.Import.FileSearch(strImportLoc, "*.dat")

            For Each file In GlobalVariables.Import.arrFileList
                arrSpl = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1).Split("\")
                lstFileList.Items.Add(file.ToString)
            Next
        ElseIf GlobalVariables.eTrain.AnalysisLab = "DeerPark" Then
            GlobalVariables.Import.arrFileList.Clear()
            GlobalVariables.Import.FileSearch(strImportLoc, "*.xls*")
            GlobalVariables.Import.FileSearch(strImportLoc, "*.xlsx")
            GlobalVariables.Import.FileSearch(strImportLoc, "*.txt*")
            GlobalVariables.Import.FileSearch(strImportLoc, "*.dat")

            For Each file In GlobalVariables.Import.arrFileList
                arrSpl = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1).Split("\")
                lstFileList.Items.Add(file.ToString)
            Next

        ElseIf GlobalVariables.eTrain.AnalysisLab <> "VISTA" Then
            GlobalVariables.Import.arrFileList.Clear()
            GlobalVariables.Import.FileSearch(strImportLoc, "*.txt*")
            GlobalVariables.Import.FileSearch(strImportLoc, "*.dat")

            For Each file In GlobalVariables.Import.arrFileList
                arrSpl = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1).Split("\")
                lstFileList.Items.Add(file.ToString)
            Next

        End If
        'Make Import Available
        Me.btnImport.Enabled = True
    End Sub

    Private Sub btnImport_Click(sender As System.Object, e As System.EventArgs) Handles btnImport.Click
        Dim strText As String
        Dim aSample As Sample
        Dim arrSpl() As String
        Dim blnCS1 As Boolean
        Dim blnCS3 As Boolean

        'Begin Import
        'Check for CS1 and CS3, 1 and only 1 of each
        Try
            If GlobalVariables.eTrain.Location = "MIDLAND" Then
                If GlobalVariables.eTrain.Team = "FAST" Then
                    blnCS1 = False
                    blnCS3 = False
                    For Each item In lstFileList.SelectedItems
                        If InStr(item.ToString, "CS1") Then
                            If blnCS1 = True Then
                                MsgBox("More than 1 CS1 has been selected, please select only 1 and try import again.")
                                Exit Sub
                            Else
                                blnCS1 = True
                            End If
                        End If
                        If InStr(item.ToString, "CS3") Then
                            If blnCS3 = True Then
                                MsgBox("More than 1 CS3 has been selected, please select only 1 and try import again.")
                                Exit Sub
                            Else
                                blnCS3 = True
                            End If
                        End If
                    Next
                    If blnCS1 = False Then
                        MsgBox("No CS1 Found, please ensure you have a CS1 selected and try import again.")
                        Exit Sub
                    End If
                    If blnCS3 = False Then
                        MsgBox("No CS3 Found, please ensure you have a CS3 selected and try import again.")
                        Exit Sub
                    End If
                End If
            End If

            If GlobalVariables.Import.Type = "CHEM" Then
                'Grab file path and import each file
                For Each item In lstFileList.SelectedItems
                    For Each file In GlobalVariables.Import.arrFileList
                        arrSpl = item.ToString.Split("\")
                        If InStr(file, item.ToString.Substring(3)) Then
                            GlobalVariables.Import.FilePath = file.ToString
                            GlobalVariables.Import.FolderPath = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1)
                            GlobalVariables.Import.SampleImport()
                        End If
                    Next
                Next
            ElseIf GlobalVariables.Import.Type = "CHEMBEVCAN" Then
                'Grab file path and import each file
                For Each item In lstFileList.SelectedItems
                    For Each file In GlobalVariables.Import.arrFileList
                        arrSpl = item.ToString.Split("\")
                        If InStr(file, arrSpl(UBound(arrSpl))) Then
                            GlobalVariables.Import.FilePath = file.ToString
                            GlobalVariables.Import.FolderPath = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1)
                            GlobalVariables.Import.SampleImport()
                        End If
                    Next
                Next
            ElseIf GlobalVariables.Import.Type = "TOC" Then
                For Each item In lstFileList.SelectedItems
                    For Each file In GlobalVariables.Import.arrFileList
                        arrSpl = item.ToString.Split("\")
                        If InStr(file, arrSpl(UBound(arrSpl))) Then
                            GlobalVariables.Import.FilePath = file.ToString
                            GlobalVariables.Import.FolderPath = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1)
                            GlobalVariables.Import.SampleImport()
                        End If
                    Next
                Next
            ElseIf GlobalVariables.Import.Type = "MASS" Then
                For Each item In lstFileList.SelectedItems
                    For Each file In GlobalVariables.Import.arrFileList
                        arrSpl = item.ToString.Split("\")
                        If InStr(file, arrSpl(UBound(arrSpl))) Then
                            GlobalVariables.Import.FilePath = file.ToString
                            GlobalVariables.Import.FolderPath = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1)
                            GlobalVariables.Import.SampleImport()
                        End If
                    Next
                Next
            ElseIf GlobalVariables.Import.Type = "TQIII" Then
                For Each item In lstFileList.SelectedItems
                    For Each file In GlobalVariables.Import.arrFileList
                        arrSpl = item.ToString.Split("\")
                        If InStr(file, arrSpl(UBound(arrSpl))) Then
                            GlobalVariables.Import.FilePath = file.ToString
                            GlobalVariables.Import.FolderPath = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1)
                            GlobalVariables.Import.SampleImport()
                        End If
                    Next
                Next
            ElseIf GlobalVariables.Import.Type = "EDD" Then
                For Each item In lstFileList.SelectedItems
                    For Each file In GlobalVariables.Import.arrFileList
                        arrSpl = item.ToString.Split("\") 'arrSpl contain file name? ()
                        If InStr(file, arrSpl(UBound(arrSpl))) Then
                            GlobalVariables.Import.FilePath = file.ToString
                            GlobalVariables.Import.FolderPath = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1)
                            GlobalVariables.Import.SampleImport()
                        End If
                    Next
                Next

            ElseIf GlobalVariables.Import.Type = "SSR" Then
                For Each item In lstFileList.SelectedItems
                    For Each file In GlobalVariables.Import.arrFileList
                        arrSpl = item.ToString.Split("\")
                        If InStr(file, arrSpl(UBound(arrSpl))) Then
                            GlobalVariables.Import.FilePath = file.ToString
                            GlobalVariables.Import.FolderPath = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1)
                            GlobalVariables.Import.SampleImport()
                        End If
                    Next
                Next
            ElseIf GlobalVariables.eTrain.Team = "CLAB" Then
                For Each item In lstFileList.SelectedItems
                    For Each file In GlobalVariables.Import.arrFileList
                        arrSpl = item.ToString.Split("\")
                        If InStr(file, arrSpl(UBound(arrSpl))) Then
                            GlobalVariables.Import.FilePath = file.ToString
                            GlobalVariables.Import.FolderPath = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1)
                            GlobalVariables.Import.SampleImport()
                        End If
                    Next
                Next
            ElseIf GlobalVariables.eTrain.Team = "NewSample" Then
                For Each item In lstFileList.SelectedItems
                    For Each file In GlobalVariables.Import.arrFileList
                        arrSpl = item.ToString.Split("\")
                        If InStr(file, arrSpl(UBound(arrSpl))) Then
                            GlobalVariables.Import.FilePath = file.ToString
                            GlobalVariables.Import.FolderPath = file.ToString.Substring(0, InStrRev(file.ToString, "\") - 1)
                            GlobalVariables.Import.SampleImport()
                        End If
                    Next
                Next
            End If
            'CAS no's
            If GlobalVariables.eTrain.Location = "MIDLAND" Then
                If GlobalVariables.eTrain.Team = "CHROM" Then
                    If GlobalVariables.Import.MidlandChromAttachCAS() Then

                    End If
                End If
            End If

            'Display imported information
            lblImportResults.Text = "Import Results: " & GlobalVariables.SampleList.Count & " Samples Imported"
            If GlobalVariables.Import.Type = "TOC" Then
                For Each aSample In GlobalVariables.SampleList
                    strText = strText & "Sample Name: " & aSample.Name & vbCrLf
                    strText = strText & "ID: " & aSample.LimsID & vbCrLf & vbCrLf
                    txtImportResults.Text = strText
                Next
            ElseIf GlobalVariables.Import.Type = "EDD" Then
                If GlobalVariables.SampleList.Count = 0 Then
                    txtImportResults.Text = "No samples detected in EDD. Please ensure EDD is formatted correctly!"
                Else
                    For Each aSample In GlobalVariables.SampleList
                        strText = strText & "Number of Compounds: " & aSample.CompoundList.Count & vbCrLf
                        strText = strText & "Sample Code: " & aSample.CompoundList(0).EDDsysSampleCode & vbCrLf
                        strText = strText & "Analysis Date: " & aSample.CompoundList(0).EDDAnalysisDate & vbCrLf & vbCrLf
                        txtImportResults.Text = strText
                    Next
                End If
            ElseIf GlobalVariables.eTrain.Team = "CLAB" Then
                If GlobalVariables.SampleList.Count = 0 Then
                    txtImportResults.Text = "No samples detected in EDD. Please ensure EDD is formatted correctly!"
                Else
                    For Each aSample In GlobalVariables.SampleList
                        strText = strText & "Sample Code: " & aSample.CompoundList(0).EDDsysSampleCode & vbCrLf
                        strText = strText & "Number of Compounds: " & aSample.CompoundList.Count & vbCrLf
                        strText = strText & "Analysis Method: " & aSample.CompoundList(0).EDDLabAnlMethodName & vbCrLf & vbCrLf
                        txtImportResults.Text = strText
                    Next
                End If
            ElseIf GlobalVariables.eTrain.Team = "NewSample" Then
                If GlobalVariables.SampleList.Count = 0 Then
                    txtImportResults.Text = "No samples detected in EDD. Please ensure EDD is formatted correctly!"
                Else
                    For Each aSample In GlobalVariables.SampleList
                        strText = strText & "Sample Code: " & aSample.CompoundList(0).EDDsysSampleCode & vbCrLf
                        strText = strText & "Number of Compounds: " & aSample.CompoundList.Count & vbCrLf
                        strText = strText & "Analysis Method: " & aSample.CompoundList(0).EDDLabAnlMethodName & vbCrLf & vbCrLf
                        txtImportResults.Text = strText
                    Next
                End If

            Else
                For Each aSample In GlobalVariables.SampleList
                    strText = strText & "Sample Name: " & aSample.Name & vbCrLf
                    strText = strText & "Internal Standards: " & aSample.InternalStdList.Count & vbCrLf
                    strText = strText & "Surrogates: " & aSample.SurrogateList.Count & vbCrLf
                    strText = strText & "Compounds: " & aSample.CompoundList.Count & vbCrLf
                    strText = strText & "Misc: " & aSample.Misc & vbCrLf & vbCrLf
                    txtImportResults.Text = strText
                Next
            End If


            'Enable Report generation if samples in list and LIMS transfer if samples and server selected
            If GlobalVariables.SampleList.Count > 0 Then
                Me.btnReport.Enabled = True
                'Enable method/instrument/sigfig controls
                nudSigFig.Enabled = True
                btnSigHelp.Enabled = True
                If Not IsNothing(GlobalVariables.eTrain.Server) Then
                    Me.btnTransLIMS.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox("Error" & vbCrLf & ex.Message)
        End Try
    End Sub

    Private Sub btnTransLIMS_Click(sender As System.Object, e As System.EventArgs) Handles btnTransLIMS.Click

        If GlobalVariables.eTrain.Server = "SEADRIFT" Then
            'Check If SF set
            If nudSigFig.Value <> 0 Then
                GlobalVariables.eTrain.SigFig = nudSigFig.Value
            Else
                MsgBox("Please set a Sig Fig amount", MsgBoxStyle.Exclamation, "eTrain Lite")
                nudSigFig.Focus()
            End If
        ElseIf GlobalVariables.eTrain.Server = "ROH" Then
            'Start here next time, as about end sub???
        End If



        ' ----> Populate selInstrument -> May be obsolete
        ' ----> Populate selProject (LIMS AnalysisLab line) Query database for integrity check? (2nd column in import is analysis method name)

        If GlobalVariables.Transfer.ToLIMS(InputBox("Please enter your UserID for LIMS Transfer", "eTrain Lite")) Then
            MsgBox("Data transfer complete!", MsgBoxStyle.Information, "eTrain Lite")
        Else
            MsgBox("Error sending data to LIMS! Data not sent.", MsgBoxStyle.Critical, "eTrain Lite")
        End If



        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                If GlobalVariables.NeedsCalculation Then
                    'Get method names for loading
                    'Clear methods list to get new list
                    GlobalVariables.MethodList.Clear()
                    'Load method names
                    GlobalVariables.Method.LoadMethodNames()

                    'Check if SF set
                    If nudSigFig.Value <> 0 Then
                        GlobalVariables.eTrain.SigFig = nudSigFig.Value
                    Else
                        MsgBox("Please set a rounding amount", MsgBoxStyle.Exclamation, "eTrain 2.0")
                        nudSigFig.Focus()
                        Exit Sub
                    End If

                    'Call up Report form
                    'With TransferForm
                    '    .lblTxt1.Text = "Associated SIS Location:"
                    '    .lbl1.Visible = True
                    '    .cbo1.Visible = True
                    '    .lbl1.Text = "Method:"
                    '    .cbo1.Items.Clear()

                    '    For Each aMethod In GlobalVariables.MethodList
                    '        .cbo1.Items.Add(aMethod.Name)
                    '    Next
                    '    .lbl2.Visible = True
                    '    .cbo2.Visible = True
                    '    .lbl2.Text = "Instrument:"
                    '    .cbo2.Items.Clear()
                    '    .lbl3.Visible = False
                    '    .cbo3.Visible = False
                    '    .lbl4.Visible = False
                    '    .cbo4.Visible = False
                    '    .lbl5.Visible = False
                    '    .cbo5.Visible = False
                    '    .lblTxt1.Visible = True
                    '    .txt1.Visible = True
                    'End With
                    'TransferForm.ShowDialog()
                Else
                    'Call up Report form
                    'With TransferForm
                    '    .lblTxt1.Text = "Associated SIS Location:"
                    '    .lbl1.Visible = True
                    '    .cbo1.Visible = True
                    '    .lbl1.Text = "Method:"
                    '    .lbl2.Visible = True
                    '    .cbo2.Visible = True
                    '    .lbl2.Text = "Instrument:"
                    '    .lbl3.Visible = False
                    '    .cbo3.Visible = False
                    '    .lbl4.Visible = False
                    '    .cbo4.Visible = False
                    '    .lbl5.Visible = False
                    '    .cbo5.Visible = False
                    '    .lblTxt1.Visible = True
                    '    .txt1.Visible = True
                    'End With
                    'TransferForm.ShowDialog()
                End If
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                'Check if SF set
                If nudSigFig.Value <> 0 Then
                    GlobalVariables.eTrain.SigFig = nudSigFig.Value
                Else
                    MsgBox("Please set a Sig Fig amount", MsgBoxStyle.Exclamation, "eTrain 2.0")
                    nudSigFig.Focus()
                    Exit Sub
                End If

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''Commented out by WT -> 10/19/2017''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                'Call up Report form
                'With TransferForm ' -> Possibly comment out... When transfer is selected, execute import and LIMS transfer in background
                '    .lblTxt1.Text = "Recovery Limits Path:"
                '    .lblTxt1.Visible = True
                '    .txt1.Visible = True
                '    .btnSISBrowse.Visible = True
                '    .lbl1.Visible = True
                '    .cbo1.Visible = True
                '    .lbl1.Text = "Source:"
                '    .lbl2.Visible = True
                '    .cbo2.Visible = True
                '    .lbl2.Text = "Source Name:"
                '    .lbl3.Visible = True
                '    .cbo3.Visible = True
                '    .lbl3.Text = "AnalysisLab:"
                '    .lbl4.Visible = True
                '    .cbo4.Visible = True
                '    .lbl4.Text = "Instrument:"
                '    .lbl5.Visible = True
                '    .cbo5.Visible = True
                '    .lbl5.Text = "Reporting Limit:"
                '    .cbo1.Enabled = True
                '    .cbo2.Enabled = False
                '    .cbo3.Enabled = False
                '    .cbo4.Enabled = False
                '    .cbo5.Enabled = False
                'End With
                'TransferForm.ShowDialog()



                ' ----> Populate selInstrument -> May be obsolete
                ' ----> Populate selProject (LIMS AnalysisLab line) Query database for integrity check? (2nd column in import is analysis method name)

                If GlobalVariables.Transfer.ToLIMS(InputBox("Please enter your UserID for LIMS Transfer", "eTrain 2.0")) Then
                    MsgBox("Data transfer complete!", MsgBoxStyle.Information, "eTrain 2.0")
                Else
                    MsgBox("Error sending data to LIMS! Data not sent.", MsgBoxStyle.Critical, "eTrain 2.0")
                End If
            End If
        End If
    End Sub

    Private Sub btnReport_Click(sender As System.Object, e As System.EventArgs) Handles btnReport.Click
        'Generate report

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.Import.Type = "DF" Then
                If nudSigFig.Value <> 0 Then
                    GlobalVariables.eTrain.SigFig = nudSigFig.Value
                Else
                    MsgBox("Please set sigfig amount", MsgBoxStyle.Exclamation, "eTrainLite")
                    nudSigFig.Focus()
                    Exit Sub
                End If

                'Call up Report form
                With ReportForm
                    .lblTxt1.Text = "Associated SIS Location:"
                    .lblTxt1.Visible = False
                    .SISBrowse.Visible = False
                    .txt1.Visible = False

                    .lbl1.Visible = False
                    .cbo1.Visible = False
                    .lbl2.Visible = False
                    .cbo2.Visible = False
                    .lbl3.Visible = False
                    .cbo3.Visible = False
                    .lbl4.Visible = False
                    .cbo4.Visible = False
                    .lbl5.Visible = False
                    .cbo5.Visible = False
                    .lbl6.Visible = False
                    .cbo6.Visible = False
                    .lblTxt1.Visible = False
                    .txt1.Visible = False
                End With
                ReportForm.ShowDialog()

            ElseIf GlobalVariables.eTrain.Team = "FAST" Then
                If GlobalVariables.NeedsCalculation Then
                    'Get method names for loading
                    'Clear methods list to get new list
                    GlobalVariables.MethodList.Clear()
                    'Load method names
                    GlobalVariables.Method.LoadMethodNames()

                    'Check if SF set
                    If nudSigFig.Value <> 0 Then
                        GlobalVariables.eTrain.SigFig = nudSigFig.Value
                    Else
                        MsgBox("Please set sigfig amount", MsgBoxStyle.Exclamation, "eTrain 2.0")
                        nudSigFig.Focus()
                        Exit Sub
                    End If

                    'Call up Report form
                    With ReportForm
                        .lblTxt1.Text = "Associated SIS Location:"
                        .lbl1.Visible = True
                        .cbo1.Visible = True
                        .lbl1.Text = "Method:"
                        .cbo1.Items.Clear()

                        For Each aMethod In GlobalVariables.MethodList
                            .cbo1.Items.Add(aMethod.Name)
                        Next
                        .lbl2.Visible = True
                        .cbo2.Visible = True
                        .lbl2.Text = "Instrument:"
                        .cbo2.Items.Clear()
                        .lbl3.Visible = False
                        .cbo3.Visible = False
                        .lbl4.Visible = False
                        .cbo4.Visible = False
                        .lbl5.Visible = False
                        .cbo5.Visible = False
                        .lbl6.Visible = False
                        .cbo6.Visible = False
                        .lblTxt1.Visible = True
                        .txt1.Visible = True
                    End With
                    ReportForm.ShowDialog()
                Else
                    'Call up Report form
                    With ReportForm
                        .lblTxt1.Text = "Associated SIS Location:"
                        .lbl1.Visible = True
                        .cbo1.Visible = True
                        .lbl1.Text = "Method:"
                        .lbl2.Visible = True
                        .cbo2.Visible = True
                        .lbl2.Text = "Instrument:"
                        .lbl3.Visible = False
                        .cbo3.Visible = False
                        .lbl4.Visible = False
                        .cbo4.Visible = False
                        .lbl5.Visible = False
                        .cbo5.Visible = False
                        .lblTxt1.Visible = True
                        .txt1.Visible = True
                    End With
                    ReportForm.ShowDialog()
                End If
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                If GlobalVariables.NeedsCalculation Then
                    'Get method names for loading
                    'Clear methods list to get new list
                    GlobalVariables.MethodList.Clear()
                    'Load method names
                    GlobalVariables.Method.LoadMethodNames()

                    'Check if SF set
                    If nudSigFig.Value <> 0 Then
                        GlobalVariables.eTrain.SigFig = nudSigFig.Value
                    Else
                        MsgBox("Please set sigfig amount", MsgBoxStyle.Exclamation, "eTrain 2.0")
                        nudSigFig.Focus()
                        Exit Sub
                    End If

                    'Call up Report form
                    With ReportForm
                        .lblTxt1.Text = "Associated SIS Location:"
                        .lbl1.Visible = True
                        .cbo1.Visible = True
                        .lbl1.Text = "Method:"
                        .cbo1.Items.Clear()

                        For Each aMethod In GlobalVariables.MethodList
                            .cbo1.Items.Add(aMethod.Name)
                        Next
                        .lbl2.Visible = True
                        .cbo2.Visible = True
                        .lbl2.Text = "Instrument:"
                        .cbo2.Items.Clear()
                        .lbl3.Visible = False
                        .cbo3.Visible = False
                        .lbl4.Visible = False
                        .cbo4.Visible = False
                        .lbl5.Visible = False
                        .cbo5.Visible = False
                        .lbl6.Visible = False
                        .cbo6.Visible = False
                        .lblTxt1.Visible = True
                        .txt1.Visible = True
                    End With
                    ReportForm.ShowDialog()
                Else
                    'Call up Report form
                    With ReportForm
                        .lblTxt1.Text = "Associated SIS Location:"
                        .lbl1.Visible = True
                        .cbo1.Visible = True
                        .lbl1.Text = "Method:"
                        .lbl2.Visible = True
                        .cbo2.Visible = True
                        .lbl2.Text = "Instrument:"
                        .lbl3.Visible = False
                        .cbo3.Visible = False
                        .lbl4.Visible = False
                        .cbo4.Visible = False
                        .lbl5.Visible = False
                        .cbo5.Visible = False
                        .lbl6.Visible = False
                        .cbo6.Visible = False
                        .lblTxt1.Visible = True
                        .txt1.Visible = True
                    End With
                    ReportForm.ShowDialog()
                End If
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                'Check if SF set
                If nudSigFig.Value <> 0 Then
                    GlobalVariables.eTrain.SigFig = nudSigFig.Value
                Else
                    MsgBox("Please set sigfig amount", MsgBoxStyle.Exclamation, "eTrain 2.0")
                    nudSigFig.Focus()
                    Exit Sub
                End If

                'Call up Report form
                With ReportForm
                    .lblTxt1.Text = "Associated SIS Location:"
                    .lbl1.Visible = True
                    .cbo1.Visible = True
                    .lbl1.Text = "Source:"
                    .lbl2.Visible = True
                    .cbo2.Visible = True
                    .lbl2.Text = "Source Name:"
                    .lbl3.Visible = True
                    .cbo3.Visible = True
                    .lbl3.Text = "Analysis:"
                    .lbl4.Visible = True
                    .cbo4.Visible = True
                    .lbl4.Text = "Instrument:"
                    .lbl5.Visible = True
                    .cbo5.Visible = True
                    .lbl5.Text = "Reporting Limit:"
                    .lbl6.Visible = True
                    .cbo6.Visible = True
                    .lbl6.Text = "Recovery Limits:"
                    .cbo1.Enabled = True
                    .cbo2.Enabled = False
                    .cbo3.Enabled = False
                    .cbo4.Enabled = False
                    .cbo5.Enabled = False
                    .cbo6.Enabled = False
                End With
                ReportForm.ShowDialog()
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                'Check if SF set
                If nudSigFig.Value <> 0 Then
                    GlobalVariables.eTrain.SigFig = nudSigFig.Value
                Else
                    MsgBox("Please set sigfig amount", MsgBoxStyle.Exclamation, "eTrain 2.0")
                    nudSigFig.Focus()
                    Exit Sub
                End If

                'Call up Report form
                With ReportForm
                    .lblTxt1.Text = "Recovery Limits Path:"
                    .lbl1.Visible = True
                    .cbo1.Visible = True
                    .lbl1.Text = "Source:"
                    .lbl2.Visible = True
                    .cbo2.Visible = True
                    .lbl2.Text = "Source Name:"
                    .lbl3.Visible = True
                    .cbo3.Visible = True
                    .lbl3.Text = "Analysis:"
                    .lbl4.Visible = True
                    .cbo4.Visible = True
                    .lbl4.Text = "Instrument:"
                    .lbl5.Visible = True
                    .cbo5.Visible = True
                    .lbl5.Text = "Reporting Limit:"
                    .cbo1.Enabled = True
                    .cbo2.Enabled = False
                    .cbo3.Enabled = False
                    .cbo4.Enabled = False
                    .cbo5.Enabled = False
                    .lbl6.Visible = False
                    .cbo6.Visible = False
                End With
                ReportForm.ShowDialog()
            End If
        End If
    End Sub

    Private Sub UpdateForm()

        If Not IsNothing(GlobalVariables.eTrain.Server) Then

            Me.tsslServer.Text = "Server: " & GlobalVariables.eTrain.Server

            Me.tsslServer.Visible = True
        End If

    End Sub

    Private Function CheckSF()
        If IsNumeric(CInt(Me.nudSigFig.Value)) Then
            GlobalVariables.eTrain.SigFig = CInt(Me.nudSigFig.Value)
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub MainMenuToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MainMenuToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub EditMethodsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)
        Dim aMethod As Method
        Dim aPermit As Permit

        'Group Differences to form
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                'load methods new
                GlobalVariables.Method.LoadMethodNames()

                With EditMethods
                    'Form setup
                    .Text = "Edit Methods - eTrain 2.0"
                    .CreateNewMethodToolStripMenuItem.Text = "&Create New Method / Project"
                    .CopyFromExistingMethodToolStripMenuItem.Visible = True
                    .UsingChemstationDataToolStripMenuItem.Visible = True
                    .tsslLocation.Text = "Location: " & GlobalVariables.eTrain.Location
                    .tsslTeam.Text = "Team: " & GlobalVariables.eTrain.Team
                    If GlobalVariables.eTrain.Server <> "" Then
                        .tsslServer.Text = "Server: " & GlobalVariables.eTrain.Server
                    Else
                        .tsslServer.Visible = False
                    End If
                    For Each aMethod In GlobalVariables.MethodList
                        .cboOption1.Items.Add(aMethod.Name)
                    Next

                    'Group Differences to form

                    .lblCboOption1.Text = "Method / Project:"
                    .lblCboOption2.Text = "Instrument:"
                    .lblCboOption3.Text = "Analyte Type:"
                    .btnOption4.Text = "Standard Books"
                    .lblTxtOption4.Text = "ETEQ:"
                    .lblTxtOption5.Text = "Report Tolerance:"
                    .btnOption2Add.Text = "Add Instrument"
                    .btnOption2Del.Text = "Delete Instrument"
                    .btnSave.Text = "Save Method"
                    .txtOption4.Visible = True
                    .txtOption5.Visible = True
                    .btnOption4.Visible = False

                    .btnOption2Add.Visible = False
                    .btnOption2Del.Visible = False
                    .btnOption3Add.Visible = False
                    .btnOption3Del.Visible = False
                    .btnAddCompound.Visible = False
                    .btnDelCompound.Visible = False
                End With
            ElseIf GlobalVariables.eTrain.Team = "CHROM" Then
                'load methods new
                GlobalVariables.Permit.LoadPermitNames()

                With EditMethods
                    'Form setup
                    .Text = "Edit Permits - eTrain 2.0"
                    .CreateNewMethodToolStripMenuItem.Text = "&Create New Permit"
                    .CopyFromExistingMethodToolStripMenuItem.Visible = False
                    .UsingChemstationDataToolStripMenuItem.Visible = False
                    .tsslLocation.Text = "Location: " & GlobalVariables.eTrain.Location
                    .tsslTeam.Text = "Team: " & GlobalVariables.eTrain.Team
                    If GlobalVariables.eTrain.Server <> "" Then
                        .tsslServer.Text = "Server: " & GlobalVariables.eTrain.Server
                    Else
                        .tsslServer.Visible = False
                    End If
                    For Each aPermit In GlobalVariables.PermitList
                        .cboOption1.Items.Add(aPermit.Name)
                    Next

                    'Group Differences to form
                    .lblCboOption1.Text = "Permit:"
                    .lblCboOption2.Text = "Project:"
                    .lblCboOption3.Text = "Instrument:"
                    .btnOption4.Text = "Default Limits"
                    .lblTxtOption4.Text = ""
                    .lblTxtOption5.Text = ""
                    .btnOption2Add.Text = "Add Project"
                    .btnOption2Del.Text = "Delete Project"
                    .btnOption3Add.Text = "Add Instrument"
                    .btnOption3Del.Text = "Delete Instrument"
                    .btnSave.Text = "Save Permit"
                    .txtOption4.Visible = False
                    .txtOption5.Visible = False

                    .btnOption2Add.Visible = False
                    .btnOption2Del.Visible = False
                    .btnOption3Add.Visible = False
                    .btnOption3Del.Visible = False
                    .btnAddCompound.Visible = False
                    .btnDelCompound.Visible = False
                End With
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                'load methods new
                GlobalVariables.Method.LoadMethodNames()

                With EditMethods
                    'Form setup
                    .Text = "Edit Methods - eTrain 2.0"
                    .CreateNewMethodToolStripMenuItem.Text = "&Create New Method / Project"
                    .CopyFromExistingMethodToolStripMenuItem.Visible = True
                    .UsingChemstationDataToolStripMenuItem.Visible = True
                    .tsslLocation.Text = "Location: " & GlobalVariables.eTrain.Location
                    .tsslTeam.Text = "Team: " & GlobalVariables.eTrain.Team
                    If GlobalVariables.eTrain.Server <> "" Then
                        .tsslServer.Text = "Server: " & GlobalVariables.eTrain.Server
                    Else
                        .tsslServer.Visible = False
                    End If
                    For Each aMethod In GlobalVariables.MethodList
                        .cboOption1.Items.Add(aMethod.Name)
                    Next

                    'Group Differences to form

                    .lblCboOption1.Text = "Method / Project:"
                    .lblCboOption2.Text = "Instrument:"
                    .lblCboOption3.Text = "Analyte Type:"
                    .btnOption4.Text = "Standard Books"
                    .lblTxtOption4.Text = "ETEQ:"
                    .lblTxtOption5.Text = "Report Tolerance:"
                    .btnOption2Add.Text = "Add Instrument"
                    .btnOption2Del.Text = "Delete Instrument"
                    .btnSave.Text = "Save Method"
                    .txtOption4.Visible = True
                    .txtOption5.Visible = True
                    .btnOption4.Visible = False

                    .btnOption2Add.Visible = False
                    .btnOption2Del.Visible = False
                    .btnOption3Add.Visible = False
                    .btnOption3Del.Visible = False
                    .btnAddCompound.Visible = False
                    .btnDelCompound.Visible = False
                End With
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                'load methods new
                GlobalVariables.Permit.LoadPermitNames()

                With EditMethods
                    'Form setup
                    .Text = "Edit Permits - eTrain 2.0"
                    .CreateNewMethodToolStripMenuItem.Text = "&Create New Permit"
                    .CopyFromExistingMethodToolStripMenuItem.Visible = False
                    .UsingChemstationDataToolStripMenuItem.Visible = False
                    .tsslLocation.Text = "Location: " & GlobalVariables.eTrain.Location
                    .tsslTeam.Text = "Team: " & GlobalVariables.eTrain.Team
                    If GlobalVariables.eTrain.Server <> "" Then
                        .tsslServer.Text = "Server: " & GlobalVariables.eTrain.Server
                    Else
                        .tsslServer.Visible = False
                    End If
                    For Each aPermit In GlobalVariables.PermitList
                        .cboOption1.Items.Add(aPermit.Name)
                    Next

                    'Group Differences to form
                    .lblCboOption1.Text = "Permit:"
                    .lblCboOption2.Text = "Project:"
                    .lblCboOption3.Text = "Instrument:"
                    .btnOption4.Text = "Default Limits"
                    .lblTxtOption4.Text = ""
                    .lblTxtOption5.Text = ""
                    .btnOption2Add.Text = "Add Project"
                    .btnOption2Del.Text = "Delete Project"
                    .btnOption3Add.Text = "Add Instrument"
                    .btnOption3Del.Text = "Delete Instrument"
                    .btnSave.Text = "Save Permit"
                    .txtOption4.Visible = False
                    .txtOption5.Visible = False

                    .btnOption2Add.Visible = False
                    .btnOption2Del.Visible = False
                    .btnOption3Add.Visible = False
                    .btnOption3Del.Visible = False
                    .btnAddCompound.Visible = False
                    .btnDelCompound.Visible = False
                End With
            End If
        End If

        EditMethods.ShowDialog()

    End Sub

    Private Sub btnSelAll_Click(sender As System.Object, e As System.EventArgs) Handles btnSelAll.Click
        For i = 0 To lstFileList.Items.Count - 1
            lstFileList.SetSelected(i, True)
        Next
    End Sub

    Private Sub btnClearList_Click(sender As System.Object, e As System.EventArgs) Handles btnClearList.Click
        lstFileList.Items.Clear()
    End Sub

    Private Sub btnClearSamples_Click(sender As System.Object, e As System.EventArgs) Handles btnClearSamples.Click
        GlobalVariables.SampleList.Clear()
        lblImportResults.Text = "Import Results: "
        txtImportResults.Text = ""
    End Sub

    Private Sub TestToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GlobalVariables.Calculations.MidlandHR(InputBox("Enter SIS location"))
    End Sub

    Private Sub btnSigHelp_Click(sender As System.Object, e As System.EventArgs) Handles btnSigHelp.Click
        MsgBox("Enter -1 for no significant figure rounding.")
    End Sub


    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Dim versionNumber As Version

        versionNumber = Assembly.GetExecutingAssembly().GetName().Version
        MsgBox("Version: " & versionNumber.ToString & vbCrLf & "Developer: Joshua Durham U411882" & vbCrLf & "Co-Developer Wyatt Towne UA20088")

    End Sub
    Private Sub SeadriftToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SeadriftToolStripMenuItem.Click
        'Set GlobalVariables
        GlobalVariables.eTrain.Server = "SEADRIFT"
        GlobalVariables.eTrain.ServerFP = "\\usmdlsdowacds1\LIMS_XFER\CHEMS\" '<- Actual path For SeaDrift server

        'Form UI
        UpdateForm()
        'Populate import type box
        btnFindFiles.Enabled = True
        cboImportType.Enabled = True
        cboImportType.Items.Clear()
        cboImportType.Items.Add("EDD") 'Added WT 9/26/2017

        'Enable LIMS transfer if samples and server selected
        If GlobalVariables.SampleList.Count > 0 And Not IsNothing(GlobalVariables.eTrain.Server) Then
            Me.btnTransLIMS.Enabled = True
        End If
    End Sub

    Private Sub ROHToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ROHToolStripMenuItem1.Click
        'Set GlobalVariables
        GlobalVariables.eTrain.Server = "ROH"
        GlobalVariables.eTrain.ServerFP = "\\usmdlsdowacds1\LIMS_XFER\ROHNA\"

        'Form UI
        UpdateForm()
        'Populate import type box
        btnFindFiles.Enabled = True
        cboImportType.Enabled = True
        cboImportType.Items.Clear()
        cboImportType.Items.Add("EDD")

        'Enable LIMS transfer if samples and server selected
        If GlobalVariables.SampleList.Count > 0 And Not IsNothing(GlobalVariables.eTrain.Server) Then
            Me.btnTransLIMS.Enabled = True
        End If
    End Sub
    Private Sub AECOMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AECOMToolStripMenuItem.Click
        'Set GlobalVariables
        GlobalVariables.eTrain.Server = "MIDLAND"
        GlobalVariables.eTrain.Location = "MIDLAND"
        GlobalVariables.eTrain.Team = "AECOM"
        GlobalVariables.eTrain.ServerFP = "\\usmdlsdowacds1\Lims_xfer\ENVMD\"
        'Form UI
        UpdateForm()
        'Populate import type box
        btnFindFiles.Enabled = True
        cboImportType.Enabled = True
        cboImportType.Items.Clear()
        cboImportType.Items.Add("SSR")

        'Enable LIMS transfer if samples and server selected
        If GlobalVariables.SampleList.Count > 0 And Not IsNothing(GlobalVariables.eTrain.Server) Then
            Me.btnTransLIMS.Enabled = True
        End If
    End Sub

    Private Sub EUROLANToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EUROLANToolStripMenuItem.Click
        'Set GlobalVariables
        GlobalVariables.eTrain.Server = "MIDLAND"
        GlobalVariables.eTrain.Location = "MIDLAND"
        GlobalVariables.eTrain.Team = "CLAB"
        GlobalVariables.eTrain.AnalysisLab = "EUROLAN"
        GlobalVariables.eTrain.ServerFP = "\\usmdlsdowacds1\Lims_xfer\ENVMD\"

        'Form UI
        UpdateForm()
        'Populate import type box
        btnFindFiles.Enabled = True
        cboImportType.Enabled = True
        cboImportType.Items.Clear()

        cboImportType.Items.Add("GRABS")
        cboImportType.Items.Add("031A")
        cboImportType.Items.Add("031B")
        cboImportType.Items.Add("031C")
        cboImportType.Items.Add("PESTICIDES")
        cboImportType.Items.Add("EUROLAN")
    End Sub
    Private Sub ALSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ALSToolStripMenuItem.Click
        'Set GlobalVariables
        GlobalVariables.eTrain.Server = "MIDLAND"
        GlobalVariables.eTrain.Location = "MIDLAND"
        GlobalVariables.eTrain.Team = "CLAB"
        GlobalVariables.eTrain.AnalysisLab = "ALS"
        GlobalVariables.eTrain.ServerFP = "\\usmdlsdowacds1\Lims_xfer\ENVMD\"

        'Form UI
        UpdateForm()
        'Populate import type box
        btnFindFiles.Enabled = True
        cboImportType.Enabled = True
        cboImportType.Items.Clear()

        cboImportType.Items.Add("BOD-PHOS")
        cboImportType.Items.Add("QUALA_ALS")
        cboImportType.Items.Add("DPS_ALS")
        cboImportType.Items.Add("XALT_ALS")
        cboImportType.Items.Add("ALS")
    End Sub
    Private Sub SGSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SGSToolStripMenuItem.Click
        'Set GlobalVariables
        GlobalVariables.eTrain.Server = "MIDLAND"
        GlobalVariables.eTrain.Location = "MIDLAND"
        GlobalVariables.eTrain.Team = "CLAB"
        GlobalVariables.eTrain.AnalysisLab = "SGS"
        GlobalVariables.eTrain.ServerFP = "\\usmdlsdowacds1\Lims_xfer\ENVMD\"

        'Form UI
        UpdateForm()
        'Populate import type box
        btnFindFiles.Enabled = True
        cboImportType.Enabled = True
        cboImportType.Items.Clear()

        cboImportType.Items.Add("METALS")
        cboImportType.Items.Add("DPS_SGS")
        cboImportType.Items.Add("SGS")
    End Sub
    Private Sub TAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TAToolStripMenuItem.Click
        'Set GlobalVariables
        GlobalVariables.eTrain.Server = "MIDLAND"
        GlobalVariables.eTrain.Location = "MIDLAND"
        GlobalVariables.eTrain.Team = "CLAB"
        GlobalVariables.eTrain.AnalysisLab = "TA"
        GlobalVariables.eTrain.ServerFP = "\\usmdlsdowacds1\Lims_xfer\ENVMD\"

        'Form UI
        UpdateForm()
        'Populate import type box
        btnFindFiles.Enabled = True
        cboImportType.Enabled = True
        cboImportType.Items.Clear()

        cboImportType.Items.Add("IONS")
        cboImportType.Items.Add("VOA_MONTHLY")
        cboImportType.Items.Add("SFE_FLUORIDE")
        cboImportType.Items.Add("1005_PS")
        cboImportType.Items.Add("SVOA_SEMIANNUAL")
        cboImportType.Items.Add("PCB_TA")
        cboImportType.Items.Add("CYANIDE")
        cboImportType.Items.Add("XALT_TA")
        cboImportType.Items.Add("CABOT")
        cboImportType.Items.Add("KENAN_TA")
        cboImportType.Items.Add("QUALA_TA")
        cboImportType.Items.Add("SEWER")
        cboImportType.Items.Add("TA")
    End Sub
    Private Sub FIBERTECToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FIBERTECToolStripMenuItem.Click
        'Set GlobalVariables
        GlobalVariables.eTrain.Server = "MIDLAND"
        GlobalVariables.eTrain.Location = "MIDLAND"
        GlobalVariables.eTrain.Team = "CLAB"
        GlobalVariables.eTrain.AnalysisLab = "FIBERTEC"
        GlobalVariables.eTrain.ServerFP = "\\usmdlsdowacds1\Lims_xfer\ENVMD\"

        'Form UI
        UpdateForm()
        'Populate import type box
        btnFindFiles.Enabled = True
        cboImportType.Enabled = True
        cboImportType.Items.Clear()

        cboImportType.Items.Add("KENAN_FIBERTEC")
        cboImportType.Items.Add("FIBERTEC")
    End Sub
    Private Sub VISTAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VISTAToolStripMenuItem.Click
        'Set GlobalVariables
        GlobalVariables.eTrain.Server = "MIDLAND"
        GlobalVariables.eTrain.Location = "MIDLAND"
        GlobalVariables.eTrain.Team = "CLAB"
        GlobalVariables.eTrain.AnalysisLab = "VISTA"
        GlobalVariables.eTrain.ServerFP = "\\usmdlsdowacds1\Lims_xfer\ENVMD\"

        'Form UI
        UpdateForm()
        'Populate import type box
        btnFindFiles.Enabled = True
        cboImportType.Enabled = True
        cboImportType.Items.Clear()

        cboImportType.Items.Add("DF")
        cboImportType.Items.Add("VISTA")
    End Sub
    ' Importing data to create a new sample in LIMS if it doesn't have a number associated with it. 
    Private Sub TANC_NEWToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles TANC_NEWToolStripMenuItem1.Click
        GlobalVariables.eTrain.Server = "MIDLAND"
        GlobalVariables.eTrain.Location = "MIDLAND"
        GlobalVariables.eTrain.Team = "NewSample"
        GlobalVariables.eTrain.AnalysisLab = "TANC_NEW"
        GlobalVariables.eTrain.ServerFP = "\\usmdlsdowacds1\Lims_xfer\ENVMD\"

        'Form UI
        UpdateForm()
        'Populate import type box
        btnFindFiles.Enabled = True
        cboImportType.Enabled = True
        cboImportType.Items.Clear()

        cboImportType.Items.Add("SEWER")

    End Sub

    Private Sub DeerParkToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeerParkToolStripMenuItem.Click
        GlobalVariables.eTrain.Server = "MIDLAND"
        GlobalVariables.eTrain.Location = "MIDLAND"
        GlobalVariables.eTrain.Team = "NewSample"
        GlobalVariables.eTrain.AnalysisLab = "DeerPark"
        GlobalVariables.eTrain.ServerFP = "\\usmdlsdowacds1\Lims_xfer\ENVMD\"

        'Form UI
        UpdateForm()
        'Populate import type box
        btnFindFiles.Enabled = True
        cboImportType.Enabled = True
        cboImportType.Items.Clear()

        cboImportType.Items.Add("DeerPark")
    End Sub
End Class