Imports System.IO
Imports Syncfusion.XlsIO
Imports System.Text.RegularExpressions

Public Class Import
    Private strFilePath As String
    Private strCCCheckFilePath As String
    Private strTMPFilePath As String
    Private strFolderPath As String
    Private strType As String
    Public Property arrFileList As New ArrayList

    'Sets/Gets
    Public Property FilePath() As String
        Get
            Return strFilePath
        End Get
        Set(ByVal value As String)
            strFilePath = value
        End Set
    End Property
    Public Property CCCheckFilePath() As String
        Get
            Return strCCCheckFilePath
        End Get
        Set(ByVal value As String)
            strCCCheckFilePath = value
        End Set
    End Property
    Public Property TMPFilePath() As String
        Get
            Return strTMPFilePath
        End Get
        Set(ByVal value As String)
            strTMPFilePath = value
        End Set
    End Property
    Public Property FolderPath() As String
        Get
            Return strFolderPath
        End Get
        Set(ByVal value As String)
            strFolderPath = value
        End Set
    End Property
    Public Property Type() As String
        Get
            Return strType
        End Get
        Set(ByVal value As String)
            strType = value
        End Set
    End Property


    'Import Samples into eTrain based on import type
    Sub SampleImport()

        'Vars for Chemstation import
        Dim line As String
        Dim lineGold As String
        Dim flg As Integer '0 - Internal Standards  1 - Surrogates  2 - Compounds
        Dim arrSplText() As String
        Dim aSample As New Sample
        Dim aStandard As Standard
        Dim aSurrogate As Surrogate
        Dim aCompound As Compound
        Dim exEngine As New ExcelEngine

        'Determine import type then begin import
        'Chemstation Import
        If Type = "CHEM" Then

            Dim sr As StreamReader = New StreamReader(GlobalVariables.Import.FilePath)

            'Set Flag
            flg = -1

            'Loop until first line of actual text
            Try
                Do
                    line = sr.ReadLine
                Loop Until Not line = ""

                'Check first line to see if QT Reviewed or not
                If InStr(line, "QT Reviewed") Then
                    aSample.QTReviewed = True
                Else
                    aSample.QTReviewed = False
                End If

                'Continue line by line splitting by ":" and loading into sample class
                'Check for spaces
                line = sr.ReadLine
                If line = "" Then
                    Do Until line <> ""
                        line = sr.ReadLine
                    Loop
                End If
                aSample.DataPath = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                line = sr.ReadLine
                If InStr(line, "...    ") Then
                    line = sr.ReadLine
                End If
                aSample.DataFile = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                aSample.LimsID = Left(aSample.DataFile, aSample.DataFile.Length - 2)
                line = sr.ReadLine
                If InStr(line, "...    ") Then
                    line = sr.ReadLine
                End If
                aSample.AcqDate = CDate(Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))))
                line = sr.ReadLine
                If InStr(line, "...    ") Then
                    line = sr.ReadLine
                End If
                aSample.Analyst = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                line = sr.ReadLine
                If InStr(line, "...    ") Then
                    line = sr.ReadLine
                End If
                aSample.Name = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))

                'Sample type assignment
                If InStr(aSample.Name, "Method Blank") Then
                    aSample.Type = "MB"
                ElseIf InStr(aSample.Name, "Lab Control Spike DUP") Or InStr(aSample.Name, "LCSD", CompareMethod.Binary) Then
                    aSample.Type = "LCSD"
                ElseIf InStr(aSample.Name, "Lab Control Spike") Or InStr(aSample.Name, "LCS", CompareMethod.Binary) Then
                    aSample.Type = "LCS"
                ElseIf InStr(aSample.Name, "MSD", CompareMethod.Binary) Then
                    aSample.Type = "MSD"
                ElseIf InStr(aSample.Name, "MS", CompareMethod.Binary) Then
                    aSample.Type = "MS"
                ElseIf InStr(aSample.Name, "Lab Blank") Then
                    aSample.Type = "LB"
                ElseIf InStr(aSample.Name, "CVS", CompareMethod.Binary) Then
                    aSample.Type = "CVS"
                ElseIf InStr(aSample.Name, "ICV", CompareMethod.Binary) Then
                    aSample.Type = "ICV"
                ElseIf InStr(aSample.Name, "DUP", CompareMethod.Binary) Then
                    aSample.Type = "DUP"
                ElseIf InStr(aSample.Name, "Standard") Then
                    aSample.Type = "STD"
                ElseIf InStr(aSample.Name, "CS", CompareMethod.Binary) Then
                    aSample.Type = "CHECK"
                Else
                    aSample.Type = "SAMPLE"
                End If

                line = sr.ReadLine
                If InStr(line, "...    ") Then
                    line = sr.ReadLine
                End If
                aSample.Misc = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))

                'Assign values from MISC if applicable
                If aSample.Misc <> "" Then
                    If GlobalVariables.eTrain.Location = "MIDLAND" Then
                        If GlobalVariables.eTrain.Team = "CHROM" Then
                            arrSplText = aSample.Misc.Split(",")
                            aSample.LimsID = arrSplText(0)
                            aSample.SampleDate = CDate(arrSplText(1))
                            aSample.DilutionFactor = arrSplText(2)
                            aSample.DetectLimitType = arrSplText(3)
                            aSample.Matrix = arrSplText(4)
                            'FAST - Aliquot, Std amt, Inj amt, LCS amt
                        ElseIf GlobalVariables.eTrain.Team = "FAST" Then
                            If InStr(aSample.Misc, ",") Then
                                arrSplText = aSample.Misc.Split(",")
                                aSample.Aliquot = arrSplText(0)
                                aSample.StdSpikeAmt = arrSplText(1)
                                aSample.InjSpikeAmt = arrSplText(2)
                                If UBound(arrSplText) > 2 Then
                                    aSample.LCSSpikeAmt = arrSplText(3)
                                End If
                            End If
                        End If
                    ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
                        If GlobalVariables.eTrain.Team = "CHROM" Then
                            arrSplText = aSample.Misc.Split(",")
                            If arrSplText(0) = "N/A" Then
                                aSample.LimsID = ""
                            Else
                                aSample.LimsID = arrSplText(0)
                            End If
                            aSample.SampleDate = CDate(arrSplText(1))
                            aSample.DilutionFactor = arrSplText(2)
                            aSample.DetectLimitType = arrSplText(3)
                            aSample.Matrix = arrSplText(4)
                        End If
                    End If
                End If

                'ALS Vial and Sample Multipler on same line
                line = sr.ReadLine
                If InStr(line, "...    ") Then
                    line = sr.ReadLine
                End If
                aSample.Vial = Trim(line.Substring(InStr(line, ":"), InStr(line, "S")))
                aSample.Multiplier = Trim(line.Substring(InStrRev(line, ":"), line.Length - InStrRev(line, ":")))

                'Blank line between top header and Quant header information
                sr.ReadLine()

                'Continue splitting until blank line & internal standards
                line = sr.ReadLine
                If InStr(line, "...    ") Then
                    line = sr.ReadLine
                End If
                aSample.QuantTime = CDate(GlobalVariables.CSImport.DateFix(Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))), False))
                line = sr.ReadLine
                If InStr(line, "...    ") Then
                    line = sr.ReadLine
                End If
                aSample.QuantMethod = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                line = sr.ReadLine
                If InStr(line, "...    ") Then
                    line = sr.ReadLine
                End If
                aSample.QuantTitle = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                line = sr.ReadLine
                If InStr(line, "...    ") Then
                    line = sr.ReadLine
                End If
                aSample.QLastUpdate = CDate(GlobalVariables.CSImport.DateFix(Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))), True))
                line = sr.ReadLine
                If InStr(line, "...    ") Then
                    line = sr.ReadLine
                End If


                aSample.ResponseVia = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                If GlobalVariables.eTrain.Team = "CHROM" Then
                    line = sr.ReadLine
                    aSample.QMethFile = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                End If

                'Loop through until end of file now and load in Standards, Surrogates, and/or Compounds
                Do Until sr.EndOfStream
                    line = sr.ReadLine

                    'Set flags based on conditions
                    If line = "" Then
                        flg = -1
                    ElseIf InStr(line, "Internal Standards") Then
                        flg = 0
                    ElseIf InStr(line, "System Monitoring Compounds") Then
                        flg = 1
                    ElseIf InStr(line, "Target Compounds") Then
                        flg = 2
                    End If

                    'Check Flag, if tripped gather information
                    If flg = 0 Then
                        GlobalVariables.CSImport.StandardLoad(aSample, line, line)
                    ElseIf flg = 1 Then
                        If InStr(line, "System Monitoring Compounds") = 0 Then
                            line = line & "  " & sr.ReadLine
                            GlobalVariables.CSImport.SurrogateLoad(aSample, line, line)
                        End If
                    ElseIf flg = 2 Then
                        GlobalVariables.CSImport.CompoundLoad(aSample, line, line)
                    End If
                Loop

                sr.Close()

                'Check for Ion Ratio File 'tmpqntrp.txt'
                If GlobalVariables.eTrain.Team = "FAST" Then
                    For Each file In Directory.GetFiles(GlobalVariables.Import.FolderPath)
                        If InStr(file, "tmpqntrp") Then
                            Try
                                GlobalVariables.Import.strTMPFilePath = file
                                sr = New StreamReader(GlobalVariables.Import.strTMPFilePath)
                                'Loop until QuantFile line
                                Do Until (InStr(line, "QuantFile"))
                                    line = sr.ReadLine
                                Loop
                                aSample.TMPQuantFile = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                                line = sr.ReadLine
                                'Loop through whole file until end, grabbing details along the way
                                Do Until sr.EndOfStream
                                    lineGold = line
                                    If InStr(line, "Compound:") Then
                                        For Each aStandard In aSample.InternalStdList
                                            'Make sure not N.D.
                                            If aStandard.Conc <> "N.D." And aStandard.Conc <> "0" Then
                                                If Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))) = aStandard.Name Then
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    If Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))) = "" Then
                                                        Exit For
                                                    End If
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    'Load in Ions and their details
                                                    'Target
                                                    GlobalVariables.CSImport.TMPStdLoad(aStandard, line, line, "T")
                                                    line = sr.ReadLine
                                                    'Q1
                                                    GlobalVariables.CSImport.TMPStdLoad(aStandard, line, line, "Q1")
                                                    line = sr.ReadLine
                                                    'Q2
                                                    GlobalVariables.CSImport.TMPStdLoad(aStandard, line, line, "Q2")
                                                    line = sr.ReadLine()
                                                    'Q3
                                                    GlobalVariables.CSImport.TMPStdLoad(aStandard, line, line, "Q3")
                                                    line = sr.ReadLine()
                                                    line = sr.ReadLine()
                                                    line = sr.ReadLine()
                                                    line = sr.ReadLine()
                                                    If InStr(line, "Rel. Std. Dev.") Then
                                                        aStandard.TMPRelStdDev = Trim(line.Substring(InStr(line, "="), line.Length - InStr(line, "=") - 1))
                                                    End If
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                        For Each aSurrogate In aSample.SurrogateList
                                            If aSurrogate.Conc <> "N.D." And aSurrogate.Conc <> "0" Then
                                                If Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))) = aSurrogate.Name Then
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    If Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))) = "" Then
                                                        Exit For
                                                    End If
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    'Load in Ions and their details
                                                    'Target
                                                    GlobalVariables.CSImport.TMPSurrLoad(aSurrogate, line, line, "T")
                                                    line = sr.ReadLine
                                                    'Q1
                                                    GlobalVariables.CSImport.TMPSurrLoad(aSurrogate, line, line, "Q1")
                                                    line = sr.ReadLine
                                                    'Q2
                                                    GlobalVariables.CSImport.TMPSurrLoad(aSurrogate, line, line, "Q2")
                                                    line = sr.ReadLine()
                                                    'Q3
                                                    GlobalVariables.CSImport.TMPSurrLoad(aSurrogate, line, line, "Q3")
                                                    line = sr.ReadLine()
                                                    line = sr.ReadLine()
                                                    line = sr.ReadLine()
                                                    line = sr.ReadLine()
                                                    If InStr(line, "Rel. Std. Dev.") Then
                                                        aSurrogate.TMPRelStdDev = Trim(line.Substring(InStr(line, "="), line.Length - InStr(line, "=") - 1))
                                                    End If
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                        For Each aCompound In aSample.CompoundList
                                            If aCompound.Conc <> "N.D." And aCompound.Conc <> "0" Then
                                                If Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))) = aCompound.Name Then
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    If Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))) = "" Then
                                                        Exit For
                                                    End If
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    line = sr.ReadLine
                                                    'Load in Ions and their details
                                                    'Target
                                                    GlobalVariables.CSImport.TMPCompLoad(aCompound, line, line, "T")
                                                    line = sr.ReadLine
                                                    'Q1
                                                    GlobalVariables.CSImport.TMPCompLoad(aCompound, line, line, "Q1")
                                                    line = sr.ReadLine
                                                    'Q2
                                                    GlobalVariables.CSImport.TMPCompLoad(aCompound, line, line, "Q2")
                                                    line = sr.ReadLine()
                                                    'Q3
                                                    GlobalVariables.CSImport.TMPCompLoad(aCompound, line, line, "Q3")
                                                    line = sr.ReadLine()
                                                    line = sr.ReadLine()
                                                    line = sr.ReadLine()
                                                    line = sr.ReadLine()
                                                    If InStr(line, "Rel. Std. Dev.") Then
                                                        aCompound.TMPRelStdDev = Trim(line.Substring(InStr(line, "="), line.Length - InStr(line, "=") - 1))
                                                    End If
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    End If
                                    line = sr.ReadLine
                                Loop
                            Catch ex As Exception
                                MsgBox("Error reading file: " & GlobalVariables.Import.strTMPFilePath & vbCrLf &
                                       "Line: " & lineGold & vbCrLf &
                                       "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                            End Try
                        End If
                    Next
                End If

                'Check for CCCheck
                If GlobalVariables.eTrain.Team = "CHROM" Then
                    For Each file In Directory.GetFiles(GlobalVariables.Import.FolderPath)
                        If InStr(file, "cccheck") Then
                            GlobalVariables.Import.strCCCheckFilePath = file
                            sr = New StreamReader(GlobalVariables.Import.strCCCheckFilePath)
                            'Look for Quant Time
                            Try
                                Do Until (InStr(line, "Quant Time:"))
                                    line = sr.ReadLine
                                Loop
                                aSample.CCCQuantTime = CDate(GlobalVariables.CSImport.DateFix(Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))), False))
                                'Look for Min. RRF
                                Do Until (InStr(line, "Min. RRF"))
                                    line = sr.ReadLine
                                Loop
                                lineGold = line
                                'Reset Line & RT
                                GlobalVariables.CSImport.Pos = 16
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                GlobalVariables.CSImport.Pos = 8
                                aSample.MinRRF = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                GlobalVariables.CSImport.Pos = 18
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                GlobalVariables.CSImport.Pos = 5
                                aSample.MinRelArea = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                GlobalVariables.CSImport.Pos = 16
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                aSample.MaxRTDev = Trim(line.Substring(0, line.Length))
                                line = sr.ReadLine
                                lineGold = line
                                GlobalVariables.CSImport.Pos = 16
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                GlobalVariables.CSImport.Pos = 5
                                aSample.MaxRRFDev = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                GlobalVariables.CSImport.Pos = 21
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                aSample.MaxRelArea = Trim(line.Substring(0, line.Length))
                                line = sr.ReadLine
                                line = sr.ReadLine
                                Do Until sr.EndOfStream
                                    GlobalVariables.CSImport.CCCheckLoad(aSample, line, line)
                                    line = sr.ReadLine
                                Loop
                            Catch ex As Exception
                                MsgBox("Error reading file: " & GlobalVariables.Import.strCCCheckFilePath & vbCrLf &
                                       "Line: " & lineGold & vbCrLf &
                                       "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                            End Try
                        End If
                    Next
                End If

                'Add sample to sample list
                GlobalVariables.SampleList.Add(aSample)
                GlobalVariables.NeedsCalculation = True

            Catch ex As Exception
                MsgBox("Error reading file: " & GlobalVariables.Import.FilePath & vbCrLf &
                "Line: " & line & vbCrLf &
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        ElseIf Type = "CHEMBEVCAN" Then

            Dim sr As StreamReader = New StreamReader(GlobalVariables.Import.FilePath)

            'Set Flag
            flg = -1

            'Loop until first line of actual text
            Try
                Do
                    line = sr.ReadLine
                Loop Until Not line = ""

                'Check first line to see if QT Reviewed or not
                If InStr(line, "QT Reviewed") Then
                    aSample.QTReviewed = True
                Else
                    aSample.QTReviewed = False
                End If

                'Continue line by line splitting by ":" and loading into sample class
                'Check for spaces
                line = sr.ReadLine
                If line = "" Then
                    Do Until line <> ""
                        line = sr.ReadLine
                    Loop
                End If
                aSample.DataPath = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                line = sr.ReadLine
                aSample.DataFile = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                aSample.LimsID = Left(aSample.DataFile, aSample.DataFile.Length - 2)
                line = sr.ReadLine
                aSample.AcqDate = CDate(Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))))
                line = sr.ReadLine
                aSample.Analyst = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                line = sr.ReadLine
                aSample.Name = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))

                'Sample type assignment
                If InStr(aSample.Name, "Method Blank") Then
                    aSample.Type = "MB"
                ElseIf InStr(aSample.Name, "Lab Control Spike DUP") Or InStr(aSample.Name, "LCSD", CompareMethod.Binary) Then
                    aSample.Type = "LCSD"
                ElseIf InStr(aSample.Name, "Lab Control Spike") Or InStr(aSample.Name, "LCS", CompareMethod.Binary) Then
                    aSample.Type = "LCS"
                ElseIf InStr(aSample.Name, "MSD", CompareMethod.Binary) Then
                    aSample.Type = "MSD"
                ElseIf InStr(aSample.Name, "MS", CompareMethod.Binary) Then
                    aSample.Type = "MS"
                ElseIf InStr(aSample.Name, "Lab Blank") Then
                    aSample.Type = "LB"
                ElseIf InStr(aSample.Name, "CVS", CompareMethod.Binary) Then
                    aSample.Type = "CVS"
                ElseIf InStr(aSample.Name, "ICV", CompareMethod.Binary) Then
                    aSample.Type = "ICV"
                ElseIf InStr(aSample.Name, "DUP", CompareMethod.Binary) Then
                    aSample.Type = "DUP"
                ElseIf InStr(aSample.Name, "CS", CompareMethod.Binary) Then
                    aSample.Type = "CHECK"
                ElseIf InStr(aSample.Name, "Standard") Then
                    aSample.Type = "STD"
                Else
                    aSample.Type = "SAMPLE"
                End If

                line = sr.ReadLine
                aSample.Misc = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))

                'Assign values from MISC if applicable
                If aSample.Misc <> "" Then
                    If GlobalVariables.eTrain.Location = "FREEPORT" Then
                        If GlobalVariables.eTrain.Team = "CHROM" Then
                            arrSplText = aSample.Misc.Split(",")
                            If arrSplText(0) = "N/A" Then
                                aSample.LimsID = ""
                            Else
                                aSample.LimsID = arrSplText(0)
                            End If
                            aSample.SampleDate = CDate(arrSplText(1))
                            aSample.DilutionFactor = arrSplText(2)
                            aSample.DetectLimitType = arrSplText(3)
                            aSample.Matrix = arrSplText(4)
                            If arrSplText(5) = "N/A" Then
                                aSample.SISSampleWeight = ""
                            Else
                                aSample.SISSampleWeight = arrSplText(5)
                            End If
                        End If
                    End If
                End If

                'ALS Vial and Sample Multipler on same line
                line = sr.ReadLine
                aSample.Vial = Trim(line.Substring(InStr(line, ":"), InStr(line, "S")))
                aSample.Multiplier = Trim(line.Substring(InStrRev(line, ":"), line.Length - InStrRev(line, ":")))

                'Blank line between top header and Quant header information
                sr.ReadLine()

                'Continue splitting until blank line & internal standards
                line = sr.ReadLine
                aSample.QuantTime = CDate(GlobalVariables.CSImport.DateFix(Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))), False))
                line = sr.ReadLine
                aSample.QuantMethod = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                line = sr.ReadLine
                aSample.QuantTitle = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                line = sr.ReadLine
                aSample.QLastUpdate = CDate(GlobalVariables.CSImport.DateFix(Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))), True))
                line = sr.ReadLine
                aSample.ResponseVia = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                If GlobalVariables.eTrain.Team = "CHROM" Then
                    line = sr.ReadLine
                    aSample.QMethFile = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                End If

                'Blank line 
                sr.ReadLine()

                'Continue splitting 
                line = sr.ReadLine
                aSample.Signals = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                line = sr.ReadLine
                aSample.VolInj = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                line = sr.ReadLine
                aSample.SigPhase = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))
                line = sr.ReadLine
                aSample.SigInfo = Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":")))

                'Loop through until end of file now and load in Standards, Surrogates, and/or Compounds
                Do Until sr.EndOfStream
                    line = sr.ReadLine

                    'Set flags based on conditions
                    If line = "" Then
                        flg = -1
                    ElseIf InStr(line, "Internal Standards") Then
                        flg = 0
                    ElseIf InStr(line, "System Monitoring Compounds") Then
                        flg = 1
                    ElseIf InStr(line, "Target Compounds") Then
                        flg = 2
                    ElseIf InStr(line, "Non Target Peaks") Then
                        flg = 3
                    End If

                    'Check Flag, if tripped gather information
                    If flg = 0 Then
                        GlobalVariables.CSImport.StandardLoad(aSample, line, line)
                    ElseIf flg = 1 Then
                        If InStr(line, "System Monitoring Compounds") = 0 Then
                            line = line & "  " & sr.ReadLine
                            GlobalVariables.CSImport.SurrogateLoad(aSample, line, line)
                        End If
                    ElseIf flg = 2 Then
                        GlobalVariables.CSImport.CompoundLoad(aSample, line, line)
                    ElseIf flg = 3 Then
                        GlobalVariables.CSImport.NonTargetPeaksLoad(aSample, line, line)
                    End If
                Loop

                sr.Close()

                'Check for CCCheck
                If GlobalVariables.eTrain.Team = "CHROM" Then
                    For Each file In Directory.GetFiles(GlobalVariables.Import.FolderPath)
                        If InStr(file, "cccheck") Then
                            GlobalVariables.Import.strCCCheckFilePath = file
                            sr = New StreamReader(GlobalVariables.Import.strCCCheckFilePath)
                            'Look for Quant Time
                            Try
                                Do Until (InStr(line, "Quant Time:"))
                                    line = sr.ReadLine
                                Loop
                                aSample.CCCQuantTime = CDate(GlobalVariables.CSImport.DateFix(Trim(line.Substring(InStr(line, ":"), line.Length - InStr(line, ":"))), False))
                                'Look for Min. RRF
                                Do Until (InStr(line, "Min. RRF"))
                                    line = sr.ReadLine
                                Loop
                                lineGold = line
                                'Reset Line & RT
                                GlobalVariables.CSImport.Pos = 16
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                GlobalVariables.CSImport.Pos = 8
                                aSample.MinRRF = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                GlobalVariables.CSImport.Pos = 18
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                GlobalVariables.CSImport.Pos = 5
                                aSample.MinRelArea = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                GlobalVariables.CSImport.Pos = 16
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                aSample.MaxRTDev = Trim(line.Substring(0, line.Length))
                                line = sr.ReadLine
                                lineGold = line
                                GlobalVariables.CSImport.Pos = 16
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                GlobalVariables.CSImport.Pos = 5
                                aSample.MaxRRFDev = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                GlobalVariables.CSImport.Pos = 21
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                aSample.MaxRelArea = Trim(line.Substring(0, line.Length))
                                line = sr.ReadLine
                                line = sr.ReadLine
                                Do Until sr.EndOfStream
                                    GlobalVariables.CSImport.CCCheckLoad(aSample, line, line)
                                    line = sr.ReadLine
                                Loop
                            Catch ex As Exception
                                MsgBox("Error reading file: " & GlobalVariables.Import.strCCCheckFilePath & vbCrLf &
                                       "Line: " & lineGold & vbCrLf &
                                       "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                            End Try
                        End If
                    Next
                End If

                'Add sample to sample list
                GlobalVariables.SampleList.Add(aSample)
                GlobalVariables.NeedsCalculation = True

            Catch ex As Exception
                MsgBox("Error reading file: " & GlobalVariables.Import.FilePath & vbCrLf &
                "Line: " & line & vbCrLf &
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        ElseIf Type = "TOC" Then
            'Loop until first line of actual text
            Try
                Dim sr As StreamReader = New StreamReader(GlobalVariables.Import.FilePath)
                'First 2 lines are placeholders
                line = sr.ReadLine
                line = sr.ReadLine

                'Load samples
                Do Until sr.EndOfStream
                    Try
                        aSample = New Sample
                        line = sr.ReadLine
                        arrSplText = line.Split(",")
                        aSample.Type = arrSplText(0)
                        aSample.Analysis = arrSplText(1)
                        aSample.Name = arrSplText(2)
                        aSample.LimsID = arrSplText(3)
                        aSample.Result = arrSplText(4)
                        aSample.Units = arrSplText(5)
                        aSample.Vial = arrSplText(6)
                        aSample.AcqDate = CDate(arrSplText(7))
                        GlobalVariables.SampleList.Add(aSample)
                        GlobalVariables.NeedsCalculation = True
                    Catch ex As Exception
                        MsgBox("Error pulling sample information!" & vbCrLf &
                            "Line: " & line & vbCrLf &
                            "Logic Error: " & ex.Message, MsgBoxStyle.Critical)

                    End Try
                Loop
            Catch ex As Exception
                MsgBox("Error reading file: " & GlobalVariables.Import.FilePath & vbCrLf &
                "Line: " & line & vbCrLf &
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        ElseIf Type = "MASS" Then

            If GlobalVariables.eTrain.Location = "MIDLAND" Then
                If GlobalVariables.eTrain.Team = "CHROM" Then
                    If GlobalVariables.MHImport.MidlandChromImport() Then
                        GlobalVariables.NeedsCalculation = True
                    Else
                        MsgBox("There was an error reading in data from the files you selected. Please clear the Sample List and Imported Samples and try again.")
                    End If
                End If
            ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then

            End If
        ElseIf Type = "TQIII" Then

            If GlobalVariables.eTrain.Location = "MIDLAND" Then
                If GlobalVariables.eTrain.Team = "HR" Then
                    If GlobalVariables.TQ3Import.MidlandHRImport() Then
                        GlobalVariables.NeedsCalculation = True
                    Else
                        MsgBox("There was an error reading in data from the files you selected. Please clear the Sample List and Imported Samples and try again.")
                    End If
                    GlobalVariables.NeedsCalculation = True
                End If
            ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then

            End If

        ElseIf Type = "EDD" Then 'Added WT 9/26/2017

            Try
                Dim aSampleTemp As New Sample
                Dim sr As StreamReader = New StreamReader(GlobalVariables.Import.FilePath)
                line = sr.ReadLine
                line = sr.ReadLine
                Do Until line = ""

                    If (InStr(line, Chr(34))) Then
                        line = Regex.Replace(line, """", "")
                    End If

                    If Not aSampleTemp.CompoundList.Count = 0 Then 'Verify that there is at least one compound in the compound list
                        If (InStr(line, aSampleTemp.CompoundList(aSampleTemp.CompoundList.Count - 1).EDDsysSampleCode)) = 0 Or (InStr(line, aSampleTemp.CompoundList(aSampleTemp.CompoundList.Count - 1).EDDLabAnlMethodName) = 0) Then 'Check if the current sample is still the same sample 
                            Dim temp As String() = aSampleTemp.CompoundList(0).EDDSysSampleCode.Split("_")
                            aSampleTemp.LimsID = temp(0) 'aSampleTemp.CompoundList(0).EDDSysSampleCode.Substring(0, 6) '7 for space? <- Note: the Lims ID does not always start with an digit <<Ask??
                            GlobalVariables.SampleList.Add(aSampleTemp)
                            aSampleTemp = New Sample
                        End If
                    End If

                    arrSplText = line.Split(vbTab)
                    If (arrSplText(34) = "TRG" Or arrSplText(34) = "Target") And arrSplText(10) = "N" Then
                        loadEDD(arrSplText, aSampleTemp)
                    End If

                    line = sr.ReadLine()

                    'If end of the file, ensure last sample is added to Global sample list
                    If line = "" And Not aSampleTemp.CompoundList.Count = 0 Then
                        If aSampleTemp.CompoundList(0).EDDSysSampleCode.Length >= 6 Then
                            Dim tempTest As String
                            tempTest = aSampleTemp.CompoundList(0).EDDSysSampleCode.Substring(0, 6)
                            aSampleTemp.LimsID = aSampleTemp.CompoundList(0).EDDSysSampleCode.Substring(0, 6)
                            GlobalVariables.SampleList.Add(aSampleTemp)
                        End If
                    End If


                Loop
            Catch ex As Exception
                MsgBox("Error pulling sample information!" & vbCrLf &
                    "Logic Error: " & ex.Message & vbCrLf &
                    "(EDD may be formatted incorrectly. Please ensure EDD format is " & vbCrLf &
                    "correct and try again.)", MsgBoxStyle.Critical)

            End Try

        ElseIf Type = "SSR" Then

            Dim exApp As IApplication
            Dim workbook As IWorkbook
            Dim aSIS As New SIS

            Dim worksheet As IWorksheet
            Dim setupCounter As Integer
            Dim j As Integer
            Dim k As Integer

            Dim wsName As String

            Dim sampDate As String
            Dim sampTemp As String
            Dim analysis As String

            Dim lowerBound As Integer
            Dim upperBound As Integer
            Dim procRow As String

            ''Get row or range of rows
            'Dim strRows As String
            'strRows = InputBox("Enter row or row range (If entering row range, seperate beginning and ending row with a comma. Example: 2,6).")

            'Dim rowArr() As String
            'rowArr = strRows.Split(",")

            'Dim count As Integer

            'If rowArr.Count = 1 Then
            '    count = rowArr(0)
            'Else
            '    count = rowArr(1) - rowArr(0)
            'End If


            Try
                exApp = exEngine.Excel
                workbook = exApp.Workbooks.Open(GlobalVariables.Import.FilePath)

                Dim wksList As New ArrayList
                wksList.Add("F-Sewer 107")
                wksList.Add("CC2N 109")
                wksList.Add("WIF LS")
                wksList.Add("CC2E 106")
                wksList.Add("CC2E 104")

                For Each worksheet In workbook.Worksheets 'Start here next time to finish.. Saving workbook naming issue.. No need for macro enabled!!

                    wsName = worksheet.Name.Trim()

                    If wksList.Contains(wsName) Then
                        'If worksheet.Name = "F-Sewer 107" And worksheet.Name <> "LIMS Setup" Then

                        j = 6
                        k = 2

                        Do While worksheet.Range(j + 1, 2).Value <> ""
                            j = j + 1
                        Loop

                        setupCounter = 1

                        Do Until wsName = workbook.Worksheets("LIMS Setup").Range(setupCounter, 1).Value
                            setupCounter = setupCounter + 1
                        Loop

                        sampTemp = workbook.Worksheets("LIMS Setup").Range(setupCounter, 4).Value

                        Do Until wsName <> workbook.Worksheets("LIMS Setup").Range(setupCounter + 2, 1).Value

                            setupCounter = setupCounter + 1
                            k = k + 1

                            aCompound = New Compound
                            aSample = New Sample

                            aCompound.Name = workbook.Worksheets("LIMS Setup").Range(setupCounter, 2).Value
                            aCompound.ReportedAmt = worksheet.Range(j, k).Value

                            analysis = workbook.Worksheets("LIMS Setup").Range(setupCounter, 6).Value

                            aSample.CompoundList.Add(aCompound)

                            aSample.SampDate = workbook.Worksheets(worksheet.Name).Range(j, 2).Value.Replace("/", "-")
                            aSample.TQ3SampleID = sampTemp & "-" & aSample.SampDate 'Double check this next time!
                            aSample.Type = sampTemp
                            aSample.Analysis = analysis

                            GlobalVariables.SampleList.Add(aSample)

                        Loop

                    End If

                Next

            Catch ex As Exception
                MsgBox("Error import SIS information!" & vbCrLf &
                         "Sub Procedure: SISImport()" & vbCrLf &
                         "Logic Error: " & ex.Message, MsgBoxStyle.Critical)

            End Try
        ElseIf Type = "EUROLAN" Then

            Try
                Dim aSampleTemp As New Sample
                Dim sr As StreamReader = New StreamReader(GlobalVariables.Import.FilePath)
                line = sr.ReadLine
                line = sr.ReadLine
                Do Until line = ""

                    If (InStr(line, Chr(34))) Then
                        line = Regex.Replace(line, """", "")
                    End If

                    If Not aSampleTemp.CompoundList.Count = 0 Then 'Verify that there is at least one compound in the compound list
                        If (InStr(line, aSampleTemp.CompoundList(aSampleTemp.CompoundList.Count - 1).EDDsysSampleCode)) = 0 Or (InStr(line, aSampleTemp.CompoundList(aSampleTemp.CompoundList.Count - 1).EDDLabAnlMethodName) = 0) Then 'Check if the current sample is still the same sample 
                            Dim temp As String() = aSampleTemp.CompoundList(0).EDDSysSampleCode.Split("_")
                            aSampleTemp.LimsID = temp(0) 'aSampleTemp.CompoundList(0).EDDSysSampleCode.Substring(0, 6) '7 for space? <- Note: the Lims ID does not always start with an digit <<Ask??
                            GlobalVariables.SampleList.Add(aSampleTemp)
                            aSampleTemp = New Sample
                        End If
                    End If

                    arrSplText = line.Split(vbTab)
                    If (arrSplText(31) = "TRG" Or arrSplText(31) = "Target") Then
                        loadEDDEUROLAN(arrSplText, aSampleTemp)
                    End If

                    line = sr.ReadLine()

                    'If end of the file, ensure last sample is added to Global sample list
                    If line = "" And Not aSampleTemp.CompoundList.Count = 0 Then
                        If aSampleTemp.CompoundList(0).EDDSysSampleCode.Length >= 6 Then
                            Dim tempTest As String
                            tempTest = aSampleTemp.CompoundList(0).EDDSysSampleCode.Substring(0, 6)
                            aSampleTemp.LimsID = aSampleTemp.CompoundList(0).EDDSysSampleCode.Substring(0, 6)
                            GlobalVariables.SampleList.Add(aSampleTemp)
                        End If
                    End If

                Loop
            Catch ex As Exception
                MsgBox("Error pulling sample information!" & vbCrLf &
                    "Logic Error: " & ex.Message & vbCrLf &
                    "(EDD may be formatted incorrectly. Please ensure EDD format is " & vbCrLf &
                    "correct and try again.)", MsgBoxStyle.Critical)

            End Try
        ElseIf Type = "ALS" Then
            Try
                Dim aSampleTemp As New Sample
                Dim sr As StreamReader = New StreamReader(GlobalVariables.Import.FilePath)
                line = sr.ReadLine
                line = sr.ReadLine
                Dim Permit As New Permit
                Do Until line = ""
                    If (InStr(line, Chr(34))) Then
                        line = Regex.Replace(line, """", "")
                    End If
                    If Not aSampleTemp.CompoundList.Count = 0 Then 'Verify that there is at least one compound in the compound list
                        If (InStr(line, aSampleTemp.CompoundList(aSampleTemp.CompoundList.Count - 1).EDDsysSampleCode)) = 0 Or (InStr(line, aSampleTemp.CompoundList(aSampleTemp.CompoundList.Count - 1).EDDLabAnlMethodName) = 0) Then 'Check if the current sample is still the same sample 
                            Dim temp As String() = aSampleTemp.CompoundList(0).EDDSysSampleCode.Split("_")
                            aSampleTemp.LimsID = temp(0) 'aSampleTemp.CompoundList(0).EDDSysSampleCode.Substring(0, 6) '7 for space? <- Note: the Lims ID does not always start with an digit <<Ask??
                            GlobalVariables.SampleList.Add(aSampleTemp)
                            aSampleTemp = New Sample
                        End If
                    End If

                    arrSplText = line.Split(vbTab)
                    If (arrSplText(34) = "TRG" Or arrSplText(34) = "Target") And arrSplText(10) = "N" Then
                        loadEDD(arrSplText, aSampleTemp)
                    End If

                    line = sr.ReadLine()

                    'If end of the file, ensure last sample is added to Global sample list
                    If line = "" And Not aSampleTemp.CompoundList.Count = 0 Then
                        If aSampleTemp.CompoundList(0).EDDSysSampleCode.Length >= 6 Then
                            Dim tempTest As String
                            tempTest = aSampleTemp.CompoundList(0).EDDSysSampleCode.Substring(0, 6)
                            aSampleTemp.LimsID = aSampleTemp.CompoundList(0).EDDSysSampleCode.Substring(0, 6)
                            GlobalVariables.SampleList.Add(aSampleTemp)
                        End If
                    End If
                Loop
                Permit.LoadEddLimsCodes()
            Catch ex As Exception
                MsgBox("Error pulling sample information!" & vbCrLf &
                    "Logic Error: " & ex.Message & vbCrLf &
                    "(EDD may be formatted incorrectly. Please ensure EDD format is " & vbCrLf &
                    "correct and try again.)", MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    Function MidlandChromAttachCAS() As Boolean
        Dim exEngine As New ExcelEngine
        Dim exApp As IApplication
        Dim aSample As Sample
        Dim aStandard As Standard
        Dim aCompound As Compound
        Dim aSurrogate As Surrogate
        Dim intC As Integer
        Dim workbook As IWorkbook
        Dim worksheet As IWorksheet
        Dim arrComponents As New ArrayList
        Dim arrSpl() As String

        Try
            If GlobalVariables.eTrain.Location = "MIDLAND" Then
                If GlobalVariables.eTrain.Team = "CHROM" Then
                    exApp = exEngine.Excel
                    workbook = exApp.Workbooks.Open("\\mdrnd\AS-Global\Special_Access\EAC\Data\eTrain\DataFiles\Midland\Chrom\CasNo.xlsx")
                    worksheet = workbook.Worksheets(0)
                    intC = 0
                    Do While worksheet.Range(intC + 2, 1).Value <> ""
                        arrComponents.Add(worksheet.Range(intC + 2, 1).Value & " | " & worksheet.Range(intC + 2, 2).Value)
                        intC = intC + 1
                    Loop
                    workbook.Close()
                    exEngine.Dispose()
                    For Each aSample In GlobalVariables.SampleList
                        For Each aStandard In aSample.InternalStdList
                            For Each item In arrComponents
                                arrSpl = item.Split("|")
                                If aStandard.Name = Trim(arrSpl(0)) Then
                                    aStandard.CasNum = Trim(arrSpl(1))
                                End If
                            Next
                        Next
                        For Each aSurrogate In aSample.SurrogateList
                            For Each item In arrComponents
                                arrSpl = item.Split("|")
                                If aSurrogate.Name = Trim(arrSpl(0)) Then
                                    aSurrogate.CasNum = Trim(arrSpl(1))
                                End If
                            Next
                        Next
                        For Each aCompound In aSample.CompoundList
                            For Each item In arrComponents
                                arrSpl = item.Split("|")
                                If aCompound.Name = Trim(arrSpl(0)) Then
                                    aCompound.CasNum = Trim(arrSpl(1))
                                End If
                            Next
                        Next
                    Next

                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox("Error attaching CAS Numbers to components " & GlobalVariables.selProject & vbCrLf &
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            If Not IsNothing(workbook) Then
                workbook.Close()
                exEngine.Dispose()
            End If
            Return False
        End Try
    End Function

    Function FreeportChromBuildMBCompoundList(ByVal strPath As String) As Boolean
        Dim line As String
        Dim arrSplText() As String
        Dim aCompound As Compound
        Dim sr As StreamReader

        Try
            'Import file with analytes and limits for method blank report
            sr = New StreamReader(strPath)
            GlobalVariables.FreeportMBCompoundList.Clear()
            Do Until sr.EndOfStream
                line = sr.ReadLine
                arrSplText = line.Split("|")
                aCompound = New Compound
                aCompound.Name = arrSplText(0)
                aCompound.ChromMBLim = arrSplText(1)
                GlobalVariables.FreeportMBCompoundList.Add(aCompound)
            Loop
            sr.Close()
            Return True
        Catch ex As Exception
            MsgBox("Error reading Method Blank file for " & GlobalVariables.selProject & vbCrLf &
                "Line: " & line & vbCrLf &
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try


    End Function

    Function MidlandChromBuildMBCompoundList(ByVal strPath As String) As Boolean
        Dim line As String
        Dim arrSplText() As String
        Dim aCompound As Compound
        Dim sr As StreamReader

        Try
            'Import file with analytes and limits for method blank report
            sr = New StreamReader(strPath)
            GlobalVariables.MidlandMBCompoundList.Clear()
            Do Until sr.EndOfStream
                line = sr.ReadLine
                arrSplText = line.Split("|")
                aCompound = New Compound
                aCompound.Name = arrSplText(0)
                aCompound.ChromMBLim = arrSplText(1)
                GlobalVariables.MidlandMBCompoundList.Add(aCompound)
            Loop
            sr.Close()
            Return True
        Catch ex As Exception
            MsgBox("Error reading Method Blank file for " & GlobalVariables.selProject & vbCrLf &
                "Line: " & line & vbCrLf &
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try


    End Function

    Function FreeportChromBuildRecLimits(ByVal strFPath As String) As Boolean
        Dim exEngine As New ExcelEngine
        Dim exApp As IApplication
        Dim workbook As IWorkbook
        Dim wks As IWorksheet
        Dim wksICV As IWorksheet
        Dim wksCVS As IWorksheet
        Dim wksLCS As IWorksheet
        Dim wksMS As IWorksheet
        Dim wksReg As IWorksheet
        Dim aSample As Sample
        Dim aCompound As Compound
        Dim aSurrogate As Surrogate
        Dim arrSheets As New ArrayList
        Dim strAName As String
        Dim intC As Integer


        Try
            'Import all, ICV, CVS, LCS, MS, Reg if sheet exists
            exApp = exEngine.Excel
            workbook = exApp.Workbooks.Open(strFPath)
            For Each wks In workbook.Worksheets
                If InStr(wks.Name, "LCS", CompareMethod.Text) Then
                    wksLCS = wks
                    arrSheets.Add(wksLCS)
                ElseIf InStr(wks.Name, "MS", CompareMethod.Text) Then
                    wksMS = wks
                    arrSheets.Add(wksMS)
                ElseIf InStr(wks.Name, "ICV", CompareMethod.Text) Then
                    wksICV = wks
                    arrSheets.Add(wksICV)
                ElseIf InStr(wks.Name, "CVS", CompareMethod.Text) Then
                    wksCVS = wks
                    arrSheets.Add(wksCVS)
                ElseIf InStr(wks.Name, "old", CompareMethod.Text) = 0 Then
                    wksReg = wks
                    arrSheets.Add(wksReg)
                End If
            Next

            'Start loading in recovery limits
            For Each wks In arrSheets
                intC = 2
                strAName = wks.Range(intC, 1).Value
                Do Until strAName = ""
                    If InStr(strAName, "(S") Then
                        For Each aSample In GlobalVariables.ReportSamList
                            For Each aSurrogate In aSample.SurrogateList
                                If aSurrogate.Name = strAName Then
                                    If InStr(wks.Name, "LCS", CompareMethod.Text) Then
                                        aSurrogate.ChromLowLCSLim = wks.Range(intC, 2).Value
                                        aSurrogate.ChromUpLCSLim = wks.Range(intC, 3).Value
                                    ElseIf InStr(wks.Name, "MS", CompareMethod.Text) Then
                                        aSurrogate.ChromLowMSLim = wks.Range(intC, 2).Value
                                        aSurrogate.ChromUpMSLim = wks.Range(intC, 3).Value
                                    ElseIf InStr(wks.Name, "ICV", CompareMethod.Text) Then
                                        aSurrogate.ChromLowICVLim = wks.Range(intC, 2).Value
                                        aSurrogate.ChromUpICVLim = wks.Range(intC, 3).Value
                                    ElseIf InStr(wks.Name, "CVS", CompareMethod.Text) Then
                                        aSurrogate.ChromLowCVSLim = wks.Range(intC, 2).Value
                                        aSurrogate.ChromUpCVSLim = wks.Range(intC, 3).Value
                                    ElseIf InStr(wks.Name, "old", CompareMethod.Text) = 0 Then
                                        aSurrogate.ChromLowContLim = wks.Range(intC, 2).Value
                                        aSurrogate.ChromUpContLim = wks.Range(intC, 3).Value
                                    End If
                                End If
                            Next
                        Next
                    Else
                        For Each aSample In GlobalVariables.ReportSamList
                            For Each aCompound In aSample.CompoundList
                                If aCompound.Name = strAName Then
                                    If InStr(wks.Name, "LCS", CompareMethod.Text) Then
                                        aCompound.ChromLowLCSLim = wks.Range(intC, 2).Value
                                        aCompound.ChromUpLCSLim = wks.Range(intC, 3).Value
                                    ElseIf InStr(wks.Name, "MS", CompareMethod.Text) Then
                                        aCompound.ChromLowMSLim = wks.Range(intC, 2).Value
                                        aCompound.ChromUpMSLim = wks.Range(intC, 3).Value
                                    ElseIf InStr(wks.Name, "ICV", CompareMethod.Text) Then
                                        aCompound.ChromLowICVLim = wks.Range(intC, 2).Value
                                        aCompound.ChromUpICVLim = wks.Range(intC, 3).Value
                                    ElseIf InStr(wks.Name, "CVS", CompareMethod.Text) Then
                                        aCompound.ChromLowCVSLim = wks.Range(intC, 2).Value
                                        aCompound.ChromUpCVSLim = wks.Range(intC, 3).Value
                                    ElseIf InStr(wks.Name, "old", CompareMethod.Text) = 0 Then
                                        aCompound.ChromLowContLim = wks.Range(intC, 2).Value
                                        aCompound.ChromUpContLim = wks.Range(intC, 3).Value
                                    End If
                                End If
                            Next
                        Next
                    End If
                    intC = intC + 1
                    strAName = wks.Range(intC, 1).Value
                Loop
            Next

            workbook.Close()
            exEngine.Dispose()
            Return True
        Catch ex As Exception
            MsgBox("Error reading Recovery Limits File: " & strFPath & vbCrLf &
                   "Sub Procedure: FreeportChromBuildRecLimits()" & vbCrLf &
                    "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        'test


    End Function

    Function MidlandChromRecLimitsNames() As Boolean
        Dim exEngine As New ExcelEngine
        Dim exApp As IApplication
        Dim workbook As IWorkbook
        Dim wks As IWorksheet
        Dim arrSheets As New ArrayList

        Try
            exApp = exEngine.Excel
            workbook = exApp.Workbooks.Open("\\mdrnd\AS-Global\Special_Access\EAC\Chrom\ReportTemplates\2009_Spike Recovery Limits.xlsx")
            For Each wks In workbook.Worksheets
                If wks.Visibility = WorksheetVisibility.Visible Then
                    GlobalVariables.MidlandChromRLimitNames.Add(wks.Name)
                End If
            Next
            workbook.Close()
            exEngine.Dispose()
            Return True
        Catch ex As Exception
            MsgBox("Error reading Recovery Limits File" & vbCrLf &
                   "Sub Procedure: MidlandChromRecLimitsNames()" & vbCrLf &
                    "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try
    End Function


    Function MidlandChromBuildRecLimits(ByVal strSelectLim As String) As Boolean
        Dim exEngine As New ExcelEngine
        Dim exApp As IApplication
        Dim workbook As IWorkbook
        Dim wks As IWorksheet
        Dim aSample As Sample
        Dim aCompound As Compound
        Dim aSurrogate As Surrogate
        Dim arrSheets As New ArrayList
        Dim strAName As String
        Dim intC As Integer


        Try
            'Import all visible sheets
            exApp = exEngine.Excel
            workbook = exApp.Workbooks.Open("\\mdrnd\AS-Global\Special_Access\EAC\Chrom\ReportTemplates\2009_Spike Recovery Limits.xlsx")
            For Each wks In workbook.Worksheets
                If wks.Name = strSelectLim Then
                    arrSheets.Add(wks)
                    Exit For
                End If
            Next

            'Start loading in recovery limits
            For Each wks In arrSheets
                intC = 2
                strAName = wks.Range(intC, 1).Value
                Do Until strAName = ""
                    For Each aSample In GlobalVariables.ReportSamList
                        For Each aCompound In aSample.CompoundList
                            If aCompound.Name = strAName Then
                                aCompound.ChromLowContLim = wks.Range(intC, 2).Value
                                aCompound.ChromUpContLim = wks.Range(intC, 3).Value
                            End If
                        Next
                        For Each aSurrogate In aSample.SurrogateList
                            If aSurrogate.Name = strAName Then
                                aSurrogate.ChromLowContLim = wks.Range(intC, 2).Value
                                aSurrogate.ChromUpContLim = wks.Range(intC, 3).Value
                            End If
                        Next
                    Next
                    intC = intC + 1
                    strAName = wks.Range(intC, 1).Value
                Loop
            Next

            workbook.Close()
            exEngine.Dispose()
            Return True
        Catch ex As Exception
            MsgBox("Error reading Recovery Limits" & vbCrLf &
                   "Sub Procedure: MidlandChromBuildRecLimits()" & vbCrLf &
                    "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try




    End Function

    Function SISImport(ByVal strSISLoc As String) As Boolean
        Dim exEngine As New ExcelEngine
        Dim exApp As IApplication
        Dim workbook As IWorkbook
        Dim aSIS As New SIS
        Dim aSample As Sample
        Dim wksSIS As IWorksheet
        Dim wksWeight As IWorksheet
        Dim wksManualInteg As IWorksheet
        Dim worksheet As IWorksheet
        Dim arrSplText() As String
        Dim i As Integer
        Dim u As Integer

        Try
            exApp = exEngine.Excel
            workbook = exApp.Workbooks.Open(strSISLoc)
            'Find applicable sheets
            For Each worksheet In workbook.Worksheets
                If worksheet.Name = "SIS" Then
                    wksSIS = worksheet
                ElseIf worksheet.Name = "Weight_Sheet" Then
                    wksWeight = worksheet
                ElseIf worksheet.Name = "GCMS_Manual_Integration" Then
                    wksManualInteg = worksheet
                End If
            Next


            'Load SIS 2.0 details
            aSIS.ProjNum = wksSIS.Range("C1").Value
            aSIS.ProjName = wksSIS.Range("C2").Value
            aSIS.Method = wksSIS.Range("C3").Value
            aSIS.Analysis = wksSIS.Range("C4").Value
            aSIS.SampMatrix = wksSIS.Range("C5").Value
            aSIS.Compliance = wksSIS.Range("C6").Value
            aSIS.SetNum = wksSIS.Range("E1").Value
            aSIS.Contact = wksSIS.Range("E2").Value
            aSIS.CostCenter = wksSIS.Range("E3").Value
            aSIS.StartDate = wksSIS.Range("E4").Value
            aSIS.EndDate = wksSIS.Range("E5").Value
            aSIS.ConfAnalysis = wksSIS.Range("E6").Value
            aSIS.PrepAnalyst = wksSIS.Range("I2").Value
            aSIS.Extraction = wksSIS.Range("I3").Value
            aSIS.CleanUpCols = wksSIS.Range("I4").Value
            aSIS.Methylation = wksSIS.Range("I5").Value
            aSIS.AddAnalyses = wksSIS.Range("I6").Value
            aSIS.Analyst = wksSIS.Range("L2").Value
            aSIS.Instrument = wksSIS.Range("L3").Value
            aSIS.Reviewer = wksSIS.Range("L4").Value

            'Get team and EOA/VOA
            aSIS.Team = wksSIS.Range("L5").Value
            If wksSIS.Range("M5").Value = "EOA" Then
                aSIS.EOA = True
                aSIS.VOA = False
            ElseIf wksSIS.Range("M5").Value = "VOA" Then
                aSIS.VOA = True
                aSIS.EOA = False
            End If

            'Get CS-Method
            aSIS.CSMethod = wksSIS.Range("L6").Value

            'Load samples
            'Look through Internal ID column
            i = 1
            Do While wksSIS.Range("A" & CStr(7 + i)).Value <> ""
                i = i + 1
            Loop

            'Account for column heading
            i = i - 1

            'Collect sample information and add to collection
            For a = 1 To i
                aSample = New Sample
                aSample.SISInternalID = wksSIS.Range("A" & CStr(7 + a)).Value
                aSample.SISLabNum = wksSIS.Range("B" & CStr(7 + a)).Value
                aSample.SISClientSampID = wksSIS.Range("C" & CStr(7 + a)).Value
                'Handle if there is end sample date
                If InStr(1, wksSIS.Range("D" & CStr(7 + a)).Value, "-", vbBinaryCompare) > 0 Then
                    arrSplText = wksSIS.Range("D" & CStr(7 + a)).Value.Split("-")
                    aSample.SISSampDate = CDate(arrSplText(0))
                    aSample.SISSampDateEnd = CDate(arrSplText(1))
                Else
                    aSample.SISSampDate = CDate(wksSIS.Range("D" & CStr(7 + a)).Value)
                End If
                aSample.SISTargetSampSize = wksSIS.Range("E" & CStr(7 + a)).Value
                aSample.SISActualSampSize = wksSIS.Range("F" & CStr(7 + a)).Value
                If InStr(wksSIS.Range("H" & CStr(7 + a)).Value, "=IF") Then
                    aSample.SISDefaultAliquot = wksSIS.Range("H" & CStr(7 + a)).FormulaNumberValue
                Else
                    aSample.SISDefaultAliquot = wksSIS.Range("H" & CStr(7 + a)).Value
                End If

                aSample.SISAnalyses = wksSIS.Range("I" & CStr(7 + a)).Value
                aSample.SISSpikeMult = wksSIS.Range("J" & CStr(7 + a)).Value
                If InStr(wksSIS.Range("K" & CStr(7 + a)).Value, "=IF") Then
                    aSample.SISDilFactor = wksSIS.Range("K" & CStr(7 + a)).FormulaNumberValue
                Else
                    aSample.SISDilFactor = wksSIS.Range("K" & CStr(7 + a)).Value
                End If
                aSample.SISFinalWeight = wksSIS.Range("G" & CStr(7 + a)).Value
                'Weights
                aSample.SISTinWeight = wksWeight.Range("J" & CStr(9 + a)).Value
                aSample.SISWetWeight = wksWeight.Range("I" & CStr(9 + a)).Value
                aSample.SISDryWeight = wksWeight.Range("K" & CStr(9 + a)).Value
                aSample.SISSampleWeight = wksWeight.Range("U" & CStr(9 + a)).Value
                aSample.SISSBottWeight = wksWeight.Range("P" & CStr(9 + a)).Value
                aSample.SISSampWetWeight = wksWeight.Range("D" & CStr(9 + a)).Value
                aSample.SISEBottWeight = wksWeight.Range("Q" & CStr(9 + a)).Value
                aSample.SISPMoisture = wksWeight.Range("N" & CStr(9 + a)).FormulaNumberValue
                aSample.SISType = "Sample"
                aSIS.SampleList.Add(aSample)
                aSample = Nothing
            Next

            'Write in acq date to manual integration for FAST
            If GlobalVariables.eTrain.Team = "FAST" Then
                u = 0
                Do Until wksManualInteg.Range("B" & CStr(5 + u)).Value = ""
                    For Each aSample In GlobalVariables.SampleList
                        If Left(aSample.DataFile, aSample.DataFile.Length - 2) = wksManualInteg.Range("B" & CStr(5 + u)).Value Then
                            wksManualInteg.Range("E" & CStr(5 + u)).Value = aSample.AcqDate
                        End If
                    Next
                    u = u + 1
                Loop
                workbook.Save()
            End If


            'Store aSIS
            GlobalVariables.SISList.Add(aSIS)

            workbook.Close()
            exEngine.Dispose()
            Return True
        Catch ex As Exception
            MsgBox("Error import SIS information!" & vbCrLf &
                 "Sub Procedure: SISImport()" & vbCrLf &
                 "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try
        Return False

    End Function

    'Read in Elution Dictionary
    Function ElutionImport(ByVal strPath As String)
        Dim line As String
        Dim aStandard As Standard
        Dim aSurrogate As Surrogate
        Dim aCompound As Compound
        Dim arrSplText() As String
        Dim sr As StreamReader

        Try
            'Read in file
            sr = New StreamReader(strPath)
            GlobalVariables.ElutionOrderSample = New Sample
            line = sr.ReadLine
            line = sr.ReadLine
            Do Until line = "Surrogates"
                aStandard = New Standard
                arrSplText = line.Split("|")
                aStandard.Name = arrSplText(0)
                GlobalVariables.ElutionOrderSample.InternalStdList.Add(aStandard)
                line = sr.ReadLine
            Loop
            line = sr.ReadLine
            Do Until line = "Compounds"
                aSurrogate = New Surrogate
                arrSplText = line.Split("|")
                aSurrogate.Name = arrSplText(0)
                GlobalVariables.ElutionOrderSample.SurrogateList.Add(aSurrogate)
                line = sr.ReadLine
            Loop
            Do Until sr.EndOfStream
                line = sr.ReadLine
                aCompound = New Compound
                arrSplText = line.Split("|")
                aCompound.Name = arrSplText(0)
                GlobalVariables.ElutionOrderSample.CompoundList.Add(aCompound)
            Loop
            sr.Close()
            sr = Nothing
            Return True
        Catch ex As Exception
            MsgBox("Error reading Elution Dictionary file." & vbCrLf &
                "Line: " & line & vbCrLf &
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try
    End Function

    'Recursive file search, looks through every folder and file and adds to a list box
    Sub FileSearch(ByVal dir As String, ByVal mask As String)
        Dim d As String
        Dim f As String
        Try
            For Each f In Directory.GetFiles(dir, mask)
                arrFileList.Add(f)
            Next
            For Each d In Directory.GetDirectories(dir)
                FileSearch(d, mask)
            Next
        Catch e As System.Exception
            Debug.WriteLine(e.Message)
        End Try
    End Sub

    'Load data into Sample object
    'Load Sample object into SampleList
    Sub loadEDD(ByVal arr() As String, aSampleTemp As Sample) 'Added WT 9/26/2017

        Dim aCompound As New Compound

        aCompound.EDDsysSampleCode = arr(0)
        'aSampleTemp.AnalysisDate = arr(1)
        aCompound.EDDLabAnlMethodName = arr(1)
        aCompound.EDDAnalysisDate = arr(2)
        aCompound.EDDAnalysisTime = arr(3)
        aCompound.EDDTotalOrDissolved = arr(4)
        aCompound.EDDColumnNumber = arr(5)
        aCompound.EDDTestType = arr(6)
        aCompound.EDDLabMatrixCode = arr(7)
        aCompound.EDDAnalysisLocation = arr(8)
        aCompound.EDDBasis = arr(9)
        aCompound.EDDSampleTypeCode = arr(10)
        aCompound.EDDEDilutionFactor = arr(11)
        aCompound.EDDPrepMethod = arr(12)
        aCompound.EDDPrepDate = arr(13)
        aCompound.EDDPrepTime = arr(14)
        aCompound.EDDLeachateMethod = arr(15)
        aCompound.EDDLeachateDate = arr(16)
        aCompound.EDDLeachateTime = arr(17)
        aCompound.EDDLabNameCode = arr(18)
        aCompound.EDDQcLevel = arr(19)
        aCompound.EDDLabSampleID = arr(20)
        aCompound.EDDPercentMoisture = arr(21)
        aCompound.EDDSubsampleAmount = arr(22)
        aCompound.EDDSubsampleAmountUnit = arr(23)
        aCompound.EDDAnalystName = arr(24)
        aCompound.EDDInstrumentID = arr(25)
        aCompound.EDDComment = arr(26)
        aCompound.EDDPreservative = arr(27)
        aCompound.EDDFinalVolume = arr(28)
        aCompound.EDDFinalVolumeUnit = arr(29)
        aCompound.EDDCasRn = arr(30)
        aCompound.EDDChemicalName = arr(31)
        aCompound.EDDResultValue = arr(32)
        aCompound.EDDResultErrorDelta = arr(33)
        aCompound.EDDResultTypeCode = arr(34)
        aCompound.EDDReportableResult = arr(35)
        aCompound.EDDDetectFlag = arr(36)
        aCompound.EDDLabQualifiers = arr(37)
        aCompound.EDDValidatorQualifiers = arr(38)
        aCompound.EDDOrganicYn = arr(39)
        aCompound.EDDMethodDetectionLimit = arr(40)
        aCompound.EDDReportingDetectionLimit = arr(41)
        aCompound.EDDQuantitationLimit = arr(42)
        aCompound.EDDResultUnit = arr(43)
        aCompound.EDDDetectionLimitUnit = arr(44)
        aCompound.EDDTicRetentionTime = arr(45)
        aCompound.EDDResultComment = arr(46)
        aCompound.EDDQcOriginalConc = arr(47)
        aCompound.EDDQcSpikeAdded = arr(48)
        aCompound.EDDQcSpikeMeasured = arr(49)
        aCompound.EDDQcSpikeRecovery = arr(50)
        aCompound.EDDQcDupOriginalConc = arr(51)
        aCompound.EDDQcDupSpikeAdded = arr(52)
        aCompound.EDDQcDupSpikeMeasured = arr(53)
        aCompound.EDDQcDupSpikeRecovery = arr(54)
        aCompound.EDDQcRpd = arr(55)
        aCompound.EDDQcSpikeLcl = arr(56)
        aCompound.EDDQcSpikeUcl = arr(57)
        aCompound.EDDQcRpdCl = arr(58)
        aCompound.EDDQcSpikeStatus = arr(59)
        aCompound.EDDQcDupSpikeStatus = arr(60)
        aCompound.EDDQcRpdStatus = arr(61)
        aCompound.EDDRlOrMdl = arr(62)


        aSampleTemp.CompoundList.Add(aCompound)
    End Sub

    Sub loadEDDEUROLAN(ByVal arr() As String, aSampleTemp As Sample) 'Added WT 9/26/2017

        Dim aCompound As New Compound

        aCompound.EDDsysSampleCode = arr(0)
        aCompound.EDDLabAnlMethodName = arr(1)
        aCompound.EDDAnalysisDate = arr(2)
        ' aCompound.EDDAnalysisTime = arr(3)
        aCompound.EDDTotalOrDissolved = arr(3)
        aCompound.EDDColumnNumber = arr(4)
        aCompound.EDDTestType = arr(5)
        aCompound.EDDLabMatrixCode = arr(6)
        aCompound.EDDAnalysisLocation = arr(7)
        aCompound.EDDBasis = arr(8)
        aCompound.EDDContainerID = arr(9)
        aCompound.EDDEDilutionFactor = arr(10)
        aCompound.EDDPrepMethod = arr(11)
        aCompound.EDDPrepDate = arr(12)
        'aCompound.EDDPrepTime = arr(14)
        aCompound.EDDLeachateMethod = arr(13)
        aCompound.EDDLeachateDate = arr(14)
        'aCompound.EDDLeachateTime = arr(17)
        aCompound.EDDLabNameCode = arr(15)
        aCompound.EDDQcLevel = arr(16)
        aCompound.EDDLabSampleID = arr(17)
        aCompound.EDDPercentMoisture = arr(18)
        aCompound.EDDSubsampleAmount = arr(19)
        aCompound.EDDSubsampleAmountUnit = arr(20)
        aCompound.EDDAnalystName = arr(21)
        aCompound.EDDInstrumentID = arr(22)
        aCompound.EDDComment = arr(23)
        aCompound.EDDPreservative = arr(24)
        aCompound.EDDFinalVolume = arr(25)
        aCompound.EDDFinalVolumeUnit = arr(26)
        aCompound.EDDCasRn = arr(27)
        aCompound.EDDChemicalName = arr(28)
        aCompound.EDDResultValue = arr(29)
        aCompound.EDDResultErrorDelta = arr(30)
        aCompound.EDDResultTypeCode = arr(31)
        aCompound.EDDReportableResult = arr(32)
        aCompound.EDDDetectFlag = arr(33)
        aCompound.EDDLabQualifiers = arr(34)
        aCompound.EDDValidatorQualifiers = arr(35)
        aCompound.EDDInterpretedQualifier = arr(36)
        aCompound.EDDOrganicYn = arr(37)
        aCompound.EDDMethodDetectionLimit = arr(38)
        aCompound.EDDReportingDetectionLimit = arr(39)
        aCompound.EDDQuantitationLimit = arr(40)
        aCompound.EDDResultUnit = arr(41)
        aCompound.EDDDetectionLimitUnit = arr(42)
        aCompound.EDDTicRetentionTime = arr(43)
        aCompound.EDDResultComment = arr(44)
        aCompound.EDDSDG = arr(45)
        aCompound.EDDQcOriginalConc = arr(46)
        aCompound.EDDQcSpikeAdded = arr(47)
        aCompound.EDDQcSpikeMeasured = arr(48)
        aCompound.EDDQcSpikeRecovery = arr(49)
        aCompound.EDDQcDupOriginalConc = arr(50)
        aCompound.EDDQcDupSpikeAdded = arr(51)
        aCompound.EDDQcDupSpikeMeasured = arr(52)
        aCompound.EDDQcDupSpikeRecovery = arr(53)
        aCompound.EDDQcRpd = arr(54)
        aCompound.EDDQcSpikeLcl = arr(55)
        aCompound.EDDQcSpikeUcl = arr(56)
        aCompound.EDDQcRpdCl = arr(57)
        aCompound.EDDQcSpikeStatus = arr(58)
        aCompound.EDDQcDupSpikeStatus = arr(59)
        aCompound.EDDQcRpdStatus = arr(60)
        aCompound.EDDCustomField2 = arr(61)
        aCompound.EDDCustomField3 = arr(62)
        aCompound.EDDCustomField4 = arr(63)
        aCompound.EDDCustomField5 = arr(64)
        aCompound.EDDUncertainty = arr(65)
        aCompound.EDDMinimumDetectableConc = arr(66)
        aCompound.EDDCountingError = arr(67)
        aCompound.EDDCriticalValue = arr(68)


        aSampleTemp.CompoundList.Add(aCompound)
    End Sub
End Class



