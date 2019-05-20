Imports System.Text.RegularExpressions

Public Class CSImport
    Private intPos As Integer

    Public Property Pos() As Integer
        Get
            Return intPos
        End Get
        Set(ByVal value As Integer)
            intPos = value
        End Set
    End Property

    'Sub to add load in Internal Std data from Chemstation import
    Sub StandardLoad(ByRef aSample As Sample, ByVal line As String, ByVal lineGold As String)
        Dim Standard As New Standard

        'Check in case a Standard is not in file
        If line = "" Or InStr(line, "-------") Or InStr(line, "Internal Standards") Then
            Exit Sub
        Else
            Try
                'Get Name
                GlobalVariables.CSImport.Pos = InStr(line, ")")
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 31
                Standard.Name = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                'Reset Line & RT
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 8
                Standard.RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                'Reset Line & QIon
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 5
                Standard.QIon = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                'Reset Line & Response & check for MI
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 10
                If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "m", CompareMethod.Text) Then
                    Standard.MI = True
                    Standard.Response = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
                Else
                    Standard.MI = False
                    Standard.Response = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                End If
                'Reset Line & Conc
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 10
                Standard.Conc = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                'Reset Line & Units
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 5
                Standard.Units = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                If aSample.Units = "" Then
                    aSample.Units = Standard.Units
                End If
                'Reset Line & DevMin
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                Standard.DevMin = Trim(line.Substring(0, line.Length))
                aSample.InternalStdList.Add(Standard)
            Catch ex As Exception
                MsgBox("Error reading file: " & GlobalVariables.Import.FilePath & vbCrLf & _
                       "Line: " & lineGold & vbCrLf & _
                       "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    'Sub to add load in Surrogate data from Chemstation import
    Sub SurrogateLoad(ByRef aSample As Sample, ByVal line As String, ByVal lineGold As String)
        Dim Surrogate As New Surrogate


        'Check in case a Surrogate is not in file
        If line = "" Or InStr(line, "-------") Or InStr(line, "System Monitoring Compounds") Then
            Exit Sub
        Else
            Try
                'Get Name
                GlobalVariables.CSImport.Pos = InStr(line, ")")
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 31
                Surrogate.Name = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                'Check for Methylated
                If InStr(Surrogate.Name, "Med)") Then
                    aSample.Methylated = True
                    Surrogate.Methylated = True
                End If
                'Reset Line & RT
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 8
                Surrogate.RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                'Reset Line & QIon
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 5
                Surrogate.QIon = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                'Reset Line & Response & check for MI
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 10
                If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "m", CompareMethod.Text) Then
                    Surrogate.MI = True
                    Surrogate.Response = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
                Else
                    Surrogate.MI = False
                    Surrogate.Response = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                End If
                'Reset Line & Conc
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 10
                Surrogate.Conc = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                'Reset Line & Units
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 5
                Surrogate.Units = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                If aSample.Units = "" Then
                    aSample.Units = Surrogate.Units
                End If
                'Reset Line & DevMin
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 8
                Surrogate.DevMin = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                'Reset Line & Spike
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 20
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 11
                Surrogate.SpkAmt = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                'Reset Line & Recovery
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 34
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                Surrogate.Recovery = Trim(line.Substring(0, line.Length - 1))
                aSample.SurrogateList.Add(Surrogate)
            Catch ex As Exception
                MsgBox("Error reading file: " & GlobalVariables.Import.FilePath & vbCrLf & _
                       "Line: " & lineGold & vbCrLf & _
                       "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    Sub NonTargetPeaksLoad(ByRef aSample As Sample, ByVal line As String, ByVal lineGold As String)
        Dim Compound As New Compound


        'Check in case a Compound is not in file
        If line = "" Or InStr(line, "-------") Or InStr(line, "Non Target Peaks") Then
            Exit Sub
        Else
            'Look for ")" and fix line for processing
            Try
                'Get Name
                If GlobalVariables.Import.Type = "CHEMBEVCAN" Then
                    'line = lineGold
                    GlobalVariables.CSImport.Pos = 4
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 26
                    Compound.Name = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    'Check for Methylated
                    If InStr(Compound.Name, "Med)") Then
                        aSample.Methylated = True
                        Compound.Methylated = True
                    End If
                    'Reset Line & RT
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 9
                    Compound.RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    'Reset Line & QIon
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 15
                    Compound.Area = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    Compound.Height = Trim(line)
                Else
                    GlobalVariables.CSImport.Pos = InStr(line, ")")
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 31
                    Compound.Name = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    'Check for Methylated
                    If InStr(Compound.Name, "Med)") Then
                        aSample.Methylated = True
                    End If
                    'Reset Line & RT
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 8
                    Compound.RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    'Reset Line & QIon
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 5
                    Compound.QIon = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    'Reset Line & Response & check for MI
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 10
                    If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "m", CompareMethod.Text) Then
                        Compound.MI = True
                        Compound.Response = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
                    Else
                        Compound.MI = False
                        Compound.Response = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    End If
                    'Reset Line & Conc
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 10
                    Compound.Conc = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    'Reset Line & Units
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 5
                    If Compound.Conc = "N.D." Then
                        If GlobalVariables.eTrain.Team = "CHROM" Then
                            Compound.Units = ""
                            Compound.Conc = "N.D."
                        Else
                            Compound.Units = ""
                            Compound.Conc = "0"
                        End If
                        If Not String.IsNullOrWhiteSpace(line) Then
                            'Reset Line & QValue
                            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                            If InStr(line, "#") Then
                                Compound.QOOR = True
                                Compound.QValue = Trim(line.Substring(0, line.Length - 1))
                            Else
                                Compound.QOOR = False
                                Compound.QValue = Trim(line.Substring(0, line.Length))
                            End If
                        End If
                    ElseIf Compound.Conc = "No Calib" Then
                        Compound.Units = ""
                        Compound.QValue = ""
                    Else
                        Compound.Units = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        If Not String.IsNullOrWhiteSpace(line) Then
                            If InStr(line, "#") Then
                                GlobalVariables.CSImport.Pos = 2
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                Compound.QOOR = True
                                Compound.QValue = Trim(line.Substring(0, line.Length))
                            Else
                                Compound.QOOR = False
                                Compound.QValue = Trim(line.Substring(0, line.Length))
                            End If
                        End If
                    End If
                End If
                If aSample.Units = "" Then
                    aSample.Units = Compound.Units
                End If
                aSample.CompoundList.Add(Compound)

            Catch ex As Exception
                MsgBox("Error reading file: " & GlobalVariables.Import.FilePath & vbCrLf & _
                      "Line: " & lineGold & vbCrLf & _
                      "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    'Sub to add load in Target Compound data from Chemstation import
    Sub CompoundLoad(ByRef aSample As Sample, ByVal line As String, ByVal lineGold As String)
        Dim Compound As New Compound
        Dim aCompound As New Compound

        'Check in case a Compound is not in file
        If line = "" Or InStr(line, "-------") Or InStr(line, "Target Compounds") Then
            Exit Sub
        Else
            'Look for ")" and fix line for processing
            Try
                If GlobalVariables.Import.Type = "CHEMBEVCAN" Then
                    GlobalVariables.CSImport.Pos = 10
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 25
                    Compound.Name = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    'Check for Methylated
                    If InStr(Compound.Name, "Med)") Then
                        aSample.Methylated = True
                    End If
                    'Reset Line & RT
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 9
                    Compound.RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    'Reset Line & QIon
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 14
                    If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "m", CompareMethod.Text) Then
                        Compound.MI = True
                        Compound.Response = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
                    Else
                        Compound.MI = False
                        Compound.Response = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    End If
                    'Reset Line & Conc
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 9
                    Compound.Conc = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    'Reset Line & Units
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 7
                    Compound.Units = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                Else
                    'Get Name
                    GlobalVariables.CSImport.Pos = InStr(line, ")")
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 31
                    Compound.Name = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    'Check for Methylated
                    If InStr(Compound.Name, "Med)") Then
                        aSample.Methylated = True
                    End If
                    'Reset Line & RT
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 8
                    Compound.RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    'Reset Line & QIon
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 5
                    Compound.QIon = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    'Reset Line & Response & check for MI
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 10
                    If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "m", CompareMethod.Text) Then
                        Compound.MI = True
                        Compound.Response = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
                    Else
                        Compound.MI = False
                        Compound.Response = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    End If
                    'Reset Line & Conc
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 10
                    Compound.Conc = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                    'Reset Line & Units
                    line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                    GlobalVariables.CSImport.Pos = 5
                    If Compound.Conc = "N.D." Then
                        If GlobalVariables.eTrain.Team = "CHROM" Then
                            Compound.Units = ""
                            Compound.Conc = "N.D."
                        Else
                            Compound.Units = ""
                            Compound.Conc = "0"
                        End If
                        If Not String.IsNullOrWhiteSpace(line) Then
                            'Reset Line & QValue
                            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                            If InStr(line, "#") Then
                                Compound.QOOR = True
                                Compound.QValue = Trim(line.Substring(0, line.Length - 1))
                            Else
                                Compound.QOOR = False
                                Compound.QValue = Trim(line.Substring(0, line.Length))
                            End If
                        End If
                    ElseIf Compound.Conc = "No Calib" Then
                        Compound.Units = ""
                        Compound.QValue = ""
                    Else
                        Compound.Units = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        If Not String.IsNullOrWhiteSpace(line) Then
                            If InStr(line, "#") Then
                                GlobalVariables.CSImport.Pos = 2
                                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                                Compound.QOOR = True
                                Compound.QValue = Trim(line.Substring(0, line.Length))
                            Else
                                Compound.QOOR = False
                                Compound.QValue = Trim(line.Substring(0, line.Length))
                            End If
                        End If
                    End If
                End If
                aSample.CompoundList.Add(Compound)

            Catch ex As Exception
                MsgBox("Error reading file: " & GlobalVariables.Import.FilePath & vbCrLf & _
                      "Line: " & lineGold & vbCrLf & _
                      "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    Sub CCCheckLoad(ByRef aSample As Sample, ByVal line As String, ByVal lineGold As String)
        Dim aStandard As Standard
        Dim aSurrogate As Surrogate
        Dim aCompound As Compound

        'Handle lines that could cause problems
        If line = "" Or InStr(line, "-------") Or InStr(line, "Compound") Then
            Exit Sub
        Else
            'Look for ")" and fix line for processing
            Try
                'Get Name
                GlobalVariables.CSImport.Pos = 9
                line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                GlobalVariables.CSImport.Pos = 29
                For Each aStandard In aSample.InternalStdList
                    If aStandard.Name = Trim(line.Substring(0, GlobalVariables.CSImport.Pos)) Then
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        GlobalVariables.CSImport.Pos = 6
                        aStandard.AvgRF = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        GlobalVariables.CSImport.Pos = 8
                        aStandard.CCRF = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        GlobalVariables.CSImport.Pos = 11
                        If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "#", CompareMethod.Text) Then
                            aStandard.PercDevOOR = True
                            aStandard.PercDev = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
                        Else
                            aStandard.PercDevOOR = False
                            aStandard.PercDev = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        End If
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        GlobalVariables.CSImport.Pos = 5
                        If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "#", CompareMethod.Text) Then
                            aStandard.PercAreaOOR = True
                            aStandard.PercArea = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
                        Else
                            aStandard.PercAreaOOR = False
                            aStandard.PercArea = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        End If
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "#", CompareMethod.Text) Then
                            aStandard.CCCDevMinOOR = True
                            aStandard.CCCDevMin = Trim(line.Substring(0, line.Length - 1))
                        Else
                            aStandard.CCCDevMinOOR = False
                            aStandard.CCCDevMin = Trim(line.Substring(0, line.Length))
                        End If
                        Exit Sub
                    End If
                Next
                For Each aSurrogate In aSample.SurrogateList
                    If aSurrogate.Name = Trim(line.Substring(0, GlobalVariables.CSImport.Pos)) Then
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        GlobalVariables.CSImport.Pos = 6
                        aSurrogate.AvgRF = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        GlobalVariables.CSImport.Pos = 8
                        aSurrogate.CCRF = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        GlobalVariables.CSImport.Pos = 11
                        If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "#", CompareMethod.Text) Then
                            aSurrogate.PercDevOOR = True
                            aSurrogate.PercDev = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
                        Else
                            aSurrogate.PercDevOOR = False
                            aSurrogate.PercDev = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        End If
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        GlobalVariables.CSImport.Pos = 5
                        If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "#", CompareMethod.Text) Then
                            aSurrogate.PercAreaOOR = True
                            aSurrogate.PercArea = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
                        Else
                            aSurrogate.PercAreaOOR = False
                            aSurrogate.PercArea = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        End If
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "#", CompareMethod.Text) Then
                            aSurrogate.CCCDevMinOOR = True
                            aSurrogate.CCCDevMin = Trim(line.Substring(0, line.Length - 1))
                        Else
                            aSurrogate.CCCDevMinOOR = False
                            aSurrogate.CCCDevMin = Trim(line.Substring(0, line.Length))
                        End If
                        Exit Sub
                    End If
                Next
                For Each aCompound In aSample.CompoundList
                    If aCompound.Name = Trim(line.Substring(0, GlobalVariables.CSImport.Pos)) Then
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        GlobalVariables.CSImport.Pos = 6
                        aCompound.AvgRF = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        GlobalVariables.CSImport.Pos = 8
                        aCompound.CCRF = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        GlobalVariables.CSImport.Pos = 11
                        If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "#", CompareMethod.Text) Then
                            aCompound.PercDevOOR = True
                            aCompound.PercDev = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
                        Else
                            aCompound.PercDevOOR = False
                            aCompound.PercDev = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        End If
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        GlobalVariables.CSImport.Pos = 5
                        If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "#", CompareMethod.Text) Then
                            aCompound.PercAreaOOR = True
                            aCompound.PercArea = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
                        Else
                            aCompound.PercAreaOOR = False
                            aCompound.PercArea = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
                        End If
                        line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
                        If InStr(line.Substring(0, GlobalVariables.CSImport.Pos), "#", CompareMethod.Text) Then
                            aCompound.CCCDevMinOOR = True
                            aCompound.CCCDevMin = Trim(line.Substring(0, line.Length - 1))
                        Else
                            aCompound.CCCDevMinOOR = False
                            aCompound.CCCDevMin = Trim(line.Substring(0, line.Length))
                        End If
                        Exit Sub
                    End If
                Next
            Catch ex As Exception
                MsgBox("Error reading file: " & GlobalVariables.Import.FilePath & vbCrLf & _
                      "Line: " & lineGold & vbCrLf & _
                      "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    Sub TMPStdLoad(ByRef aStandard As Standard, ByVal line As String, ByVal lineGold As String, ByVal IType As String)

        If IType = "T" Then
            GlobalVariables.CSImport.Pos = 4
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aStandard.TSignal = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aStandard.TRatios = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 16
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aStandard.TRT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aStandard.RTLLim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 12
            aStandard.TResp = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aStandard.TIntType = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
        ElseIf IType = "Q1" Then
            GlobalVariables.CSImport.Pos = 4
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aStandard.Q1Signal = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aStandard.Q1Ratios = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            If aStandard.Q1Ratios.Length >= 7 Then
                GlobalVariables.CSImport.Pos = aStandard.Q1Ratios.Length + 1
            End If
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aStandard.Q1SLLim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aStandard.Q1SULim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aStandard.Q1RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 12
            aStandard.Q1Resp = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aStandard.Q1IntType = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
        ElseIf IType = "Q2" Then
            GlobalVariables.CSImport.Pos = 4
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aStandard.Q2Signal = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aStandard.Q2Ratios = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aStandard.Q1SLLim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aStandard.Q1SULim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aStandard.Q2RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aStandard.RTULim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 12
            aStandard.Q2Resp = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aStandard.Q2IntType = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
        ElseIf IType = "Q3" Then
            GlobalVariables.CSImport.Pos = 4
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aStandard.Q3Signal = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aStandard.Q3Ratios = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aStandard.Q1SLLim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aStandard.Q1SULim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aStandard.Q3RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 12
            aStandard.Q3Resp = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aStandard.Q3IntType = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
        End If
    End Sub


    Sub TMPSurrLoad(ByRef aSurrogate As Surrogate, ByVal line As String, ByVal lineGold As String, ByVal IType As String)

        If IType = "T" Then
            GlobalVariables.CSImport.Pos = 4
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aSurrogate.TSignal = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aSurrogate.TRatios = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 16
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aSurrogate.TRT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aSurrogate.RTLLim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 12
            aSurrogate.TResp = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aSurrogate.TIntType = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
        ElseIf IType = "Q1" Then
            GlobalVariables.CSImport.Pos = 4
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aSurrogate.Q1Signal = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aSurrogate.Q1Ratios = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aSurrogate.Q1SLLim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aSurrogate.Q1SULim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aSurrogate.Q1RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 12
            aSurrogate.Q1Resp = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aSurrogate.Q1IntType = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
        ElseIf IType = "Q2" Then
            GlobalVariables.CSImport.Pos = 4
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aSurrogate.Q2Signal = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aSurrogate.Q2Ratios = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aSurrogate.Q1SLLim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aSurrogate.Q1SULim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aSurrogate.Q2RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aSurrogate.RTULim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 12
            aSurrogate.Q2Resp = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aSurrogate.Q2IntType = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
        ElseIf IType = "Q3" Then
            GlobalVariables.CSImport.Pos = 4
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aSurrogate.Q3Signal = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aSurrogate.Q3Ratios = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aSurrogate.Q1SLLim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aSurrogate.Q1SULim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aSurrogate.Q3RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 12
            aSurrogate.Q3Resp = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aSurrogate.Q3IntType = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
        End If
    End Sub

    Sub TMPCompLoad(ByRef aCompound As Compound, ByVal line As String, ByVal lineGold As String, ByVal IType As String)

        If IType = "T" Then
            GlobalVariables.CSImport.Pos = 4
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aCompound.TSignal = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aCompound.TRatios = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 16
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aCompound.TRT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aCompound.RTLLim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 12
            aCompound.TResp = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aCompound.TIntType = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
        ElseIf IType = "Q1" Then
            GlobalVariables.CSImport.Pos = 4
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aCompound.Q1Signal = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aCompound.Q1Ratios = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aCompound.Q1SLLim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aCompound.Q1SULim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aCompound.Q1RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 12
            aCompound.Q1Resp = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aCompound.Q1IntType = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
        ElseIf IType = "Q2" Then
            GlobalVariables.CSImport.Pos = 4
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aCompound.Q2Signal = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aCompound.Q2Ratios = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aCompound.Q2SLLim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aCompound.Q2SULim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aCompound.Q2RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aCompound.RTULim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 12
            aCompound.Q2Resp = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aCompound.Q2IntType = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
        ElseIf IType = "Q3" Then
            GlobalVariables.CSImport.Pos = 4
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aCompound.Q3Signal = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 7
            aCompound.Q3Ratios = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aCompound.Q3SLLim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos - 1))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aCompound.Q3SULim = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 6
            aCompound.Q3RT = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 12
            aCompound.Q3Resp = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 1
            line = line.Substring(GlobalVariables.CSImport.Pos, line.Length - GlobalVariables.CSImport.Pos)
            GlobalVariables.CSImport.Pos = 8
            aCompound.Q3IntType = Trim(line.Substring(0, GlobalVariables.CSImport.Pos))
        End If
    End Sub

    'Function to format date correctly in ChemStation files
    Function DateFix(ByVal line As String, ByVal HasDay As Boolean) As String
        Dim arrSplText() As String
        Dim nDate As String

        arrSplText = line.Split(" ")
        'Handle if day is in front
        If HasDay Then
            nDate = arrSplText(1) & " " & arrSplText(2) & " " & arrSplText(4) & " " & arrSplText(3)
        Else
            nDate = arrSplText(0) & " " & arrSplText(1) & " " & arrSplText(3) & " " & arrSplText(2)
        End If

        Return nDate

    End Function
End Class
