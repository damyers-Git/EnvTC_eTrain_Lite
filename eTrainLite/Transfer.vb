Imports System.Data
Imports System.Data.Odbc
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Transfer


    'Sends data to LIMS via textfile
    Function ToLIMS(ByVal strUserID As String)
        Dim strPath As String
        Dim aSample As Sample
        Dim aSIS As SIS
        Dim aSISSample As Sample
        Dim objWriter As StreamWriter
        Dim aStandard As Standard
        Dim aSurrogate As Surrogate
        Dim aCompound As Compound
        Dim d As DateTime
        Dim intFileCounter As Integer
        Dim strIDText As String
        Dim strUni As String

        intFileCounter = 0
        d = DateTime.Now


        For Each aSample In GlobalVariables.SampleList
            If aSample.Include Then
                Try

                    If GlobalVariables.eTrain.Server = "SEADRIFT" Then
                        strPath = "C:\TEST\test" & intFileCounter.ToString("000") & ".txt"
                        'strUni = d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt"
                        'strPath = GlobalVariables.eTrain.ServerFP & d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt" '<- Original path
                        objWriter = New System.IO.StreamWriter(strPath)
                        objWriter.WriteLine("$INDTMODE = S")

                        'Dim EDDLIMSMethodName As String
                        'EDDLIMSMethodName = getLIMSMethodName(aSample.CompoundList(0).EDDLabAnlMethodName, aSample.CompoundList(0).EDDDetectionLimitUnit)


                        Do While True
                            If Not Regex.IsMatch(aSample.LimsID, "^[0-9 ]+$") Then
                                aSample.LimsID = InputBox("No LIMS ID located in EDD. Please enter a LIMS ID:", "LIMS ID")
                                Continue Do
                            End If
                            Exit Do
                        Loop


                        If aSample.LimsID <> "" Then
                            objWriter.WriteLine("$SAMPLEID = " & aSample.LimsID)
                        Else
                            strIDText = aSample.Name & "-" & d.ToString("dd-MM-yyyy-HHmmss-") & intFileCounter.ToString("000") & "-" & aSample.Type 'Verify
                            objWriter.WriteLine("$SAMPLEID = " & strIDText)
                        End If

                        If aSample.LimsID <> "" Then
                            'objWriter.WriteLine("$SAMPTMPL = LAB_EDD")
                        End If
                        objWriter.WriteLine("$ANALYSIS = " & matchAnalysisName(aSample.CompoundList(0).EDDLabAnlMethodName)) 'matchAnalysisName(aSample.CompoundList(0).EDDLabAnlMethodName))
                        objWriter.WriteLine("$REPLNUMB = 0")
                        objWriter.WriteLine("$OPERATOR = EDD")
                        objWriter.WriteLine("$ANALYSTN = " & strUserID)

                        If aSample.LimsID <> "" Then
                            objWriter.WriteLine("$NEWSAMPL = FALSE")
                        Else
                            objWriter.WriteLine("$NEWSAMPL = TRUE")

                        End If

                        objWriter.WriteLine("$INSTRMNT = ") '& aSample.CompoundList(0).EDDLabAnlMethodName)
                        objWriter.WriteLine("$SOURCE_N = 2")
                        objWriter.WriteLine("$SOURCE_1 = ")
                        objWriter.WriteLine("$SOURCE_2 = Contact D. Myers 989-636-6204, J. Durham 989-638-9406, or W. Towne 989-633-1975")
                        objWriter.WriteLine("$SAMP_FLD = dow_field_02") '?" & aSample.DetectLimitType)
                        objWriter.WriteLine("$SAMP_FLD = dow_field_03") '?" & aSample.AcqDate) 

                        For Each aCompound In aSample.CompoundList

                            If aCompound.EDDDetectFlag = "Y" Then
                                objWriter.WriteLine("?" & aCompound.EDDChemicalName & "  ?N  ?" & GlobalVariables.Calculations.FormatSF(aCompound.EDDResultValue) & "?    ?    ?" & aCompound.EDDEDilutionFactor)
                                'objWriter.WriteLine("?" & "Component 1" & "  ?N  ?" & "123456" & "?    ?    ?" & aCompound.EDDEDilutionFactor) '<- Line Created to test ACTUAL transfer to LIMS
                            Else
                                objWriter.WriteLine("?" & aCompound.EDDChemicalName & "  ?N  ?" & "0.00" & "?    ?    ?" & aCompound.EDDEDilutionFactor)
                                'objWriter.WriteLine("?" & "Component 1" & "  ?N  ?" & "0.00" & "?    ?    ?" & aCompound.EDDEDilutionFactor) '<- Line Created to test ACTUAL transfer to LIMS
                            End If

                        Next
                        'Close file
                        objWriter.Close() 'Added WT 10/13/2017 -> Close File
                        'File.Copy(strPath & strUni, GlobalVariables.eTrain.ServerFP & strUni)
                    ElseIf GlobalVariables.eTrain.Server = "ROH" Then
                        'strPath = "C:\Users\u411882\Desktop\TestFolder\eTrainLite\"
                        'strUni = d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt"
                        'strPath = GlobalVariables.eTrain.ServerFP & d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt" '<- Original path
                        'strPath = "S:\TEST\test" & intFileCounter.ToString("000") & ".txt"
                        strPath = "\\usmdlsdowacds1\LIMS_XFER\ROHNA\" & d.ToString("ddMMyy") & d.ToString("HHmm") & "-" & intFileCounter.ToString("000") & ".txt"
                        objWriter = New System.IO.StreamWriter(strPath)
                        objWriter.WriteLine("$INDTMODE = S")

                        'Dim EDDLIMSMethodName As String
                        'EDDLIMSMethodName = getLIMSMethodName(aSample.CompoundList(0).EDDLabAnlMethodName, aSample.CompoundList(0).EDDDetectionLimitUnit)

                        Do While True
                            If Not Regex.IsMatch(aSample.LimsID, "^[0-9 ]+$") Then
                                aSample.LimsID = InputBox("No LIMS ID located in EDD. Please enter a LIMS ID:", "LIMS ID")
                                Continue Do
                            End If
                            Exit Do
                        Loop

                        If aSample.LimsID <> "" Then
                            objWriter.WriteLine("$SAMPLEID = " & aSample.LimsID)
                        Else
                            strIDText = aSample.Name & "-" & d.ToString("dd-MM-yyyy-HHmmss-") & intFileCounter.ToString("000") & "-" & aSample.Type 'Verify
                            objWriter.WriteLine("$SAMPLEID = " & strIDText)
                        End If

                        If aSample.LimsID <> "" Then
                            'objWriter.WriteLine("$SAMPTMPL = LAB_EDD")
                        End If
                        objWriter.WriteLine("$ANALYSIS = " & matchAnalysisName(aSample.CompoundList(0).EDDLabAnlMethodName)) 'matchAnalysisName(aSample.CompoundList(0).EDDLabAnlMethodName)) 'aSample.CompoundList(0).EDDLabAnlMethodName)
                        objWriter.WriteLine("$REPLNUMB = 0")
                        objWriter.WriteLine("$OPERATOR = EDD")
                        objWriter.WriteLine("$ANALYSTN = " & strUserID)

                        If aSample.LimsID <> "" Then
                            objWriter.WriteLine("$NEWSAMPL = FALSE")
                        Else
                            objWriter.WriteLine("$NEWSAMPL = TRUE")

                        End If

                        objWriter.WriteLine("$INSTRMNT = ") '& aSample.CompoundList(0).EDDInstrumentID)
                        objWriter.WriteLine("$SOURCE_N = 2")
                        objWriter.WriteLine("$SOURCE_1 = ")
                        objWriter.WriteLine("$SOURCE_2 = Contact D. Myers 989-636-6204, J. Durham 989-638-9406, or W. Towne 989-633-1975")
                        objWriter.WriteLine("$SAMP_FLD = dow_field_02") '?" & aSample.DetectLimitType)
                        objWriter.WriteLine("$SAMP_FLD = dow_field_03") '?" & aSample.AcqDate) 

                        For Each aCompound In aSample.CompoundList

                            If aCompound.EDDDetectFlag = "Y" Then
                                objWriter.WriteLine("?" & matchComponentName(aCompound.EDDChemicalName) & "  ?N  ?" & GlobalVariables.Calculations.FormatSF(aCompound.EDDResultValue) & "?    ?    ?" & aCompound.EDDEDilutionFactor)
                                'objWriter.WriteLine("?" & "Component 1" & "  ?N  ?" & "123456" & "?    ?    ?" & aCompound.EDDEDilutionFactor) '<- Line Created to test ACTUAL transfer to LIMS
                            Else
                                objWriter.WriteLine("?" & matchComponentName(aCompound.EDDChemicalName) & "  ?N  ?" & "0.00" & "?    ?    ?" & aCompound.EDDEDilutionFactor)
                                'objWriter.WriteLine("?" & "Component 1" & "  ?N  ?" & "0.00" & "?    ?    ?" & aCompound.EDDEDilutionFactor) '<- Line Created to test ACTUAL transfer to LIMS
                            End If

                        Next
                        'Close file
                        objWriter.Close() 'Added WT 10/13/2017 -> Close File
                        'File.Copy(strPath & strUni, GlobalVariables.eTrain.ServerFP & strUni)

                    ElseIf GlobalVariables.eTrain.Location = "MIDLAND" Then
                        If GlobalVariables.eTrain.Server = "MIDLAND" Then
                            If GlobalVariables.eTrain.Team = "CLAB" Then
                                If GlobalVariables.Import.Type = "GRABS" Then

                                    d = DateTime.Now

                                    strPath = GlobalVariables.eTrain.ServerFP & d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt"
                                    'strPath = "C:\Users\nb98715\Desktop\CLab_Test\" & d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt"
                                    objWriter = New System.IO.StreamWriter(strPath)

                                    'Header info
                                    objWriter.WriteLine("$IDNTMODE = S")
                                    objWriter.WriteLine("$SAMPLEID = " & aSample.LimsID)
                                    objWriter.WriteLine("$ANALYSIS = " & aSample.Analysis)
                                    objWriter.WriteLine("$REPLNUMB = 0")
                                    objWriter.WriteLine("$OPERATOR = CONTLAB") ' Change to something else?
                                    objWriter.WriteLine("$ANALYSTN = " & strUserID)
                                    objWriter.WriteLine("$NEWSAMPL = FALSE")
                                    objWriter.WriteLine("$INSTRMNT = ") '& aSample.Instrument)
                                    objWriter.WriteLine("$SOURCE_N = 2")
                                    objWriter.WriteLine("$SOURCE_1 = MIOPS WWTP Grabs Contract Lab Data")
                                    objWriter.WriteLine("$SOURCE_2 = CONTACT W. Bodeis 989-636-5245")
                                    objWriter.WriteLine("$SAMP_FLD = dow_field_02?") '& aSample.DetectLimitType)
                                    objWriter.WriteLine("$SAMP_FLD = dow_field_03?") '& aSample.AcqDate)

                                    For Each aCompound In aSample.CompoundList
                                        ' Skipping methylChlorpyrifos beacuse it isn't in LIMS and should be added into the chlorpyrifos (Dursban) amount in the verifyCLabData() method call.
                                        If aCompound.EDDCasRn = "5598-13-0" Or aCompound.EDDChemicalName = "Chlorpyrifos, Methyl" Then
                                            Continue For
                                        End If
                                        ' Dilution factor set to 1 because the DF calculation is done by the lab to the reported value so it doesn't need it applied a second time. 
                                        objWriter.WriteLine("?" & aCompound.EDDChemicalName & "  ?N  ?" & aCompound.EDDResultValue & "  ?  ?" & aCompound.EDDReportingDetectionLimit & "  ?1")
                                    Next
                                    objWriter.Close()
                                    intFileCounter = intFileCounter + 1
                                ElseIf GlobalVariables.Import.Type = "031A" Then

                                    d = DateTime.Now

                                    strPath = GlobalVariables.eTrain.ServerFP & d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt"
                                    'strPath = "C:\Users\nb98715\Desktop\CLab_Temp\" & d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt"
                                    objWriter = New System.IO.StreamWriter(strPath)

                                    'Header info
                                    objWriter.WriteLine("$IDNTMODE = S")
                                    objWriter.WriteLine("$SAMPLEID = " & aSample.LimsID)
                                    objWriter.WriteLine("$ANALYSIS = " & aSample.Analysis)
                                    objWriter.WriteLine("$REPLNUMB = 0")
                                    objWriter.WriteLine("$OPERATOR = CONTLAB") ' Change to something else?
                                    objWriter.WriteLine("$ANALYSTN = " & strUserID)
                                    objWriter.WriteLine("$NEWSAMPL = FALSE")
                                    objWriter.WriteLine("$INSTRMNT = ") '& aSample.Instrument)
                                    objWriter.WriteLine("$SOURCE_N = 2")
                                    objWriter.WriteLine("$SOURCE_1 = MIOPS NPDES 031A Contract Lab Data")
                                    objWriter.WriteLine("$SOURCE_2 = CONTACT W. Bodeis 989-636-5245")
                                    objWriter.WriteLine("$SAMP_FLD = dow_field_02?") '& aSample.DetectLimitType)
                                    objWriter.WriteLine("$SAMP_FLD = dow_field_03?") '& aSample.AcqDate)

                                    For Each aCompound In aSample.CompoundList
                                        ' Dilution factor set to 1 because the DF calculation is done by the lab to the reported value so it doesn't need it applied a second time. 
                                        objWriter.WriteLine("?" & aCompound.EDDChemicalName & "  ?N  ?" & aCompound.EDDResultValue & "  ?  ?" & aCompound.EDDReportingDetectionLimit & "  ?1")
                                    Next
                                    objWriter.Close()
                                    intFileCounter = intFileCounter + 1
                                ElseIf GlobalVariables.Import.Type = "031B" Then

                                    d = DateTime.Now

                                    strPath = GlobalVariables.eTrain.ServerFP & d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt"
                                    'strPath = "C:\Users\nb98715\Desktop\CLab_Temp\" & d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt"
                                    objWriter = New System.IO.StreamWriter(strPath)

                                    'Header info
                                    objWriter.WriteLine("$IDNTMODE = S")
                                    objWriter.WriteLine("$SAMPLEID = " & aSample.LimsID)
                                    objWriter.WriteLine("$ANALYSIS = " & aSample.Analysis)
                                    objWriter.WriteLine("$REPLNUMB = 0")
                                    objWriter.WriteLine("$OPERATOR = CONTLAB") ' Change to something else?
                                    objWriter.WriteLine("$ANALYSTN = " & strUserID)
                                    objWriter.WriteLine("$NEWSAMPL = FALSE")
                                    objWriter.WriteLine("$INSTRMNT = ") '& aSample.Instrument)
                                    objWriter.WriteLine("$SOURCE_N = 2")
                                    objWriter.WriteLine("$SOURCE_1 = MIOPS NPDES 031B Contract Lab Data")
                                    objWriter.WriteLine("$SOURCE_2 = CONTACT W. Bodeis 989-636-5245")
                                    objWriter.WriteLine("$SAMP_FLD = dow_field_02?") '& aSample.DetectLimitType)
                                    objWriter.WriteLine("$SAMP_FLD = dow_field_03?") '& aSample.AcqDate)

                                    For Each aCompound In aSample.CompoundList
                                        ' Dilution factor set to 1 because the DF calculation is done by the lab to the reported value so it doesn't need it applied a second time. 
                                        objWriter.WriteLine("?" & aCompound.EDDChemicalName & "  ?N  ?" & aCompound.EDDResultValue & "  ?  ?" & aCompound.EDDReportingDetectionLimit & "  ?1")
                                    Next
                                    objWriter.Close()
                                    intFileCounter = intFileCounter + 1
                                ElseIf GlobalVariables.Import.Type = "031C" Then

                                    d = DateTime.Now

                                    strPath = GlobalVariables.eTrain.ServerFP & d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt"
                                    'strPath = "C:\Users\nb98715\Desktop\CLab_Temp\" & d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt"
                                    objWriter = New System.IO.StreamWriter(strPath)

                                    'Header info
                                    objWriter.WriteLine("$IDNTMODE = S")
                                    objWriter.WriteLine("$SAMPLEID = " & aSample.LimsID)
                                    objWriter.WriteLine("$ANALYSIS = " & aSample.Analysis)
                                    objWriter.WriteLine("$REPLNUMB = 0")
                                    objWriter.WriteLine("$OPERATOR = CONTLAB") ' Change to something else?
                                    objWriter.WriteLine("$ANALYSTN = " & strUserID)
                                    objWriter.WriteLine("$NEWSAMPL = FALSE")
                                    objWriter.WriteLine("$INSTRMNT = ") '& aSample.Instrument)
                                    objWriter.WriteLine("$SOURCE_N = 2")
                                    objWriter.WriteLine("$SOURCE_1 = MIOPS NPDES 031C Contract Lab Data")
                                    objWriter.WriteLine("$SOURCE_2 = CONTACT W. Bodeis 989-636-5245")
                                    objWriter.WriteLine("$SAMP_FLD = dow_field_02?") '& aSample.DetectLimitType)
                                    objWriter.WriteLine("$SAMP_FLD = dow_field_03?") '& aSample.AcqDate)

                                    For Each aCompound In aSample.CompoundList
                                        ' Dilution factor set to 1 because the DF calculation is done by the lab to the reported value so it doesn't need it applied a second time. 
                                        objWriter.WriteLine("?" & aCompound.EDDChemicalName & "  ?N  ?" & aCompound.EDDResultValue & "  ?  ?" & aCompound.EDDReportingDetectionLimit & "  ?1")
                                    Next
                                    objWriter.Close()
                                    intFileCounter = intFileCounter + 1
                                ElseIf GlobalVariables.Import.Type = "DF" Or GlobalVariables.eTrain.AnalysisLab = "VISTA" Then
                                    ' Variable only used for reporting the vista data. 
                                    ' Total TEQ and 2378-OCDD TEQ scores are reported as the same thing so the variable is used vs making the method call twice. 
                                    ' The check2378TEQ method was left and commented out in case anything changes in the future. 
                                    Dim dblTEQScore As Double
                                    d = DateTime.Now

                                    strPath = GlobalVariables.eTrain.ServerFP & d.ToString("ddMMyy") & d.ToString("HHmm") & "-" & intFileCounter.ToString("000") & ".txt"
                                    'strPath = "C:\Users\nb98715\Desktop\CLab_Test\" & d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt"
                                    objWriter = New System.IO.StreamWriter(strPath)

                                    'Header info
                                    objWriter.WriteLine("$IDNTMODE = S")
                                    objWriter.WriteLine("$NEWSAMPL = FALSE")
                                    objWriter.WriteLine("$SAMPLEID = " & aSample.LimsID)
                                    objWriter.WriteLine("$ANALYSIS = " & aSample.Analysis)
                                    objWriter.WriteLine("$REPLNUMB = 0")
                                    objWriter.WriteLine("$OPERATOR = BATCH")
                                    objWriter.WriteLine("$ANALYSTN = " & strUserID)
                                    objWriter.WriteLine("$INSTRMNT = _VISTA")
                                    objWriter.WriteLine("$SOURCE_N = 2")
                                    objWriter.WriteLine("$SOURCE_1 = Midland HR Data Transfer")
                                    objWriter.WriteLine("$SOURCE_2 = CONTACT W. Bodeis 989-636-5245")
                                    objWriter.WriteLine("$SAMP_FLD = description?" & aSample.CompoundList(0).EDDLabSampleID)
                                    objWriter.WriteLine("$SAMP_FLD = job_name?")
                                    objWriter.WriteLine("$SAMP_FLD = sample_name?" & aSample.CompoundList(0).EDDsysSampleCode)

                                    For Each aCompound In aSample.CompoundList
                                        objWriter.WriteLine("?" & aCompound.EDDChemicalName & "?N? " & aCompound.EDDResultValue)
                                        If aCompound.EDDLabQualifiers = "" Then
                                            objWriter.WriteLine("?" & aCompound.EDDChemicalName & " Flags?T?NONE")
                                        Else
                                            objWriter.WriteLine("?" & aCompound.EDDChemicalName & " Flags?T?" & aCompound.EDDLabQualifiers)
                                        End If
                                    Next
                                    dblTEQScore = calculateTotalTEQ(aSample.LimsID)
                                    objWriter.WriteLine("?Total TEQ (ppq)?N?" & dblTEQScore)
                                    objWriter.WriteLine("?Total TEQ (ppq)?T?")
                                    objWriter.WriteLine("?Total TEQ (ppq) Flags?T?NONE")

                                    objWriter.WriteLine("?2378-TCDD_TEQ (ND=0)?N?" & dblTEQScore)
                                    objWriter.WriteLine("?2378-TCDD_TEQ (ND=0)?T?")
                                    'objWriter.WriteLine("?2378-TCDD_TEQ (ND=0) Flags?T?NONE")

                                    objWriter.Close()
                                    intFileCounter = intFileCounter + 1
                                    ' For the rest of the CLab data getting sent into LIMS
                                Else

                                    d = DateTime.Now

                                    strPath = GlobalVariables.eTrain.ServerFP & d.ToString("ddMMyy") & d.ToString("HHmm") & "-" & intFileCounter.ToString("000") & ".txt"
                                    'strPath = "C: \Users\nb98715\Desktop\CLab_Test\" & d.ToString("ddMMyy") & "-" & d.ToString("HHmm") & intFileCounter.ToString("000") & ".txt"
                                    objWriter = New System.IO.StreamWriter(strPath)

                                    'Header info
                                    objWriter.WriteLine("$IDNTMODE = S")
                                    objWriter.WriteLine("$SAMPLEID = " & aSample.LimsID)
                                    'objWriter.WriteLine("$SAMPTMPL = " & aSample.Type)
                                    objWriter.WriteLine("$ANALYSIS = " & aSample.Analysis)
                                    objWriter.WriteLine("$REPLNUMB = 0")
                                    objWriter.WriteLine("$OPERATOR = CONTLAB")
                                    objWriter.WriteLine("$ANALYSTN = " & strUserID)
                                    objWriter.WriteLine("$NEWSAMPL = FALSE")
                                    objWriter.WriteLine("$INSTRMNT = ")
                                    objWriter.WriteLine("$SOURCE_N = 2")
                                    objWriter.WriteLine("$SOURCE_1 = MIOPS Contract Lab Data")
                                    objWriter.WriteLine("$SOURCE_2 = CONTACT W. Bodeis 989-636-5245")
                                    objWriter.WriteLine("$SAMP_FLD = dow_field_02?")
                                    objWriter.WriteLine("$SAMP_FLD = dow_field_03?") '& LIMSDate(aSample.CompoundList(0).EDDAnalysisDate))
                                    'objWriter.WriteLine("SAMP_FLD = " & LIMSDate(aSample.CompoundList(0).EDDAnalysisDate))
                                    For Each aCompound In aSample.CompoundList
                                        objWriter.WriteLine("?" & aCompound.EDDChemicalName & "  ?N  ?" & aCompound.EDDResultValue & "  ?  ?" & aCompound.EDDLabQualifiers & "  ?" & aCompound.EDDEDilutionFactor)
                                    Next
                                    objWriter.Close()
                                    intFileCounter = intFileCounter + 1
                                End If

                            ElseIf GlobalVariables.eTrain.Team = "AECOM" Then

                                'Construct sample files with data from import
                                'This is not Will now
                                If Not aSample.CompoundList.Item(0).ReportedAmt = "" And IsNumeric(aSample.CompoundList.Item(0).ReportedAmt) Then

                                    d = DateTime.Now

                                    strPath = GlobalVariables.eTrain.ServerFP & d.ToString("ddMMyy") & d.ToString("HHmm") & "-" & intFileCounter.ToString("000") & ".txt"
                                    'strPath = "C:\Users\ua20088\Documents\TEST\" & d.ToString("ddMMyy") & d.ToString("HHmm") & "-" & intFileCounter.ToString("000") & ".txt"
                                    objWriter = New System.IO.StreamWriter(strPath)

                                    objWriter.WriteLine("$IDNTMODE = S")
                                    objWriter.WriteLine("$SAMPLEID = " & aSample.LimsID)
                                    objWriter.WriteLine("$SAMPTMPL = " & aSample.Type)
                                    objWriter.WriteLine("$ANALYSIS = " & aSample.Analysis)
                                    objWriter.WriteLine("$REPLNUMB = 0")
                                    objWriter.WriteLine("$OPERATOR = BATCH")
                                    objWriter.WriteLine("$ANALYSTN = ETRAIN")
                                    objWriter.WriteLine("NEWSAMPL = True")
                                    objWriter.WriteLine("SOURCE_N = 2")
                                    objWriter.WriteLine("SOURCE_1 = MIOPS Sewer Study Result Summary")
                                    objWriter.WriteLine("SOURCE_2 = CONTACT D. MEYERS 989-636-6204 Wyatt Towne 989-633-1975")
                                    objWriter.WriteLine("SAMP_FLD = " & LIMSDate(aSample.SampDate))
                                    For Each aCompound In aSample.CompoundList
                                        objWriter.WriteLine("?" & aSample.CompoundList.Item(0).Name.Split("(")(0).Trim() & "  ?N  ?  " & aSample.CompoundList.Item(0).ReportedAmt & "?")
                                    Next
                                    objWriter.Close()
                                    intFileCounter = intFileCounter + 1

                                End If
                            End If
                        End If
                    End If
                Catch ex As Exception
                    MsgBox("Error writing file for LIMS Transfer!" & vbCrLf &
                                   "Sub Procedure: ToLIMS()" & vbCrLf &
                                   "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                    'objWriter.Close() 
                    'IO.File.Delete(strPath)
                    Return False
                End Try
                intFileCounter += 1
            End If
        Next







        Return True
    End Function
    ' Dates could come in with either a "-" or an "/" so it is checked for both possibilites, then split.
    ' Months could be written with a leading zero or by itself so it is converted to a single digit.
    ' Converting to an integer will take off the leading zero. 
    ' EX:  06 -> 6  for June.
    Function LIMSDate(ByVal sDate As String) As String
        Dim arrSpl() As String
        If sDate.Contains("/") Then
            arrSpl = sDate.Split("/")
        ElseIf sDate.Contains("-") Then
            arrSpl = sDate.Split("-")
        End If
        If Convert.ToInt32(arrSpl(0)) = "1" Then
            arrSpl(0) = "JAN"
        ElseIf Convert.ToInt32(arrSpl(0)) = "2" Then
            arrSpl(0) = "FEB"
        ElseIf Convert.ToInt32(arrSpl(0)) = "3" Then
            arrSpl(0) = "MAR"
        ElseIf Convert.ToInt32(arrSpl(0)) = "4" Then
            arrSpl(0) = "APR"
        ElseIf Convert.ToInt32(arrSpl(0)) = "5" Then
            arrSpl(0) = "MAY"
        ElseIf Convert.ToInt32(arrSpl(0)) = "6" Then
            arrSpl(0) = "JUN"
        ElseIf Convert.ToInt32(arrSpl(0)) = "7" Then
            arrSpl(0) = "JUL"
        ElseIf Convert.ToInt32(arrSpl(0)) = "8" Then
            arrSpl(0) = "AUG"
        ElseIf Convert.ToInt32(arrSpl(0)) = "9" Then
            arrSpl(0) = "SEP"
        ElseIf arrSpl(0) = "10" Then
            arrSpl(0) = "OCT"
        ElseIf arrSpl(0) = "11" Then
            arrSpl(0) = "NOV"
        ElseIf arrSpl(0) = "12" Then
            arrSpl(0) = "DEC"
        End If

        sDate = arrSpl(1) & "-" & arrSpl(0) & "-" & arrSpl(2)
        Return sDate

    End Function

    'Added wmtowne -> 10/31/2017 To match LIMS Method name to EDD Method name
    Function getLIMSMethodName(EDDMethod As String, EDDUnit As String)
        Dim line As String
        Dim strArr() As String 'array to hold elements on the current line of input

        Dim reader As Object
        reader = New StreamReader("\\mdrnd\AS-Global\Special_Access\EAC\Data\eTrainLite\Methods\eTrainLiteMethods.txt") 'Text file on file share containing LIMS method names 

        line = reader.ReadLine() 'Read next line

        While Not line = ""
            strArr = line.Split("|")

            If strArr.Length = 3 Then 'If strArr has 3 elements, compare method names and unit
                If EDDMethod = strArr(1) And EDDUnit = strArr(2) Then
                    Return strArr(0)
                End If
            Else 'If strArr has 2 elements, only compare method names
                If EDDMethod = strArr(1) Then
                    Return strArr(0)
                End If
            End If

            line = reader.ReadLine() 'Read next line
        End While

        Return "" 'If no matching EDD Method name is found, return null

    End Function


    'Match the respective lab analysis name variation (analysisName) with the analysis name used in LIMS
    Function matchAnalysisName(analysisName As String) As String

        '*************** NEW MATCH USING METHOD FILE ON FILESHARE ***************

        Dim filePath As String
        Dim sr As StreamReader
        Dim line As String
        Dim tempArr() As String
        Dim count As Integer
        Dim aName As String

        'filePath = "S:/TempMethods.txt" 'Path/link to method file on fileshare goes here
        If GlobalVariables.eTrain.Server = "ROH" Then
            filePath = "\\Usfrpsdowa120\nwa\FAS\QADATA\DeerPark\eTrain\Methods.txt"
        Else
            filePath = "\\Seasv02\analyticalsv\Data\Analytical Natural Work Teams\Lab ENV\Automated Data Transfer\Methods.txt"
        End If

        sr = New StreamReader(filePath)
        aName = ""

        Do Until sr.EndOfStream()
            line = sr.ReadLine()
            tempArr = line.Split("|")
            For count = 1 To tempArr.Length - 1
                tempArr(count) = tempArr(count).Trim()
                If tempArr(count) = analysisName Then
                    aName = tempArr(0)
                    Exit Do
                End If
            Next
        Loop

        matchAnalysisName = aName
        sr.Close()


        '**************************** MATCH USING CONNECTION STRING *************************************

        ''Declare local variables
        'Dim sConn As String
        'Dim sSQL As String
        'Dim objConn
        'Dim odAdapter
        'Dim aCount As Integer
        'Dim rCount As Integer
        'Dim dtUnits As DataTable
        'Dim dvUnits As DataView
        'Dim tempArr As String()

        'If GlobalVariables.eTrain.Server = "SEADRIFT" Then
        '    sConn = "DRIVER={Microsoft ODBC for Oracle};UID=FGLLIMS_CHEMS;PWD=lg#ch3ms;SERVER=PPT108P.nam.dow.com;"
        'ElseIf GlobalVariables.eTrain.Server = "ROH" Then
        '    sConn = "DRIVER={Microsoft ODBC for Oracle};UID=FGLLIMS_ROHNA;PWD=lg#R0hna;SERVER=PPT105P.nam.dow.com;"
        'End If


        ''sSQL = "SELECT PHRASE.PHRASE_ID, PHRASE.PHRASE_TEXT, FROM FGLNWA_CHEMS.PHRASE PHRASE WHERE (PHRASE.PHRASE_TYPE='ETRAINLITE')"

        'dtUnits = New DataTable()

        ''Connect and fill dtLimits for later use
        'Try

        '    Dim queryString As String = "SELECT PHRASE.PHRASE_ID, PHRASE.PHRASE_TEXT, FROM FGLNWA_CHEMS.PHRASE PHRASE WHERE (PHRASE.PHRASE_TYPE='ETRAINLITE')"
        '    Dim command As New OdbcCommand(queryString)

        '    Using connection As New OdbcConnection(sConn)
        '        command.Connection = connection
        '        connection.Open()
        '        odAdapter = New OdbcDataAdapter(queryString, sConn)
        '        odAdapter.Fill(dtUnits)
        '        command.ExecuteNonQuery()
        '        ' The connection is automatically closed at 
        '        ' the end of the Using block.
        '    End Using


        '    'objConn = New OdbcConnection(sConn)
        '    'objConn.Open()
        '    'odAdapter = New OdbcDataAdapter(sSQL, sConn)
        '    'odAdapter.Fill(dtUnits)
        '    'objConn.Close()

        'Catch ex As Exception
        '    MsgBox("Error connecting to LIMS!" & vbCrLf &
        '           "Sub Procedure: matchAnalysisName()" & vbCrLf &
        '           "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
        '    Return False
        'End Try

        ''Get datatable into view and sort then move back to table
        'dvUnits = New DataView(dtUnits)
        'If GlobalVariables.eTrain.Location = "MIDLAND" Then
        '    dvUnits.Sort = "ANALYSIS ASC"
        'ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
        '    dvUnits.Sort = "ANALYSIS ASC"
        'End If

        'dtUnits = dvUnits.ToTable

        'rCount = 0

        'Do Until rCount = dtUnits.Rows.Count - 1

        '    tempArr = dtUnits.Rows(rCount)(2).ToString().Split("|") 'In case of multiple naming conventions for analysis from same lab

        '    For aCount = 0 To tempArr.Length - 1

        '        If analysisName = tempArr(aCount) Then
        '            matchAnalysisName = dtUnits.Rows(rCount)(1)
        '            Exit Do
        '        End If

        '    Next

        '    rCount = rCount + 1
        'Loop

        'matchAnalysisName = ""

    End Function


    'Match the respective lab analysis name variation (analysisName) with the analysis name used in LIMS
    Function matchComponentName(compName As String) As String

        '*************** NEW MATCH USING METHOD FILE ON FILESHARE ***************

        Dim filePath As String
        Dim sr As StreamReader
        Dim line As String
        Dim tempArr() As String
        Dim count As Integer
        Dim aName As String

        Try

            'filePath = "S:/TempMethods.txt" 'Path/link to method file on fileshare goes here
            If GlobalVariables.eTrain.Server = "ROH" Then
                filePath = "\\Usfrpsdowa120\nwa\FAS\QADATA\DeerPark\eTrain\Components.txt"
            Else
                filePath = "\\Seasv02\analyticalsv\Data\Analytical Natural Work Teams\Lab ENV\Automated Data Transfer\Components.txt"
            End If

            sr = New StreamReader(filePath)
            aName = ""

            Do Until sr.EndOfStream()
                line = sr.ReadLine()
                tempArr = line.Split("|")
                For count = 1 To tempArr.Length - 1
                    tempArr(count) = tempArr(count).Trim()
                    If tempArr(count) = compName Then
                        aName = tempArr(0)
                        Exit Do
                    End If
                Next
            Loop

            matchComponentName = aName
            sr.Close()
        Catch ex As Exception
            MsgBox("Error: Something went wrong when attempting to read from Transfer Cross Check file!" & vbCrLf & ex.StackTrace)
        End Try

    End Function

    Function calculateTotalTEQ(limsID As String) As Double
        Dim aCompound As Compound
        Dim aSample As Sample
        Dim dblTeqScore = 0.0

        For Each aSample In GlobalVariables.SampleList
            If aSample.LimsID = limsID Then
                For Each aCompound In aSample.CompoundList
                    For i As Integer = 0 To GlobalVariables.befAndTefScores.Count - 1
                        If GlobalVariables.befAndTefScores(i).Contains(aCompound.EDDChemicalName) And Convert.ToDouble(aCompound.EDDResultValue) >= Convert.ToDouble(GlobalVariables.befAndTefScores(i).Item(2)) Then
                            dblTeqScore += Convert.ToDouble(aCompound.EDDResultValue) * Convert.ToDouble(GlobalVariables.befAndTefScores(i).Item(3)) * Convert.ToDouble(GlobalVariables.befAndTefScores(i).Item(4))
                        End If
                    Next
                Next
            Else
                Continue For
            End If
        Next
        Return dblTeqScore
    End Function
End Class
