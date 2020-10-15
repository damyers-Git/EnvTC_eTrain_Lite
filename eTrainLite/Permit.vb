Imports System.IO
Imports System.Data.Odbc
Imports Syncfusion.XlsIO

Public Class Permit

    Private strName As String
    Private strType As String
    Private dtCreatedDate As Date
    Public Property ProjectList As New ArrayList

    Public Property Name() As String
        Get
            Return strName
        End Get
        Set(ByVal value As String)
            strName = value
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
    Public Property CreatedDate() As Date
        Get
            Return dtCreatedDate
        End Get
        Set(ByVal value As Date)
            dtCreatedDate = value
        End Set
    End Property

    Function LoadNonCompliance() As Boolean
        Dim exEngine As New ExcelEngine
        Dim exApp As IApplication
        Dim workbook As IWorkbook
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aInstrument As New mInstrument
        Dim aCompound As mCompound
        Dim aSurrogate As mSurrogate
        Dim worksheet As IWorksheet
        Dim i As Integer

        If GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                Try
                    exApp = exEngine.Excel
                    workbook = exApp.Workbooks.Open("\\usfrpndowd002\EHS_Analytical\Unapproved\Responsible_Care_EHS_Analytical\e-Train_Files\FreeportChromNonComplianceInfo.xlsx")
                    aPermit = New Permit
                    aPermit.Name = "NonCompliance"
                    'Find applicable sheets
                    For Each worksheet In workbook.Worksheets
                        aProject = New Project
                        aProject.Name = worksheet.Name
                        aInstrument = New mInstrument
                        'Start at row 5
                        i = 0

                        Do Until worksheet.Range("B" & CStr(5 + i)).Value = ""
                            If worksheet.Range("A" & CStr(5 + i)).Value <> "" Then
                                If aInstrument.Name <> "" Then
                                    aProject.mInstrumentList.Add(aInstrument)
                                End If
                                aInstrument = New mInstrument
                                aInstrument.Name = worksheet.Range("A" & CStr(5 + i)).Value
                            End If
                            If worksheet.Range("D" & CStr(5 + i)).Value = "Compound" Then
                                aCompound = New mCompound
                                aCompound.Name = worksheet.Range("B" & CStr(5 + i)).Value
                                aCompound.CAS = worksheet.Range("C" & CStr(5 + i)).Value
                                If worksheet.Range("E" & CStr(5 + i)).Value <> "" Then
                                    aCompound.Conc = worksheet.Range("E" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("E" & CStr(5 + i)).Value, " ") - 1)
                                    aCompound.Units = worksheet.Range("E" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("E" & CStr(5 + i)).Value, " "), worksheet.Range("E" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("E" & CStr(5 + i)).Value, " "))
                                End If
                                If worksheet.Range("F" & CStr(5 + i)).Value <> "" Then
                                    aCompound.SurLLim = worksheet.Range("F" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("F" & CStr(5 + i)).Value, "-") - 1)
                                    aCompound.SurULim = worksheet.Range("F" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("F" & CStr(5 + i)).Value, "-"), worksheet.Range("F" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("F" & CStr(5 + i)).Value, "-"))
                                End If
                                If worksheet.Range("G" & CStr(5 + i)).Value <> "" Then
                                    aCompound.MSLLim = worksheet.Range("G" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("G" & CStr(5 + i)).Value, "-") - 1)
                                    aCompound.MSULim = worksheet.Range("G" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("G" & CStr(5 + i)).Value, "-"), worksheet.Range("G" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("G" & CStr(5 + i)).Value, "-"))
                                End If
                                If worksheet.Range("H" & CStr(5 + i)).Value <> "" Then
                                    aCompound.LCSLLim = worksheet.Range("H" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("H" & CStr(5 + i)).Value, "-") - 1)
                                    aCompound.LCSULim = worksheet.Range("H" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("H" & CStr(5 + i)).Value, "-"), worksheet.Range("H" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("H" & CStr(5 + i)).Value, "-"))
                                End If
                                aCompound.MDL = worksheet.Range("I" & CStr(5 + i)).Value
                                aCompound.PQL = worksheet.Range("J" & CStr(5 + i)).Value
                                aCompound.RL = worksheet.Range("K" & CStr(5 + i)).Value
                                aInstrument.mCompoundList.Add(aCompound)
                            ElseIf worksheet.Range("D" & CStr(5 + i)).Value = "Surrogate" Then
                                aSurrogate = New mSurrogate
                                aSurrogate.Name = worksheet.Range("B" & CStr(5 + i)).Value
                                aSurrogate.CAS = worksheet.Range("C" & CStr(5 + i)).Value
                                If worksheet.Range("E" & CStr(5 + i)).Value <> "" Then
                                    aSurrogate.Conc = worksheet.Range("E" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("E" & CStr(5 + i)).Value, " ") - 1)
                                    aSurrogate.Units = worksheet.Range("E" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("E" & CStr(5 + i)).Value, " "), worksheet.Range("E" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("E" & CStr(5 + i)).Value, " "))
                                End If
                                If worksheet.Range("F" & CStr(5 + i)).Value <> "" Then
                                    aSurrogate.RecLLim = worksheet.Range("F" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("F" & CStr(5 + i)).Value, "-") - 1)
                                    aSurrogate.RecULim = worksheet.Range("F" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("F" & CStr(5 + i)).Value, "-"), worksheet.Range("F" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("F" & CStr(5 + i)).Value, "-"))
                                End If
                                If worksheet.Range("G" & CStr(5 + i)).Value <> "" Then
                                    aSurrogate.MSLLim = worksheet.Range("G" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("G" & CStr(5 + i)).Value, "-") - 1)
                                    aSurrogate.MSULim = worksheet.Range("G" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("G" & CStr(5 + i)).Value, "-"), worksheet.Range("G" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("G" & CStr(5 + i)).Value, "-"))
                                End If
                                If worksheet.Range("H" & CStr(5 + i)).Value <> "" Then
                                    aSurrogate.LCSLLim = worksheet.Range("H" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("H" & CStr(5 + i)).Value, "-") - 1)
                                    aSurrogate.LCSULim = worksheet.Range("H" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("H" & CStr(5 + i)).Value, "-"), worksheet.Range("H" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("H" & CStr(5 + i)).Value, "-"))
                                End If
                                aSurrogate.MDL = worksheet.Range("I" & CStr(5 + i)).Value
                                aSurrogate.PQL = worksheet.Range("J" & CStr(5 + i)).Value
                                aSurrogate.RL = worksheet.Range("K" & CStr(5 + i)).Value
                                aInstrument.mSurrogateList.Add(aSurrogate)
                            End If
                            i = i + 1
                        Loop
                        aProject.mInstrumentList.Add(aInstrument)
                        'aInstrument = Nothing
                        aPermit.ProjectList.Add(aProject)

                    Next
                    workbook.Close()
                    exEngine.Dispose()
                    GlobalVariables.PermitList.Add(aPermit)
                    Return True
                Catch ex As Exception
                    MsgBox("Error import Non Compliance information!" & vbCrLf &
                         "Sub Procedure: LoadNonCompliance()" & vbCrLf &
                         "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                    Return False
                End Try
                Return False
            End If
        ElseIf GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                Try
                    exApp = exEngine.Excel
                    workbook = exApp.Workbooks.Open("\\Helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Midland\Chrom\MidlandChromNonComplianceInfo.xlsx")
                    aPermit = New Permit
                    aPermit.Name = "NonCompliance"
                    'Find applicable sheets
                    For Each worksheet In workbook.Worksheets
                        aProject = New Project
                        aProject.Name = worksheet.Name
                        aInstrument = New mInstrument
                        'Start at row 5
                        i = 0
                        Do Until worksheet.Range("B" & CStr(5 + i)).Value = ""
                            If worksheet.Range("A" & CStr(5 + i)).Value <> "" Then
                                If aInstrument.Name <> "" Then
                                    aProject.mInstrumentList.Add(aInstrument)
                                End If
                                aInstrument = New mInstrument
                                aInstrument.Name = worksheet.Range("A" & CStr(5 + i)).Value
                            End If
                            If InStr(worksheet.Range("B" & CStr(5 + i)).Value, "(SS)", CompareMethod.Binary) Then
                                aSurrogate = New mSurrogate
                                aSurrogate.Name = worksheet.Range("B" & CStr(5 + i)).Value
                                aSurrogate.CAS = worksheet.Range("C" & CStr(5 + i)).Value
                                If worksheet.Range("D" & CStr(5 + i)).Value <> "" Then
                                    aSurrogate.Conc = worksheet.Range("D" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("D" & CStr(5 + i)).Value, " ") - 1)
                                    aSurrogate.Units = worksheet.Range("D" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("D" & CStr(5 + i)).Value, " "), worksheet.Range("D" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("D" & CStr(5 + i)).Value, " "))
                                End If
                                If worksheet.Range("E" & CStr(5 + i)).Value <> "" Then
                                    aSurrogate.RecLLim = worksheet.Range("E" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("E" & CStr(5 + i)).Value, "-") - 1)
                                    aSurrogate.RecULim = worksheet.Range("E" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("E" & CStr(5 + i)).Value, "-"), worksheet.Range("E" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("E" & CStr(5 + i)).Value, "-"))
                                End If
                                If worksheet.Range("F" & CStr(5 + i)).Value <> "" Then
                                    aSurrogate.MSLLim = worksheet.Range("F" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("F" & CStr(5 + i)).Value, "-") - 1)
                                    aSurrogate.MSULim = worksheet.Range("F" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("F" & CStr(5 + i)).Value, "-"), worksheet.Range("F" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("F" & CStr(5 + i)).Value, "-"))
                                End If
                                If worksheet.Range("G" & CStr(5 + i)).Value <> "" Then
                                    aSurrogate.LCSLLim = worksheet.Range("G" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("G" & CStr(5 + i)).Value, "-") - 1)
                                    aSurrogate.LCSULim = worksheet.Range("G" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("G" & CStr(5 + i)).Value, "-"), worksheet.Range("G" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("G" & CStr(5 + i)).Value, "-"))
                                End If
                                aSurrogate.MDL = worksheet.Range("H" & CStr(5 + i)).Value
                                aSurrogate.PQL = worksheet.Range("I" & CStr(5 + i)).Value
                                aSurrogate.RL = worksheet.Range("J" & CStr(5 + i)).Value
                                aInstrument.mSurrogateList.Add(aSurrogate)
                            Else
                                aCompound = New mCompound
                                aCompound.Name = worksheet.Range("B" & CStr(5 + i)).Value
                                aCompound.CAS = worksheet.Range("C" & CStr(5 + i)).Value
                                If worksheet.Range("D" & CStr(5 + i)).Value <> "" Then
                                    aCompound.Conc = worksheet.Range("D" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("D" & CStr(5 + i)).Value, " ") - 1)
                                    aCompound.Units = worksheet.Range("D" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("D" & CStr(5 + i)).Value, " "), worksheet.Range("D" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("D" & CStr(5 + i)).Value, " "))
                                End If
                                If worksheet.Range("E" & CStr(5 + i)).Value <> "" Then
                                    aCompound.SurLLim = worksheet.Range("E" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("E" & CStr(5 + i)).Value, "-") - 1)
                                    aCompound.SurULim = worksheet.Range("E" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("E" & CStr(5 + i)).Value, "-"), worksheet.Range("E" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("E" & CStr(5 + i)).Value, "-"))
                                End If
                                If worksheet.Range("F" & CStr(5 + i)).Value <> "" Then
                                    aCompound.MSLLim = worksheet.Range("F" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("F" & CStr(5 + i)).Value, "-") - 1)
                                    aCompound.MSULim = worksheet.Range("F" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("F" & CStr(5 + i)).Value, "-"), worksheet.Range("F" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("F" & CStr(5 + i)).Value, "-"))
                                End If
                                If worksheet.Range("G" & CStr(5 + i)).Value <> "" Then
                                    aCompound.LCSLLim = worksheet.Range("G" & CStr(5 + i)).Value.Substring(0, InStr(worksheet.Range("G" & CStr(5 + i)).Value, "-") - 1)
                                    aCompound.LCSULim = worksheet.Range("G" & CStr(5 + i)).Value.Substring(InStr(worksheet.Range("G" & CStr(5 + i)).Value, "-"), worksheet.Range("G" & CStr(5 + i)).Value.Length - InStr(worksheet.Range("G" & CStr(5 + i)).Value, "-"))
                                End If
                                aCompound.MDL = worksheet.Range("H" & CStr(5 + i)).Value
                                aCompound.PQL = worksheet.Range("I" & CStr(5 + i)).Value
                                aCompound.RL = worksheet.Range("J" & CStr(5 + i)).Value
                                aInstrument.mCompoundList.Add(aCompound)
                            End If
                            i = i + 1
                        Loop
                        aProject.mInstrumentList.Add(aInstrument)
                        'aInstrument = Nothing
                        aPermit.ProjectList.Add(aProject)

                    Next
                    workbook.Close()
                    exEngine.Dispose()
                    GlobalVariables.PermitList.Add(aPermit)
                    Return True
                Catch ex As Exception
                    MsgBox("Error import Non Compliance information!" & vbCrLf &
                         "Sub Procedure: LoadNonCompliance()" & vbCrLf &
                         "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                    Return False
                End Try
                Return False
            End If
        End If



    End Function


    Function LoadLimsUnits() As Boolean
        Dim sConn As String
        Dim sSQL As String
        Dim dtUnits As New DataTable
        Dim dvUnits As DataView
        Dim aPermit As Permit
        Dim aProject As Project
        Dim rCount As Integer
        Dim objConn As OdbcConnection
        Dim odAdapter As OdbcDataAdapter


        'Connection based on location
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            sConn = "DRIVER={Microsoft ODBC for Oracle};UID=FGLLIMS_EnvMD;PWD=lg#En3#;SERVER=PPT87P.nam.dow.com;"
            'SQL statement
            sSQL = "SELECT COMPONENT_VIEW.ANALYSIS, COMPONENT_VIEW.UNITS FROM LIMS_ENVMD.COMPONENT_VIEW " &
                   "WHERE COMPONENT_VIEW.ANALYSIS = 'VOC' OR COMPONENT_VIEW.ANALYSIS = 'EOA'"
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            sConn = "DRIVER={Microsoft ODBC for Oracle};UID=FGLLIMS_ENVTX;PWD=lg#Tx1#;SERVER=PPT87P.nam.dow.com;"
            'sConn = "DRIVER={Microsoft ODBC for Oracle};UID=FGLLIMS_ENVTX;PWD=lg#Tx1#;SERVER=PPT85P.nam.dow.com;"
            'Units
            sSQL = "SELECT COMPONENT_VIEW.ANALYSIS, COMPONENT_VIEW.UNITS FROM LIMS_ENVTX.COMPONENT_VIEW " ' _
            '"WHERE COMPONENT_VIEW.ANALYSIS = 'TPH_DUP' OR COMPONENT_VIEW.ANALYSIS = 'M624H_DUP' OR COMPONENT_VIEW.ANALYSIS = 'HS_FID_DUP'"
        End If

        'Connect and fill dtLimits for later use test 
        Try
            objConn = New OdbcConnection(sConn)
            objConn.Open()
            odAdapter = New OdbcDataAdapter(sSQL, sConn)
            odAdapter.Fill(dtUnits)
            objConn.Close()

        Catch ex As Exception
            MsgBox("Error connecting to LIMS!" & vbCrLf &
                   "Sub Procedure: LoadLimsUnits()" & vbCrLf &
                   "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        'Get datatable into view and sort
        dvUnits = New DataView(dtUnits)
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            dvUnits.Sort = "ANALYSIS ASC"
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            dvUnits.Sort = "ANALYSIS ASC"
        End If
        dtUnits = dvUnits.ToTable

        rCount = 0
        For Each aPermit In GlobalVariables.PermitList
            Do Until rCount = dtUnits.Rows.Count - 1
                For Each aProject In aPermit.ProjectList
                    If aProject.Name = dtUnits.Rows(rCount)(0).ToString() Then
                        aProject.LimsUnits = dtUnits.Rows(rCount)(1).ToString()
                        Exit For
                    End If
                Next
                rCount = rCount + 1
            Loop
        Next
        Return True

    End Function

    'This generates an a list of limits for each compound
    Function LoadLimsLimit() As Boolean
        Dim sConn As String
        Dim sSQL As String
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aInstrument As mInstrument
        Dim aCompound As mCompound
        Dim rCount As Integer
        Dim dtLimits As New DataTable
        Dim dvLimits As DataView
        Dim dtUnits As New DataTable
        Dim objConn As OdbcConnection
        Dim odAdapter As OdbcDataAdapter

        'Connection based on location
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            sConn = "DRIVER={Microsoft ODBC for Oracle};UID=FGLLIMS_EnvMD;PWD=lg#En3#;SERVER=PPT87P.nam.dow.com;"
            'SQL statement
            sSQL = "SELECT DOW_COMP_LIMIT_ENTRY.ANALYSIS_ID, DOW_COMP_LIMIT_ENTRY.INSTRUMENT, DOW_COMP_LIMIT_ENTRY.COMPONENT_NAME, " &
                "DOW_COMP_LIMIT_ENTRY.MDL, DOW_COMP_LIMIT_ENTRY.RL, DOW_COMP_LIMIT_ENTRY.PQL FROM LIMS_ENVMD.DOW_COMP_LIMIT_ENTRY DOW_COMP_LIMIT_ENTRY " &
            "WHERE DOW_COMP_LIMIT_ENTRY.ANALYSIS_ID = 'VOC' AND DOW_COMP_LIMIT_ENTRY.ANALYSIS_VERSION = '         4' OR DOW_COMP_LIMIT_ENTRY.ANALYSIS_ID = 'EOA' AND DOW_COMP_LIMIT_ENTRY.ANALYSIS_VERSION = '         3'"
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            sConn = "DRIVER={Microsoft ODBC for Oracle};UID=FGLLIMS_ENVTX;PWD=lg#Tx1#;SERVER=PPT85P.nam.dow.com;"
            'sConn = "DRIVER={Microsoft ODBC for Oracle};UID=FGLLIMS_ENVTX;PWD=lg#Tx1#;SERVER=PPT87P.nam.dow.com;"
            'Limits
            sSQL = "SELECT DOW_COMP_LIMIT_ENTRY.ANALYSIS_ID, DOW_COMP_LIMIT_ENTRY.INSTRUMENT, DOW_COMP_LIMIT_ENTRY.COMPONENT_NAME, " &
                "DOW_COMP_LIMIT_ENTRY.MDL, DOW_COMP_LIMIT_ENTRY.RL, DOW_COMP_LIMIT_ENTRY.PQL FROM LIMS_ENVTX.DOW_COMP_LIMIT_ENTRY " '& _
            ' "WHERE DOW_COMP_LIMIT_ENTRY.ANALYSIS_ID = 'TPH_DUP' OR DOW_COMP_LIMIT_ENTRY.ANALYSIS_ID = 'M624H_DUP' OR DOW_COMP_LIMIT_ENTRY.ANALYSIS_ID = 'HS_FID_DUP'"
        End If

        'Connect and fill dtLimits for later use
        Try
            objConn = New OdbcConnection(sConn)
            objConn.Open()
            odAdapter = New OdbcDataAdapter(sSQL, sConn)
            odAdapter.Fill(dtLimits)
            objConn.Close()
        Catch ex As Exception
            MsgBox("Error connecting to LIMS!" & vbCrLf &
                   "Sub Procedure: LoadLimsLimit()" & vbCrLf &
                   "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        'Load into Permit List
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            'Check if Permit by LIMS is already used
            For Each aPermit In GlobalVariables.PermitList
                If aPermit.Name = "LIMS" Then
                    Return True
                End If
            Next

            'Get datatable into view and sort
            dvLimits = New DataView(dtLimits)
            dvLimits.Sort = "ANALYSIS_ID ASC, INSTRUMENT ASC"
            dtLimits = dvLimits.ToTable
            rCount = 0
            aPermit = New Permit
            aPermit.Name = "LIMS"
            Do Until rCount = dtLimits.Rows.Count - 1
                'Set starting AnalysisLab/Project
                aProject = New Project
                aProject.Name = dtLimits.Rows(rCount)(0).ToString()
                Do Until dtLimits.Rows(rCount)(0).ToString() <> aProject.Name Or rCount = dtLimits.Rows.Count - 1
                    aInstrument = New mInstrument
                    aInstrument.Name = dtLimits.Rows(rCount)(1).ToString()
                    Do Until dtLimits.Rows(rCount)(1).ToString() <> aInstrument.Name Or dtLimits.Rows(rCount)(0).ToString() <> aProject.Name Or rCount = dtLimits.Rows.Count - 1
                        aCompound = New mCompound
                        aCompound.Name = dtLimits.Rows(rCount)(2).ToString()
                        aCompound.MDL = dtLimits.Rows(rCount)(3).ToString()
                        aCompound.RL = dtLimits.Rows(rCount)(4).ToString()
                        aCompound.PQL = dtLimits.Rows(rCount)(5).ToString()
                        aInstrument.mCompoundList.Add(aCompound)
                        rCount = rCount + 1
                    Loop
                    aProject.mInstrumentList.Add(aInstrument)
                Loop
                aPermit.ProjectList.Add(aProject)
            Loop
            GlobalVariables.PermitList.Add(aPermit)
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            'Check if Permit by LIMS is already used
            For Each aPermit In GlobalVariables.PermitList
                If aPermit.Name = "LIMS" Then
                    Return True
                End If
            Next

            'Get datatable into view and sort
            dvLimits = New DataView(dtLimits)
            dvLimits.Sort = "ANALYSIS_ID ASC, INSTRUMENT ASC"
            dtLimits = dvLimits.ToTable
            rCount = 0
            aPermit = New Permit
            aPermit.Name = "LIMS"
            Do Until rCount = dtLimits.Rows.Count - 1
                'Set starting AnalysisLab/Project
                aProject = New Project
                aProject.Name = dtLimits.Rows(rCount)(0).ToString()
                Do Until dtLimits.Rows(rCount)(0).ToString() <> aProject.Name Or rCount = dtLimits.Rows.Count - 1
                    aInstrument = New mInstrument
                    aInstrument.Name = dtLimits.Rows(rCount)(1).ToString()
                    Do Until dtLimits.Rows(rCount)(1).ToString() <> aInstrument.Name Or dtLimits.Rows(rCount)(0).ToString() <> aProject.Name Or rCount = dtLimits.Rows.Count - 1
                        aCompound = New mCompound
                        aCompound.Name = dtLimits.Rows(rCount)(2).ToString()
                        aCompound.MDL = dtLimits.Rows(rCount)(3).ToString()
                        aCompound.RL = dtLimits.Rows(rCount)(4).ToString()
                        aCompound.PQL = dtLimits.Rows(rCount)(5).ToString()
                        aInstrument.mCompoundList.Add(aCompound)
                        rCount = rCount + 1
                    Loop
                    aProject.mInstrumentList.Add(aInstrument)
                Loop
                aPermit.ProjectList.Add(aProject)
            Loop
            GlobalVariables.PermitList.Add(aPermit)
        End If

        'Load in units from lims
        If GlobalVariables.Permit.LoadLimsUnits() Then
            Return True
        Else
            Return False
        End If

    End Function
    Function LoadPermitNames() As Boolean
        Dim strFileNames() As String
        Dim aPermit As Permit

        'Check if permits are already populated then clear if they are, assuming team change or updated permit from file
        If GlobalVariables.PermitList.Count <> 0 Then
            GlobalVariables.PermitList.Clear()
        End If

        Try
            If GlobalVariables.eTrain.Location = "MIDLAND" Then
                If GlobalVariables.eTrain.Team = "CHROM" Then
                    'Gets listing of file names in folder and uses them as permit names, adds them to Permitlist
                    strFileNames = Directory.GetFiles(GlobalVariables.eTrain.DataFileFP & "Midland\Chrom\Projects_Methods\")
                    For Each f In strFileNames
                        aPermit = New Permit
                        aPermit.Name = Path.GetFileNameWithoutExtension(f)
                        GlobalVariables.PermitList.Add(aPermit)
                    Next
                End If
                Return True
            ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
                If GlobalVariables.eTrain.Team = "CHROM" Then
                    'Gets listing of file names in folder and uses them as permit names, adds them to Permitlist
                    strFileNames = Directory.GetFiles(GlobalVariables.eTrain.DataFileFP & "Freeport\Chrom\Projects_Methods\")
                    For Each f In strFileNames
                        aPermit = New Permit
                        aPermit.Name = Path.GetFileNameWithoutExtension(f)
                        GlobalVariables.PermitList.Add(aPermit)
                    Next
                End If
                Return True
            End If

        Catch ex As Exception
            MsgBox("Error getting Permit names!" & vbCrLf &
                   "Sub Procedure: LoadPermitNames()" & vbCrLf &
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try
    End Function

    Sub LoadPermit(ByVal strName As String)
        Dim strLine As String
        Dim arrSplit() As String
        Dim aPermit As Permit
        Dim aProject As Project
        Dim aInstrument As mInstrument
        Dim aAnalyte As mCompound

        Try
            If GlobalVariables.eTrain.Location = "MIDLAND" Then
                If GlobalVariables.eTrain.Team = "CHROM" Then
                    'Figure out which Permit is being loaded
                    For Each aPermit In GlobalVariables.PermitList
                        If strName = aPermit.Name Then
                            'check to see if already loaded
                            If aPermit.ProjectList.Count > 0 Then
                                Exit Sub
                            End If
                            Dim sr As StreamReader = New StreamReader(GlobalVariables.eTrain.DataFileFP & "Midland\Chrom\Projects_Methods\" & strName & ".et2")
                            strLine = sr.ReadLine
                            'Name check
                            If aPermit.Name = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))) Then
                                strLine = sr.ReadLine
                                'Date
                                aPermit.CreatedDate = CDate(Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))))
                                strLine = sr.ReadLine
                                'Project
                                Do Until sr.EndOfStream
                                    aProject = New Project
                                    aProject.Name = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    strLine = sr.ReadLine
                                    'Reviewed
                                    aProject.Reviewed = CBool(Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))))
                                    'Date reviewed
                                    strLine = sr.ReadLine
                                    aProject.ReviewedDate = CDate(Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))))
                                    'Default RL
                                    strLine = sr.ReadLine
                                    aProject.DefRL = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    'Default MDL
                                    strLine = sr.ReadLine
                                    aProject.DefMDL = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    'Default PQL
                                    strLine = sr.ReadLine
                                    aProject.DefPQL = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    strLine = sr.ReadLine
                                    'Instrument
                                    Do Until InStr(strLine, "Project:", CompareMethod.Binary) Or sr.EndOfStream
                                        aInstrument = New mInstrument
                                        aInstrument.Name = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                        strLine = sr.ReadLine
                                        Do Until InStr(strLine, "Project:", CompareMethod.Binary) Or InStr(strLine, "Instrument:", CompareMethod.Binary)
                                            'Analytes - using mCompound class
                                            If strLine = "" Then
                                                Exit Do
                                            End If
                                            arrSplit = strLine.Split("|")
                                            aAnalyte = New mCompound
                                            aAnalyte.Name = arrSplit(0)
                                            aAnalyte.CAS = arrSplit(1)
                                            aAnalyte.RL = arrSplit(2)
                                            aAnalyte.MDL = arrSplit(3)
                                            aAnalyte.PQL = arrSplit(4)
                                            aAnalyte.RecLLim = arrSplit(5)
                                            aAnalyte.RecULim = arrSplit(6)
                                            aInstrument.mCompoundList.Add(aAnalyte)
                                            strLine = sr.ReadLine
                                        Loop
                                        aProject.mInstrumentList.Add(aInstrument)
                                    Loop
                                    aPermit.ProjectList.Add(aProject)
                                Loop
                            End If

                            sr.Close()
                            sr.Dispose()
                        End If
                    Next
                End If
            ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
                If GlobalVariables.eTrain.Team = "CHROM" Then
                    'Figure out which Permit is being loaded
                    For Each aPermit In GlobalVariables.PermitList
                        If strName = aPermit.Name Then
                            'check to see if already loaded
                            If aPermit.ProjectList.Count > 0 Then
                                Exit Sub
                            End If
                            Dim sr As StreamReader = New StreamReader(GlobalVariables.eTrain.DataFileFP & "Freeport\Chrom\Projects_Methods\" & strName & ".et2")
                            strLine = sr.ReadLine
                            'Name check
                            If aPermit.Name = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))) Then
                                strLine = sr.ReadLine
                                'Date
                                aPermit.CreatedDate = CDate(Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))))
                                strLine = sr.ReadLine
                                'Project
                                Do Until sr.EndOfStream
                                    aProject = New Project
                                    aProject.Name = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    strLine = sr.ReadLine
                                    'Reviewed
                                    aProject.Reviewed = CBool(Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))))
                                    'Date reviewed
                                    strLine = sr.ReadLine
                                    aProject.ReviewedDate = CDate(Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))))
                                    'Default RL
                                    strLine = sr.ReadLine
                                    aProject.DefRL = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    'Default MDL
                                    strLine = sr.ReadLine
                                    aProject.DefMDL = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    'Default PQL
                                    strLine = sr.ReadLine
                                    aProject.DefPQL = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    strLine = sr.ReadLine
                                    'Instrument
                                    Do Until InStr(strLine, "Project:", CompareMethod.Binary) Or sr.EndOfStream
                                        aInstrument = New mInstrument
                                        aInstrument.Name = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                        strLine = sr.ReadLine
                                        Do Until InStr(strLine, "Project:", CompareMethod.Binary) Or InStr(strLine, "Instrument:", CompareMethod.Binary)
                                            'Analytes - using mCompound class
                                            If strLine = "" Then
                                                Exit Do
                                            End If
                                            arrSplit = strLine.Split("|")
                                            aAnalyte = New mCompound
                                            aAnalyte.Name = arrSplit(0)
                                            aAnalyte.CAS = arrSplit(1)
                                            aAnalyte.RL = arrSplit(2)
                                            aAnalyte.MDL = arrSplit(3)
                                            aAnalyte.PQL = arrSplit(4)
                                            aAnalyte.RecLLim = arrSplit(5)
                                            aAnalyte.RecULim = arrSplit(6)
                                            aInstrument.mCompoundList.Add(aAnalyte)
                                            strLine = sr.ReadLine
                                        Loop
                                        aProject.mInstrumentList.Add(aInstrument)
                                    Loop
                                    aPermit.ProjectList.Add(aProject)
                                Loop
                            End If

                            sr.Close()
                            sr.Dispose()
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox("Error reading Permit file!" & vbCrLf &
                   "Sub Procedure: LoadPermit()" & vbCrLf &
                "Line: " & strLine & vbCrLf &
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    'True return means save successful, false return save cancelled/unsuccessful nothing changed
    Function SavePermit(ByVal aPermit As Permit) As Boolean

        Dim strFileLoc As String

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                Try
                    strFileLoc = "\\Helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Midland\Chrom\Projects_Methods\" & aPermit.Name & ".et2"
                    If Not File.Exists(strFileLoc) Then
                        'New Method
                        If GlobalVariables.Permit.WritePermit(aPermit, strFileLoc) Then
                            Return True
                        Else
                            Return False
                        End If
                    Else
                        'Method already exists, update
                        If MsgBox("A Permit already exists with this name, do you intend to update it?", MsgBoxStyle.YesNo, "eTrain 2.0") = MsgBoxResult.Yes Then
                            If GlobalVariables.Permit.WritePermit(aPermit, strFileLoc) Then
                                Return True
                            Else
                                Return False
                            End If
                        End If
                    End If

                Catch ex As Exception
                    MsgBox("Error Saving Permit File" & vbCrLf &
                     "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                    Return False
                End Try

            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                Try
                    strFileLoc = "\\Helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\Projects_Methods\" & aPermit.Name & ".et2"
                    If Not File.Exists(strFileLoc) Then
                        'New Method
                        If GlobalVariables.Permit.WritePermit(aPermit, strFileLoc) Then
                            Return True
                        Else
                            Return False
                        End If
                    Else
                        'Method already exists, update
                        If MsgBox("A Permit already exists with this name, do you intend to update it?", MsgBoxStyle.YesNo, "eTrain 2.0") = MsgBoxResult.Yes Then
                            If GlobalVariables.Permit.WritePermit(aPermit, strFileLoc) Then
                                Return True
                            Else
                                Return False
                            End If
                        End If
                    End If

                Catch ex As Exception
                    MsgBox("Error Saving Permit File" & vbCrLf &
                           "Sub Procedure: SavePermit()" & vbCrLf &
                     "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                    Return False
                End Try

            End If
        End If
        Return False
    End Function

    Function WritePermit(ByVal aPermit As Permit, ByVal strFileLoc As String)
        Dim sr As StreamWriter
        Dim curDate As Date
        Dim aInstrument As mInstrument
        Dim aProject As Project
        Dim aAnalyte As mCompound

        curDate = DateTime.Now

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                Try
                    'Backup old file if there
                    If File.Exists(strFileLoc) Then
                        File.Copy(strFileLoc, "\\Helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Midland\Chrom\Projects_Methods\Backups\" & aPermit.Name & "_" &
                                  curDate.Month & curDate.Day & curDate.Year & "_" & curDate.Hour & curDate.Minute & ".et2")
                    End If
                    'Begin write
                    sr = New StreamWriter(strFileLoc)
                    sr.WriteLine("Permit Name: " & aPermit.Name)
                    sr.WriteLine("Permit Date: " & CStr(aPermit.CreatedDate))
                    For Each aProject In aPermit.ProjectList
                        sr.WriteLine("Project: " & aProject.Name)
                        sr.WriteLine("Reviewed: " & CStr(aProject.Reviewed))
                        sr.WriteLine("Reviewed Date: " & CStr(aProject.ReviewedDate))
                        sr.WriteLine("Default RL: " & CStr(aProject.DefRL))
                        sr.WriteLine("Default MDL: " & CStr(aProject.DefMDL))
                        sr.WriteLine("Default PQL: " & CStr(aProject.DefPQL))
                        For Each aInstrument In aProject.mInstrumentList
                            sr.WriteLine("Instrument: " & aInstrument.Name)
                            For Each aAnalyte In aInstrument.mCompoundList
                                sr.WriteLine(aAnalyte.Name & "|" & aAnalyte.CAS & "|" & aAnalyte.RL & "|" & aAnalyte.MDL & "|" & aAnalyte.PQL & "|" & aAnalyte.RecLLim & "|" & aAnalyte.RecULim)
                            Next
                        Next
                    Next
                    sr.Close()
                    sr.Dispose()
                    Return True
                Catch ex As Exception
                    MsgBox("Error Writing Permit File" & vbCrLf &
                     "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                    Return False
                End Try
            End If
        ElseIf GlobalVariables.eTrain.Location = "FREEPORT" Then
            If GlobalVariables.eTrain.Team = "CHROM" Then
                Try
                    'Backup old file if there
                    If File.Exists(strFileLoc) Then
                        File.Copy(strFileLoc, "\\Helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Freeport\Chrom\Projects_Methods\Backups\" & aPermit.Name & "_" &
                                  curDate.Month & curDate.Day & curDate.Year & "_" & curDate.Hour & curDate.Minute & ".et2")
                    End If
                    'Begin write
                    sr = New StreamWriter(strFileLoc)
                    sr.WriteLine("Permit Name: " & aPermit.Name)
                    sr.WriteLine("Permit Date: " & CStr(aPermit.CreatedDate))
                    For Each aProject In aPermit.ProjectList
                        sr.WriteLine("Project: " & aProject.Name)
                        sr.WriteLine("Reviewed: " & CStr(aProject.Reviewed))
                        sr.WriteLine("Reviewed Date: " & CStr(aProject.ReviewedDate))
                        sr.WriteLine("Default RL: " & CStr(aProject.DefRL))
                        sr.WriteLine("Default MDL: " & CStr(aProject.DefMDL))
                        sr.WriteLine("Default PQL: " & CStr(aProject.DefPQL))
                        For Each aInstrument In aProject.mInstrumentList
                            sr.WriteLine("Instrument: " & aInstrument.Name)
                            For Each aAnalyte In aInstrument.mCompoundList
                                sr.WriteLine(aAnalyte.Name & "|" & aAnalyte.CAS & "|" & aAnalyte.RL & "|" & aAnalyte.MDL & "|" & aAnalyte.PQL & "|" & aAnalyte.RecLLim & "|" & aAnalyte.RecULim)
                            Next
                        Next
                    Next
                    sr.Close()
                    sr.Dispose()
                    Return True
                Catch ex As Exception
                    MsgBox("Error Writing Permit File" & vbCrLf &
                           "Sub Procedure: WritePermit()" & vbCrLf &
                     "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                    Return False
                End Try
            End If
        End If
    End Function

    Function loadLimsInformation() As Boolean ' Added 6/12/19 WB & WT
        Dim sConn As String
        Dim sSQL As String
        Dim dtLimits As New DataTable
        Dim dvLimits As DataView
        Dim dtUnits As New DataTable
        Dim objConn As OdbcConnection
        Dim odAdapter As OdbcDataAdapter
        ' Connection based on location
        ' Only Midland since that is where all the CLab EDDs will be connecting through for the time being.
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            sConn = "DRIVER={Microsoft ODBC for Oracle};uid=FGLLIMS_ENVMD;Pwd=lg#En3#;SERVER=PPT107P.NAM.DOW.COM;"
            'SQL statement
            sSQL = "SELECT DOW_COMPONENT_CODE.COMPONENT_NAME, DOW_COMPONENT_CODE.CAS_NAME FROM LIMS_ENVMD.DOW_COMPONENT_CODE"
        End If
        'Connect and fill dtLimits for later use
        Try
            objConn = New OdbcConnection(sConn)
            objConn.Open()
            odAdapter = New OdbcDataAdapter(sSQL, sConn)
            odAdapter.Fill(dtLimits)
            objConn.Close()
        Catch ex As Exception
            MsgBox("Error connecting To LIMS!" & vbCrLf &
                   "Sub Procedure: loadLimsInformation()" & vbCrLf &
                   "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try

        Try
            'Load into Permit List
            If GlobalVariables.eTrain.Location = "MIDLAND" Then
                'Get datatable into view and sort
                dvLimits = New DataView(dtLimits)
                dtLimits = dvLimits.ToTable
                ' Supplementing analytes and their CAS number into the dictionary for those not in the DOW_COMPONENT_CODE query.
                ' Adding them here first so the names are correct versus what is in LIMS.
                ' Example: '2,4-D' is actually '2,4-D (Med)' for all the samples that I have seen come through. 
                For Each line As String In IO.File.ReadAllLines("\\mdrnd\AS-Global\Special_Access\EAC\Data\eTrainLite\Name_CAS\casComponentCustom.txt")
                    Dim parts() As String = line.Split("|")
                    If Not GlobalVariables.compNameToCASDic.ContainsKey(parts(0)) Then
                        GlobalVariables.compNameToCASDic.Add(parts(0), parts(1))
                    End If
                Next
                ' Adding the rest of the compounds from LIMS into the dictionary. 
                For Each row As DataRow In dtLimits.Rows
                    If Not GlobalVariables.compNameToCASDic.ContainsKey(row(0).ToString()) Then
                        GlobalVariables.compNameToCASDic.Add(row(0).ToString(), row(1).ToString())
                    End If
                Next
                ' Only importing the method that the data will go into so there will (hypothetically) be half the number imported versus the whole table.
                ' Mostly for the LIMS samples that use a dup sample and then a calcluation is made to populate the reported sample. 
                If GlobalVariables.eTrain.Team <> "NewSample" Then
                    For Each line As String In IO.File.ReadAllLines("\\mdrnd\AS-Global\Special_Access\EAC\Data\eTrainLite\Methods\" & GlobalVariables.Import.Type & "_LIMS.txt")
                        Dim parts() As String = line.Split("|")
                        If Not GlobalVariables.limsAnalysisMethod.ContainsKey(parts(0)) Then
                            GlobalVariables.limsAnalysisMethod.Add(parts(0), parts(1))
                        End If
                    Next
                    ' Importing the lab's analysis method and the corresponding LIMS method. 
                    For Each line As String In IO.File.ReadAllLines("\\mdrnd\AS-Global\Special_Access\EAC\Data\eTrainLite\Methods\" & GlobalVariables.Import.Type & "_EDD.txt")
                        Dim parts() As String = line.Split("|")
                        If Not GlobalVariables.eddAnalysisMethod.ContainsKey(parts(0)) Then
                            GlobalVariables.eddAnalysisMethod.Add(parts(0), parts(1))
                        End If
                    Next
                Else
                    ' Methods
                    For Each line As String In IO.File.ReadAllLines("\\mdrnd\AS-Global\Special_Access\EAC\Data\eTrainLite\NewSample\" & GlobalVariables.Import.Type & ".txt")
                        Dim parts() As String = line.Split("|")
                        GlobalVariables.newSampleMethods.Add(parts(0), parts(1))
                    Next
                    ' Analyte Information
                    For Each line As String In IO.File.ReadAllLines("\\mdrnd\AS-Global\Special_Access\EAC\Data\eTrainLite\NewSample\" & GlobalVariables.Import.Type & "_ANALYTE.txt")
                        Dim parts() As String = line.Split("|")
                        GlobalVariables.newSampleAnalytes.Add(parts(0), parts(1))
                    Next
                    For Each line As String In IO.File.ReadAllLines("\\mdrnd\AS-Global\Special_Access\EAC\Data\eTrainLite\NewSample\LABS.txt")
                        Dim parts() As String = line.Split("|")
                        GlobalVariables.newSampleLabs.Add(parts(0), parts(1))
                    Next
                End If
            End If
            ' Creating a list of lists of the possible recovery units.
            ' Sometimes reported slightly different than what they are in LIMS even though they might mean the same thing. 
            ' Example: ppm and mg/L
            For Each line As String In IO.File.ReadAllLines("\\mdrnd\AS-Global\Special_Access\EAC\Data\eTrainLite\Units\RecoveryUnits.txt")
                Dim count As Integer = 0
                Dim tempList As New List(Of String)
                Dim parts() As String = line.Split("|")
                For i As Integer = 0 To parts.Length - 1
                    tempList.Add(parts(i))
                Next
                GlobalVariables.recoveryUnits.Add(tempList)
            Next
            ' Creating a list of lists of the BEF and TEF values for the D&F data.
            For Each line As String In IO.File.ReadAllLines("\\mdrnd\AS-Global\Special_Access\EAC\Data\eTrainLite\Units\BefAndTef.txt")
                Dim count As Integer = 0
                Dim tempList As New List(Of String)
                    Dim parts() As String = line.Split("|")
                For i As Integer = 0 To parts.Length - 1
                    tempList.Add(parts(i))
                Next
                GlobalVariables.befAndTefScores.Add(tempList)

            Next
            'End If
        Catch ex As Exception
            MsgBox("Error importing methods and dictionaries!" & vbCrLf &
                   "Sub Procedure: loadLimsInformation()" & vbCrLf &
                   "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
        ' Querying LIMS for each sample that was pulled in from the EDD.
        ' Queried for each sample because the LIMS number is the unique identifier to pull in the compound information.
        If GlobalVariables.eTrain.Team <> "NewSample" Then
            For Each limsSample As Sample In GlobalVariables.SampleList
                ' Changing the Sample.type to whatever the first compound in the sampleList is so that it can get transfered to LIMS. 
                Try
                    limsSample.Analysis = GlobalVariables.eddAnalysisMethod.Item(limsSample.CompoundList(0).EDDLabAnlMethodName)
                    If GlobalVariables.Permit.loadLimsCompounds(limsSample.LimsID) Then
                        Continue For
                    Else
                        MsgBox("Error connecting to LIMS for ID: " & limsSample.LimsID & vbCrLf &
                           "Sub Procedure: loadLimsInformation()", MsgBoxStyle.Critical)
                        Return False
                    End If
                Catch ex As Exception
                    MsgBox("Error verify analysis method" & vbCrLf &
                       "Sub Procedure: LoadLimsAnalysisAndUnits()" & vbCrLf &
                       "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                End Try
            Next
            verifyCLabData()
            ' For new samples being created
        Else
            verifyNewSample()
        End If
        Return True
    End Function
    Function loadLimsCompounds(limsID As String) As Boolean
        Dim sConn As String
        Dim sSQL As String
        Dim dtUnits As New DataTable
        Dim dtUnitsUnique As New DataTable
        Dim dvUnits As DataView
        Dim objConn As OdbcConnection
        Dim odAdapter As OdbcDataAdapter
        Dim limsCompound As Compound

        'Connection based on location
        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            sConn = "DRIVER={Microsoft ODBC for Oracle};SERVER=PPT107P.nam.dow.com;UID=FGLLIMS_ENVMD;PWD=lg#En3#;"
            'SQL statement
            sSQL = "SELECT RESULT.NAME, RESULT.UNITS, TEST.ANALYSIS, RESULT.ORDER_NUMBER FROM LIMS_ENVMD.RESULT JOIN LIMS_ENVMD.TEST ON RESULT.TEST_NUMBER = TEST.TEST_NUMBER WHERE TEST.SAMPLE=" & limsID
        End If
        Try
            objConn = New OdbcConnection(sConn)
            objConn.Open()
            odAdapter = New OdbcDataAdapter(sSQL, sConn)
            odAdapter.Fill(dtUnits)
            objConn.Close()

        Catch ex As Exception
            MsgBox("Error connecting to LIMS!" & vbCrLf &
                   "Sub Procedure: LoadLimsAnalysisAndUnits()" & vbCrLf &
                   "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try
        ' Get datatable into view and sort.
        dvUnits = New DataView(dtUnits)
        ' Putting the dataview into the table. 
        dtUnits = dvUnits.ToTable
        ' Creating an arrayList of each compound based on the LIMS number from the method call.
         For Each row As DataRow In dtUnits.Rows
            ' Skipping over the check put in place for the WWTP Grabs, and various NPDES sample sets so they go in as a sample to be checked against.
            If row(0).ToString = "Limit Check" Then
                Continue For
            ElseIf GlobalVariables.limsAnalysisMethod.ContainsKey(row(2).ToString) Then
                limsCompound = New Compound
                limsCompound.EDDsysSampleCode = limsID
                limsCompound.EDDChemicalName = row(0).ToString
                limsCompound.EDDResultUnit = row(1).ToString
                limsCompound.EDDLabAnlMethodName = row(2).ToString
                If GlobalVariables.compNameToCASDic.ContainsKey(limsCompound.EDDChemicalName) Then
                    limsCompound.EDDCasRn = GlobalVariables.compNameToCASDic.Item(limsCompound.EDDChemicalName)
                End If
                GlobalVariables.limsCompoundInformation.Add(limsCompound)
                'Console.WriteLine(limsCompound.EDDCasRn & " " & limsCompound.EDDsysSampleCode & " " & limsCompound.EDDChemicalName & " " & limsCompound.EDDResultUnit & " " & limsCompound.EDDLabAnlMethodName)
            End If
        Next
        Return True
    End Function

    ' Verifying the data inported from the EDD matches the given values pulled from LIMS.
    Sub verifyCLabData()
        Dim EDDSample As Sample
        Dim EDDCompound As Compound
        Dim blnLimsSampleImported As Boolean
        ' Used for the grabs data to add the methyl-Chlorpyrifos (Reldan) and ethyl-Chlorpyrifos (Dursban) recoveries together.
        ' LIMS only has a spot for EC, which is why the methyl is looked for in the samples and added together. 
        Dim strMethylChlorRecovery As String
        Dim blnMethylChlorInSample As Boolean
        Dim blnMethylChorInSet = False
        Dim strMethylChlorValues As String
        ' Variable for skipping the unit check if the sample is a 1005 pressed solids
        Dim bln1005PS As Boolean
        Try
            If GlobalVariables.Import.Type = "1005_PS" Then
                bln1005PS = True
            End If
            ' Sample based on LIMS ID.
            For Each EDDSample In GlobalVariables.SampleList
                ' Resetting this for each sample. 
                blnMethylChlorInSample = False
                ' Each target analyte within the given sample.
                For Each EDDCompound In EDDSample.CompoundList
                    ' Labs use different CAS numbers compounds and they have to match to get entered into LIMS so I am having eTL change it to what it is in LIMS otherwise it would get switched.
                    ' m/p-Xylenes
                    If EDDCompound.EDDCasRn = "136777-61-2" Or EDDCompound.EDDCasRn = "179601-23-1" Or EDDCompound.EDDCasRn = "M/P-XYLENE" Or EDDCompound.EDDCasRn = "XYLMP" Then
                        EDDCompound.EDDCasRn = "MPXYL"
                        ' MCPP
                    ElseIf EDDCompound.EDDCasRn = "93-65-2" Then
                        EDDCompound.EDDCasRn = "7085-19-0"
                    End If
                    ' Looping through each of the compounds made from the LIMS queries to compare to the EDD values.
                    For Each LIMSCompound In GlobalVariables.limsCompoundInformation
                        ' Checking to make sure the sample information was pulled from LIMS and created the samples to check against. 
                        blnLimsSampleImported = True
                        ' Trying to figure out and correct edge case scenarios for checking the information that was imported. 
                        ' The .Contains is for the SGS (and now other labs do it too) metals samples. Their EDDs add a "T" to the end of the CAS numbers so they would be skipped otherwise.
                        ' Allegedly, the T is for "total" metals and the lab wouldn't remove it from the EDD when I asked. 
                        ' Nested if statements because an "and" with the .contains will cause the 1005s to fail when verifying the data. IDK I have no idea why it mattered. 
                        If EDDSample.Analysis = "E200_8_DUP" Or EDDSample.Analysis = "E200MGL_DU" Or EDDSample.Analysis = "E245_HG_UG" Or EDDSample.Analysis = "E245_1_MER" Then
                            If EDDCompound.EDDCasRn.Contains(LIMSCompound.EDDCasRn) Then
                                EDDCompound.EDDCasRn = LIMSCompound.EDDCasRn
                                'ElseIf EDDCompound.EDDChemicalName = LIMSCompound.EDDChemicalName Then
                                '    EDDCompound.EDDChemicalName = LIMSCompound.EDDChemicalName
                            End If
                        End If
                        ' Changing the cas # for TSS reported by Fibertec  for the Kenan samples.
                        If GlobalVariables.Import.Type = "KENAN_FIBERTEC" And EDDCompound.EDDCasRn = "SS" Then
                            EDDCompound.EDDCasRn = "TSS"

                            ' Changing all the stuff from the Cabot data since they have been sending two different types of EDDs.
                        ElseIf GlobalVariables.Import.Type = "CABOT" Then
                            If EDDCompound.EDDCasRn = "TDS" Or EDDCompound.EDDCasRn = "E-10173" Then
                                EDDCompound.EDDCasRn = "0000-000-26"
                                EDDCompound.EDDChemicalName = "Total Dissolved Solids"
                            End If
                            If EDDCompound.EDDCasRn = "TON" Or EDDCompound.EDDChemicalName.ToLower = "total organic nitrogen" Then
                                EDDCompound.EDDChemicalName = "Organic Nitrogen"
                                EDDCompound.EDDCasRn = "STL00111"
                            End If
                            ' Changes being made for the 1005s
                        ElseIf GlobalVariables.Import.Type = "1005_PS" Then
                            If EDDCompound.EDDCasRn = "15831-10-4" And EDDCompound.EDDChemicalName = "3 & 4-Methylphenol" Then
                                EDDCompound.EDDChemicalName = "3/4-Methylphenol"
                            End If
                        ElseIf GlobalVariables.Import.Type = "SEWER" Then
                            If EDDSample.Analysis = "BOD" And EDDCompound.EDDChemicalName <> "BOD" Then
                                EDDCompound.EDDChemicalName = "BOD"
                            ElseIf EDDSample.Analysis = "AMMONIA" And EDDCompound.EDDChemicalName <> "Ammonia" Then
                                EDDCompound.EDDChemicalName = "Ammonia"
                            ElseIf EDDSample.Analysis = "NITRO-TOT" And EDDCompound.EDDChemicalName <> "Nitrogen" Then
                                EDDCompound.EDDChemicalName = "Nitrogen"
                            End If
                        End If
                        ' For changing the LIMS's cas # to something if it didn't have one to pull in because that will cause the analyte to be skipped completely. 
                        If LIMSCompound.EDDCasRn = "" Or String.IsNullOrEmpty(LIMSCompound.EDDCasRn) Then
                            If GlobalVariables.compNameToCASDic.ContainsKey(LIMSCompound.EDDChemicalName) Then
                                LIMSCompound.EDDCasRn = GlobalVariables.compNameToCASDic(LIMSCompound.EDDChemicalName)
                            End If
                        End If

                        ' Using the LIMS number and CAS number as the identifiers to compare the EDD and LIMS results.
                        If LIMSCompound.EDDsysSampleCode = EDDSample.LimsID And LIMSCompound.EDDCasRn = EDDCompound.EDDCasRn Then
                            ' Changing the EDD name to the one used in LIMS.
                            If LIMSCompound.EDDChemicalName <> EDDCompound.EDDChemicalName Then
                                EDDCompound.EDDChemicalName = LIMSCompound.EDDChemicalName
                            End If
                            ' Chlorine and Chloride have the same CAS number, so their names are changed based on the lab method used to what they need to be reported in LIMS.
                            If bln1005PS = True Then
                                If EDDCompound.EDDLabAnlMethodName = "SW9056" Then
                                    EDDCompound.EDDChemicalName = "Chlorine"
                                ElseIf EDDCompound.EDDLabAnlMethodName = "SW9056A" Then
                                    EDDCompound.EDDChemicalName = "Chloride"
                                End If
                                If EDDCompound.EDDCasRn = "87-68-3" Then
                                    EDDCompound.EDDChemicalName = "Hexachlorobutadiene"
                                End If
                            End If

                            ' Checking that the units are the same.
                            ' Example: ppm and mg/L
                            ' Skipping if it was the 1005 pressed solids because everything is reported in weighted values. 
                            If Not checkRecoveryUnits(LIMSCompound.EDDResultUnit.ToLower, EDDCompound.EDDResultUnit.ToLower) And bln1005PS = False Then
                                MsgBox("Unit mismatch between EDD and LIMS values." & vbCrLf &
                                    "LIMS ID: " & EDDSample.LimsID & " - Analyte: " & EDDCompound.EDDChemicalName & vbCrLf &
                                    "Verify in the report and in the text file located at:" & vbCrLf &
                                    "\\mdrnd\AS-Global\Special_Access\EAC\Data\eTrainLite\Units\", MsgBoxStyle.OkOnly)
                            End If
                            ' Changing analytes with no recovery to 0 so that the data will fill into LIMS when submitted. 
                            ' It is supposed to be done if there is no value, but I have had varying results so I just do it outright to not worry about it. 
                            If String.IsNullOrEmpty(EDDCompound.EDDResultValue) Or EDDCompound.EDDResultValue = "" Or EDDCompound.EDDResultValue = vbTab And bln1005PS = True Then
                                EDDCompound.EDDResultValue = 0
                            End If
                            ' Checking the EDD Flags for everything besides the grabs.
                            If GlobalVariables.Import.Type <> "GRABS" And GlobalVariables.Import.Type <> "031B" And GlobalVariables.Import.Type <> "031C" And GlobalVariables.Import.Type <> "DF" Then
                                ' Changing the blank MDL to 0 so it can be converted from a string and not throw an exception. 
                                If EDDCompound.EDDMethodDetectionLimit = "" Then
                                    EDDCompound.EDDMethodDetectionLimit = 0
                                End If
                                If Convert.ToDouble(EDDCompound.EDDResultValue) = 0 Then
                                    If EDDCompound.EDDLabQualifiers = "" Then
                                        EDDCompound.EDDLabQualifiers = "U"
                                    End If

                                ElseIf Convert.ToDouble(EDDCompound.EDDResultValue) >= Convert.ToDouble(EDDCompound.EDDMethodDetectionLimit) And Convert.ToDouble(EDDCompound.EDDResultValue) < Convert.ToDouble(EDDCompound.EDDReportingDetectionLimit) Then
                                    If EDDCompound.EDDLabQualifiers = "" Then
                                        EDDCompound.EDDLabQualifiers = "J"
                                    End If
                                End If
                            End If
                            ' Exiting the loop since the matching analyte was found. 
                            Exit For
                        End If
                        'Console.WriteLine(EDDCompound.EDDCasRn & " " & EDDCompound.EDDsysSampleCode & " " & EDDCompound.EDDChemicalName & " " & EDDCompound.EDDResultUnit & " " & EDDCompound.EDDLabAnlMethodName)
                    Next
                    ' Checking for methyl-Chlorpyrifos.
                    ' Will be skipped by the inner for becaues it doesn't exist in LIMS so it should be outside the LIMScompoud loop and inside of the one for EDDCompound.
                    ' Value isn't set to 0 beacuse it won't be included in the LIMS compounds so it is checked against the empty value.
                    If GlobalVariables.Import.Type = "GRABS" And EDDCompound.EDDCasRn = "5598-13-0" And EDDCompound.EDDResultValue <> "" Then ' Should this be the RL of the lab?
                        strMethylChlorRecovery = EDDCompound.EDDResultValue
                        blnMethylChlorInSample = True
                        blnMethylChorInSet = True
                    End If
                Next
                ' Combining the two values together for methyl and ethyl-Chlorpyrifos if they were detected in the sample.
                If blnMethylChlorInSample = True Then
                    For Each EDDCompound In EDDSample.CompoundList
                        If EDDCompound.EDDCasRn = "2921-88-2" Then
                            ' Try just to make sure that a letter isn't put into the value position and break the program. 
                            Try
                                If EDDCompound.EDDResultValue = "" Then
                                    EDDCompound.EDDResultValue = 0
                                End If
                                ' Everything is stored as strings so they're converted to doubles, added together, and then converted back to a string.
                                strMethylChlorValues = Convert.ToString(Convert.ToDouble(strMethylChlorRecovery) + Convert.ToDouble(EDDCompound.EDDResultValue))
                                EDDCompound.EDDResultValue = strMethylChlorValues
                            Catch ex As Exception
                                MsgBox("Error converting methyl-Chlorpyrifos recovery value." & vbCrLf &
                                       "Sub Procedure: verifyCLabData()" & vbCrLf &
                                       "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                            End Try
                            Exit For
                        End If
                    Next
                End If
            Next
            ' In case the methods for the analysis aren't in the server location, the LIMS samples might not get pulled.
            ' The checks won't be performed since there is nothing the verify the data against.
            ' Various messages for like when the  unit check is skipped or if chlorpyrifos was detected in a grab sample. 
            If blnLimsSampleImported = False Then
                MsgBox("No samples data was pulled in from LIMS to verfiy the data in the EDD." & vbCrLf &
                       "Make sure the analysis and LIMS methods are in the text files located at" & vbCrLf &
                       "\\mdrnd\AS-Global\Special_Access\EAC\Data\eTrainLite\Methods\" & vbCrLf &
                       "Also make sure the samples in the EDD start with their LIMS number.", MsgBoxStyle.Critical)
            End If
            If bln1005PS = True Then
                MsgBox("Samples for 1005 Pressed Solids." & vbCrLf &
                       "Verify that the reported units are correct in the EDD." & vbCrLf &
                       "They were not checked by eTrainLite.", MsgBoxStyle.OkOnly)
            End If
            If blnMethylChorInSet = True Then
                MsgBox("methyl-Chlorpyrifos was detected in at least one of the grabs samples." & vbCrLf &
                       "Verify that the reported value was added correctly to the ethyl-chlorpyrifos (Dursban) recovery.", MsgBoxStyle.OkOnly)
            End If
        Catch ex As Exception
            MsgBox("Error verifying data." & vbCrLf &
                   "Sub Procedure: verifyCLabData()" & vbCrLf &
                   "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Function checkRecoveryUnits(limsUnit As String, eddUnit As String) As Boolean
        For i As Integer = 0 To GlobalVariables.recoveryUnits.Count - 1
            If GlobalVariables.recoveryUnits(i).Contains(limsUnit) And GlobalVariables.recoveryUnits(i).Contains(eddUnit) Then
                Return True
            End If
        Next
        Return False
    End Function
    Sub verifyNewSample()
        Dim EDDSample As Sample
        Dim EDDCompound As Compound

        For Each EDDSample In GlobalVariables.SampleList
            For Each EDDCompound In EDDSample.CompoundList
                If GlobalVariables.newSampleMethods.ContainsKey(EDDCompound.EDDLabAnlMethodName) Then
                    EDDCompound.EDDLabAnlMethodName = GlobalVariables.newSampleMethods(EDDCompound.EDDLabAnlMethodName)
                End If
                If GlobalVariables.newSampleAnalytes.ContainsKey(EDDCompound.EDDChemicalName) Then
                    EDDCompound.EDDChemicalName = GlobalVariables.newSampleAnalytes(EDDCompound.EDDChemicalName)
                End If
                If String.IsNullOrEmpty(EDDCompound.EDDResultValue) Or EDDCompound.EDDResultValue = "" Or EDDCompound.EDDResultValue = vbTab Then
                    EDDCompound.EDDResultValue = 0
                End If
            Next
        Next
    End Sub
End Class