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
                'Set starting Analysis/Project
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
                'Set starting Analysis/Project
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
End Class
