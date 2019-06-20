Imports System.IO

Public Class Method
    Private strName As String
    Private dtCreatedDate As Date
    Private strRptTolerance As String
    Private strETEQ As String
    Private blnLoaded As Boolean
    Public Property mInstrumentList As New ArrayList
    Public Property RefBookList As New ArrayList

    ' For contract lab EDD information
    Private strUnits As String
    Private strMethodName As String

    Public Sub New()
        'Constructor
        blnLoaded = False
    End Sub

    'Sets/Gets
    Public Property Name() As String
        Get
            Return strName
        End Get
        Set(ByVal value As String)
            strName = value
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
    Public Property RptTolerance() As String
        Get
            Return strRptTolerance
        End Get
        Set(ByVal value As String)
            strRptTolerance = value
        End Set
    End Property
    Public Property ETEQ() As String
        Get
            Return strETEQ
        End Get
        Set(ByVal value As String)
            strETEQ = value
        End Set
    End Property
    Public Property Loaded() As Boolean
        Get
            Return blnLoaded
        End Get
        Set(ByVal value As Boolean)
            blnLoaded = value
        End Set
    End Property
    Public Property Units() As String
        Get
            Return strUnits
        End Get
        Set(ByVal value As String)
            strUnits = value
        End Set
    End Property
    Public Property MethodName() As String
        Get
            Return strMethodName
        End Get
        Set(ByVal value As String)
            strMethodName = value
        End Set
    End Property

    Sub LoadMethodNames()
        Dim strFileNames() As String
        Dim aMethod As Method

        'Check if methods are already populated then clear if they are, assuming team change or updated methods from file
        If GlobalVariables.MethodList.Count <> 0 Then
            GlobalVariables.MethodList.Clear()
        End If

        Try
            If GlobalVariables.eTrain.Location = "MIDLAND" Then
                If GlobalVariables.eTrain.Team = "FAST" Then
                    'Gets listing of file names in folder and uses them as Method names, adds them to methodlist
                    strFileNames = Directory.GetFiles(GlobalVariables.eTrain.DataFileFP & "Midland\FAST\Projects_Methods\")
                    For Each f In strFileNames
                        aMethod = New Method
                        aMethod.Name = Path.GetFileNameWithoutExtension(f)
                        GlobalVariables.MethodList.Add(aMethod)
                    Next
                ElseIf GlobalVariables.eTrain.Team = "HR" Then
                    'Gets listing of file names in folder and uses them as Method names, adds them to methodlist
                    strFileNames = Directory.GetFiles(GlobalVariables.eTrain.DataFileFP & "Midland\HR\Projects_Methods\")
                    For Each f In strFileNames
                        aMethod = New Method
                        aMethod.Name = Path.GetFileNameWithoutExtension(f)
                        GlobalVariables.MethodList.Add(aMethod)
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox("Error getting Method names!" & vbCrLf & _
                   "Sub Procedure: LoadMethodNames()" & vbCrLf & _
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Sub LoadMethod(ByVal strName As String)
        Dim strLine As String
        Dim arrSplit() As String
        Dim arrSpl() As String
        Dim arrSpl2() As String
        Dim arrSpl3() As String
        Dim aMethod As Method
        Dim aInstrument As mInstrument
        Dim aCompound As mCompound
        Dim aStandard As mStandard
        Dim aRefBook As RefBook

        Try
            If GlobalVariables.eTrain.Location = "MIDLAND" Then
                If GlobalVariables.eTrain.Team = "FAST" Then
                    'Figure out which method is being loaded
                    For Each aMethod In GlobalVariables.MethodList
                        If strName = aMethod.Name Then
                            If Not aMethod.Loaded Then
                                Dim sr As StreamReader = New StreamReader(GlobalVariables.eTrain.DataFileFP & "Midland\FAST\Projects_Methods\" & strName & ".et2")
                                strLine = sr.ReadLine
                                'Name check
                                If aMethod.Name = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))) Then
                                    strLine = sr.ReadLine
                                    'Date
                                    aMethod.CreatedDate = CDate(Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))))
                                    'Book info
                                    strLine = sr.ReadLine
                                    strLine = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    'Standard reference book 
                                    If strLine <> "" Then
                                        If InStr(strLine, "|") Then
                                            arrSpl = strLine.Split("|")
                                            For i = 0 To UBound(arrSpl)
                                                aRefBook = New RefBook
                                                arrSpl3 = arrSpl(i).Split("_")
                                                aRefBook.Name = Trim(arrSpl3(0))
                                                aRefBook.Expiration = CDate(Trim(arrSpl3(1)))
                                                aRefBook.Note = Trim(arrSpl3(2))
                                                arrSpl2 = aRefBook.Name.Split("-")
                                                aRefBook.Type = "13C"
                                                aRefBook.Book = Trim(arrSpl2(0))
                                                aRefBook.BookPg = Trim(arrSpl2(1))
                                                aRefBook.Num = Trim(arrSpl2(2))
                                                aRefBook.Section = Trim(arrSpl2(3))
                                                aMethod.RefBookList.Add(aRefBook)
                                            Next
                                        Else
                                            aRefBook = New RefBook
                                            arrSpl3 = strLine.Split("_")
                                            aRefBook.Name = Trim(arrSpl3(0))
                                            aRefBook.Expiration = CDate(Trim(arrSpl3(1)))
                                            aRefBook.Note = Trim(arrSpl3(2))
                                            arrSpl2 = aRefBook.Name.Split("-")
                                            aRefBook.Type = "13C"
                                            aRefBook.Book = Trim(arrSpl2(0))
                                            aRefBook.BookPg = Trim(arrSpl2(1))
                                            aRefBook.Num = Trim(arrSpl2(2))
                                            aRefBook.Section = Trim(arrSpl2(3))
                                            aMethod.RefBookList.Add(aRefBook)
                                        End If

                                    End If

                                    strLine = sr.ReadLine
                                    strLine = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    'Injection reference book 
                                    If strLine <> "" Then
                                        If InStr(strLine, "|") Then
                                            arrSpl = strLine.Split("|")
                                            For i = 0 To UBound(arrSpl)
                                                aRefBook = New RefBook
                                                arrSpl3 = arrSpl(i).Split("_")
                                                aRefBook.Name = Trim(arrSpl3(0))
                                                aRefBook.Expiration = CDate(Trim(arrSpl3(1)))
                                                aRefBook.Note = Trim(arrSpl3(2))
                                                arrSpl2 = aRefBook.Name.Split("-")
                                                aRefBook.Type = "Injection"
                                                aRefBook.Book = Trim(arrSpl2(0))
                                                aRefBook.BookPg = Trim(arrSpl2(1))
                                                aRefBook.Num = Trim(arrSpl2(2))
                                                aRefBook.Section = Trim(arrSpl2(3))
                                                aMethod.RefBookList.Add(aRefBook)
                                            Next
                                        Else
                                            aRefBook = New RefBook
                                            arrSpl3 = strLine.Split("_")
                                            aRefBook.Name = Trim(arrSpl3(0))
                                            aRefBook.Expiration = CDate(Trim(arrSpl3(1)))
                                            aRefBook.Note = Trim(arrSpl3(2))
                                            arrSpl2 = aRefBook.Name.Split("-")
                                            aRefBook.Type = "Injection"
                                            aRefBook.Book = Trim(arrSpl2(0))
                                            aRefBook.BookPg = Trim(arrSpl2(1))
                                            aRefBook.Num = Trim(arrSpl2(2))
                                            aRefBook.Section = Trim(arrSpl2(3))
                                            aMethod.RefBookList.Add(aRefBook)
                                        End If

                                    End If

                                    strLine = sr.ReadLine
                                    strLine = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    'LCS reference book 
                                    If strLine <> "" Then
                                        If InStr(strLine, "|") Then
                                            arrSpl = strLine.Split("|")
                                            For i = 0 To UBound(arrSpl)
                                                aRefBook = New RefBook
                                                arrSpl3 = arrSpl(i).Split("_")
                                                aRefBook.Name = Trim(arrSpl3(0))
                                                aRefBook.Expiration = CDate(Trim(arrSpl3(1)))
                                                aRefBook.Note = Trim(arrSpl3(2))
                                                arrSpl2 = aRefBook.Name.Split("-")
                                                aRefBook.Type = "LCS"
                                                aRefBook.Book = Trim(arrSpl2(0))
                                                aRefBook.BookPg = Trim(arrSpl2(1))
                                                aRefBook.Num = Trim(arrSpl2(2))
                                                aRefBook.Section = Trim(arrSpl2(3))
                                                aMethod.RefBookList.Add(aRefBook)
                                            Next
                                        Else
                                            aRefBook = New RefBook
                                            arrSpl3 = strLine.Split("_")
                                            aRefBook.Name = Trim(arrSpl3(0))
                                            aRefBook.Expiration = CDate(Trim(arrSpl3(1)))
                                            aRefBook.Note = Trim(arrSpl3(2))
                                            arrSpl2 = aRefBook.Name.Split("-")
                                            aRefBook.Type = "LCS"
                                            aRefBook.Book = Trim(arrSpl2(0))
                                            aRefBook.BookPg = Trim(arrSpl2(1))
                                            aRefBook.Num = Trim(arrSpl2(2))
                                            aRefBook.Section = Trim(arrSpl2(3))
                                            aMethod.RefBookList.Add(aRefBook)
                                        End If
                                    End If

                                    'Report Tolerance
                                    strLine = sr.ReadLine
                                    aMethod.RptTolerance = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    strLine = sr.ReadLine
                                    aMethod.ETEQ = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    'Instrument
                                    strLine = sr.ReadLine
                                    strLine = sr.ReadLine
                                    Do Until sr.EndOfStream
                                        aInstrument = New mInstrument
                                        aInstrument.Name = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                        strLine = sr.ReadLine
                                        'Reviewed
                                        aInstrument.Reviewed = CBool(Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))))
                                        'Date reviewed
                                        strLine = sr.ReadLine
                                        aInstrument.ReviewedDate = CDate(Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))))
                                        strLine = sr.ReadLine
                                        If InStr(strLine, "Standards", CompareMethod.Binary) Then
                                            strLine = sr.ReadLine()
                                            Do Until InStr(strLine, "Compounds", CompareMethod.Binary)
                                                'Set standard info
                                                arrSplit = strLine.Split(",")
                                                aStandard = New mStandard
                                                aStandard.Name = arrSplit(0)
                                                aStandard.Type = arrSplit(1)
                                                aStandard.AvgArea = arrSplit(2)
                                                aStandard.CalAmt = arrSplit(3)
                                                aStandard.Conc = arrSplit(4)
                                                aStandard.RecLowLim = arrSplit(5)
                                                aStandard.RecUpLim = arrSplit(6)
                                                aStandard.IonTarget = arrSplit(7)
                                                aStandard.IonQual = arrSplit(8)
                                                aStandard.AbundTarget = arrSplit(9)
                                                aStandard.AbundQual = arrSplit(10)
                                                aInstrument.mStandardList.Add(aStandard)
                                                strLine = sr.ReadLine
                                            Loop
                                            strLine = sr.ReadLine()
                                            Do Until InStr(strLine, "Instrument:", CompareMethod.Binary) Or strLine = Nothing
                                                'Set compound info
                                                arrSplit = strLine.Split(",")
                                                aCompound = New mCompound
                                                aCompound.Name = arrSplit(0)
                                                aCompound.RRF = arrSplit(1)
                                                aCompound.RSD = arrSplit(2)
                                                aCompound.MaxPeakArea = arrSplit(3)
                                                aCompound.Conc = arrSplit(4)
                                                aCompound.CS3Amt = arrSplit(5)
                                                aCompound.TEF = arrSplit(6)
                                                aCompound.Ion = arrSplit(7)
                                                aCompound.Abundance = arrSplit(8)
                                                aCompound.LCSLLim = arrSplit(9)
                                                aCompound.LCSULim = arrSplit(10)
                                                aCompound.Assoc13C = arrSplit(11)
                                                aInstrument.mCompoundList.Add(aCompound)
                                                strLine = sr.ReadLine
                                            Loop
                                            aMethod.mInstrumentList.Add(aInstrument)
                                        End If
                                    Loop
                                    aMethod.Loaded = True
                                End If

                                sr.Close()
                                sr.Dispose()
                            End If
                        End If
                    Next
                ElseIf GlobalVariables.eTrain.Team = "HR" Then
                    'Figure out which method is being loaded
                    For Each aMethod In GlobalVariables.MethodList
                        If strName = aMethod.Name Then
                            If Not aMethod.Loaded Then
                                Dim sr As StreamReader = New StreamReader(GlobalVariables.eTrain.DataFileFP & "Midland\HR\Projects_Methods\" & strName & ".et2")
                                strLine = sr.ReadLine
                                'Name check
                                If aMethod.Name = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))) Then
                                    strLine = sr.ReadLine
                                    'Date
                                    aMethod.CreatedDate = CDate(Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))))
                                    'Book info
                                    strLine = sr.ReadLine
                                    strLine = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    'Standard reference book 
                                    If strLine <> "" Then
                                        If InStr(strLine, "|") Then
                                            arrSpl = strLine.Split("|")
                                            For i = 0 To UBound(arrSpl)
                                                aRefBook = New RefBook
                                                arrSpl3 = arrSpl(i).Split("_")
                                                aRefBook.Name = Trim(arrSpl3(0))
                                                aRefBook.Expiration = CDate(Trim(arrSpl3(1)))
                                                aRefBook.Note = Trim(arrSpl3(2))
                                                arrSpl2 = aRefBook.Name.Split("-")
                                                aRefBook.Type = "13C"
                                                aRefBook.Book = Trim(arrSpl2(0))
                                                aRefBook.BookPg = Trim(arrSpl2(1))
                                                aRefBook.Num = Trim(arrSpl2(2))
                                                aRefBook.Section = Trim(arrSpl2(3))
                                                aMethod.RefBookList.Add(aRefBook)
                                            Next
                                        Else
                                            aRefBook = New RefBook
                                            arrSpl3 = strLine.Split("_")
                                            aRefBook.Name = Trim(arrSpl3(0))
                                            aRefBook.Expiration = CDate(Trim(arrSpl3(1)))
                                            aRefBook.Note = Trim(arrSpl3(2))
                                            arrSpl2 = aRefBook.Name.Split("-")
                                            aRefBook.Type = "13C"
                                            aRefBook.Book = Trim(arrSpl2(0))
                                            aRefBook.BookPg = Trim(arrSpl2(1))
                                            aRefBook.Num = Trim(arrSpl2(2))
                                            aRefBook.Section = Trim(arrSpl2(3))
                                            aMethod.RefBookList.Add(aRefBook)
                                        End If

                                    End If

                                    strLine = sr.ReadLine
                                    strLine = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    'Injection reference book 
                                    If strLine <> "" Then
                                        If InStr(strLine, "|") Then
                                            arrSpl = strLine.Split("|")
                                            For i = 0 To UBound(arrSpl)
                                                aRefBook = New RefBook
                                                arrSpl3 = arrSpl(i).Split("_")
                                                aRefBook.Name = Trim(arrSpl3(0))
                                                aRefBook.Expiration = CDate(Trim(arrSpl3(1)))
                                                aRefBook.Note = Trim(arrSpl3(2))
                                                arrSpl2 = aRefBook.Name.Split("-")
                                                aRefBook.Type = "Injection"
                                                aRefBook.Book = Trim(arrSpl2(0))
                                                aRefBook.BookPg = Trim(arrSpl2(1))
                                                aRefBook.Num = Trim(arrSpl2(2))
                                                aRefBook.Section = Trim(arrSpl2(3))
                                                aMethod.RefBookList.Add(aRefBook)
                                            Next
                                        Else
                                            aRefBook = New RefBook
                                            arrSpl3 = strLine.Split("_")
                                            aRefBook.Name = Trim(arrSpl3(0))
                                            aRefBook.Expiration = CDate(Trim(arrSpl3(1)))
                                            aRefBook.Note = Trim(arrSpl3(2))
                                            arrSpl2 = aRefBook.Name.Split("-")
                                            aRefBook.Type = "Injection"
                                            aRefBook.Book = Trim(arrSpl2(0))
                                            aRefBook.BookPg = Trim(arrSpl2(1))
                                            aRefBook.Num = Trim(arrSpl2(2))
                                            aRefBook.Section = Trim(arrSpl2(3))
                                            aMethod.RefBookList.Add(aRefBook)
                                        End If

                                    End If

                                    strLine = sr.ReadLine
                                    strLine = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    'LCS reference book 
                                    If strLine <> "" Then
                                        If InStr(strLine, "|") Then
                                            arrSpl = strLine.Split("|")
                                            For i = 0 To UBound(arrSpl)
                                                aRefBook = New RefBook
                                                arrSpl3 = arrSpl(i).Split("_")
                                                aRefBook.Name = Trim(arrSpl3(0))
                                                aRefBook.Expiration = CDate(Trim(arrSpl3(1)))
                                                aRefBook.Note = Trim(arrSpl3(2))
                                                arrSpl2 = aRefBook.Name.Split("-")
                                                aRefBook.Type = "LCS"
                                                aRefBook.Book = Trim(arrSpl2(0))
                                                aRefBook.BookPg = Trim(arrSpl2(1))
                                                aRefBook.Num = Trim(arrSpl2(2))
                                                aRefBook.Section = Trim(arrSpl2(3))
                                                aMethod.RefBookList.Add(aRefBook)
                                            Next
                                        Else
                                            aRefBook = New RefBook
                                            arrSpl3 = strLine.Split("_")
                                            aRefBook.Name = Trim(arrSpl3(0))
                                            aRefBook.Expiration = CDate(Trim(arrSpl3(1)))
                                            aRefBook.Note = Trim(arrSpl3(2))
                                            arrSpl2 = aRefBook.Name.Split("-")
                                            aRefBook.Type = "LCS"
                                            aRefBook.Book = Trim(arrSpl2(0))
                                            aRefBook.BookPg = Trim(arrSpl2(1))
                                            aRefBook.Num = Trim(arrSpl2(2))
                                            aRefBook.Section = Trim(arrSpl2(3))
                                            aMethod.RefBookList.Add(aRefBook)
                                        End If
                                    End If

                                    'Report Tolerance
                                    strLine = sr.ReadLine
                                    aMethod.RptTolerance = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    strLine = sr.ReadLine
                                    aMethod.ETEQ = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                    'Instrument
                                    strLine = sr.ReadLine
                                    strLine = sr.ReadLine
                                    Do Until sr.EndOfStream
                                        aInstrument = New mInstrument
                                        aInstrument.Name = Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":")))
                                        strLine = sr.ReadLine
                                        'Reviewed
                                        aInstrument.Reviewed = CBool(Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))))
                                        'Date reviewed
                                        strLine = sr.ReadLine
                                        aInstrument.ReviewedDate = CDate(Trim(strLine.Substring(InStr(strLine, ":"), strLine.Length - InStr(strLine, ":"))))
                                        strLine = sr.ReadLine
                                        If InStr(strLine, "Standards", CompareMethod.Binary) Then
                                            strLine = sr.ReadLine()
                                            Do Until InStr(strLine, "Compounds", CompareMethod.Binary)
                                                'Set standard info
                                                arrSplit = strLine.Split(",")
                                                aStandard = New mStandard
                                                aStandard.Name = arrSplit(0)
                                                aStandard.Type = arrSplit(1)
                                                aStandard.AvgArea = arrSplit(2)
                                                aStandard.CalAmt = arrSplit(3)
                                                aStandard.Conc = arrSplit(4)
                                                aStandard.RecLowLim = arrSplit(5)
                                                aStandard.RecUpLim = arrSplit(6)
                                                aStandard.IonTarget = arrSplit(7)
                                                aStandard.IonQual = arrSplit(8)
                                                aStandard.AbundTarget = arrSplit(9)
                                                aStandard.AbundQual = arrSplit(10)
                                                aInstrument.mStandardList.Add(aStandard)
                                                strLine = sr.ReadLine
                                            Loop
                                            strLine = sr.ReadLine()
                                            Do Until InStr(strLine, "Instrument:", CompareMethod.Binary) Or strLine = Nothing
                                                'Set compound info
                                                arrSplit = strLine.Split(",")
                                                aCompound = New mCompound
                                                aCompound.Name = arrSplit(0)
                                                aCompound.RRF = arrSplit(1)
                                                aCompound.RSD = arrSplit(2)
                                                aCompound.MaxPeakArea = arrSplit(3)
                                                aCompound.Conc = arrSplit(4)
                                                aCompound.CalAmt = arrSplit(5)
                                                aCompound.TEF = arrSplit(6)
                                                aCompound.Ion = arrSplit(7)
                                                aCompound.Abundance = arrSplit(8)
                                                aCompound.LCSLLim = arrSplit(9)
                                                aCompound.LCSULim = arrSplit(10)
                                                aCompound.Assoc13C = arrSplit(11)
                                                aInstrument.mCompoundList.Add(aCompound)
                                                strLine = sr.ReadLine
                                            Loop
                                            aMethod.mInstrumentList.Add(aInstrument)
                                        End If
                                    Loop
                                    aMethod.Loaded = True
                                End If

                                sr.Close()
                                sr.Dispose()
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox("Error reading Method file!" & vbCrLf & _
                   "Sub Procedure: LoadMethod()" & vbCrLf & _
                "Line: " & strLine & vbCrLf & _
                "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    'Reads in associated 13c's for method edit form
    Sub LoadAssoc13cFile()
        Dim sr As StreamReader
        Dim strFileLoc As String

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                Try
                    strFileLoc = "\\Helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Midland\FAST\Associated13Cs.et2"
                    If File.Exists(strFileLoc) Then
                        'Begin Write
                        sr = New StreamReader(strFileLoc)
                        Do Until sr.EndOfStream
                            GlobalVariables.Associated13Cs.Add(sr.ReadLine)
                        Loop
                    Else
                        MsgBox("Error reading Associated 13C database file, file not found.", MsgBoxStyle.Critical)
                    End If
                Catch ex As Exception
                    MsgBox("Error reading Associated 13C database file." & vbCrLf & _
                           "Sub Procedure: LoadAssoc13cFile()" & vbCrLf & _
                    "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                End Try
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                Try
                    strFileLoc = "\\Helium\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Midland\HR\Associated13Cs.et2"
                    If File.Exists(strFileLoc) Then
                        'Begin Write
                        sr = New StreamReader(strFileLoc)
                        Do Until sr.EndOfStream
                            GlobalVariables.Associated13Cs.Add(sr.ReadLine)
                        Loop
                    Else
                        MsgBox("Error reading Associated 13C database file, file not found.", MsgBoxStyle.Critical)
                    End If
                Catch ex As Exception
                    MsgBox("Error reading Associated 13C database file." & vbCrLf & _
                           "Sub Procedure: LoadAssoc13cFile()" & vbCrLf & _
                    "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                End Try
            End If
        End If
    End Sub

    Sub Associate13cs(ByVal aInstrument As mInstrument)
        Dim aSample As Sample
        Dim aCompound As Compound
        Dim amCompound As mCompound

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                For Each aSample In GlobalVariables.SampleList
                    For Each aCompound In aSample.CompoundList
                        For Each amCompound In aInstrument.mCompoundList
                            If aCompound.Name = amCompound.Name Then
                                aCompound.MidF13CAssoc = amCompound.Assoc13C
                            End If
                        Next
                    Next
                Next
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                For Each aSample In GlobalVariables.SampleList
                    For Each aCompound In aSample.CompoundList
                        For Each amCompound In aInstrument.mCompoundList
                            If aCompound.Name = amCompound.Name Then
                                aCompound.MidF13CAssoc = amCompound.Assoc13C
                            End If
                        Next
                    Next
                Next
            End If
        End If
    End Sub

    'True return means save successful, false return save cancelled/unsuccessful nothing changed
    Function SaveMethod(ByVal aMethod As Method) As Boolean

        Dim strFileLoc As String

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                Try
                    strFileLoc = "\\Mdrnd\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Midland\FAST\Projects_Methods\" & aMethod.Name & ".et2"
                    If Not File.Exists(strFileLoc) Then
                        'New Method
                        If GlobalVariables.Method.WriteMethod(aMethod, strFileLoc) Then
                            Return True
                        Else
                            Return False
                        End If
                    Else
                        'Method already exists, update
                        If MsgBox("A Method already exists with this name, do you intend to update it?", MsgBoxStyle.YesNo, "eTrain 2.0") = MsgBoxResult.Yes Then
                            If GlobalVariables.Method.WriteMethod(aMethod, strFileLoc) Then
                                Return True
                            Else
                                Return False
                            End If
                        End If
                    End If

                Catch ex As Exception
                    MsgBox("Error Saving Method File" & vbCrLf & _
                           "Sub Procedure: SaveMethod()" & vbCrLf & _
                     "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                    Return False
                End Try
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                Try
                    strFileLoc = "\\Mdrnd\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Midland\HR\Projects_Methods\" & aMethod.Name & ".et2"
                    If Not File.Exists(strFileLoc) Then
                        'New Method
                        If GlobalVariables.Method.WriteMethod(aMethod, strFileLoc) Then
                            Return True
                        Else
                            Return False
                        End If
                    Else
                        'Method already exists, update
                        If MsgBox("A Method already exists with this name, do you intend to update it?", MsgBoxStyle.YesNo, "eTrain 2.0") = MsgBoxResult.Yes Then
                            If GlobalVariables.Method.WriteMethod(aMethod, strFileLoc) Then
                                Return True
                            Else
                                Return False
                            End If
                        End If
                    End If

                Catch ex As Exception
                    MsgBox("Error Saving Method File" & vbCrLf & _
                           "Sub Procedure: SaveMethod()" & vbCrLf & _
                     "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                    Return False
                End Try

            End If
        End If
        Return False
    End Function

    Function WriteMethod(ByVal aMethod As Method, ByVal strFileLoc As String)
        Dim sr As StreamWriter
        Dim curDate As Date
        Dim aInstrument As mInstrument
        Dim aStandard As mStandard
        Dim aCompound As mCompound
        Dim aRefBook As RefBook
        Dim strStdRefBook As String
        Dim strInjRefBook As String
        Dim strLCSRefBook As String

        curDate = DateTime.Now

        If GlobalVariables.eTrain.Location = "MIDLAND" Then
            If GlobalVariables.eTrain.Team = "FAST" Then
                Try
                    strStdRefBook = ""
                    strInjRefBook = ""
                    strLCSRefBook = ""
                    'Backup old file if there
                    If File.Exists(strFileLoc) Then
                        File.Copy(strFileLoc, "\\Mdrnd\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Midland\FAST\Projects_Methods\Backups\" & aMethod.Name & "_" & _
                                  curDate.Month & curDate.Day & curDate.Year & "_" & curDate.Hour & curDate.Minute & ".et2")
                    End If
                    'Begin write
                    sr = New StreamWriter(strFileLoc)
                    sr.WriteLine("Method Name: " & aMethod.Name)
                    sr.WriteLine("Method Date: " & CStr(aMethod.CreatedDate))
                    For Each aRefBook In aMethod.RefBookList
                        If aRefBook.Type = "13C" Then
                            If strStdRefBook = "" Then
                                strStdRefBook = aRefBook.Name & "_" & CStr(aRefBook.Expiration) & "_" & CStr(aRefBook.Note)
                            Else
                                strStdRefBook = strStdRefBook & " | " & aRefBook.Name & "_" & CStr(aRefBook.Expiration) & "_" & CStr(aRefBook.Note)
                            End If
                        End If
                        If aRefBook.Type = "Injection" Then
                            If strInjRefBook = "" Then
                                strInjRefBook = aRefBook.Name & "_" & CStr(aRefBook.Expiration) & "_" & CStr(aRefBook.Note)
                            Else
                                strInjRefBook = strInjRefBook & " | " & aRefBook.Name & "_" & CStr(aRefBook.Expiration) & "_" & CStr(aRefBook.Note)
                            End If
                        End If
                        If aRefBook.Type = "LCS" Then
                            If strLCSRefBook = "" Then
                                strLCSRefBook = aRefBook.Name & "_" & CStr(aRefBook.Expiration) & "_" & CStr(aRefBook.Note)
                            Else
                                strLCSRefBook = strLCSRefBook & " | " & aRefBook.Name & "_" & CStr(aRefBook.Expiration) & "_" & CStr(aRefBook.Note)
                            End If
                        End If
                    Next
                    sr.WriteLine("13C Book Info: " & strStdRefBook)
                    sr.WriteLine("Inj Book Info: " & strInjRefBook)
                    sr.WriteLine("LCS Book Info: " & strLCSRefBook)
                    sr.WriteLine("Report Tolerance: " & aMethod.RptTolerance)
                    sr.WriteLine("ETEQ: " & aMethod.ETEQ)
                    sr.WriteLine("Cal Information By Instrument")
                    For Each aInstrument In aMethod.mInstrumentList
                        sr.WriteLine("Instrument: " & aInstrument.Name)
                        sr.WriteLine("Reviewed: " & CStr(aInstrument.Reviewed))
                        sr.WriteLine("Reviewed Date: " & CStr(aInstrument.ReviewedDate))
                        sr.WriteLine("Standards")
                        For Each aStandard In aInstrument.mStandardList
                            sr.WriteLine(aStandard.Name & "," & aStandard.Type & "," & aStandard.AvgArea & "," & aStandard.CalAmt & "," & aStandard.Conc & "," & aStandard.RecLowLim & "," & aStandard.RecUpLim & "," & aStandard.IonTarget & "," & aStandard.IonQual & "," & aStandard.AbundTarget & "," & aStandard.AbundQual)
                        Next
                        sr.WriteLine("Compounds")
                        For Each aCompound In aInstrument.mCompoundList
                            sr.WriteLine(aCompound.Name & "," & aCompound.RRF & "," & aCompound.RSD & "," & aCompound.MaxPeakArea & "," & aCompound.Conc & "," & aCompound.CS3Amt & "," & aCompound.TEF & "," & aCompound.Ion & "," & aCompound.Abundance & "," & aCompound.LCSLLim & "," & aCompound.LCSULim & "," & aCompound.Assoc13C)
                        Next
                    Next
                    sr.Close()
                    sr.Dispose()
                    Return True
                Catch ex As Exception
                    MsgBox("Error Writing Method File" & vbCrLf & _
                           "Sub Procedure: WriteMethod()" & vbCrLf & _
                     "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                    Return False
                End Try
            ElseIf GlobalVariables.eTrain.Team = "HR" Then
                Try
                    strStdRefBook = ""
                    strInjRefBook = ""
                    strLCSRefBook = ""
                    'Backup old file if there
                    If File.Exists(strFileLoc) Then
                        File.Copy(strFileLoc, "\\Mdrnd\as-global\Special_Access\EAC\Data\eTrain\DataFiles\Midland\HR\Projects_Methods\Backups\" & aMethod.Name & "_" & _
                                  curDate.Month & curDate.Day & curDate.Year & "_" & curDate.Hour & curDate.Minute & ".et2")
                    End If
                    'Begin write
                    sr = New StreamWriter(strFileLoc)
                    sr.WriteLine("Method Name: " & aMethod.Name)
                    sr.WriteLine("Method Date: " & CStr(aMethod.CreatedDate))
                    For Each aRefBook In aMethod.RefBookList
                        If aRefBook.Type = "13C" Then
                            If strStdRefBook = "" Then
                                strStdRefBook = aRefBook.Name & "_" & CStr(aRefBook.Expiration) & "_" & CStr(aRefBook.Note)
                            Else
                                strStdRefBook = strStdRefBook & " | " & aRefBook.Name & "_" & CStr(aRefBook.Expiration) & "_" & CStr(aRefBook.Note)
                            End If
                        End If
                        If aRefBook.Type = "Injection" Then
                            If strInjRefBook = "" Then
                                strInjRefBook = aRefBook.Name & "_" & CStr(aRefBook.Expiration) & "_" & CStr(aRefBook.Note)
                            Else
                                strInjRefBook = strInjRefBook & " | " & aRefBook.Name & "_" & CStr(aRefBook.Expiration) & "_" & CStr(aRefBook.Note)
                            End If
                        End If
                        If aRefBook.Type = "LCS" Then
                            If strLCSRefBook = "" Then
                                strLCSRefBook = aRefBook.Name & "_" & CStr(aRefBook.Expiration) & "_" & CStr(aRefBook.Note)
                            Else
                                strLCSRefBook = strLCSRefBook & " | " & aRefBook.Name & "_" & CStr(aRefBook.Expiration) & "_" & CStr(aRefBook.Note)
                            End If
                        End If
                    Next
                    sr.WriteLine("13C Book Info: " & strStdRefBook)
                    sr.WriteLine("Inj Book Info: " & strInjRefBook)
                    sr.WriteLine("LCS Book Info: " & strLCSRefBook)
                    sr.WriteLine("Report Tolerance: " & aMethod.RptTolerance)
                    sr.WriteLine("ETEQ: " & aMethod.ETEQ)
                    sr.WriteLine("Cal Information By Instrument")
                    For Each aInstrument In aMethod.mInstrumentList
                        sr.WriteLine("Instrument: " & aInstrument.Name)
                        sr.WriteLine("Reviewed: " & CStr(aInstrument.Reviewed))
                        sr.WriteLine("Reviewed Date: " & CStr(aInstrument.ReviewedDate))
                        sr.WriteLine("Standards")
                        For Each aStandard In aInstrument.mStandardList
                            sr.WriteLine(aStandard.Name & "," & aStandard.Type & "," & aStandard.AvgArea & "," & aStandard.CalAmt & "," & aStandard.Conc & "," & aStandard.RecLowLim & "," & aStandard.RecUpLim & "," & aStandard.IonTarget & "," & aStandard.IonQual & "," & aStandard.AbundTarget & "," & aStandard.AbundQual)
                        Next
                        sr.WriteLine("Compounds")
                        For Each aCompound In aInstrument.mCompoundList
                            sr.WriteLine(aCompound.Name & "," & aCompound.RRF & "," & aCompound.RSD & "," & aCompound.MaxPeakArea & "," & aCompound.Conc & "," & aCompound.CalAmt & "," & aCompound.TEF & "," & aCompound.Ion & "," & aCompound.Abundance & "," & aCompound.LCSLLim & "," & aCompound.LCSULim & "," & aCompound.Assoc13C)
                        Next
                    Next
                    sr.Close()
                    sr.Dispose()
                    Return True
                Catch ex As Exception
                    MsgBox("Error Writing Method File" & vbCrLf & _
                           "Sub Procedure: WriteMethod()" & vbCrLf & _
                     "Logic Error: " & ex.Message, MsgBoxStyle.Critical)
                    Return False
                End Try
            End If
        End If
    End Function

End Class
