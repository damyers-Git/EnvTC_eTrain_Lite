'Class for method compounds

Public Class mCompound
    Private strName As String
    Private blnCalculated As Boolean

    'FAST Variables
    Private strRRF As String
    Private strRSD As String
    Private strConc As String
    Private strMaxPeakArea As String
    Private strAssoc13C As String
    Private strCS3Amt As String
    Private strTEF As String
    Private strAbundance As String
    Private strIon As String
    Private strUnits As String
    Private strLCSLLim As String
    Private strLCSULim As String


    'HR Variables
    Private strCalAmt As String
    Private strAvgArea As String

    'Chrom Variables
    Private strMDL As String
    Private strPQL As String
    Private strRL As String
    Private strCAS As String
    Private strRecLLim As String
    Private strRecULim As String
    Private strSurLLim As String
    Private strSurULim As String
    Private strMSLLim As String
    Private strMSULim As String
    Public Property AliasList As New ArrayList

    Public Sub New()
        'Constructor
        blnCalculated = False

    End Sub


    Public Property Name() As String
        Get
            Return strName
        End Get
        Set(ByVal value As String)
            strName = value
        End Set
    End Property
    Public Property RRF() As String
        Get
            Return strRRF
        End Get
        Set(ByVal value As String)
            strRRF = value
        End Set
    End Property
    Public Property RSD() As String
        Get
            Return strRSD
        End Get
        Set(ByVal value As String)
            strRSD = value
        End Set
    End Property
    Public Property Conc() As String
        Get
            Return strConc
        End Get
        Set(ByVal value As String)
            strConc = value
        End Set
    End Property
    Public Property MaxPeakArea() As String
        Get
            Return strMaxPeakArea
        End Get
        Set(ByVal value As String)
            strMaxPeakArea = value
        End Set
    End Property
    Public Property Assoc13C() As String
        Get
            Return strAssoc13C
        End Get
        Set(ByVal value As String)
            strAssoc13C = value
        End Set
    End Property
    Public Property CS3Amt() As String
        Get
            Return strCS3Amt
        End Get
        Set(ByVal value As String)
            strCS3Amt = value
        End Set
    End Property
    Public Property TEF() As String
        Get
            Return strTEF
        End Get
        Set(ByVal value As String)
            strTEF = value
        End Set
    End Property
    Public Property Abundance() As String
        Get
            Return strAbundance
        End Get
        Set(ByVal value As String)
            strAbundance = value
        End Set
    End Property
    Public Property Ion() As String
        Get
            Return strIon
        End Get
        Set(ByVal value As String)
            strIon = value
        End Set
    End Property
    Public Property LCSLLim() As String
        Get
            Return strLCSLLim
        End Get
        Set(ByVal value As String)
            strLCSLLim = value
        End Set
    End Property
    Public Property LCSULim() As String
        Get
            Return strLCSULim
        End Get
        Set(ByVal value As String)
            strLCSULim = value
        End Set
    End Property

    Public Property MDL() As String
        Get
            Return strMDL
        End Get
        Set(ByVal value As String)
            strMDL = value
        End Set
    End Property
    Public Property PQL() As String
        Get
            Return strPQL
        End Get
        Set(ByVal value As String)
            strPQL = value
        End Set
    End Property
    Public Property RL() As String
        Get
            Return strRL
        End Get
        Set(ByVal value As String)
            strRL = value
        End Set
    End Property
    Public Property AvgArea() As String
        Get
            Return strAvgArea
        End Get
        Set(ByVal value As String)
            strAvgArea = value
        End Set
    End Property
    Public Property CAS() As String
        Get
            Return strCAS
        End Get
        Set(ByVal value As String)
            strCAS = value
        End Set
    End Property
    Public Property RecLLim() As String
        Get
            Return strRecLLim
        End Get
        Set(ByVal value As String)
            strRecLLim = value
        End Set
    End Property
    Public Property RecULim() As String
        Get
            Return strRecULim
        End Get
        Set(ByVal value As String)
            strRecULim = value
        End Set
    End Property
    Public Property SurLLim() As String
        Get
            Return strSurLLim
        End Get
        Set(ByVal value As String)
            strSurLLim = value
        End Set
    End Property
    Public Property SurULim() As String
        Get
            Return strSurULim
        End Get
        Set(ByVal value As String)
            strSurULim = value
        End Set
    End Property
    Public Property MSLLim() As String
        Get
            Return strMSLLim
        End Get
        Set(ByVal value As String)
            strMSLLim = value
        End Set
    End Property
    Public Property MSULim() As String
        Get
            Return strMSULim
        End Get
        Set(ByVal value As String)
            strMSULim = value
        End Set
    End Property
    Public Property CalAmt() As String
        Get
            Return strCalAmt
        End Get
        Set(ByVal value As String)
            strCalAmt = value
        End Set
    End Property
    Public Property Calculated() As Boolean
        Get
            Return blnCalculated
        End Get
        Set(ByVal value As Boolean)
            blnCalculated = value
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
    'Function returns true if a match is found between a Compound and a method compound
    Function DetermineMatch(ByVal aCompound As Compound) As Boolean
        If Me.AliasList.Count > 0 Then
            For Each item In Me.AliasList
                If aCompound.Name = Me.AliasList(item) Then
                    'Sets the method compound name temporarily to the found alias
                    strName = Me.AliasList(item)
                    Return True
                End If
            Next
        End If
        Return False
    End Function
End Class
