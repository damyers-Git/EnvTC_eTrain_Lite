'Class for method compounds

Public Class mStandard
    Private strName As String

    'FAST
    Private strAvgArea As String
    Private strCalAmt As String
    Private strConc As String
    Private strRecLowLim As String
    Private strRecUpLim As String
    Private strType As String
    Private strAbundTarget As String
    Private strAbundQual As String
    Private strIonTarget As String
    Private strIonQual As String

    'Chrom Variables
    Private strRL As String
    Private strMDL As String
    Private strPQL As String

    Public Property Name() As String
        Get
            Return strName
        End Get
        Set(ByVal value As String)
            strName = value
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
    Public Property CalAmt() As String
        Get
            Return strCalAmt
        End Get
        Set(ByVal value As String)
            strCalAmt = value
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
    Public Property RecLowLim() As String
        Get
            Return strRecLowLim
        End Get
        Set(ByVal value As String)
            strRecLowLim = value
        End Set
    End Property
    Public Property RecUpLim() As String
        Get
            Return strRecUpLim
        End Get
        Set(ByVal value As String)
            strRecUpLim = value
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
    Public Property AbundTarget() As String
        Get
            Return strAbundTarget
        End Get
        Set(ByVal value As String)
            strAbundTarget = value
        End Set
    End Property
    Public Property AbundQual() As String
        Get
            Return strAbundQual
        End Get
        Set(ByVal value As String)
            strAbundQual = value
        End Set
    End Property
    Public Property IonTarget() As String
        Get
            Return strIonTarget
        End Get
        Set(ByVal value As String)
            strIonTarget = value
        End Set
    End Property
    Public Property IonQual() As String
        Get
            Return strIonQual
        End Get
        Set(ByVal value As String)
            strIonQual = value
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
End Class
