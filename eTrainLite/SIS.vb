Public Class SIS
    Private strProjNum As String
    Private strProjName As String
    Private strMethod As String
    Private strAnalysis As String
    Private strSampMatrix As String
    Private strCompliance As String
    Private strSetNum As String
    Private strContact As String
    Private strCostCenter As String
    Private dtStartDate As Date
    Private dtEndDate As Date
    Private strConfAnalysis As String
    Private strPrepAnalyst As String
    Private strExtraction As String
    Private strCleanUpCols As String
    Private strMethylation As String
    Private strAddAnalyses As String
    Private strAnalyst As String
    Private strInstrument As String
    Private strReviewer As String
    Private strTeam As String
    Private blnEOA As Boolean
    Private blnVOA As Boolean
    Private strCSMethod As String
    Public Property SampleList As New ArrayList

    Public Sub New()
        'Constructor

    End Sub

    'Sets/Gets
    Public Property ProjNum() As String
        Get
            Return strProjNum
        End Get
        Set(ByVal value As String)
            strProjNum = value
        End Set
    End Property
    Public Property ProjName() As String
        Get
            Return strProjName
        End Get
        Set(ByVal value As String)
            strProjName = value
        End Set
    End Property
    Public Property Method() As String
        Get
            Return strMethod
        End Get
        Set(ByVal value As String)
            strMethod = value
        End Set
    End Property
    Public Property Analysis() As String
        Get
            Return strAnalysis
        End Get
        Set(ByVal value As String)
            strAnalysis = value
        End Set
    End Property
    Public Property SampMatrix() As String
        Get
            Return strSampMatrix
        End Get
        Set(ByVal value As String)
            strSampMatrix = value
        End Set
    End Property
    Public Property Compliance() As String
        Get
            Return strCompliance
        End Get
        Set(ByVal value As String)
            strCompliance = value
        End Set
    End Property
    Public Property SetNum() As String
        Get
            Return strSetNum
        End Get
        Set(ByVal value As String)
            strSetNum = value
        End Set
    End Property
    Public Property Contact() As String
        Get
            Return strContact
        End Get
        Set(ByVal value As String)
            strContact = value
        End Set
    End Property
    Public Property CostCenter() As String
        Get
            Return strCostCenter
        End Get
        Set(ByVal value As String)
            strCostCenter = value
        End Set
    End Property
    Public Property StartDate() As Date
        Get
            Return dtStartDate
        End Get
        Set(ByVal value As Date)
            dtStartDate = value
        End Set
    End Property
    Public Property EndDate() As Date
        Get
            Return dtEndDate
        End Get
        Set(ByVal value As Date)
            dtEndDate = value
        End Set
    End Property
    Public Property ConfAnalysis() As String
        Get
            Return strConfAnalysis
        End Get
        Set(ByVal value As String)
            strConfAnalysis = value
        End Set
    End Property
    Public Property PrepAnalyst() As String
        Get
            Return strPrepAnalyst
        End Get
        Set(ByVal value As String)
            strPrepAnalyst = value
        End Set
    End Property
    Public Property Extraction() As String
        Get
            Return strExtraction
        End Get
        Set(ByVal value As String)
            strExtraction = value
        End Set
    End Property
    Public Property CleanUpCols() As String
        Get
            Return strCleanUpCols
        End Get
        Set(ByVal value As String)
            strCleanUpCols = value
        End Set
    End Property
    Public Property Methylation() As String
        Get
            Return strMethylation
        End Get
        Set(ByVal value As String)
            strMethylation = value
        End Set
    End Property
    Public Property AddAnalyses() As String
        Get
            Return strAddAnalyses
        End Get
        Set(ByVal value As String)
            strAddAnalyses = value
        End Set
    End Property
    Public Property Analyst() As String
        Get
            Return strAnalyst
        End Get
        Set(ByVal value As String)
            strAnalyst = value
        End Set
    End Property
    Public Property Instrument() As String
        Get
            Return strInstrument
        End Get
        Set(ByVal value As String)
            strInstrument = value
        End Set
    End Property
    Public Property Reviewer() As String
        Get
            Return strReviewer
        End Get
        Set(ByVal value As String)
            strReviewer = value
        End Set
    End Property
    Public Property Team() As String
        Get
            Return strTeam
        End Get
        Set(ByVal value As String)
            strTeam = value
        End Set
    End Property
    Public Property EOA() As Boolean
        Get
            Return blnEOA
        End Get
        Set(ByVal value As Boolean)
            blnEOA = value
        End Set
    End Property
    Public Property VOA() As Boolean
        Get
            Return blnVOA
        End Get
        Set(ByVal value As Boolean)
            blnVOA = value
        End Set
    End Property
    Public Property CSMethod() As String
        Get
            Return strCSMethod
        End Get
        Set(ByVal value As String)
            strCSMethod = value
        End Set
    End Property
End Class
