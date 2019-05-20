Public Class Project

    Private strName As String
    Private blnReviewed As Boolean
    Private dtReviewedDate As Date
    'Private strRL As String
    Private strDefRL As String
    Private strDefMDL As String
    Private strDefPQL As String
    Private strLimsUnits As String
    Public Property mInstrumentList As New ArrayList

    Public Sub New()
        'Constructor
        strLimsUnits = ""
    End Sub
    Public Property Name() As String
        Get
            Return strName
        End Get
        Set(ByVal value As String)
            strName = value
        End Set
    End Property
    Public Property Reviewed() As Boolean
        Get
            Return blnReviewed
        End Get
        Set(ByVal value As Boolean)
            blnReviewed = value
        End Set
    End Property
    Public Property ReviewedDate() As Date
        Get
            Return dtReviewedDate
        End Get
        Set(ByVal value As Date)
            dtReviewedDate = value
        End Set
    End Property
    Public Property DefRL() As String
        Get
            Return strDefRL
        End Get
        Set(ByVal value As String)
            strDefRL = value
        End Set
    End Property
    Public Property DefMDL() As String
        Get
            Return strDefMDL
        End Get
        Set(ByVal value As String)
            strDefMDL = value
        End Set
    End Property
    Public Property DefPQL() As String
        Get
            Return strDefPQL
        End Get
        Set(ByVal value As String)
            strDefPQL = value
        End Set
    End Property
    Public Property LimsUnits() As String
        Get
            Return strLimsUnits
        End Get
        Set(ByVal value As String)
            strLimsUnits = value
        End Set
    End Property
    'Public Property RL() As String
    '    Get
    '        Return strRL
    '    End Get
    '    Set(ByVal value As String)
    '        strRL = value
    '    End Set
    'End Property
End Class
