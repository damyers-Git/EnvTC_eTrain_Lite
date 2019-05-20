Public Class RefBook
    Private strBook As String
    Private strBookPg As String
    Private strNum As String
    Private strSection As String
    Private strType As String
    Private strName As String
    Private strNote As String
    Private dtExpiration As Date

    Public Sub New()
        'Constructor
        strName = "XX-X0-0-0" 'Defaults
        dtExpiration = CDate("1/1/1970")  'Default date
        strNote = "NONE"
    End Sub
    Public Property Book() As String
        Get
            Return strBook
        End Get
        Set(ByVal value As String)
            strBook = value
        End Set
    End Property
    Public Property BookPg() As String
        Get
            Return strBookPg
        End Get
        Set(ByVal value As String)
            strBookPg = value
        End Set
    End Property
    Public Property Num() As String
        Get
            Return strNum
        End Get
        Set(ByVal value As String)
            strNum = value
        End Set
    End Property
    Public Property Section() As String
        Get
            Return strSection
        End Get
        Set(ByVal value As String)
            strSection = value
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
    Public Property Name() As String
        Get
            Return strName
        End Get
        Set(ByVal value As String)
            strName = value
        End Set
    End Property
    Public Property Note() As String
        Get
            Return strNote
        End Get
        Set(ByVal value As String)
            strNote = value
        End Set
    End Property
    Public Property Expiration() As Date
        Get
            Return dtExpiration
        End Get
        Set(ByVal value As Date)
            dtExpiration = value
        End Set
    End Property
End Class
