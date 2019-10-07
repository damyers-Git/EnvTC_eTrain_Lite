Imports System.IO

Public Class eTrain
    Private Property strTeam As String
    Private Property strLocation As String
    Private Property strServer As String
    Private Property strServerFP As String
    Private Property strDataFilesFP As String
    Private Property intSigFig As Integer

    'Const
    Public Sub New()
        strDataFilesFP = "\\Mdrnd\as-global\Special_Access\EAC\Data\eTrain\DataFiles\"
    End Sub

    'Sets/Gets
    Public Property Team() As String
        Get
            Return strTeam
        End Get
        Set(ByVal value As String)
            strTeam = value
        End Set
    End Property
    Public Property Location() As String
        Get
            Return strLocation
        End Get
        Set(ByVal value As String)
            strLocation = value
        End Set
    End Property
    Public Property Server() As String
        Get
            Return strServer
        End Get
        Set(ByVal value As String)
            strServer = value
        End Set
    End Property
    Public Property ServerFP() As String
        Get
            Return strServerFP
        End Get
        Set(ByVal value As String)
            strServerFP = value
        End Set
    End Property
    Public Property DataFileFP() As String
        Get
            Return strDataFilesFP
        End Get
        Set(value As String)
            strDataFilesFP = value
        End Set
    End Property
    Public Property SigFig() As Integer
        Get
            Return intSigFig
        End Get
        Set(ByVal value As Integer)
            intSigFig = value
        End Set
    End Property


    Function ChooseFolder(ByVal rootFolder As String, ByVal Desc As String) As String
        Dim dl As New FolderBrowserDialog()
        Dim f As String

        'Set starting root folder
        dl.RootFolder = Environment.SpecialFolder.Desktop
        dl.SelectedPath = rootFolder

        'Set Windows description
        If Desc = "" Then
            Desc = "Select a Folder"
        End If
        dl.Description = Desc

        If dl.ShowDialog = Windows.Forms.DialogResult.OK Then
            f = dl.SelectedPath
        Else
            f = "NULL"
        End If

        Return f
    End Function



End Class


