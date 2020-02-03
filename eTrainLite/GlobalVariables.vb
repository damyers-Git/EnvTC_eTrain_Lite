Imports Syncfusion.XlsIO

Public Class GlobalVariables
    Public Shared Property eTrain As New eTrain
    Public Shared Property Import As New Import
    Public Shared Property Transfer As New Transfer
    Public Shared Property Report As New Report
    Public Shared Property Method As New Method
    Public Shared Property Permit As New Permit
    Public Shared Property Project As New Project
    Public Shared Property mInstrument As New mInstrument
    Public Shared Property mCompound As New mCompound
    Public Shared Property mStandard As New mStandard
    Public Shared Property Calculations As New Calculations
    Public Shared Property CSImport As New CSImport
    Public Shared Property MHImport As New MHImport
    Public Shared Property TQ3Import As New TQ3Import
    Public Shared Property SIS As New SIS
    Public Shared Property RefBook As New RefBook
    Public Shared Property SISList As New ArrayList
    Public Shared Property SampleList As New ArrayList
    Public Shared Property ReportSamList As New ArrayList
    Public Shared Property TempReportSamList As New ArrayList
    Public Shared Property TheoComps As New ArrayList
    Public Shared Property MethodList As New ArrayList
    Public Shared Property PermitList As New ArrayList
    Public Shared Property Associated13Cs As New ArrayList
    Public Shared Property FreeportMBCompoundList As New ArrayList
    Public Shared Property MidlandMBCompoundList As New ArrayList
    Public Shared Property MidlandHRAvgAreaCompList As New ArrayList
    Public Shared Property MidlandChromRLimitNames As New ArrayList
    Public Shared Property compNameToCASDic As New Dictionary(Of String, String)
    Public Shared Property limsAnalysisMethod As New Dictionary(Of String, String)
    Public Shared Property eddAnalysisMethod As New Dictionary(Of String, String)
    Public Shared Property limsCompoundInformation As New ArrayList
    Public Shared Property methodNameAndUnits As New ArrayList
    Public Shared Property limsMethodNames As New Dictionary(Of String, String)
    Public Shared Property recoveryUnits As New List(Of List(Of String))
    Public Shared Property befAndTefScores As New List(Of List(Of String))
    Public Shared Property NeedsCalculation As Boolean
    Public Shared Property NeedsUnitConversion As Boolean
    Public Shared Property CustomReportError As Boolean
    Public Shared selMethod As Method
    Public Shared selPermit As Permit
    Public Shared selProject As String
    Public Shared selInstrument As String
    Public Shared workbook As IWorkbook
    Public Shared selLimit As String
    Public Shared selLimitPath As String
    Public Shared strFreeportAnalysis As String
    Public Shared ContinueTransfer As Boolean
    Public Shared ContinueReport As Boolean
    Public Shared ElutionOrderSample As Sample
End Class
