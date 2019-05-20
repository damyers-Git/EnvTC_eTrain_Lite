Public Class Sample

    Private intUniqueID As Integer
    Private blnQTReviewed As Boolean
    Private strDataPath As String
    Private strDataFile As String
    Private dtAcqDate As Date
    Private strAnalyst As String
    Private strAnalysis As String
    Private strName As String
    Private strMisc As String
    Private strLimsID As String
    Private dtSampleDate As Date
    Private strDilutionFactor As String
    Private strDetectLimitType As String
    Private strInstrument As String
    Private strMatrix As String
    Private strVial As String
    Private strMultiplier As String
    Private strQuantMethod As String
    Private strQuantTitle As String
    Private dtQLastUpdate As Date
    Private strResponseVia As String
    Private strQMethFile As String
    Private strStdSpikeAmt As String
    Private strInjSpikeAmt As String
    Private strLCSSpikeAmt As String
    Private strAliquot As String
    Private blnCalculated As Boolean
    Private strType As String
    Private blnInclude As Boolean
    Private strResult As String
    Private strUnits As String
    Private strReportedUnits As String
    Private strSignals As String
    Private strVolInj As String
    Private strSigPhase As String
    Private strSigInfo As String
    Private blnReported As Boolean
    Private blnParent As Boolean
    Private strSampDate As String
    Public Property InternalStdList As New ArrayList
    Public Property SurrogateList As New ArrayList
    Public Property CompoundList As New ArrayList
    Public Property CompoundNameList As New ArrayList

    '   'EDD File Values 
    '   Private strSysSampleCode As String 'Added (EDD values) WT 9/26/2017
    '   Private strLabAnlMethodName As String
    '   Private strAnalysisDate As String
    '   Private strAnalysisTime As String '?
    'Private strTotalOrDissolved As String
    'Private strColumnNumber As String
    'Private strTestType As String
    'Private strLabMatrixCode As String
    'Private strAnalysisLocation As String
    '   Private strBasis As String
    '   Private strSampleTypeCode As String
    '   'Private strContainerID As String 
    '   Private strEDilutionFactor As String
    'Private strPrepMethod As String
    '   Private strPrepDate As String
    '   Private strPrepTime As String
    'Private strLeachateMethod As String
    '   Private dtLeachateDate As String
    '   Private strLeachateTime As String
    'Private strLabNameCode As String
    'Private strQcLevel As String
    'Private strLabSampleID As String
    'Private strPercentMoisture As String
    'Private strSubsampleAmount As String
    'Private strSubsampleAmountUnit As String
    'Private strAnalystName As String
    '   Private strInstrumentID As String
    '   Private strComment As String
    '   Private strPreservative As String
    '   'Private strCommentPreservative As String
    '   Private strFinalVolume As String
    'Private strFinalVolumeUnit As String
    'Private strCasRn As String
    'Private strChemicalName As String
    'Private strResultValue As String
    'Private strResultErrorDelta As String
    'Private strResultTypeCode As String
    'Private strReportableResult As String
    'Private strDetectFlag As String 'Boolean?
    'Private strLabQualifiers As String
    'Private strValidatorQualifiers As String
    'Private strMethodDetectionLimit As String
    'Private strOrganicYn As String
    'Private strReportingDetectionLimit As String
    'Private strQuantitationLimt As String
    'Private strResultUnit As String
    'Private strDetectionLimitUnit As String
    'Private strTicRetentionTime As String
    'Private strResultComment As String
    'Private strQcOriginalConc As String
    'Private strQcSpikeAdded As String
    'Private strQcSpikeMeasured As String
    'Private strQcSpikeRecovery As String
    'Private strQcDupOriginalConc As String
    'Private strQcDupSpikeAdded As String
    'Private strQcDupSpikeMeasured As String
    'Private strQcDupSpikeRecovery As String
    'Private strQcRpd As String
    'Private strQcSpikeLcl As String
    'Private strQcSpikeUcl As String
    'Private strQcRpdCl As String
    'Private strQcSpikeStatus As String
    'Private strQcDupSpikeStatus As String
    'Private strQcRpdStatus As String
    'Private strRlOrMdl As String

    'TQ3 File Values
    Private strTQ3QuanFile As String
    Private strTQ3DataFile As String
    Private strTQ3ResponseFile As String
    Private strTQ3Entries As String
    Private strTQ3SampleID As String
    Private strTQ3Study As String
    Private strTQ3Client As String
    Private strTQ3Laboratory As String
    Private strTQ3Operator As String
    Private strTQ3Phone As String
    Private strTQ3Barcode As String
    Private strTQ3QUALCompatMode As String
    Private strTQ3InjectionVol As String
    Private strTQ3SampleVol As String
    Private strTQ3SampleWeight As String
    Private strTQ3DilutionFactor As String
    Private strTQ3DetLimitFactor As String
    Private strTQ3DisplayQuantStatusArea As String
    Private strTQ3DisplayQuantStatusHeight As String
    Private strTQ3SumQMRM1 As String
    Private strTQ3SumQMRM2 As String
    Private strTQ3SinglePointRF As String
    Private strTQ3AvgRF As String
    Private strTQ3RFvsArea As String
    Private strTQ3AreaRatiovsConc As String
    Private strTQ3LinearFit As String
    Private strTQ3SquareFit As String
    Private strTQ3NonWeightedRegress As String
    Private strTQ3RegressWeighted1Amt As String
    Private strTQ3RegressWeighted1Resp As String
    Private strTQ3WeightedRegressFactor As String

    'EPATEMP File Values
    Private dtQuantTime As Date

    'TMPQNTRP File Values
    Private strTMPQuantFile As String

    'FAST Values
    Private dblMidFInjAmt As Double
    Private strMidFETEQ0 As String
    Private strMidFETEQ05 As String
    Private strMidFETEQLOD As String

    'Chrom Values
    Private blnMethylated As Boolean
    Private strChromSpikeAmt As String
    Private blnSpikeCalculated As Boolean

    'Trace Values
    Private strAvgCalArea As String

    'SIS Values
    Private strSISInternalID As String
    Private strSISLabNum As String
    Private strSISClientSampID As String
    Private dtSISSampDate As Date
    Private dtSISSampDateEnd As Date
    Private strSISTargetSampSize As String
    Private strSISActualSampSize As String
    Private strSISDefaultAliquot As String
    Private strSISAnalyses As String
    Private strSISSpikeMult As String
    Private strSISDilFactor As String
    Private strSISFinalWeight As String
    Private strSISTinWeight As String
    Private strSISWetWeight As String
    Private strSISDryWeight As String
    Private strSISSampleWeight As String
    Private strSISSBottWeight As String
    Private strSISSampWetWeight As String
    Private strSISEBottWeight As String
    Private strSISPMoisture As String
    Private strSISType As String

    'CCCHECK File Values
    Private dtCCCQuantTime As Date
    Private strMinRRF As String
    Private strMinRelArea As String
    Private strMaxRTDev As String
    Private strMaxRRFDev As String
    Private strMaxRelArea As String

    Public Sub New()
        'Constructor
        blnCalculated = False
        strMidFETEQ0 = "0"
        strMidFETEQ05 = "0"
        strMidFETEQLOD = "0"
        dtSISSampDate = CDate("1/1/1970")
        dtSISSampDateEnd = CDate("1/1/1970")
        blnInclude = True
        blnMethylated = False
        blnReported = False
        blnSpikeCalculated = False
        intUniqueID = GlobalVariables.SampleList.Count
    End Sub

    'Sets/Gets
    'Public Property EDDsysSampleCode() As String 'Added (EDD's) WT 9/26/2017
    '    Get
    '        Return strSysSampleCode
    '    End Get
    '    Set(ByVal value As String)
    '        strSysSampleCode = value
    '    End Set
    'End Property

    'Public Property EDDLabAnlMethodName() As String
    '    Get
    '        Return strLabAnlMethodName
    '    End Get
    '    Set(ByVal value As String)
    '        strLabAnlMethodName = value
    '    End Set
    'End Property

    'Public Property EDDAnalysisDate() As String
    '    Get
    '        Return strAnalysisDate
    '    End Get
    '    Set(ByVal value As String)
    '        strAnalysisDate = value
    '    End Set
    'End Property

    'Public Property EDDAnalysisTime() As String
    '    Get
    '        Return strAnalysisTime
    '    End Get
    '    Set(ByVal value As String)
    '        strAnalysisTime = value
    '    End Set
    'End Property

    'Public Property EDDTotalOrDissolved() As String
    '    Get
    '        Return strTotalOrDissolved
    '    End Get
    '    Set(ByVal value As String)
    '        strTotalOrDissolved = value
    '    End Set
    'End Property

    'Public Property EDDColumnNumber() As String
    '    Get
    '        Return strColumnNumber
    '    End Get
    '    Set(ByVal value As String)
    '        strColumnNumber = value
    '    End Set
    'End Property

    'Public Property EDDTestType() As String
    '    Get
    '        Return strTestType
    '    End Get
    '    Set(ByVal value As String)
    '        strTestType = value
    '    End Set
    'End Property

    'Public Property EDDLabMatrixCode() As String
    '    Get
    '        Return strLabMatrixCode
    '    End Get
    '    Set(ByVal value As String)
    '        strLabMatrixCode = value
    '    End Set
    'End Property

    'Public Property EDDAnalysisLocation() As String
    '    Get
    '        Return strAnalysisLocation
    '    End Get
    '    Set(ByVal value As String)
    '        strAnalysisLocation = value
    '    End Set
    'End Property

    'Public Property EDDBasis() As String
    '    Get
    '        Return strBasis
    '    End Get
    '    Set(ByVal value As String)
    '        strBasis = value
    '    End Set
    'End Property

    'Public Property EDDSampleTypeCode() As String
    '    Get
    '        Return strSampleTypeCode
    '    End Get
    '    Set(ByVal value As String)
    '        strSampleTypeCode = value
    '    End Set
    'End Property

    ''Public Property EDDContainerID() As String
    ''    Get
    ''        Return strContainerID
    ''    End Get
    ''    Set(ByVal value As String)
    ''        strContainerID = value
    ''    End Set
    ''End Property

    'Public Property EDDEDilutionFactor() As String
    '    Get
    '        Return strEDilutionFactor
    '    End Get
    '    Set(ByVal value As String)
    '        strEDilutionFactor = value
    '    End Set
    'End Property

    'Public Property EDDPrepMethod() As String
    '    Get
    '        Return strPrepMethod
    '    End Get
    '    Set(ByVal value As String)
    '        strPrepMethod = value
    '    End Set
    'End Property

    'Public Property EDDPrepDate() As String
    '    Get
    '        Return strPrepDate
    '    End Get
    '    Set(ByVal value As String)
    '        strPrepDate = value
    '    End Set
    'End Property

    'Public Property EDDPrepTime() As String
    '    Get
    '        Return strPrepTime
    '    End Get
    '    Set(ByVal value As String)
    '        strPrepTime = value
    '    End Set
    'End Property

    'Public Property EDDLeachateMethod() As String
    '    Get
    '        Return strLeachateMethod
    '    End Get
    '    Set(ByVal value As String)
    '        strLeachateMethod = value
    '    End Set
    'End Property

    'Public Property EDDLeachateDate() As String

    'Public Property EDDLeachateTime() As String
    '    Get
    '        Return strLeachateTime
    '    End Get
    '    Set(ByVal value As String)
    '        strLeachateTime = value
    '    End Set
    'End Property

    'Public Property EDDLabNameCode() As String
    '    Get
    '        Return strLabNameCode
    '    End Get
    '    Set(ByVal value As String)
    '        strLabNameCode = value
    '    End Set
    'End Property

    'Public Property EDDQcLevel() As String
    '    Get
    '        Return strQcLevel
    '    End Get
    '    Set(value As String)
    '        strQcLevel = value
    '    End Set
    'End Property

    'Public Property EDDLabSampleID() As String
    '    Get
    '        Return strLabSampleID
    '    End Get
    '    Set(ByVal value As String)
    '        strLabSampleID = value
    '    End Set
    'End Property

    'Public Property EDDPercentMoisture() As String
    '    Get
    '        Return strPercentMoisture
    '    End Get
    '    Set(ByVal value As String)
    '        strPercentMoisture = value
    '    End Set
    'End Property

    'Public Property EDDSubsampleAmount() As String
    '    Get
    '        Return strSubsampleAmount
    '    End Get
    '    Set(ByVal value As String)
    '        strSubsampleAmount = value
    '    End Set
    'End Property

    'Public Property EDDSubsampleAmountUnit() As String
    '    Get
    '        Return strSubsampleAmountUnit
    '    End Get
    '    Set(ByVal value As String)
    '        strSubsampleAmountUnit = value
    '    End Set
    'End Property

    'Public Property EDDAnalystName() As String
    '    Get
    '        Return strAnalystName
    '    End Get
    '    Set(ByVal value As String)
    '        strAnalystName = value
    '    End Set
    'End Property

    'Public Property EDDInstrumentID() As String
    '    Get
    '        Return strInstrumentID
    '    End Get
    '    Set(ByVal value As String)
    '        strInstrumentID = value
    '    End Set
    'End Property

    'Public Property EDDComment() As String
    '    Get
    '        Return strComment
    '    End Get
    '    Set(ByVal value As String)
    '        strComment = value
    '    End Set
    'End Property

    'Public Property EDDPreservative() As String
    '    Get
    '        Return strPreservative
    '    End Get
    '    Set(ByVal value As String)
    '        strPreservative = value
    '    End Set
    'End Property


    ''Public Property EDDCommentPreservative() As String
    ''    Get
    ''        Return strCommentPreservative
    ''    End Get
    ''    Set(ByVal value As String)
    ''        strCommentPreservative = value
    ''    End Set
    ''End Property

    'Public Property EDDFinalVolume() As String
    '    Get
    '        Return strFinalVolume
    '    End Get
    '    Set(ByVal value As String)
    '        strFinalVolume = value
    '    End Set
    'End Property

    'Public Property EDDFinalVolumeUnit() As String
    '    Get
    '        Return strFinalVolumeUnit
    '    End Get
    '    Set(ByVal value As String)
    '        strFinalVolumeUnit = value
    '    End Set
    'End Property

    'Public Property EDDCasRn() As String
    '    Get
    '        Return strCasRn
    '    End Get
    '    Set(ByVal value As String)
    '        strCasRn = value
    '    End Set
    'End Property

    'Public Property EDDChemicalName() As String
    '    Get
    '        Return strChemicalName
    '    End Get
    '    Set(ByVal value As String)
    '        strChemicalName = value
    '    End Set
    'End Property

    'Public Property EDDResultValue() As String
    '    Get
    '        Return strResultValue
    '    End Get
    '    Set(ByVal value As String)
    '        strResultValue = value
    '    End Set
    'End Property


    'Public Property EDDResultErrorDelta() As String
    '    Get
    '        Return strResultErrorDelta
    '    End Get
    '    Set(ByVal value As String)
    '        strResultErrorDelta = value
    '    End Set
    'End Property

    'Public Property EDDResultTypeCode() As String
    '    Get
    '        Return strResultTypeCode
    '    End Get
    '    Set(ByVal value As String)
    '        strResultTypeCode = value
    '    End Set
    'End Property

    'Public Property EDDReportableResult() As String
    '    Get
    '        Return strReportableResult
    '    End Get
    '    Set(ByVal value As String)
    '        strReportableResult = value
    '    End Set
    'End Property

    'Public Property EDDDetectFlag() As String 'Boolean?
    '    Get
    '        Return strDetectFlag
    '    End Get
    '    Set(ByVal value As String)
    '        strDetectFlag = value
    '    End Set
    'End Property

    'Public Property EDDLabQualifiers() As String
    '    Get
    '        Return strLabQualifiers
    '    End Get
    '    Set(ByVal value As String)
    '        strLabQualifiers = value
    '    End Set
    'End Property

    'Public Property EDDValidatorQualifiers() As String
    '    Get
    '        Return strValidatorQualifiers
    '    End Get
    '    Set(ByVal value As String)
    '        strValidatorQualifiers = value
    '    End Set
    'End Property

    'Public Property EDDMethodDetectionLimit() As String
    '    Get
    '        Return strMethodDetectionLimit
    '    End Get
    '    Set(ByVal value As String)
    '        strMethodDetectionLimit = value
    '    End Set
    'End Property

    'Public Property EDDOrganicYn() As String
    '    Get
    '        Return strOrganicYn
    '    End Get
    '    Set(ByVal value As String)
    '        strOrganicYn = value
    '    End Set
    'End Property

    'Public Property EDDReportingDetectionLimit() As String
    '    Get
    '        Return strReportingDetectionLimit
    '    End Get
    '    Set(ByVal value As String)
    '        strReportingDetectionLimit = value
    '    End Set
    'End Property

    'Public Property EDDQuantitationLimit() As String
    '    Get
    '        Return strQuantitationLimt
    '    End Get
    '    Set(value As String)
    '        strQuantitationLimt = value
    '    End Set
    'End Property

    'Public Property EDDResultUnit() As String
    '    Get
    '        Return strResultUnit
    '    End Get
    '    Set(ByVal value As String)
    '        strResultUnit = value
    '    End Set
    'End Property

    'Public Property EDDDetectionLimitUnit() As String
    '    Get
    '        Return strDetectionLimitUnit
    '    End Get
    '    Set(ByVal value As String)
    '        strDetectionLimitUnit = value
    '    End Set
    'End Property

    'Public Property EDDTicRetentionTime() As String
    '    Get
    '        Return strTicRetentionTime
    '    End Get
    '    Set(ByVal value As String)
    '        strTicRetentionTime = value
    '    End Set
    'End Property

    'Public Property EDDResultComment() As String
    '    Get
    '        Return strResultComment
    '    End Get
    '    Set(ByVal value As String)
    '        strResultComment = value
    '    End Set
    'End Property

    'Public Property EDDQcOriginalConc() As String
    '    Get
    '        Return strQcOriginalConc
    '    End Get
    '    Set(ByVal value As String)
    '        strQcOriginalConc = value
    '    End Set
    'End Property

    'Public Property EDDQcSpikeAdded() As String
    '    Get
    '        Return strQcSpikeAdded
    '    End Get
    '    Set(ByVal value As String)
    '        strQcSpikeAdded = value
    '    End Set
    'End Property

    'Public Property EDDQcSpikeMeasured() As String
    '    Get
    '        Return strQcSpikeMeasured
    '    End Get
    '    Set(value As String)
    '        strQcSpikeMeasured = value
    '    End Set
    'End Property

    'Public Property EDDQcSpikeRecovery() As String
    '    Get
    '        Return strQcSpikeRecovery
    '    End Get
    '    Set(ByVal value As String)
    '        strQcSpikeRecovery = value
    '    End Set
    'End Property

    'Public Property EDDQcDupOriginalConc() As String
    '    Get
    '        Return strQcDupOriginalConc
    '    End Get
    '    Set(ByVal value As String)
    '        strQcDupOriginalConc = value
    '    End Set
    'End Property

    'Public Property EDDQcDupSpikeAdded() As String
    '    Get
    '        Return strQcDupSpikeAdded
    '    End Get
    '    Set(ByVal value As String)
    '        strQcDupSpikeAdded = value
    '    End Set
    'End Property

    'Public Property EDDQcDupSpikeMeasured() As String
    '    Get
    '        Return strQcDupSpikeMeasured
    '    End Get
    '    Set(ByVal value As String)
    '        strQcDupSpikeMeasured = value
    '    End Set
    'End Property

    'Public Property EDDQcDupSpikeRecovery() As String
    '    Get
    '        Return strQcDupSpikeRecovery
    '    End Get
    '    Set(ByVal value As String)
    '        strQcDupSpikeRecovery = value
    '    End Set
    'End Property

    'Public Property EDDQcRpd() As String
    '    Get
    '        Return strQcRpd
    '    End Get
    '    Set(ByVal value As String)
    '        strQcRpd = value
    '    End Set
    'End Property

    'Public Property EDDQcSpikeLcl() As String
    '    Get
    '        Return strQcSpikeLcl
    '    End Get
    '    Set(ByVal value As String)
    '        strQcSpikeLcl = value
    '    End Set
    'End Property

    'Public Property EDDQcSpikeUcl() As String
    '    Get
    '        Return strQcSpikeUcl
    '    End Get
    '    Set(ByVal value As String)
    '        strQcSpikeUcl = value
    '    End Set
    'End Property

    'Public Property EDDQcRpdCl() As String
    '    Get
    '        Return strQcRpdCl
    '    End Get
    '    Set(ByVal value As String)
    '        strQcRpdCl = value
    '    End Set
    'End Property

    'Public Property EDDQcSpikeStatus() As String
    '    Get
    '        Return strQcSpikeStatus
    '    End Get
    '    Set(ByVal value As String)
    '        strQcSpikeStatus = value
    '    End Set
    'End Property

    'Public Property EDDQcDupSpikeStatus() As String
    '    Get
    '        Return strQcDupSpikeStatus
    '    End Get
    '    Set(ByVal value As String)
    '        strQcDupSpikeStatus = value
    '    End Set
    'End Property

    'Public Property EDDQcRpdStatus() As String
    '    Get
    '        Return strQcRpdStatus
    '    End Get
    '    Set(ByVal value As String)
    '        strQcRpdStatus = value
    '    End Set
    'End Property

    'Public Property EDDRlOrMdl() As String
    '    Get
    '        Return strRlOrMdl
    '    End Get
    '    Set(ByVal value As String)
    '        strRlOrMdl = value
    '    End Set
    'End Property '<- End add WT

    Public Property SampDate() As String
        Get
            Return strSampDate
        End Get
        Set(ByVal value As String)
            strSampDate = value
        End Set
    End Property

    Public Property UniqueID() As Integer
        Get
            Return intUniqueID
        End Get
        Set(ByVal value As Integer)
            intUniqueID = value
        End Set
    End Property
    Public Property QTReviewed() As Boolean
        Get
            Return blnQTReviewed
        End Get
        Set(ByVal value As Boolean)
            blnQTReviewed = value
        End Set
    End Property
    Public Property DataPath() As String
        Get
            Return strDataPath
        End Get
        Set(ByVal value As String)
            strDataPath = value
        End Set
    End Property
    Public Property DataFile() As String
        Get
            Return strDataFile
        End Get
        Set(ByVal value As String)
            strDataFile = value
        End Set
    End Property
    Public Property Result() As String
        Get
            Return strResult
        End Get
        Set(ByVal value As String)
            strResult = value
        End Set
    End Property
    Public Property AcqDate() As Date
        Get
            Return dtAcqDate
        End Get
        Set(ByVal value As Date)
            dtAcqDate = value
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
    Public Property Analysis() As String
        Get
            Return strAnalysis
        End Get
        Set(ByVal value As String)
            strAnalysis = value
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
    Public Property Misc() As String
        Get
            Return strMisc
        End Get
        Set(ByVal value As String)
            strMisc = value
        End Set
    End Property
    Public Property LimsID() As String
        Get
            Return strLimsID
        End Get
        Set(ByVal value As String)
            strLimsID = value
        End Set
    End Property
    Public Property SampleDate() As Date
        Get
            Return dtSampleDate
        End Get
        Set(ByVal value As Date)
            dtSampleDate = value
        End Set
    End Property
    Public Property DilutionFactor() As String
        Get
            Return strDilutionFactor
        End Get
        Set(ByVal value As String)
            strDilutionFactor = value
        End Set
    End Property
    Public Property DetectLimitType() As String
        Get
            Return strDetectLimitType
        End Get
        Set(ByVal value As String)
            strDetectLimitType = value
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
    Public Property Matrix() As String
        Get
            Return strMatrix
        End Get
        Set(ByVal value As String)
            strMatrix = value
        End Set
    End Property
    Public Property Vial() As String
        Get
            Return strVial
        End Get
        Set(ByVal value As String)
            strVial = value
        End Set
    End Property
    Public Property Multiplier() As String
        Get
            Return strMultiplier
        End Get
        Set(ByVal value As String)
            strMultiplier = value
        End Set
    End Property
    Public Property QuantTime() As Date
        Get
            Return dtQuantTime
        End Get
        Set(ByVal value As Date)
            dtQuantTime = value
        End Set
    End Property
    Public Property QuantMethod() As String
        Get
            Return strQuantMethod
        End Get
        Set(ByVal value As String)
            strQuantMethod = value
        End Set
    End Property
    Public Property QuantTitle() As String
        Get
            Return strQuantTitle
        End Get
        Set(ByVal value As String)
            strQuantTitle = value
        End Set
    End Property
    Public Property QLastUpdate() As Date
        Get
            Return dtQLastUpdate
        End Get
        Set(ByVal value As Date)
            dtQLastUpdate = value
        End Set
    End Property
    Public Property ResponseVia() As String
        Get
            Return strResponseVia
        End Get
        Set(ByVal value As String)
            strResponseVia = value
        End Set
    End Property
    Public Property QMethFile() As String
        Get
            Return strQMethFile
        End Get
        Set(ByVal value As String)
            strQMethFile = value
        End Set
    End Property
    Public Property CCCQuantTime() As Date
        Get
            Return dtCCCQuantTime
        End Get
        Set(ByVal value As Date)
            dtCCCQuantTime = value
        End Set
    End Property
    Public Property MinRRF() As String
        Get
            Return strMinRRF
        End Get
        Set(ByVal value As String)
            strMinRRF = value
        End Set
    End Property
    Public Property MinRelArea() As String
        Get
            Return strMinRelArea
        End Get
        Set(ByVal value As String)
            strMinRelArea = value
        End Set
    End Property
    Public Property MaxRTDev() As String
        Get
            Return strMaxRTDev
        End Get
        Set(ByVal value As String)
            strMaxRTDev = value
        End Set
    End Property
    Public Property MaxRRFDev() As String
        Get
            Return strMaxRRFDev
        End Get
        Set(ByVal value As String)
            strMaxRRFDev = value
        End Set
    End Property
    Public Property MaxRelArea() As String
        Get
            Return strMaxRelArea
        End Get
        Set(ByVal value As String)
            strMaxRelArea = value
        End Set
    End Property
    Public Property TMPQuantFile() As String
        Get
            Return strTMPQuantFile
        End Get
        Set(ByVal value As String)
            strTMPQuantFile = value
        End Set
    End Property
    Public Property StdSpikeAmt() As String
        Get
            Return strStdSpikeAmt
        End Get
        Set(ByVal value As String)
            strStdSpikeAmt = value
        End Set
    End Property
    Public Property InjSpikeAmt() As String
        Get
            Return strInjSpikeAmt
        End Get
        Set(ByVal value As String)
            strInjSpikeAmt = value
        End Set
    End Property
    Public Property LCSSpikeAmt() As String
        Get
            Return strLCSSpikeAmt
        End Get
        Set(ByVal value As String)
            strLCSSpikeAmt = value
        End Set
    End Property
    Public Property Aliquot() As String
        Get
            Return strAliquot
        End Get
        Set(ByVal value As String)
            strAliquot = value
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
    Public Property Parent() As Boolean
        Get
            Return blnParent
        End Get
        Set(ByVal value As Boolean)
            blnParent = value
        End Set
    End Property
    Public Property SpikeCalculated() As Boolean
        Get
            Return blnSpikeCalculated
        End Get
        Set(ByVal value As Boolean)
            blnSpikeCalculated = value
        End Set
    End Property
    Public Property MidFInjAmt() As Double
        Get
            Return dblMidFInjAmt
        End Get
        Set(ByVal value As Double)
            dblMidFInjAmt = value
        End Set
    End Property
    Public Property SISInternalID() As String
        Get
            Return strSISInternalID
        End Get
        Set(ByVal value As String)
            strSISInternalID = value
        End Set
    End Property
    Public Property SISLabNum() As String
        Get
            Return strSISLabNum
        End Get
        Set(ByVal value As String)
            strSISLabNum = value
        End Set
    End Property
    Public Property SISClientSampID() As String
        Get
            Return strSISClientSampID
        End Get
        Set(ByVal value As String)
            strSISClientSampID = value
        End Set
    End Property
    Public Property SISSampDate() As Date
        Get
            Return dtSISSampDate
        End Get
        Set(ByVal value As Date)
            dtSISSampDate = value
        End Set
    End Property
    Public Property SISSampDateEnd() As Date
        Get
            Return dtSISSampDateEnd
        End Get
        Set(ByVal value As Date)
            dtSISSampDateEnd = value
        End Set
    End Property
    Public Property SISTargetSampSize() As String
        Get
            Return strSISTargetSampSize
        End Get
        Set(ByVal value As String)
            strSISTargetSampSize = value
        End Set
    End Property
    Public Property SISActualSampSize() As String
        Get
            Return strSISActualSampSize
        End Get
        Set(ByVal value As String)
            strSISActualSampSize = value
        End Set
    End Property
    Public Property SISDefaultAliquot() As String
        Get
            Return strSISDefaultAliquot
        End Get
        Set(ByVal value As String)
            strSISDefaultAliquot = value
        End Set
    End Property
    Public Property SISAnalyses() As String
        Get
            Return strSISAnalyses
        End Get
        Set(ByVal value As String)
            strSISAnalyses = value
        End Set
    End Property
    Public Property SISSpikeMult() As String
        Get
            Return strSISSpikeMult
        End Get
        Set(ByVal value As String)
            strSISSpikeMult = value
        End Set
    End Property
    Public Property SISDilFactor() As String
        Get
            Return strSISDilFactor
        End Get
        Set(ByVal value As String)
            strSISDilFactor = value
        End Set
    End Property
    Public Property SISFinalWeight() As String
        Get
            Return strSISFinalWeight
        End Get
        Set(ByVal value As String)
            strSISFinalWeight = value
        End Set
    End Property
    Public Property SISTinWeight() As String
        Get
            Return strSISTinWeight
        End Get
        Set(ByVal value As String)
            strSISTinWeight = value
        End Set
    End Property
    Public Property SISWetWeight() As String
        Get
            Return strSISWetWeight
        End Get
        Set(ByVal value As String)
            strSISWetWeight = value
        End Set
    End Property
    Public Property SISDryWeight() As String
        Get
            Return strSISDryWeight
        End Get
        Set(ByVal value As String)
            strSISDryWeight = value
        End Set
    End Property
    Public Property SISSampleWeight() As String
        Get
            Return strSISSampleWeight
        End Get
        Set(ByVal value As String)
            strSISSampleWeight = value
        End Set
    End Property
    Public Property SISSBottWeight() As String
        Get
            Return strSISSBottWeight
        End Get
        Set(ByVal value As String)
            strSISSBottWeight = value
        End Set
    End Property
    Public Property SISSampWetWeight() As String
        Get
            Return strSISSampWetWeight
        End Get
        Set(ByVal value As String)
            strSISSampWetWeight = value
        End Set
    End Property
    Public Property SISEBottWeight() As String
        Get
            Return strSISEBottWeight
        End Get
        Set(ByVal value As String)
            strSISEBottWeight = value
        End Set
    End Property
    Public Property SISPMoisture() As String
        Get
            Return strSISPMoisture
        End Get
        Set(ByVal value As String)
            strSISPMoisture = value
        End Set
    End Property
    Public Property SISType() As String
        Get
            Return strSISType
        End Get
        Set(ByVal value As String)
            strSISType = value
        End Set
    End Property
    Public Property MidFETEQ0() As String
        Get
            Return strMidFETEQ0
        End Get
        Set(ByVal value As String)
            strMidFETEQ0 = value
        End Set
    End Property
    Public Property MidFETEQ05() As String
        Get
            Return strMidFETEQ05
        End Get
        Set(ByVal value As String)
            strMidFETEQ05 = value
        End Set
    End Property
    Public Property MidFETEQLOD() As String
        Get
            Return strMidFETEQLOD
        End Get
        Set(ByVal value As String)
            strMidFETEQLOD = value
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
    Public Property Include() As Boolean
        Get
            Return blnInclude
        End Get
        Set(ByVal value As Boolean)
            blnInclude = value
        End Set
    End Property
    Public Property Reported() As Boolean
        Get
            Return blnReported
        End Get
        Set(ByVal value As Boolean)
            blnReported = value
        End Set
    End Property
    Public Property Methylated() As Boolean
        Get
            Return blnMethylated
        End Get
        Set(ByVal value As Boolean)
            blnMethylated = value
        End Set
    End Property
    Public Property ChromSpikeAmt() As String
        Get
            Return strChromSpikeAmt
        End Get
        Set(ByVal value As String)
            strChromSpikeAmt = value
        End Set
    End Property
    Public Property Signals() As String
        Get
            Return strSignals
        End Get
        Set(ByVal value As String)
            strSignals = value
        End Set
    End Property
    Public Property VolInj() As String
        Get
            Return strVolInj
        End Get
        Set(ByVal value As String)
            strVolInj = value
        End Set
    End Property
    Public Property SigPhase() As String
        Get
            Return strSigPhase
        End Get
        Set(ByVal value As String)
            strSigPhase = value
        End Set
    End Property
    Public Property SigInfo() As String
        Get
            Return strSigInfo
        End Get
        Set(ByVal value As String)
            strSigInfo = value
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
    Public Property ReportedUnits() As String
        Get
            Return strReportedUnits
        End Get
        Set(ByVal value As String)
            strReportedUnits = value
        End Set
    End Property
    Public Property TQ3QuanFile() As String
        Get
            Return strTQ3QuanFile
        End Get
        Set(ByVal value As String)
            strTQ3QuanFile = value
        End Set
    End Property
    Public Property TQ3DataFile() As String
        Get
            Return strTQ3DataFile
        End Get
        Set(ByVal value As String)
            strTQ3DataFile = value
        End Set
    End Property
    Public Property TQ3ResponseFile() As String
        Get
            Return strTQ3ResponseFile
        End Get
        Set(ByVal value As String)
            strTQ3ResponseFile = value
        End Set
    End Property
    Public Property TQ3Entries() As String
        Get
            Return strTQ3Entries
        End Get
        Set(ByVal value As String)
            strTQ3Entries = value
        End Set
    End Property
    Public Property TQ3SampleID() As String
        Get
            Return strTQ3SampleID
        End Get
        Set(ByVal value As String)
            strTQ3SampleID = value
        End Set
    End Property
    Public Property TQ3Study() As String
        Get
            Return strTQ3Study
        End Get
        Set(ByVal value As String)
            strTQ3Study = value
        End Set
    End Property
    Public Property TQ3Client() As String
        Get
            Return strTQ3Client
        End Get
        Set(ByVal value As String)
            strTQ3Client = value
        End Set
    End Property
    Public Property TQ3Laboratory() As String
        Get
            Return strTQ3Laboratory
        End Get
        Set(ByVal value As String)
            strTQ3Laboratory = value
        End Set
    End Property
    Public Property TQ3Operator() As String
        Get
            Return strTQ3Operator
        End Get
        Set(ByVal value As String)
            strTQ3Operator = value
        End Set
    End Property
    Public Property TQ3Phone() As String
        Get
            Return strTQ3Phone
        End Get
        Set(ByVal value As String)
            strTQ3Phone = value
        End Set
    End Property
    Public Property TQ3Barcode() As String
        Get
            Return strTQ3Barcode
        End Get
        Set(ByVal value As String)
            strTQ3Barcode = value
        End Set
    End Property
    Public Property TQ3QUALCompatMode() As String
        Get
            Return strTQ3QUALCompatMode
        End Get
        Set(ByVal value As String)
            strTQ3QUALCompatMode = value
        End Set
    End Property
    Public Property TQ3InjectionVol() As String
        Get
            Return strTQ3InjectionVol
        End Get
        Set(ByVal value As String)
            strTQ3InjectionVol = value
        End Set
    End Property
    Public Property TQ3SampleVol() As String
        Get
            Return strTQ3SampleVol
        End Get
        Set(ByVal value As String)
            strTQ3SampleVol = value
        End Set
    End Property
    Public Property TQ3SampleWeight() As String
        Get
            Return strTQ3SampleWeight
        End Get
        Set(ByVal value As String)
            strTQ3SampleWeight = value
        End Set
    End Property
    Public Property TQ3DilutionFactor() As String
        Get
            Return strTQ3DilutionFactor
        End Get
        Set(ByVal value As String)
            strTQ3DilutionFactor = value
        End Set
    End Property
    Public Property TQ3DetLimitFactor() As String
        Get
            Return strTQ3DetLimitFactor
        End Get
        Set(ByVal value As String)
            strTQ3DetLimitFactor = value
        End Set
    End Property
    Public Property TQ3DisplayQuantStatusArea() As String
        Get
            Return strTQ3DisplayQuantStatusArea
        End Get
        Set(ByVal value As String)
            strTQ3DisplayQuantStatusArea = value
        End Set
    End Property
    Public Property TQ3DisplayQuantStatusHeight() As String
        Get
            Return strTQ3DisplayQuantStatusHeight
        End Get
        Set(ByVal value As String)
            strTQ3DisplayQuantStatusHeight = value
        End Set
    End Property
    Public Property TQ3SumQMRM1() As String
        Get
            Return strTQ3SumQMRM1
        End Get
        Set(ByVal value As String)
            strTQ3SumQMRM1 = value
        End Set
    End Property
    Public Property TQ3SumQMRM2() As String
        Get
            Return strTQ3SumQMRM2
        End Get
        Set(ByVal value As String)
            strTQ3SumQMRM2 = value
        End Set
    End Property
    Public Property TQ3SinglePointRF() As String
        Get
            Return strTQ3SinglePointRF
        End Get
        Set(ByVal value As String)
            strTQ3SinglePointRF = value
        End Set
    End Property
    Public Property TQ3AvgRF() As String
        Get
            Return strTQ3AvgRF
        End Get
        Set(ByVal value As String)
            strTQ3AvgRF = value
        End Set
    End Property
    Public Property TQ3RFvsArea() As String
        Get
            Return strTQ3RFvsArea
        End Get
        Set(ByVal value As String)
            strTQ3RFvsArea = value
        End Set
    End Property
    Public Property TQ3AreaRatiovsConc() As String
        Get
            Return strTQ3AreaRatiovsConc
        End Get
        Set(ByVal value As String)
            strTQ3AreaRatiovsConc = value
        End Set
    End Property
    Public Property TQ3LinearFit() As String
        Get
            Return strTQ3LinearFit
        End Get
        Set(ByVal value As String)
            strTQ3LinearFit = value
        End Set
    End Property
    Public Property TQ3SquareFit() As String
        Get
            Return strTQ3SquareFit
        End Get
        Set(ByVal value As String)
            strTQ3SquareFit = value
        End Set
    End Property
    Public Property TQ3NonWeightedRegress() As String
        Get
            Return strTQ3NonWeightedRegress
        End Get
        Set(ByVal value As String)
            strTQ3NonWeightedRegress = value
        End Set
    End Property
    Public Property TQ3RegressWeighted1Amt() As String
        Get
            Return strTQ3RegressWeighted1Amt
        End Get
        Set(ByVal value As String)
            strTQ3RegressWeighted1Amt = value
        End Set
    End Property
    Public Property TQ3RegressWeighted1Resp() As String
        Get
            Return strTQ3RegressWeighted1Resp
        End Get
        Set(ByVal value As String)
            strTQ3RegressWeighted1Resp = value
        End Set
    End Property
    Public Property TQ3WeightedRegressFactor() As String
        Get
            Return strTQ3WeightedRegressFactor
        End Get
        Set(ByVal value As String)
            strTQ3WeightedRegressFactor = value
        End Set
    End Property
    Public Property AvgCalArea() As String
        Get
            Return strAvgCalArea
        End Get
        Set(ByVal value As String)
            strAvgCalArea = value
        End Set
    End Property

    Public Sub CopyMerge(ByVal aNewSample As Sample)
        aNewSample.QTReviewed = blnQTReviewed
        aNewSample.DataPath = strDataPath
        aNewSample.DataFile = strDataFile
        aNewSample.AcqDate = dtAcqDate
        aNewSample.Analyst = strAnalyst
        aNewSample.Analysis = strAnalysis
        aNewSample.Name = strName
        aNewSample.Misc = strMisc
        aNewSample.LimsID = strLimsID
        aNewSample.SampleDate = dtSampleDate
        aNewSample.DilutionFactor = strDilutionFactor
        aNewSample.DetectLimitType = strDetectLimitType
        aNewSample.Instrument = strInstrument
        aNewSample.Matrix = strMatrix
        aNewSample.Vial = strVial
        aNewSample.Multiplier = strMultiplier
        aNewSample.QuantMethod = strQuantMethod
        aNewSample.QuantTitle = strQuantTitle
        aNewSample.QLastUpdate = dtQLastUpdate
        aNewSample.ResponseVia = strResponseVia
        aNewSample.QMethFile = strQMethFile
        aNewSample.StdSpikeAmt = strStdSpikeAmt
        aNewSample.InjSpikeAmt = strInjSpikeAmt
        aNewSample.LCSSpikeAmt = strLCSSpikeAmt
        aNewSample.Aliquot = strAliquot
        aNewSample.Calculated = blnCalculated
        aNewSample.Type = strType
        aNewSample.Include = blnInclude
        aNewSample.Result = strResult
        aNewSample.Units = strUnits
        aNewSample.ReportedUnits = strReportedUnits
        aNewSample.Signals = strSignals
        aNewSample.VolInj = strVolInj
        aNewSample.SigPhase = strSigPhase
        aNewSample.SigInfo = strSigInfo
        aNewSample.Reported = blnReported
        aNewSample.Parent = blnParent


        'TQ3 File Values
        aNewSample.TQ3QuanFile = strTQ3QuanFile
        aNewSample.TQ3DataFile = strTQ3DataFile
        aNewSample.TQ3ResponseFile = strTQ3ResponseFile
        aNewSample.TQ3Entries = strTQ3Entries
        aNewSample.TQ3SampleID = strTQ3SampleID
        aNewSample.TQ3Study = strTQ3Study
        aNewSample.TQ3Client = strTQ3Client
        aNewSample.TQ3Laboratory = strTQ3Laboratory
        aNewSample.TQ3Operator = strTQ3Operator
        aNewSample.TQ3Phone = strTQ3Phone
        aNewSample.TQ3Barcode = strTQ3Barcode
        aNewSample.TQ3QUALCompatMode = strTQ3QUALCompatMode
        aNewSample.TQ3InjectionVol = strTQ3InjectionVol
        aNewSample.TQ3SampleVol = strTQ3SampleVol
        aNewSample.TQ3SampleWeight = strTQ3SampleWeight
        aNewSample.TQ3DilutionFactor = strTQ3DilutionFactor
        aNewSample.TQ3DetLimitFactor = strTQ3DetLimitFactor
        aNewSample.TQ3DisplayQuantStatusArea = strTQ3DisplayQuantStatusArea
        aNewSample.TQ3DisplayQuantStatusHeight = strTQ3DisplayQuantStatusHeight
        aNewSample.TQ3SumQMRM1 = strTQ3SumQMRM1
        aNewSample.TQ3SumQMRM2 = strTQ3SumQMRM2
        aNewSample.TQ3SinglePointRF = strTQ3SinglePointRF
        aNewSample.TQ3AvgRF = strTQ3AvgRF
        aNewSample.TQ3RFvsArea = strTQ3RFvsArea
        aNewSample.TQ3AreaRatiovsConc = strTQ3AreaRatiovsConc
        aNewSample.TQ3LinearFit = strTQ3LinearFit
        aNewSample.TQ3SquareFit = strTQ3SquareFit
        aNewSample.TQ3NonWeightedRegress = strTQ3NonWeightedRegress
        aNewSample.TQ3RegressWeighted1Amt = strTQ3RegressWeighted1Amt
        aNewSample.TQ3RegressWeighted1Resp = strTQ3RegressWeighted1Resp
        aNewSample.TQ3WeightedRegressFactor = strTQ3WeightedRegressFactor

        'EPATEMP File Values
        aNewSample.QuantTime = dtQuantTime

        'TMPQNTRP File Values
        aNewSample.TMPQuantFile = strTMPQuantFile

        'FAST Values
        aNewSample.MidFInjAmt = dblMidFInjAmt
        aNewSample.MidFETEQ0 = strMidFETEQ0
        aNewSample.MidFETEQ05 = strMidFETEQ05
        aNewSample.MidFETEQLOD = strMidFETEQLOD

        'Chrom Values
        aNewSample.Methylated = blnMethylated
        aNewSample.ChromSpikeAmt = strChromSpikeAmt
        aNewSample.SpikeCalculated = blnSpikeCalculated

        'Trace Values
        aNewSample.AvgCalArea = strAvgCalArea

        'SIS Values
        aNewSample.SISInternalID = strSISInternalID
        aNewSample.SISLabNum = strSISLabNum
        aNewSample.SISClientSampID = strSISClientSampID
        aNewSample.SISSampDate = dtSISSampDate
        aNewSample.SISSampDate = dtSISSampDateEnd
        aNewSample.SISTargetSampSize = strSISTargetSampSize
        aNewSample.SISActualSampSize = strSISActualSampSize
        aNewSample.SISDefaultAliquot = strSISDefaultAliquot
        aNewSample.SISAnalyses = strSISAnalyses
        aNewSample.SISSpikeMult = strSISSpikeMult
        aNewSample.SISDilFactor = strSISDilFactor
        aNewSample.SISFinalWeight = strSISFinalWeight
        aNewSample.SISTinWeight = strSISTinWeight
        aNewSample.SISWetWeight = strSISWetWeight
        aNewSample.SISDryWeight = strSISDryWeight
        aNewSample.SISSampleWeight = strSISSampleWeight
        aNewSample.SISSBottWeight = strSISSBottWeight
        aNewSample.SISSampWetWeight = strSISSampWetWeight
        aNewSample.SISEBottWeight = strSISEBottWeight
        aNewSample.SISPMoisture = strSISPMoisture
        aNewSample.SISType = strSISType

        'CCCHECK File Values
        aNewSample.CCCQuantTime = dtCCCQuantTime
        aNewSample.MinRRF = strMinRRF
        aNewSample.MinRelArea = strMinRelArea
        aNewSample.MaxRTDev = strMaxRTDev
        aNewSample.MaxRRFDev = strMaxRRFDev
        aNewSample.MaxRelArea = strMaxRelArea
    End Sub

    'Elution Sort Standards
    Public Sub ESortStandards()

        Dim aStandard1 As Standard
        Dim aStandard2 As Standard
        Dim aTempStandard As Standard
        Dim iCount As Integer
        Dim intStart As Integer
        Dim intEnd As Integer
        Dim blnSwapped As Boolean


        'Record rank
        For Each aStandard2 In SurrogateList
            iCount = 0
            For Each aStandard1 In GlobalVariables.ElutionOrderSample.InternalStdList
                If aStandard2.Name = aStandard1.Name Then
                    aStandard2.Index = iCount
                    Exit For
                End If
                iCount += 1
            Next
        Next

        'If not found in Dictionary
        iCount = 1
        For Each aStandard2 In InternalStdList
            If aStandard2.Index = -1 Then
                aStandard2.Index = GlobalVariables.ElutionOrderSample.InternalStdList.Count + iCount
                iCount += 1
            End If
        Next

        'Sort
        intStart = 0
        intEnd = InternalStdList.Count - 1
        blnSwapped = True

        Do While blnSwapped
            'Reset swapped flag
            blnSwapped = False
            For i = 0 To intEnd
                If i < intEnd Then
                    aStandard1 = InternalStdList.Item(i)
                    aStandard2 = InternalStdList.Item(i + 1)
                    If aStandard1.Index > aStandard2.Index Then
                        aTempStandard = aStandard1
                        aStandard1 = aStandard2
                        aStandard1 = aTempStandard
                        blnSwapped = True
                    End If
                    InternalStdList.Item(i) = aStandard1
                    InternalStdList.Item(i + 1) = aStandard2
                End If

            Next

            blnSwapped = False

            For i = intEnd - 1 To intStart Step -1
                If i >= intStart Then
                    aStandard1 = InternalStdList.Item(i)
                    aStandard2 = InternalStdList.Item(i + 1)
                    If aStandard1.Index > aStandard2.Index Then
                        aTempStandard = aStandard1
                        aStandard1 = aStandard2
                        aStandard2 = aTempStandard
                        blnSwapped = True
                    End If
                    InternalStdList.Item(i) = aStandard1
                    InternalStdList.Item(i + 1) = aStandard2
                End If

            Next
            intStart = intStart + 1
        Loop

    End Sub
    'Elution Sort Standards
    Public Sub ESortSurrogates()

        Dim aSurrogate1 As Surrogate
        Dim aSurrogate2 As Surrogate
        Dim aTempSurrogate As Surrogate
        Dim iCount As Integer
        Dim intStart As Integer
        Dim intEnd As Integer
        Dim blnSwapped As Boolean


        'Record rank
        For Each aSurrogate2 In SurrogateList
            iCount = 0
            For Each aSurrogate1 In GlobalVariables.ElutionOrderSample.SurrogateList
                If aSurrogate2.Name = aSurrogate1.Name Then
                    aSurrogate2.Index = iCount
                    Exit For
                End If
                iCount += 1
            Next
        Next

        'If not found in Dictionary
        iCount = 1
        For Each aSurrogate2 In SurrogateList
            If aSurrogate2.Index = -1 Then
                aSurrogate2.Index = GlobalVariables.ElutionOrderSample.SurrogateList.Count + iCount
                iCount += 1
            End If
        Next

        'Sort
        intStart = 0
        intEnd = SurrogateList.Count - 1
        blnSwapped = True

        Do While blnSwapped
            'Reset swapped flag
            blnSwapped = False
            For i = 0 To intEnd
                If i < intEnd Then
                    aSurrogate1 = SurrogateList.Item(i)
                    aSurrogate2 = SurrogateList.Item(i + 1)
                    If aSurrogate1.Index > aSurrogate2.Index Then
                        aTempSurrogate = aSurrogate1
                        aSurrogate1 = aSurrogate2
                        aSurrogate1 = aTempSurrogate
                        blnSwapped = True
                    End If
                    SurrogateList.Item(i) = aSurrogate1
                    SurrogateList.Item(i + 1) = aSurrogate2
                End If

            Next

            blnSwapped = False

            For i = intEnd - 1 To intStart Step -1
                If i >= intStart Then
                    aSurrogate1 = SurrogateList.Item(i)
                    aSurrogate2 = SurrogateList.Item(i + 1)
                    If aSurrogate1.Index > aSurrogate2.Index Then
                        aTempSurrogate = aSurrogate1
                        aSurrogate1 = aSurrogate2
                        aSurrogate2 = aTempSurrogate
                        blnSwapped = True
                    End If
                    SurrogateList.Item(i) = aSurrogate1
                    SurrogateList.Item(i + 1) = aSurrogate2
                End If

            Next
            intStart = intStart + 1
        Loop


    End Sub
    'Elution Sort Standards
    Public Sub ESortCompounds()

        Dim aCompound1 As Compound
        Dim aCompound2 As Compound
        Dim aTempCompound As Compound
        Dim iCount As Integer
        Dim intStart As Integer
        Dim intEnd As Integer
        Dim blnSwapped As Boolean


        'Record rank
        For Each aCompound2 In CompoundList
            iCount = 0
            For Each aCompound1 In GlobalVariables.ElutionOrderSample.CompoundList
                If aCompound2.Name = aCompound1.Name Then
                    aCompound2.Index = iCount
                    Exit For
                End If
                iCount += 1
            Next
        Next

        'If not found in Dictionary
        iCount = 1
        For Each aCompound2 In CompoundList
            If aCompound2.Index = -1 Then
                aCompound2.Index = GlobalVariables.ElutionOrderSample.CompoundList.Count + iCount
                iCount += 1
            End If
        Next

        'Sort
        intStart = 0
        intEnd = CompoundList.Count - 1
        blnSwapped = True

        Do While blnSwapped
            'Reset swapped flag
            blnSwapped = False
            For i = 0 To intEnd
                If i < intEnd Then
                    aCompound1 = CompoundList.Item(i)
                    aCompound2 = CompoundList.Item(i + 1)
                    If aCompound1.Index > aCompound2.Index Then
                        aTempCompound = aCompound1
                        aCompound1 = aCompound2
                        aCompound1 = aTempCompound
                        blnSwapped = True
                    End If
                    CompoundList.Item(i) = aCompound1
                    CompoundList.Item(i + 1) = aCompound2
                End If

            Next

            blnSwapped = False

            For i = intEnd - 1 To intStart Step -1
                If i >= intStart Then
                    aCompound1 = CompoundList.Item(i)
                    aCompound2 = CompoundList.Item(i + 1)
                    If aCompound1.Index > aCompound2.Index Then
                        aTempCompound = aCompound1
                        aCompound1 = aCompound2
                        aCompound2 = aTempCompound
                        blnSwapped = True
                    End If
                    CompoundList.Item(i) = aCompound1
                    CompoundList.Item(i + 1) = aCompound2
                End If

            Next
            intStart = intStart + 1
        Loop

    End Sub


    Public Sub SortStandards()
        'Done using Cocktail Shaker Sort
        Dim aStandard1 As Standard
        Dim aStandard2 As Standard
        Dim aTempStandard As Standard
        Dim blnSwapped As Boolean
        Dim intStart As Integer
        Dim intEnd As Integer

        intStart = 0
        intEnd = InternalStdList.Count - 1
        blnSwapped = True

        Do While blnSwapped
            'Reset swapped flag
            blnSwapped = False
            For i = 0 To intEnd
                If i < intEnd Then
                    aStandard1 = InternalStdList.Item(i)
                    aStandard2 = InternalStdList.Item(i + 1)
                    If aStandard1.Name > aStandard2.Name Then
                        aTempStandard = aStandard1
                        aStandard1 = aStandard2
                        aStandard2 = aTempStandard
                        blnSwapped = True
                    End If
                    InternalStdList.Item(i) = aStandard1
                    InternalStdList.Item(i + 1) = aStandard2
                End If

            Next

            blnSwapped = False

            For i = intEnd - 1 To intStart Step -1
                If i >= intStart Then
                    aStandard1 = InternalStdList.Item(i)
                    aStandard2 = InternalStdList.Item(i + 1)
                    If aStandard1.Name > aStandard2.Name Then
                        aTempStandard = aStandard1
                        aStandard1 = aStandard2
                        aStandard2 = aTempStandard
                        blnSwapped = True
                    End If
                    InternalStdList.Item(i) = aStandard1
                    InternalStdList.Item(i + 1) = aStandard2
                End If

            Next
            intStart = intStart + 1
        Loop
    End Sub
    Public Sub SortSSurrogates()
        'Done using Cocktail Shaker Sort
        Dim aSurrogate1 As Surrogate
        Dim aSurrogate2 As Surrogate
        Dim aTempSurrogate As Surrogate
        Dim blnSwapped As Boolean
        Dim intStart As Integer
        Dim intEnd As Integer

        intStart = 0
        intEnd = SurrogateList.Count - 1
        blnSwapped = True

        Do While blnSwapped
            'Reset swapped flag
            blnSwapped = False

            For i = 0 To intEnd
                If i < intEnd Then
                    aSurrogate1 = SurrogateList.Item(i)
                    aSurrogate2 = SurrogateList.Item(i + 1)
                    If aSurrogate1.Name > aSurrogate2.Name Then
                        aTempSurrogate = aSurrogate1
                        aSurrogate1 = aSurrogate2
                        aSurrogate2 = aTempSurrogate
                        blnSwapped = True
                    End If
                    SurrogateList.Item(i) = aSurrogate1
                    SurrogateList.Item(i + 1) = aSurrogate2
                End If

            Next

            blnSwapped = False

            For i = intEnd - 1 To intStart Step -1
                If i >= intStart Then
                    aSurrogate1 = SurrogateList.Item(i)
                    aSurrogate2 = SurrogateList.Item(i + 1)
                    If aSurrogate1.Name > aSurrogate2.Name Then
                        aTempSurrogate = aSurrogate1
                        aSurrogate1 = aSurrogate2
                        aSurrogate2 = aTempSurrogate
                        blnSwapped = True
                    End If
                    SurrogateList.Item(i) = aSurrogate1
                    SurrogateList.Item(i + 1) = aSurrogate2
                End If

            Next
            intStart = intStart + 1
        Loop
    End Sub
    Public Sub SortCompounds()
        'Done using Cocktail Shaker Sort
        Dim aCompound1 As Compound
        Dim aCompound2 As Compound
        Dim aTempCompound As Compound
        Dim blnSwapped As Boolean
        Dim intStart As Integer
        Dim intEnd As Integer

        intStart = 0
        intEnd = CompoundList.Count - 1
        blnSwapped = True

        Do While blnSwapped
            'Reset swapped flag
            blnSwapped = False

            For i = 0 To intEnd
                If i < intEnd Then
                    aCompound1 = CompoundList.Item(i)
                    aCompound2 = CompoundList.Item(i + 1)
                    If aCompound1.Name > aCompound2.Name Then
                        aTempCompound = aCompound1
                        aCompound1 = aCompound2
                        aCompound2 = aTempCompound
                        blnSwapped = True
                    End If
                    CompoundList.Item(i) = aCompound1
                    CompoundList.Item(i + 1) = aCompound2
                End If

            Next

            blnSwapped = False

            For i = intEnd - 1 To intStart Step -1
                If i >= intStart Then
                    aCompound1 = CompoundList.Item(i)
                    aCompound2 = CompoundList.Item(i + 1)
                    If aCompound1.Name > aCompound2.Name Then
                        aTempCompound = aCompound1
                        aCompound1 = aCompound2
                        aCompound2 = aTempCompound
                        blnSwapped = True
                    End If
                    CompoundList.Item(i) = aCompound1
                    CompoundList.Item(i + 1) = aCompound2
                End If

            Next
            intStart = intStart + 1
        Loop
    End Sub
End Class
