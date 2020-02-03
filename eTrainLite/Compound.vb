Public Class Compound

    'Chemstation
    Private strName As String
    Private strRT As String
    Private strQIon As String
    Private strResponse As String
    Private strConc As String
    Private strUnits As String
    Private strQValue As String
    Private strRLimit As String
    Private blnQOOR As Boolean
    Private blnMI As Boolean
    Private blnSS As Boolean
    Private strAvgRF As String
    Private strCCRF As String
    Private strPercDev As String
    Private strPercArea As String
    Private strCCCDevMin As String
    Private blnPercDevOOR As Boolean
    Private blnPercAreaOOR As Boolean
    Private blnCCCDevMinOOR As Boolean
    Private blnWritten As Boolean
    Private blnWriteToReport As Boolean
    Private strReportedAmt As String
    Private strReportedUnits As String
    Private strCASNum As String
    Private strArea As String
    Private strHeight As String

    'EDD File Values 
    Private strSysSampleCode As String 'Added (EDD values) WT 10/12/2017
    Private strLabAnlMethodName As String
    Private strAnalysisDate As String
    Private strAnalysisTime As String '?
    Private strTotalOrDissolved As String
    Private strColumnNumber As String
    Private strTestType As String
    Private strLabMatrixCode As String
    Private strAnalysisLocation As String
    Private strBasis As String
    Private strSampleTypeCode As String
    Private strContainerID As String
    Private strEDilutionFactor As String
    Private strPrepMethod As String
    Private strPrepDate As String
    Private strPrepTime As String
    Private strLeachateMethod As String
    Private dtLeachateDate As String
    Private strLeachateTime As String
    Private strLabNameCode As String
    Private strQcLevel As String
    Private strLabSampleID As String
    Private strPercentMoisture As String
    Private strSubsampleAmount As String
    Private strSubsampleAmountUnit As String
    Private strAnalystName As String
    Private strInstrumentID As String
    Private strComment As String
    Private strPreservative As String
    'Private strCommentPreservative As String
    Private strFinalVolume As String
    Private strFinalVolumeUnit As String
    Private strCasRn As String
    Private strChemicalName As String
    Private strResultValue As String
    Private strResultErrorDelta As String
    Private strResultTypeCode As String
    Private strReportableResult As String
    Private strDetectFlag As String 'Boolean?
    Private strLabQualifiers As String
    Private strValidatorQualifiers As String
    Private strInterpretedQualifier As String
    Private strMethodDetectionLimit As String
    Private strOrganicYn As String
    Private strReportingDetectionLimit As String
    Private strQuantitationLimt As String
    Private strResultUnit As String
    Private strDetectionLimitUnit As String
    Private strTicRetentionTime As String
    Private strResultComment As String
    Private strSDG As String
    Private strQcOriginalConc As String
    Private strQcSpikeAdded As String
    Private strQcSpikeMeasured As String
    Private strQcSpikeRecovery As String
    Private strQcDupOriginalConc As String
    Private strQcDupSpikeAdded As String
    Private strQcDupSpikeMeasured As String
    Private strQcDupSpikeRecovery As String
    Private strQcRpd As String
    Private strQcSpikeLcl As String
    Private strQcSpikeUcl As String
    Private strQcRpdCl As String
    Private strQcSpikeStatus As String
    Private strQcDupSpikeStatus As String
    Private strQcRpdStatus As String
    Private strRlOrMdl As String
    Private strCustomField2 As String
    Private strCustomField3 As String
    Private strCustomField4 As String
    Private strCustomField5 As String
    Private strUncertainty As String
    Private strMinimumDetectableConc As String
    Private strCountingError As String
    Private strCriticalValue As String
    ' Added for Vista samples/data
    Private strTEQType As String
    Private strTEQMin As String
    Private strTEQMax As String
    Private strTEQRisk As String
    Private strPercentSolids As String
    Private strPercentLipids As String
    Private strEMPC As String

    'TQ3
    Private strTQ3QuanMass As String
    Private strTQ3QMHeight As String
    Private strTQ3QMArea As String
    Private strTQ3QMAreaAvg As String
    Private strTQ3QMSigNoi As String
    Private strTQ3QMNoise As String
    Private strTQ3RatioMass As String
    Private strTQ3RM1Height As String
    Private strTQ3RM1Area As String
    Private strTQ3RM1SigNoi As String
    Private strTQ3RM1Noise As String
    Private strTQ3R1R2 As String
    Private strTQ3SpecAmt As String
    Private strTQ3MI1 As String
    Private strTQ3MI2 As String

    'MidlandFAST
    Private dblMidFIonRatio As Double
    Private blnMidFIonRatioInLim As Boolean
    Private blnMidFIsTarget As Boolean
    Private blnMidFIsQual As Boolean
    Private blnMidFQC1 As Boolean
    Private blnMidFQC2 As Boolean
    Private blnMidFQC3 As Boolean
    Private blnMidFQC4 As Boolean
    Private blnMidFQC5 As Boolean
    Private blnMidFQC6 As Boolean
    Private blnMidFQC7 As Boolean
    Private blnMidFQC8 As Boolean
    Private blnMidFQC9 As Boolean
    Private blnMidFQC10 As Boolean
    Private blnMidFQC11 As Boolean
    Private blnMidFQC12 As Boolean
    Private strMidFLLim As String
    Private strMidFULim As String
    Private dblMidFLoq As Double
    Private dblMidF13CAmt As Double
    Private dblMidFRF As Double
    Private strMidF13CAssoc As String
    Private strMidFReportedAmt As String
    Private strMidFReportedLOQAmt As String
    Private strMidFLCSAmtAdded As String
    Private strMidFLCSAmtRecovered As String
    Private strMidFLCSPercRecovered As String
    Private strMidFLCSLowLim As String
    Private strMidFLCSHighLim As String
    Private strMidFLCSTolerance As String
    Private strMidFCS3TotalAmt As String
    Private strMidFCS3AmtRecovered As String
    Private strMidFCS3LowLim As String
    Private strMidFCS3HighLim As String
    Private strMidFCS3Tolerance As String
    Private strMidFFinalWeight As String
    Private strMidFTEQFinalWeight As String
    Private blnMidFNonDetect As Boolean
    Private blnMidFETEQ0 As Boolean
    Private blnMidFETEQ05 As Boolean
    Private blnMidFETEQLOD As Boolean
    Private strMidFFlags As String

    'Chrom
    Private strChromLowContLim As String
    Private strChromUpContLim As String
    Private strChromLowICVLim As String
    Private strChromUpICVLim As String
    Private strChromLowCVSLim As String
    Private strChromUpCVSLim As String
    Private strChromMBLim As String
    Private strChromLowMSLim As String
    Private strChromUpMSLim As String
    Private strChromLowLCSLim As String
    Private strChromUpLCSLim As String
    Private strChromReportLimit As String
    Private strChromAdjustedLimit As String
    Private strChromCorrectedSpike As String
    Private strChromSpikeRecovery As String
    Private blnChromSpikePass As Boolean
    Private strChromRPD As String
    Private strChromRPDLimit As String
    Private blnChromICVPass As Boolean
    Private blnChromCVSPass As Boolean
    Private blnMethylated As Boolean
    Private strChromAdjustConc As String

    'TMPQNTRP 
    Private strTSignal As String
    Private strQ1Signal As String
    Private strQ2Signal As String
    Private strQ3Signal As String
    Private strTRatios As String
    Private strQ1Ratios As String
    Private strQ2Ratios As String
    Private strQ3Ratios As String
    Private strQ1SLLim As String
    Private strQ1SULim As String
    Private strQ2SLLim As String
    Private strQ2SULim As String
    Private strQ3SLLim As String
    Private strQ3SULim As String
    Private strTResp As String
    Private strQ1Resp As String
    Private strQ2Resp As String
    Private strQ3Resp As String
    Private strTRT As String
    Private strQ1RT As String
    Private strQ2RT As String
    Private strQ3RT As String
    Private strRTLLim As String
    Private strRTULim As String
    Private strTIntType As String
    Private strQ1IntType As String
    Private strQ2IntType As String
    Private strQ3IntType As String
    Private strTMPRelStdDev As String

    Private blnKeep As Boolean
    Private intIndex As Integer

    Public Sub New()
        'Constructor

        'Default values
        dblMidFIonRatio = -1
        dblMidFLoq = -1
        strMidFReportedAmt = "-1"
        strMidFReportedLOQAmt = "-1"
        blnMidFIonRatioInLim = False
        blnMidFIsTarget = False
        blnMidFIsQual = False
        blnMidFQC1 = False
        blnMidFQC2 = False
        blnMidFQC3 = False
        blnMidFQC4 = False
        blnMidFQC5 = False
        blnMidFQC6 = False
        blnMidFQC7 = False
        blnMidFQC8 = False
        blnMidFQC9 = False
        blnMidFQC10 = False
        blnMidFQC11 = False
        blnMidFQC12 = False
        blnWritten = False
        blnMidFNonDetect = False
        blnMidFETEQ0 = False
        blnMidFETEQ05 = False
        blnMidFETEQLOD = False
        blnWriteToReport = True
        blnChromSpikePass = False
        blnChromICVPass = False
        blnChromCVSPass = False
        blnMethylated = False
        blnKeep = True
        intIndex = -1

    End Sub

    'Sets/Gets
    Public Property EDDsysSampleCode() As String 'Added (EDD's) WT 10/12/2017
        Get
            Return strSysSampleCode
        End Get
        Set(ByVal value As String)
            strSysSampleCode = value
        End Set
    End Property

    Public Property EDDLabAnlMethodName() As String
        Get
            Return strLabAnlMethodName
        End Get
        Set(ByVal value As String)
            strLabAnlMethodName = value
        End Set
    End Property

    Public Property EDDAnalysisDate() As String
        Get
            Return strAnalysisDate
        End Get
        Set(ByVal value As String)
            strAnalysisDate = value
        End Set
    End Property

    Public Property EDDAnalysisTime() As String
        Get
            Return strAnalysisTime
        End Get
        Set(ByVal value As String)
            strAnalysisTime = value
        End Set
    End Property

    Public Property EDDTotalOrDissolved() As String
        Get
            Return strTotalOrDissolved
        End Get
        Set(ByVal value As String)
            strTotalOrDissolved = value
        End Set
    End Property

    Public Property EDDColumnNumber() As String
        Get
            Return strColumnNumber
        End Get
        Set(ByVal value As String)
            strColumnNumber = value
        End Set
    End Property

    Public Property EDDTestType() As String
        Get
            Return strTestType
        End Get
        Set(ByVal value As String)
            strTestType = value
        End Set
    End Property

    Public Property EDDLabMatrixCode() As String
        Get
            Return strLabMatrixCode
        End Get
        Set(ByVal value As String)
            strLabMatrixCode = value
        End Set
    End Property

    Public Property EDDAnalysisLocation() As String
        Get
            Return strAnalysisLocation
        End Get
        Set(ByVal value As String)
            strAnalysisLocation = value
        End Set
    End Property

    Public Property EDDBasis() As String
        Get
            Return strBasis
        End Get
        Set(ByVal value As String)
            strBasis = value
        End Set
    End Property

    Public Property EDDSampleTypeCode() As String
        Get
            Return strSampleTypeCode
        End Get
        Set(ByVal value As String)
            strSampleTypeCode = value
        End Set
    End Property

    Public Property EDDContainerID() As String
        Get
            Return strContainerID
        End Get
        Set(ByVal value As String)
            strContainerID = value
        End Set
    End Property

    Public Property EDDEDilutionFactor() As String
        Get
            Return strEDilutionFactor
        End Get
        Set(ByVal value As String)
            strEDilutionFactor = value
        End Set
    End Property

    Public Property EDDPrepMethod() As String
        Get
            Return strPrepMethod
        End Get
        Set(ByVal value As String)
            strPrepMethod = value
        End Set
    End Property

    Public Property EDDPrepDate() As String
        Get
            Return strPrepDate
        End Get
        Set(ByVal value As String)
            strPrepDate = value
        End Set
    End Property

    Public Property EDDPrepTime() As String
        Get
            Return strPrepTime
        End Get
        Set(ByVal value As String)
            strPrepTime = value
        End Set
    End Property

    Public Property EDDLeachateMethod() As String
        Get
            Return strLeachateMethod
        End Get
        Set(ByVal value As String)
            strLeachateMethod = value
        End Set
    End Property

    Public Property EDDLeachateDate() As String

    Public Property EDDLeachateTime() As String
        Get
            Return strLeachateTime
        End Get
        Set(ByVal value As String)
            strLeachateTime = value
        End Set
    End Property

    Public Property EDDLabNameCode() As String
        Get
            Return strLabNameCode
        End Get
        Set(ByVal value As String)
            strLabNameCode = value
        End Set
    End Property

    Public Property EDDQcLevel() As String
        Get
            Return strQcLevel
        End Get
        Set(value As String)
            strQcLevel = value
        End Set
    End Property

    Public Property EDDLabSampleID() As String
        Get
            Return strLabSampleID
        End Get
        Set(ByVal value As String)
            strLabSampleID = value
        End Set
    End Property

    Public Property EDDPercentMoisture() As String
        Get
            Return strPercentMoisture
        End Get
        Set(ByVal value As String)
            strPercentMoisture = value
        End Set
    End Property

    Public Property EDDSubsampleAmount() As String
        Get
            Return strSubsampleAmount
        End Get
        Set(ByVal value As String)
            strSubsampleAmount = value
        End Set
    End Property

    Public Property EDDSubsampleAmountUnit() As String
        Get
            Return strSubsampleAmountUnit
        End Get
        Set(ByVal value As String)
            strSubsampleAmountUnit = value
        End Set
    End Property

    Public Property EDDAnalystName() As String
        Get
            Return strAnalystName
        End Get
        Set(ByVal value As String)
            strAnalystName = value
        End Set
    End Property

    Public Property EDDInstrumentID() As String
        Get
            Return strInstrumentID
        End Get
        Set(ByVal value As String)
            strInstrumentID = value
        End Set
    End Property

    Public Property EDDComment() As String
        Get
            Return strComment
        End Get
        Set(ByVal value As String)
            strComment = value
        End Set
    End Property

    Public Property EDDPreservative() As String
        Get
            Return strPreservative
        End Get
        Set(ByVal value As String)
            strPreservative = value
        End Set
    End Property


    'Public Property EDDCommentPreservative() As String
    '    Get
    '        Return strCommentPreservative
    '    End Get
    '    Set(ByVal value As String)
    '        strCommentPreservative = value
    '    End Set
    'End Property

    Public Property EDDFinalVolume() As String
        Get
            Return strFinalVolume
        End Get
        Set(ByVal value As String)
            strFinalVolume = value
        End Set
    End Property

    Public Property EDDFinalVolumeUnit() As String
        Get
            Return strFinalVolumeUnit
        End Get
        Set(ByVal value As String)
            strFinalVolumeUnit = value
        End Set
    End Property

    Public Property EDDCasRn() As String
        Get
            Return strCasRn
        End Get
        Set(ByVal value As String)
            strCasRn = value
        End Set
    End Property

    Public Property EDDChemicalName() As String
        Get
            Return strChemicalName
        End Get
        Set(ByVal value As String)
            strChemicalName = value
        End Set
    End Property

    Public Property EDDResultValue() As String
        Get
            Return strResultValue
        End Get
        Set(ByVal value As String)
            strResultValue = value
        End Set
    End Property


    Public Property EDDResultErrorDelta() As String
        Get
            Return strResultErrorDelta
        End Get
        Set(ByVal value As String)
            strResultErrorDelta = value
        End Set
    End Property

    Public Property EDDResultTypeCode() As String
        Get
            Return strResultTypeCode
        End Get
        Set(ByVal value As String)
            strResultTypeCode = value
        End Set
    End Property

    Public Property EDDReportableResult() As String
        Get
            Return strReportableResult
        End Get
        Set(ByVal value As String)
            strReportableResult = value
        End Set
    End Property

    Public Property EDDDetectFlag() As String 'Boolean?
        Get
            Return strDetectFlag
        End Get
        Set(ByVal value As String)
            strDetectFlag = value
        End Set
    End Property

    Public Property EDDLabQualifiers() As String
        Get
            Return strLabQualifiers
        End Get
        Set(ByVal value As String)
            strLabQualifiers = value
        End Set
    End Property

    Public Property EDDValidatorQualifiers() As String
        Get
            Return strValidatorQualifiers
        End Get
        Set(ByVal value As String)
            strValidatorQualifiers = value
        End Set
    End Property
    Public Property EDDInterpretedQualifier() As String
        Get
            Return strInterpretedQualifier
        End Get
        Set(ByVal value As String)
            strInterpretedQualifier = value
        End Set
    End Property
    Public Property EDDSDG() As String
        Get
            Return strSDG
        End Get
        Set(ByVal value As String)
            strSDG = value
        End Set
    End Property
    Public Property EDDCustomField2() As String
        Get
            Return strCustomField2
        End Get
        Set(ByVal value As String)
            strCustomField2 = value
        End Set
    End Property
    Public Property EDDCustomField3() As String
        Get
            Return strCustomField3
        End Get
        Set(ByVal value As String)
            strCustomField3 = value
        End Set
    End Property
    Public Property EDDCustomField4() As String
        Get
            Return strCustomField4
        End Get
        Set(ByVal value As String)
            strCustomField4 = value
        End Set
    End Property
    Public Property EDDCustomField5() As String
        Get
            Return strCustomField5
        End Get
        Set(ByVal value As String)
            strCustomField5 = value
        End Set
    End Property
    Public Property EDDUncertainty() As String
        Get
            Return strUncertainty
        End Get
        Set(ByVal value As String)
            strUncertainty = value
        End Set
    End Property
    Public Property EDDMinimumDetectableConc() As String
        Get
            Return strMinimumDetectableConc
        End Get
        Set(ByVal value As String)
            strMinimumDetectableConc = value
        End Set
    End Property
    Public Property EDDCountingError() As String
        Get
            Return strCountingError
        End Get
        Set(ByVal value As String)
            strCountingError = value
        End Set
    End Property
    Public Property EDDCriticalValue() As String
        Get
            Return strCriticalValue
        End Get
        Set(ByVal value As String)
            strCriticalValue = value
        End Set
    End Property
    Public Property EDDMethodDetectionLimit() As String
        Get
            Return strMethodDetectionLimit
        End Get
        Set(ByVal value As String)
            strMethodDetectionLimit = value
        End Set
    End Property

    Public Property EDDOrganicYn() As String
        Get
            Return strOrganicYn
        End Get
        Set(ByVal value As String)
            strOrganicYn = value
        End Set
    End Property

    Public Property EDDReportingDetectionLimit() As String
        Get
            Return strReportingDetectionLimit
        End Get
        Set(ByVal value As String)
            strReportingDetectionLimit = value
        End Set
    End Property

    Public Property EDDQuantitationLimit() As String
        Get
            Return strQuantitationLimt
        End Get
        Set(value As String)
            strQuantitationLimt = value
        End Set
    End Property

    Public Property EDDResultUnit() As String
        Get
            Return strResultUnit
        End Get
        Set(ByVal value As String)
            strResultUnit = value
        End Set
    End Property

    Public Property EDDDetectionLimitUnit() As String
        Get
            Return strDetectionLimitUnit
        End Get
        Set(ByVal value As String)
            strDetectionLimitUnit = value
        End Set
    End Property

    Public Property EDDTicRetentionTime() As String
        Get
            Return strTicRetentionTime
        End Get
        Set(ByVal value As String)
            strTicRetentionTime = value
        End Set
    End Property

    Public Property EDDResultComment() As String
        Get
            Return strResultComment
        End Get
        Set(ByVal value As String)
            strResultComment = value
        End Set
    End Property

    Public Property EDDQcOriginalConc() As String
        Get
            Return strQcOriginalConc
        End Get
        Set(ByVal value As String)
            strQcOriginalConc = value
        End Set
    End Property

    Public Property EDDQcSpikeAdded() As String
        Get
            Return strQcSpikeAdded
        End Get
        Set(ByVal value As String)
            strQcSpikeAdded = value
        End Set
    End Property

    Public Property EDDQcSpikeMeasured() As String
        Get
            Return strQcSpikeMeasured
        End Get
        Set(value As String)
            strQcSpikeMeasured = value
        End Set
    End Property

    Public Property EDDQcSpikeRecovery() As String
        Get
            Return strQcSpikeRecovery
        End Get
        Set(ByVal value As String)
            strQcSpikeRecovery = value
        End Set
    End Property

    Public Property EDDQcDupOriginalConc() As String
        Get
            Return strQcDupOriginalConc
        End Get
        Set(ByVal value As String)
            strQcDupOriginalConc = value
        End Set
    End Property

    Public Property EDDQcDupSpikeAdded() As String
        Get
            Return strQcDupSpikeAdded
        End Get
        Set(ByVal value As String)
            strQcDupSpikeAdded = value
        End Set
    End Property

    Public Property EDDQcDupSpikeMeasured() As String
        Get
            Return strQcDupSpikeMeasured
        End Get
        Set(ByVal value As String)
            strQcDupSpikeMeasured = value
        End Set
    End Property

    Public Property EDDQcDupSpikeRecovery() As String
        Get
            Return strQcDupSpikeRecovery
        End Get
        Set(ByVal value As String)
            strQcDupSpikeRecovery = value
        End Set
    End Property

    Public Property EDDQcRpd() As String
        Get
            Return strQcRpd
        End Get
        Set(ByVal value As String)
            strQcRpd = value
        End Set
    End Property

    Public Property EDDQcSpikeLcl() As String
        Get
            Return strQcSpikeLcl
        End Get
        Set(ByVal value As String)
            strQcSpikeLcl = value
        End Set
    End Property

    Public Property EDDQcSpikeUcl() As String
        Get
            Return strQcSpikeUcl
        End Get
        Set(ByVal value As String)
            strQcSpikeUcl = value
        End Set
    End Property

    Public Property EDDQcRpdCl() As String
        Get
            Return strQcRpdCl
        End Get
        Set(ByVal value As String)
            strQcRpdCl = value
        End Set
    End Property

    Public Property EDDQcSpikeStatus() As String
        Get
            Return strQcSpikeStatus
        End Get
        Set(ByVal value As String)
            strQcSpikeStatus = value
        End Set
    End Property

    Public Property EDDQcDupSpikeStatus() As String
        Get
            Return strQcDupSpikeStatus
        End Get
        Set(ByVal value As String)
            strQcDupSpikeStatus = value
        End Set
    End Property

    Public Property EDDQcRpdStatus() As String
        Get
            Return strQcRpdStatus
        End Get
        Set(ByVal value As String)
            strQcRpdStatus = value
        End Set
    End Property

    Public Property EDDRlOrMdl() As String
        Get
            Return strRlOrMdl
        End Get
        Set(ByVal value As String)
            strRlOrMdl = value
        End Set
    End Property '<- End add WT

    Public Property EDDTEQType() As String
        Get
            Return strTEQType
        End Get
        Set(ByVal value As String)
            strTEQType = value
        End Set
    End Property

    Public Property EDDTEQMin() As String
        Get
            Return strTEQMin
        End Get
        Set(ByVal value As String)
            strTEQMin = value
        End Set
    End Property

    Public Property EDDTEQMax() As String
        Get
            Return strTEQMax
        End Get
        Set(ByVal value As String)
            strTEQMax = value
        End Set
    End Property

    Public Property EDDTEQRisk() As String
        Get
            Return strTEQRisk
        End Get
        Set(ByVal value As String)
            strTEQRisk = value
        End Set
    End Property

    Public Property EDDPercentSolids() As String
        Get
            Return strPercentSolids
        End Get
        Set(ByVal value As String)
            strPercentSolids = value
        End Set
    End Property

    Public Property EDDPercentLipids() As String
        Get
            Return strPercentLipids
        End Get
        Set(ByVal value As String)
            strPercentLipids = value
        End Set
    End Property
    Public Property EDDEMPC() As String
        Get
            Return strEMPC
        End Get
        Set(ByVal value As String)
            strEMPC = value
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
    Public Property RT() As String
        Get
            Return strRT
        End Get
        Set(ByVal value As String)
            strRT = value
        End Set
    End Property
    Public Property QIon() As String
        Get
            Return strQIon
        End Get
        Set(ByVal value As String)
            strQIon = value
        End Set
    End Property
    Public Property Response() As String
        Get
            Return strResponse
        End Get
        Set(ByVal value As String)
            strResponse = value
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
    Public Property Units() As String
        Get
            Return strUnits
        End Get
        Set(ByVal value As String)
            strUnits = value
        End Set
    End Property
    Public Property QValue() As String
        Get
            Return strQValue
        End Get
        Set(ByVal value As String)
            strQValue = value
        End Set
    End Property
    Public Property CasNum() As String
        Get
            Return strCasNum
        End Get
        Set(ByVal value As String)
            strCasNum = value
        End Set
    End Property
    Public Property RLimit() As String
        Get
            Return strRLimit
        End Get
        Set(ByVal value As String)
            strRLimit = value
        End Set
    End Property
    Public Property QOOR() As Boolean
        Get
            Return blnQOOR
        End Get
        Set(ByVal value As Boolean)
            blnQOOR = value
        End Set
    End Property
    Public Property MI() As Boolean
        Get
            Return blnMI
        End Get
        Set(ByVal value As Boolean)
            blnMI = value
        End Set
    End Property
    Public Property SS() As Boolean
        Get
            Return blnSS
        End Get
        Set(ByVal value As Boolean)
            blnSS = value
        End Set
    End Property
    Public Property AvgRF() As String
        Get
            Return strAvgRF
        End Get
        Set(ByVal value As String)
            strAvgRF = value
        End Set
    End Property
    Public Property CCRF() As String
        Get
            Return strCCRF
        End Get
        Set(ByVal value As String)
            strCCRF = value
        End Set
    End Property
    Public Property PercDev() As String
        Get
            Return strPercDev
        End Get
        Set(ByVal value As String)
            strPercDev = value
        End Set
    End Property
    Public Property PercArea() As String
        Get
            Return strPercArea
        End Get
        Set(ByVal value As String)
            strPercArea = value
        End Set
    End Property
    Public Property CCCDevMin() As String
        Get
            Return strCCCDevMin
        End Get
        Set(ByVal value As String)
            strCCCDevMin = value
        End Set
    End Property
    Public Property PercDevOOR() As Boolean
        Get
            Return blnPercDevOOR
        End Get
        Set(ByVal value As Boolean)
            blnPercDevOOR = value
        End Set
    End Property
    Public Property PercAreaOOR() As Boolean
        Get
            Return blnPercAreaOOR
        End Get
        Set(ByVal value As Boolean)
            blnPercAreaOOR = value
        End Set
    End Property
    Public Property CCCDevMinOOR() As Boolean
        Get
            Return blnCCCDevMinOOR
        End Get
        Set(ByVal value As Boolean)
            blnCCCDevMinOOR = value
        End Set
    End Property
    Public Property TSignal() As String
        Get
            Return strTSignal
        End Get
        Set(ByVal value As String)
            strTSignal = value
        End Set
    End Property
    Public Property Q1Signal() As String
        Get
            Return strQ1Signal
        End Get
        Set(ByVal value As String)
            strQ1Signal = value
        End Set
    End Property
    Public Property Q2Signal() As String
        Get
            Return strQ2Signal
        End Get
        Set(ByVal value As String)
            strQ2Signal = value
        End Set
    End Property
    Public Property Q3Signal() As String
        Get
            Return strQ3Signal
        End Get
        Set(ByVal value As String)
            strQ3Signal = value
        End Set
    End Property
    Public Property TRatios() As String
        Get
            Return strTRatios
        End Get
        Set(ByVal value As String)
            strTRatios = value
        End Set
    End Property
    Public Property Q1Ratios() As String
        Get
            Return strQ1Ratios
        End Get
        Set(ByVal value As String)
            strQ1Ratios = value
        End Set
    End Property
    Public Property Q2Ratios() As String
        Get
            Return strQ2Ratios
        End Get
        Set(ByVal value As String)
            strQ2Ratios = value
        End Set
    End Property
    Public Property Q3Ratios() As String
        Get
            Return strQ3Ratios
        End Get
        Set(ByVal value As String)
            strQ3Ratios = value
        End Set
    End Property
    Public Property Q1SLLim() As String
        Get
            Return strQ1SLLim
        End Get
        Set(ByVal value As String)
            strQ1SLLim = value
        End Set
    End Property
    Public Property Q1SULim() As String
        Get
            Return strQ1SULim
        End Get
        Set(ByVal value As String)
            strQ1SULim = value
        End Set
    End Property
    Public Property Q2SLLim() As String
        Get
            Return strQ2SLLim
        End Get
        Set(ByVal value As String)
            strQ2SLLim = value
        End Set
    End Property
    Public Property Q2SULim() As String
        Get
            Return strQ2SULim
        End Get
        Set(ByVal value As String)
            strQ2SULim = value
        End Set
    End Property
    Public Property Q3SLLim() As String
        Get
            Return strQ3SLLim
        End Get
        Set(ByVal value As String)
            strQ3SLLim = value
        End Set
    End Property
    Public Property Q3SULim() As String
        Get
            Return strQ3SULim
        End Get
        Set(ByVal value As String)
            strQ3SULim = value
        End Set
    End Property
    Public Property TRT() As String
        Get
            Return strTRT
        End Get
        Set(ByVal value As String)
            strTRT = value
        End Set
    End Property
    Public Property Q1RT() As String
        Get
            Return strQ1RT
        End Get
        Set(ByVal value As String)
            strQ1RT = value
        End Set
    End Property
    Public Property Q2RT() As String
        Get
            Return strQ2RT
        End Get
        Set(ByVal value As String)
            strQ2RT = value
        End Set
    End Property
    Public Property Q3RT() As String
        Get
            Return strQ3RT
        End Get
        Set(ByVal value As String)
            strQ3RT = value
        End Set
    End Property
    Public Property RTLLim() As String
        Get
            Return strRTLLim
        End Get
        Set(ByVal value As String)
            strRTLLim = value
        End Set
    End Property
    Public Property RTULim() As String
        Get
            Return strRTULim
        End Get
        Set(ByVal value As String)
            strRTULim = value
        End Set
    End Property
    Public Property TIntType() As String
        Get
            Return strTIntType
        End Get
        Set(ByVal value As String)
            strTIntType = value
        End Set
    End Property
    Public Property Q1IntType() As String
        Get
            Return strQ1IntType
        End Get
        Set(ByVal value As String)
            strQ1IntType = value
        End Set
    End Property
    Public Property Q2IntType() As String
        Get
            Return strQ2IntType
        End Get
        Set(ByVal value As String)
            strQ2IntType = value
        End Set
    End Property
    Public Property Q3IntType() As String
        Get
            Return strQ3IntType
        End Get
        Set(ByVal value As String)
            strQ3IntType = value
        End Set
    End Property
    Public Property TMPRelStdDev() As String
        Get
            Return strTMPRelStdDev
        End Get
        Set(ByVal value As String)
            strTMPRelStdDev = value
        End Set
    End Property
    Public Property TResp() As String
        Get
            Return strTResp
        End Get
        Set(ByVal value As String)
            strTResp = value
        End Set
    End Property
    Public Property Q1Resp() As String
        Get
            Return strQ1Resp
        End Get
        Set(ByVal value As String)
            strQ1Resp = value
        End Set
    End Property
    Public Property Q2Resp() As String
        Get
            Return strQ2Resp
        End Get
        Set(ByVal value As String)
            strQ2Resp = value
        End Set
    End Property
    Public Property Q3Resp() As String
        Get
            Return strQ3Resp
        End Get
        Set(ByVal value As String)
            strQ3Resp = value
        End Set
    End Property
    Public Property MidFIonRatio() As Double
        Get
            Return dblMidFIonRatio
        End Get
        Set(ByVal value As Double)
            dblMidFIonRatio = value
        End Set
    End Property
    Public Property MidFIonRatioInLim() As Boolean
        Get
            Return blnMidFIonRatioInLim
        End Get
        Set(ByVal value As Boolean)
            blnMidFIonRatioInLim = value
        End Set
    End Property
    Public Property MidFIsTarget() As Boolean
        Get
            Return blnMidFIsTarget
        End Get
        Set(ByVal value As Boolean)
            blnMidFIsTarget = value
        End Set
    End Property
    Public Property MidFIsQual() As Boolean
        Get
            Return blnMidFIsQual
        End Get
        Set(ByVal value As Boolean)
            blnMidFIsQual = value
        End Set
    End Property
    Public Property MidFQC1() As Boolean
        Get
            Return blnMidFQC1
        End Get
        Set(ByVal value As Boolean)
            blnMidFQC1 = value
        End Set
    End Property
    Public Property MidFQC2() As Boolean
        Get
            Return blnMidFQC2
        End Get
        Set(ByVal value As Boolean)
            blnMidFQC2 = value
        End Set
    End Property
    Public Property MidFQC3() As Boolean
        Get
            Return blnMidFQC3
        End Get
        Set(ByVal value As Boolean)
            blnMidFQC3 = value
        End Set
    End Property
    Public Property MidFQC4() As Boolean
        Get
            Return blnMidFQC4
        End Get
        Set(ByVal value As Boolean)
            blnMidFQC4 = value
        End Set
    End Property
    Public Property MidFQC5() As Boolean
        Get
            Return blnMidFQC5
        End Get
        Set(ByVal value As Boolean)
            blnMidFQC5 = value
        End Set
    End Property
    Public Property MidFQC6() As Boolean
        Get
            Return blnMidFQC6
        End Get
        Set(ByVal value As Boolean)
            blnMidFQC6 = value
        End Set
    End Property
    Public Property MidFQC7() As Boolean
        Get
            Return blnMidFQC7
        End Get
        Set(ByVal value As Boolean)
            blnMidFQC7 = value
        End Set
    End Property
    Public Property MidFQC8() As Boolean
        Get
            Return blnMidFQC8
        End Get
        Set(ByVal value As Boolean)
            blnMidFQC8 = value
        End Set
    End Property
    Public Property MidFQC9() As Boolean
        Get
            Return blnMidFQC9
        End Get
        Set(ByVal value As Boolean)
            blnMidFQC9 = value
        End Set
    End Property
    Public Property MidFQC10() As Boolean
        Get
            Return blnMidFQC10
        End Get
        Set(ByVal value As Boolean)
            blnMidFQC10 = value
        End Set
    End Property
    Public Property MidFQC11() As Boolean
        Get
            Return blnMidFQC11
        End Get
        Set(ByVal value As Boolean)
            blnMidFQC11 = value
        End Set
    End Property
    Public Property MidFQC12() As Boolean
        Get
            Return blnMidFQC12
        End Get
        Set(ByVal value As Boolean)
            blnMidFQC12 = value
        End Set
    End Property
    Public Property MidFLLim() As String
        Get
            Return strMidFLLim
        End Get
        Set(ByVal value As String)
            strMidFLLim = value
        End Set
    End Property
    Public Property MidFULim() As String
        Get
            Return strMidFULim
        End Get
        Set(ByVal value As String)
            strMidFULim = value
        End Set
    End Property
    Public Property MidFLoq() As Double
        Get
            Return dblMidFLoq
        End Get
        Set(ByVal value As Double)
            dblMidFLoq = value
        End Set
    End Property
    Public Property MidF13CAmt() As Double
        Get
            Return dblMidF13CAmt
        End Get
        Set(ByVal value As Double)
            dblMidF13CAmt = value
        End Set
    End Property
    Public Property MidFRF() As Double
        Get
            Return dblMidFRF
        End Get
        Set(ByVal value As Double)
            dblMidFRF = value
        End Set
    End Property
    Public Property MidFFlags() As String
        Get
            Return strMidFFlags
        End Get
        Set(ByVal value As String)
            strMidFFlags = value
        End Set
    End Property
    Public Property Written() As Boolean
        Get
            Return blnWritten
        End Get
        Set(ByVal value As Boolean)
            blnWritten = value
        End Set
    End Property
    Public Property WriteToReport() As Boolean
        Get
            Return blnWriteToReport
        End Get
        Set(ByVal value As Boolean)
            blnWriteToReport = value
        End Set
    End Property
    Public Property ReportedAmt() As String
        Get
            Return strReportedAmt
        End Get
        Set(ByVal value As String)
            strReportedAmt = value
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
    Public Property MidF13CAssoc() As String
        Get
            Return strMidF13CAssoc
        End Get
        Set(ByVal value As String)
            strMidF13CAssoc = value
        End Set
    End Property
    Public Property MidFReportedAmt() As String
        Get
            Return strMidFReportedAmt
        End Get
        Set(ByVal value As String)
            strMidFReportedAmt = value
        End Set
    End Property
    Public Property MidFReportedLOQAmt() As String
        Get
            Return strMidFReportedLOQAmt
        End Get
        Set(ByVal value As String)
            strMidFReportedLOQAmt = value
        End Set
    End Property
    Public Property MidFLCSAmtAdded() As String
        Get
            Return strMidFLCSAmtAdded
        End Get
        Set(ByVal value As String)
            strMidFLCSAmtAdded = value
        End Set
    End Property
    Public Property MidFLCSAmtRecovered() As String
        Get
            Return strMidFLCSAmtRecovered
        End Get
        Set(ByVal value As String)
            strMidFLCSAmtRecovered = value
        End Set
    End Property
    Public Property MidFLCSPercRecovered() As String
        Get
            Return strMidFLCSPercRecovered
        End Get
        Set(ByVal value As String)
            strMidFLCSPercRecovered = value
        End Set
    End Property
    Public Property MidFLCSLowLim() As String
        Get
            Return strMidFLCSLowLim
        End Get
        Set(ByVal value As String)
            strMidFLCSLowLim = value
        End Set
    End Property
    Public Property MidFLCSHighLim() As String
        Get
            Return strMidFLCSHighLim
        End Get
        Set(ByVal value As String)
            strMidFLCSHighLim = value
        End Set
    End Property
    Public Property MidFLCSTolerance() As String
        Get
            Return strMidFLCSTolerance
        End Get
        Set(ByVal value As String)
            strMidFLCSTolerance = value
        End Set
    End Property
    Public Property MidFCS3TotalAmt() As String
        Get
            Return strMidFCS3TotalAmt
        End Get
        Set(ByVal value As String)
            strMidFCS3TotalAmt = value
        End Set
    End Property
    Public Property MidFCS3AmtRecovered() As String
        Get
            Return strMidFCS3AmtRecovered
        End Get
        Set(ByVal value As String)
            strMidFCS3AmtRecovered = value
        End Set
    End Property
    Public Property MidFCS3LowLim() As String
        Get
            Return strMidFCS3LowLim
        End Get
        Set(ByVal value As String)
            strMidFCS3LowLim = value
        End Set
    End Property
    Public Property MidFCS3HighLim() As String
        Get
            Return strMidFCS3HighLim
        End Get
        Set(ByVal value As String)
            strMidFCS3HighLim = value
        End Set
    End Property
    Public Property MidFCS3Tolerance() As String
        Get
            Return strMidFCS3Tolerance
        End Get
        Set(ByVal value As String)
            strMidFCS3Tolerance = value
        End Set
    End Property
    Public Property MidFFinalWeight() As String
        Get
            Return strMidFFinalWeight
        End Get
        Set(ByVal value As String)
            strMidFFinalWeight = value
        End Set
    End Property
    Public Property MidFTEQFinalWeight() As String
        Get
            Return strMidFTEQFinalWeight
        End Get
        Set(ByVal value As String)
            strMidFTEQFinalWeight = value
        End Set
    End Property
    Public Property MidFETEQ0() As Boolean
        Get
            Return blnMidFETEQ0
        End Get
        Set(ByVal value As Boolean)
            blnMidFETEQ0 = value
        End Set
    End Property
    Public Property MidFETEQ05() As Boolean
        Get
            Return blnMidFETEQ05
        End Get
        Set(ByVal value As Boolean)
            blnMidFETEQ05 = value
        End Set
    End Property
    Public Property MidFETEQLOD() As Boolean
        Get
            Return blnMidFETEQLOD
        End Get
        Set(ByVal value As Boolean)
            blnMidFETEQLOD = value
        End Set
    End Property
    Public Property MidFNonDetect() As Boolean
        Get
            Return blnMidFNonDetect
        End Get
        Set(ByVal value As Boolean)
            blnMidFNonDetect = value
        End Set
    End Property
    Public Property ChromLowLCSLim() As String
        Get
            Return strChromLowLCSLim
        End Get
        Set(ByVal value As String)
            strChromLowLCSLim = value
        End Set
    End Property
    Public Property ChromUpLCSLim() As String
        Get
            Return strChromUpLCSLim
        End Get
        Set(ByVal value As String)
            strChromUpLCSLim = value
        End Set
    End Property
    Public Property ChromLowMSLim() As String
        Get
            Return strChromLowMSLim
        End Get
        Set(ByVal value As String)
            strChromLowMSLim = value
        End Set
    End Property
    Public Property ChromUpMSLim() As String
        Get
            Return strChromUpMSLim
        End Get
        Set(ByVal value As String)
            strChromUpMSLim = value
        End Set
    End Property
    Public Property ChromLowICVLim() As String
        Get
            Return strChromLowICVLim
        End Get
        Set(ByVal value As String)
            strChromLowICVLim = value
        End Set
    End Property
    Public Property ChromUpICVLim() As String
        Get
            Return strChromUpICVLim
        End Get
        Set(ByVal value As String)
            strChromUpICVLim = value
        End Set
    End Property
    Public Property ChromLowCVSLim() As String
        Get
            Return strChromLowCVSLim
        End Get
        Set(ByVal value As String)
            strChromLowCVSLim = value
        End Set
    End Property
    Public Property ChromUpCVSLim() As String
        Get
            Return strChromUpCVSLim
        End Get
        Set(ByVal value As String)
            strChromUpCVSLim = value
        End Set
    End Property
    Public Property ChromMBLim() As String
        Get
            Return strChromMBLim
        End Get
        Set(ByVal value As String)
            strChromMBLim = value
        End Set
    End Property
    Public Property ChromLowContLim() As String
        Get
            Return strChromLowContLim
        End Get
        Set(ByVal value As String)
            strChromLowContLim = value
        End Set
    End Property
    Public Property ChromUpContLim() As String
        Get
            Return strChromUpContLim
        End Get
        Set(ByVal value As String)
            strChromUpContLim = value
        End Set
    End Property
    Public Property ChromReportLimit() As String
        Get
            Return strChromReportLimit
        End Get
        Set(ByVal value As String)
            strChromReportLimit = value
        End Set
    End Property
    Public Property ChromAdjustedLimit() As String
        Get
            Return strChromAdjustedLimit
        End Get
        Set(ByVal value As String)
            strChromAdjustedLimit = value
        End Set
    End Property
    Public Property ChromCorrectedSpike() As String
        Get
            Return strChromCorrectedSpike
        End Get
        Set(ByVal value As String)
            strChromCorrectedSpike = value
        End Set
    End Property
    Public Property ChromSpikeRecovery() As String
        Get
            Return strChromSpikeRecovery
        End Get
        Set(ByVal value As String)
            strChromSpikeRecovery = value
        End Set
    End Property
    Public Property ChromSpikePass() As Boolean
        Get
            Return blnChromSpikePass
        End Get
        Set(ByVal value As Boolean)
            blnChromSpikePass = value
        End Set
    End Property
    Public Property ChromICVPass() As Boolean
        Get
            Return blnChromICVPass
        End Get
        Set(ByVal value As Boolean)
            blnChromICVPass = value
        End Set
    End Property
    Public Property ChromCVSPass() As Boolean
        Get
            Return blnChromCVSPass
        End Get
        Set(ByVal value As Boolean)
            blnChromCVSPass = value
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
    Public Property Keep() As Boolean
        Get
            Return blnKeep
        End Get
        Set(ByVal value As Boolean)
            blnKeep = value
        End Set
    End Property
    Public Property ChromRPD() As String
        Get
            Return strChromRPD
        End Get
        Set(ByVal value As String)
            strChromRPD = value
        End Set
    End Property
    Public Property ChromRPDLimit() As String
        Get
            Return strChromRPDLimit
        End Get
        Set(ByVal value As String)
            strChromRPDLimit = value
        End Set
    End Property
    Public Property Area() As String
        Get
            Return strArea
        End Get
        Set(ByVal value As String)
            strArea = value
        End Set
    End Property
    Public Property Height() As String
        Get
            Return strHeight
        End Get
        Set(ByVal value As String)
            strHeight = value
        End Set
    End Property
    Public Property TQ3QuanMass() As String
        Get
            Return strTQ3QuanMass
        End Get
        Set(ByVal value As String)
            strTQ3QuanMass = value
        End Set
    End Property
    Public Property TQ3QMHeight() As String
        Get
            Return strTQ3QMHeight
        End Get
        Set(ByVal value As String)
            strTQ3QMHeight = value
        End Set
    End Property
    Public Property TQ3QMArea() As String
        Get
            Return strTQ3QMArea
        End Get
        Set(ByVal value As String)
            strTQ3QMArea = value
        End Set
    End Property
    Public Property TQ3QMAreaAvg() As String
        Get
            Return strTQ3QMAreaAvg
        End Get
        Set(ByVal value As String)
            strTQ3QMAreaAvg = value
        End Set
    End Property
    Public Property TQ3QMSigNoi() As String
        Get
            Return strTQ3QMSigNoi
        End Get
        Set(ByVal value As String)
            strTQ3QMSigNoi = value
        End Set
    End Property
    Public Property TQ3QMNoise() As String
        Get
            Return strTQ3QMNoise
        End Get
        Set(ByVal value As String)
            strTQ3QMNoise = value
        End Set
    End Property
    Public Property TQ3RatioMass() As String
        Get
            Return strTQ3RatioMass
        End Get
        Set(ByVal value As String)
            strTQ3RatioMass = value
        End Set
    End Property
    Public Property TQ3RM1Height() As String
        Get
            Return strTQ3RM1Height
        End Get
        Set(ByVal value As String)
            strTQ3RM1Height = value
        End Set
    End Property
    Public Property TQ3RM1Area() As String
        Get
            Return strTQ3RM1Area
        End Get
        Set(ByVal value As String)
            strTQ3RM1Area = value
        End Set
    End Property
    Public Property TQ3RM1SigNoi() As String
        Get
            Return strTQ3RM1SigNoi
        End Get
        Set(ByVal value As String)
            strTQ3RM1SigNoi = value
        End Set
    End Property
    Public Property TQ3RM1Noise() As String
        Get
            Return strTQ3RM1Noise
        End Get
        Set(ByVal value As String)
            strTQ3RM1Noise = value
        End Set
    End Property
    Public Property TQ3R1R2() As String
        Get
            Return strTQ3R1R2
        End Get
        Set(ByVal value As String)
            strTQ3R1R2 = value
        End Set
    End Property
    Public Property TQ3SpecAmt() As String
        Get
            Return strTQ3SpecAmt
        End Get
        Set(ByVal value As String)
            strTQ3SpecAmt = value
        End Set
    End Property
    Public Property TQ3MI1() As String
        Get
            Return strTQ3MI1
        End Get
        Set(ByVal value As String)
            strTQ3MI1 = value
        End Set
    End Property
    Public Property TQ3MI2() As String
        Get
            Return strTQ3MI2
        End Get
        Set(ByVal value As String)
            strTQ3MI2 = value
        End Set
    End Property
    Public Property ChromAdjustConc() As String
        Get
            Return strChromAdjustConc
        End Get
        Set(ByVal value As String)
            strChromAdjustConc = value
        End Set
    End Property
    Public Property Index() As Integer
        Get
            Return intIndex
        End Get
        Set(ByVal value As Integer)
            intIndex = value
        End Set
    End Property

    Public Sub CopyDetails(ByVal aCompound As Compound, ByVal aCompound2 As Compound)
        strRT = aCompound.RT
        strQIon = aCompound.QIon
        strResponse = aCompound.Response
        strConc = aCompound.Conc
        strUnits = aCompound.Units
        strQValue = aCompound.QValue
        strRLimit = aCompound.RLimit
        blnQOOR = aCompound.QOOR
        blnMI = aCompound.MI
        blnSS = aCompound.SS
        strAvgRF = aCompound.AvgRF
        strCCRF = aCompound.CCRF
        strPercDev = aCompound.PercDev
        strPercArea = aCompound.PercArea
        strCCCDevMin = aCompound.CCCDevMin
        blnPercDevOOR = aCompound.PercDevOOR
        blnPercAreaOOR = aCompound.PercAreaOOR
        blnCCCDevMinOOR = aCompound.CCCDevMinOOR
        blnWritten = aCompound.Written
        blnWriteToReport = aCompound.WriteToReport
        strReportedAmt = aCompound.ReportedAmt

        'MidlandFAST
        dblMidFIonRatio = aCompound.MidFIonRatio
        blnMidFIonRatioInLim = aCompound.MidFIonRatioInLim
        blnMidFIsTarget = aCompound.MidFIsTarget
        blnMidFIsQual = aCompound.MidFIsQual
        If aCompound.MidFQC1 Or aCompound2.MidFQC1 Then
            blnMidFQC1 = True
        End If
        If aCompound.MidFQC2 Or aCompound2.MidFQC2 Then
            blnMidFQC2 = True
        End If
        If aCompound.MidFQC3 Or aCompound2.MidFQC3 Then
            blnMidFQC3 = True
        End If
        If aCompound.MidFQC4 Or aCompound2.MidFQC4 Then
            blnMidFQC4 = True
        End If
        If aCompound.MidFQC5 Or aCompound2.MidFQC5 Then
            blnMidFQC5 = True
        End If
        If aCompound.MidFQC6 Or aCompound2.MidFQC6 Then
            blnMidFQC6 = True
        End If
        If aCompound.MidFQC7 Or aCompound2.MidFQC7 Then
            blnMidFQC7 = True
        End If
        If aCompound.MidFQC8 Or aCompound2.MidFQC8 Then
            blnMidFQC8 = True
        End If
        If aCompound.MidFQC9 Or aCompound2.MidFQC9 Then
            blnMidFQC9 = True
        End If
        If aCompound.MidFQC10 Or aCompound2.MidFQC10 Then
            blnMidFQC10 = True
        End If
        If aCompound.MidFQC11 Or aCompound2.MidFQC11 Then
            blnMidFQC11 = True
        End If
        strMidFLLim = aCompound.MidFLLim
        strMidFULim = aCompound.MidFULim
        dblMidFLoq = aCompound.MidFLoq
        dblMidF13CAmt = aCompound.MidF13CAmt
        dblMidFRF = aCompound.MidFRF
        strMidF13CAssoc = aCompound.MidF13CAssoc
        strMidFReportedAmt = aCompound.MidFReportedAmt
        strMidFReportedLOQAmt = aCompound.MidFReportedLOQAmt
        strMidFLCSAmtAdded = aCompound.MidFLCSAmtAdded
        strMidFLCSAmtRecovered = aCompound.MidFLCSAmtRecovered
        strMidFLCSPercRecovered = aCompound.MidFLCSPercRecovered
        strMidFLCSLowLim = aCompound.MidFLCSLowLim
        strMidFLCSHighLim = aCompound.MidFLCSHighLim
        strMidFLCSTolerance = aCompound.MidFLCSTolerance
        strMidFCS3TotalAmt = aCompound.MidFCS3TotalAmt
        strMidFCS3AmtRecovered = aCompound.MidFCS3AmtRecovered
        strMidFCS3LowLim = aCompound.MidFCS3LowLim
        strMidFCS3HighLim = aCompound.MidFCS3HighLim
        strMidFCS3Tolerance = aCompound.MidFCS3Tolerance
        strMidFFinalWeight = aCompound.MidFFinalWeight
        strMidFTEQFinalWeight = aCompound.MidFTEQFinalWeight
        blnMidFNonDetect = aCompound.MidFNonDetect
        blnMidFETEQ0 = aCompound.MidFETEQ0
        blnMidFETEQ05 = aCompound.MidFETEQ05
        blnMidFETEQLOD = aCompound.MidFETEQLOD
        strMidFFlags = aCompound.MidFFlags

        strChromLowContLim = aCompound.ChromLowContLim
        strChromUpContLim = aCompound.ChromUpContLim

        'TMPQNTRP 
        strTSignal = aCompound.TSignal
        strQ1Signal = aCompound.Q1Signal
        strQ2Signal = aCompound.Q2Signal
        strQ3Signal = aCompound.Q3Signal
        strTRatios = aCompound.TRatios
        strQ1Ratios = aCompound.Q1Ratios
        strQ2Ratios = aCompound.Q2Ratios
        strQ3Ratios = aCompound.Q3Ratios
        strQ1SLLim = aCompound.Q1SLLim
        strQ1SULim = aCompound.Q1SULim
        strQ2SLLim = aCompound.Q2SLLim
        strQ2SULim = aCompound.Q2SULim
        strQ3SLLim = aCompound.Q3SLLim
        strQ3SULim = aCompound.Q3SULim
        strTResp = aCompound.TResp
        strQ1Resp = aCompound.Q1Resp
        strQ2Resp = aCompound.Q2Resp
        strQ3Resp = aCompound.Q3Resp
        strTRT = aCompound.TRT
        strQ1RT = aCompound.Q1RT
        strQ2RT = aCompound.Q2RT
        strQ3RT = aCompound.Q3RT
        strRTLLim = aCompound.RTLLim
        strRTULim = aCompound.RTULim
        strTIntType = aCompound.TIntType
        strQ1IntType = aCompound.Q1IntType
        strQ2IntType = aCompound.Q2IntType
        strQ3IntType = aCompound.Q3IntType
        strTMPRelStdDev = aCompound.TMPRelStdDev
    End Sub
End Class
