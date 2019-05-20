Public Class Standard

    'Chemstation
    Private strName As String
    Private strRT As String
    Private strQIon As String
    Private strResponse As String
    Private strConc As String
    Private strUnits As String
    Private strDevMin As String
    Private strCasNum As String
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
    Private dblMidFTargetAmt As Double
    Private dblMidFQualAmt As Double
    Private dblMidFLoq As Double
    Private dblMidF13CAmt As Double
    Private dblMidF13CRecovery As Double
    Private blnMidF13CRecoveryInLim As Boolean
    Private dblMidFRF As Double
    Private strMidFFlags As String

    'MidlandChrom
    Private strMidCAdjConc As String
    Private strMidCAdjRLimit As String
    Private blnMidCExceedance As Boolean

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
        dblMidFIonRatio = -1.0
        blnMidFIonRatioInLim = False
        blnMidF13CRecoveryInLim = False
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
        blnKeep = True
        blnWriteToReport = True
        dblMidF13CAmt = -1
        dblMidF13CRecovery = -1
        intIndex = -1
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
    Public Property Keep() As Boolean
        Get
            Return blnKeep
        End Get
        Set(ByVal value As Boolean)
            blnKeep = value
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
    Public Property DevMin() As String
        Get
            Return strDevMin
        End Get
        Set(ByVal value As String)
            strDevMin = value
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
    Public Property MidCAdjConc() As String
        Get
            Return strMidCAdjConc
        End Get
        Set(ByVal value As String)
            strMidCAdjConc = value
        End Set
    End Property
    Public Property MidCAdjRLimit() As String
        Get
            Return strMidCAdjRLimit
        End Get
        Set(ByVal value As String)
            strMidCAdjRLimit = value
        End Set
    End Property
    Public Property MidCExeedance() As Boolean
        Get
            Return blnMidCExceedance
        End Get
        Set(ByVal value As Boolean)
            blnMidCExceedance = value
        End Set
    End Property
    Public Property MidFTargetAmt() As Double
        Get
            Return dblMidFTargetAmt
        End Get
        Set(ByVal value As Double)
            dblMidFTargetAmt = value
        End Set
    End Property
    Public Property MidFQualAmt() As Double
        Get
            Return dblMidFQualAmt
        End Get
        Set(ByVal value As Double)
            dblMidFQualAmt = value
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
    Public Property MidF13CRecovery() As Double
        Get
            Return dblMidF13CRecovery
        End Get
        Set(ByVal value As Double)
            dblMidF13CRecovery = value
        End Set
    End Property
    Public Property MidF13CRecoveryInLim() As Boolean
        Get
            Return blnMidF13CRecoveryInLim
        End Get
        Set(ByVal value As Boolean)
            blnMidF13CRecoveryInLim = value
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
    Public Property Index() As Integer
        Get
            Return intIndex
        End Get
        Set(ByVal value As Integer)
            intIndex = value
        End Set
    End Property
End Class
