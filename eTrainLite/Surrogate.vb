Public Class Surrogate
    Private strName As String
    Private strRT As String
    Private strQIon As String
    Private strResponse As String
    Private strConc As String
    Private strUnits As String
    Private strDevMin As String
    Private strSpkAmt As String
    Private strRecovery As String
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
    Private strCasNum As String

    'Chrom
    Private strChromLowContLim As String
    Private strChromUpContLim As String
    Private strChromLowICVLim As String
    Private strChromUpICVLim As String
    Private strChromLowCVSLim As String
    Private strChromUpCVSLim As String
    Private strChromLowMSLim As String
    Private strChromUpMSLim As String
    Private strChromLowLCSLim As String
    Private strChromUpLCSLim As String
    Private blnMethylated As Boolean

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
        blnWriteToReport = True
        blnWritten = False
        blnMethylated = False
        blnKeep = True
        strRecovery = "NA"
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
    Public Property CasNum() As String
        Get
            Return strCasNum
        End Get
        Set(ByVal value As String)
            strCasNum = value
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
    Public Property SpkAmt() As String
        Get
            Return strSpkAmt
        End Get
        Set(ByVal value As String)
            strSpkAmt = value
        End Set
    End Property
    Public Property Recovery() As String
        Get
            Return strRecovery
        End Get
        Set(ByVal value As String)
            strRecovery = value
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
    Public Property Methylated() As Boolean
        Get
            Return blnMethylated
        End Get
        Set(ByVal value As Boolean)
            blnMethylated = value
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

