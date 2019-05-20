'Class for Method-Instrument

Public Class mInstrument
    Private strName As String
    Private blnReviewed As Boolean
    Private dtReviewedDate As Date
    Public Property mCompoundList As New ArrayList
    Public Property mSurrogateList As New ArrayList
    Public Property mStandardList As New ArrayList

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

    Sub CopyMethodInfo(ByVal aExistInstrument As mInstrument)
        Dim aStandard As mStandard
        Dim aCompound As mCompound
        Dim aExistStandard As mStandard
        Dim aExistCompound As mCompound

        'Copy everything except calibration information
        For Each aExistStandard In aExistInstrument.mStandardList
            aStandard = New mStandard
            aStandard.Name = aExistStandard.Name
            aStandard.Type = aExistStandard.Type
            aStandard.Conc = aExistStandard.Conc
            aStandard.RecLowLim = aExistStandard.RecLowLim
            aStandard.RecUpLim = aExistStandard.RecLowLim
            aStandard.IonTarget = aExistStandard.IonTarget
            aStandard.IonQual = aExistStandard.IonQual
            aStandard.AbundTarget = aExistStandard.AbundTarget
            aStandard.AbundQual = aExistStandard.AbundQual
            mStandardList.Add(aStandard)
        Next
        For Each aExistCompound In aExistInstrument.mCompoundList
            aCompound = New mCompound
            aCompound.Name = aExistCompound.Name
            aCompound.Conc = aExistCompound.Conc
            aCompound.CS3Amt = aExistCompound.CS3Amt
            aCompound.TEF = aExistCompound.TEF
            aCompound.Ion = aExistCompound.Ion
            aCompound.Abundance = aExistCompound.Abundance
            aCompound.LCSLLim = aExistCompound.LCSLLim
            aCompound.LCSULim = aExistCompound.LCSULim
            aCompound.Assoc13C = aExistCompound.Assoc13C
            mCompoundList.Add(aCompound)
        Next

    End Sub
End Class
