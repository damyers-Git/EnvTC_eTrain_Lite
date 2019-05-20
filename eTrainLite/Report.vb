Imports Syncfusion.XlsIO

Public Class Report
	Private strType As String
	Private strSavLoc As String
	Private strRName As String
	Private strReportLimit As String
	Private strMatrix As String
	Private strPermit As String
	Private strMethod As String
	Private strInstrument As String

	'Sets/Gets
	Public Property Type() As String
		Get
			Return strType
		End Get
		Set(ByVal value As String)
			strType = value
		End Set
	End Property
	Public Property SavLoc() As String
		Get
			Return strSavLoc
		End Get
		Set(ByVal value As String)
			strSavLoc = value
		End Set
	End Property
	Public Property RName() As String
		Get
			Return strRName
		End Get
		Set(ByVal value As String)
			strRName = value
		End Set
	End Property
	Public Property ReportLimit() As String
		Get
			Return strReportLimit
		End Get
		Set(ByVal value As String)
			strReportLimit = value
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
	Public Property Permit() As String
		Get
			Return strPermit
		End Get
		Set(ByVal value As String)
			strPermit = value
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
	Public Property Instrument() As String
		Get
			Return strInstrument
		End Get
		Set(ByVal value As String)
			strInstrument = value
		End Set
	End Property

	'IDL Report for Midland HR
	Sub MidlandHRIDLReport()
		'Dim exEngine As New ExcelEngine
		'Dim exApp As IApplication
		'Dim workbook As IWorkbook
		'Dim worksheet As IWorksheet
		'Dim aSample As Sample





	End Sub


	'LCS Report for Midland FAST
	Sub MidlandFASTLCSReport()
		Dim exEngine As New ExcelEngine
		Dim exApp As IApplication
		Dim aSample As Sample
		Dim aLCSSample As Sample
		Dim aCompound1 As Compound
		Dim workbook As IWorkbook
		Dim worksheet As IWorksheet
		Dim fDialog As New FolderBrowserDialog
		Dim aMethod As Method
		Dim aInstrument As mInstrument
		Dim amCompound As mCompound
		Dim strSaveLoc As String
		Dim intCmpdCount As Integer
		Dim arrSpl() As String

		Try
			'Get LCS Sample
			aLCSSample = Nothing
			For Each aSample In GlobalVariables.ReportSamList
				If aSample.Include Then
					If InStr(aSample.DataFile, "P") Then
						aLCSSample = aSample
						'Begin creating report
						strSaveLoc = fDialog.SelectedPath
						exApp = exEngine.Excel
						workbook = exApp.Workbooks.Create(1)
						worksheet = workbook.Worksheets(0)
						worksheet.Range("B1").Value = "LCS Check Summary"
						worksheet.Range("A3").Value = "Data Path: " & aLCSSample.DataPath
						worksheet.Range("A4").Value = "Data File: " & aLCSSample.DataFile
						worksheet.Range("A5").Value = "Acq On: " & aLCSSample.AcqDate
						worksheet.Range("A6").Value = "Operator: " & aLCSSample.Analyst
						worksheet.Range("A7").Value = "Sample: " & aLCSSample.Name
						worksheet.Range("A8").Value = "Misc: " & aLCSSample.Misc
						worksheet.Range("A9").Value = "ALS Vial: " & aLCSSample.Vial & " Spike Multiplier: " & aLCSSample.Multiplier
						worksheet.Range("A11").Value = "Quant Time: " & aLCSSample.QuantTime
						worksheet.Range("A12").Value = "Quant Method: " & aLCSSample.QuantMethod
						worksheet.Range("A19").Value = "Tolerance = " & GlobalVariables.selMethod.RptTolerance
						worksheet.Range("A20").Value = "Analyte Name"
						worksheet.Range("B20").Value = "Amount Added (ppt)"
						worksheet.Range("C20").Value = "Recovered Amount (ppt)"
						worksheet.Range("D20").Value = "% Recovery"
						worksheet.Range("E20").Value = "Limit Low (ppt)"
						worksheet.Range("F20").Value = "Limit High (ppt)"
						worksheet.Range("G20").Value = "Comments"
						worksheet.Range("A20:G20").CellStyle.Font.Bold = True
						worksheet.Range("A20:G20").BorderAround()
						worksheet.Range("A20:G20").BorderInside()
						worksheet.Range("A20:G20").AutofitColumns()

						intCmpdCount = 0
						'Compound data

						For Each aCompound1 In aLCSSample.CompoundList
							If InStr(aCompound1.Name, "(INJ)") = 0 Then
								For Each aMethod In GlobalVariables.MethodList
									If aMethod.Name = GlobalVariables.selMethod.Name Then
										For Each aInstrument In aMethod.mInstrumentList
											If aInstrument.Name = GlobalVariables.selInstrument Then
												For Each amCompound In aInstrument.mCompoundList
													If aCompound1.Name = amCompound.Name Then
														If aCompound1.WriteToReport Then
															worksheet.Range("A" & CStr(21 + intCmpdCount)).Value = aCompound1.Name
															worksheet.Range("B" & CStr(21 + intCmpdCount)).Value = aCompound1.MidFLCSAmtAdded
															worksheet.Range("C" & CStr(21 + intCmpdCount)).Value = aCompound1.MidFLCSAmtRecovered
															worksheet.Range("D" & CStr(21 + intCmpdCount)).Value = aCompound1.MidFLCSPercRecovered
															worksheet.Range("E" & CStr(21 + intCmpdCount)).Value = amCompound.LCSLLim
															worksheet.Range("F" & CStr(21 + intCmpdCount)).Value = amCompound.LCSULim
															If CDbl(worksheet.Range("D" & CStr(21 + intCmpdCount)).Value) >= CDbl(worksheet.Range("E" & CStr(21 + intCmpdCount)).Value) And
																CDbl(worksheet.Range("D" & CStr(21 + intCmpdCount)).Value) <= CDbl(worksheet.Range("F" & CStr(21 + intCmpdCount)).Value) Then
																worksheet.Range("G" & CStr(21 + intCmpdCount)).Value = "Pass"
															Else
																worksheet.Range("G" & CStr(21 + intCmpdCount)).Value = "Fail"
															End If
															intCmpdCount = intCmpdCount + 1
														End If
													End If
												Next
											End If
										Next
									End If
								Next

							End If
						Next
						arrSpl = aLCSSample.DataFile.Split(".")

						'Start setting up save
						workbook.Version = ExcelVersion.Excel2010
						workbook.SaveAs(GlobalVariables.Report.SavLoc & "\" & GlobalVariables.Report.RName & ".xlsx")
						workbook.Close()
						exEngine.Dispose()

					End If
				End If

			Next
		Catch ex As Exception
			MsgBox("Error creating LCS Report" & vbCrLf &
			   "Sub Procedure: MidlandFASTLCSReport()" & vbCrLf &
			   "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
		End Try

	End Sub

	'CS3Report for Midland FAST
	Sub MidlandFASTCS3Report()
		Dim exEngine As New ExcelEngine
		Dim exApp As IApplication
		Dim aSample As Sample
		Dim aCS3Sample As Sample
		Dim aCompound As Compound
		Dim workbook As IWorkbook
		Dim worksheet As IWorksheet
		Dim fDialog As New FolderBrowserDialog
		Dim strSaveLoc As String
		Dim intCmpdCount As Integer
		Dim arrSpl() As String

		Try
			'Get CS3 Sample
			aCS3Sample = Nothing
			For Each aSample In GlobalVariables.ReportSamList
				If aSample.Include Then
					If InStr(aSample.DataFile, "CS3") Then
						aCS3Sample = aSample
						'Begin creating report
						strSaveLoc = fDialog.SelectedPath
						exApp = exEngine.Excel
						workbook = exApp.Workbooks.Create(1)
						worksheet = workbook.Worksheets(0)
						worksheet.Range("B1").Value = "CS3 Check Summary"
						worksheet.Range("A3").Value = "Data Path: " & aCS3Sample.DataPath
						worksheet.Range("A4").Value = "Data File: " & aCS3Sample.DataFile
						worksheet.Range("A5").Value = "Acq On: " & aCS3Sample.AcqDate
						worksheet.Range("A6").Value = "Operator: " & aCS3Sample.Analyst
						worksheet.Range("A7").Value = "Sample: " & aCS3Sample.Name
						worksheet.Range("A8").Value = "Misc: " & aCS3Sample.Misc
						worksheet.Range("A9").Value = "ALS Vial: " & aCS3Sample.Vial & " Spike Multiplier: " & aCS3Sample.Multiplier
						worksheet.Range("A11").Value = "Quant Time: " & aCS3Sample.QuantTime
						worksheet.Range("A12").Value = "Quant Method: " & aCS3Sample.QuantMethod
						worksheet.Range("A19").Value = "Tolerance = " & GlobalVariables.selMethod.RptTolerance
						worksheet.Range("A20").Value = "Analyte Name"
						worksheet.Range("B20").Value = "Total Amount (ng)"
						worksheet.Range("C20").Value = "Recovered Amount (ng)"
						worksheet.Range("D20").Value = "Limit Low (ppt)"
						worksheet.Range("E20").Value = "Limit High (ppt)"
						worksheet.Range("F20").Value = "Comments"
						worksheet.Range("A20:F20").CellStyle.Font.Bold = True
						'worksheet.Range("A20:F20").CellStyle.Borders.LineStyle = ExcelLineStyle.Thick

						intCmpdCount = 0
						'Compound data
						For Each aCompound In aCS3Sample.CompoundList
							If InStr(aCompound.Name, "(INJ)") = 0 Then
								If aCompound.WriteToReport Then
									worksheet.Range("A" & CStr(21 + intCmpdCount)).Value = aCompound.Name
									worksheet.Range("B" & CStr(21 + intCmpdCount)).Value = aCompound.MidFCS3TotalAmt
									worksheet.Range("C" & CStr(21 + intCmpdCount)).Value = aCompound.MidFCS3AmtRecovered
									worksheet.Range("D" & CStr(21 + intCmpdCount)).Value = aCompound.MidFCS3LowLim
									worksheet.Range("E" & CStr(21 + intCmpdCount)).Value = aCompound.MidFCS3HighLim
									If CDbl(worksheet.Range("C" & CStr(21 + intCmpdCount)).Value) >= CDbl(worksheet.Range("D" & CStr(21 + intCmpdCount)).Value) And
										CDbl(worksheet.Range("C" & CStr(21 + intCmpdCount)).Value) <= CDbl(worksheet.Range("E" & CStr(21 + intCmpdCount)).Value) Then
										worksheet.Range("F" & CStr(21 + intCmpdCount)).Value = "Pass"
									Else
										worksheet.Range("F" & CStr(21 + intCmpdCount)).Value = "Fail"
									End If
									intCmpdCount = intCmpdCount + 1
								End If
							End If
						Next
						arrSpl = aCS3Sample.DataFile.Split(".")

						'Start setting up save
						workbook.Version = ExcelVersion.Excel2010
						workbook.SaveAs(GlobalVariables.Report.SavLoc & "\" & GlobalVariables.Report.RName & ".xlsx")
						workbook.Close()
						exEngine.Dispose()

					End If
				End If

			Next
		Catch ex As Exception
			MsgBox("Error creating CS3 Report" & vbCrLf &
			   "Sub Procedure: MidlandFASTCS3Report()" & vbCrLf &
			   "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
		End Try

	End Sub

	'Final Data Report for Midland FAST
	Sub MidlandFASTFinalDataReport(ByVal strSISLoc As String)
		Dim exEngine As New ExcelEngine
		Dim exApp As IApplication
		Dim workbook As IWorkbook
		Dim worksheet As IWorksheet
		Dim aSIS As SIS
		Dim aSISSample As Sample
		Dim aSample As Sample
		Dim aCompound As Compound
		Dim aStandard As Standard
		Dim dt As Date = Date.Today
		Dim intCountCol As Integer
		Dim intCountRow As Integer
		Dim rng As String
		Dim strEARLNum As String

		strEARLNum = ""
		Try
			exApp = exEngine.Excel
			workbook = exApp.Workbooks.Create(1)
			workbook.StandardFont = "Tahoma"
			workbook.StandardFontSize = 8
			worksheet = workbook.Worksheets(0)
			worksheet.Name = "Final Data"

			'Begin sheet setup
			worksheet.Range("A2").Value = "Report #:"
			worksheet.Range("A3").Value = "Issued By:"
			worksheet.Range("C3").Value = "Date:"
			worksheet.Range("D3").Value = dt.ToString("MM/dd/yyyy")
			worksheet.Range("A5").Value = "Lab ID"
			worksheet.Range("B5").Value = "Sample Description"
			workbook.ActiveSheet.SetColumnWidth(2, 48)
			'ETEQ Factor
			worksheet.Range("F2:F3").Merge()
			worksheet.Range("F2").CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
			worksheet.Range("F2").CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
			worksheet.Range("F2").Value = "ETEQ Factor: "
			worksheet.Range("G2:H3").Merge()
			worksheet.Range("G2").CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
			worksheet.Range("G2").CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
			worksheet.Range("G2:H3").CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thin
			worksheet.Range("G2:H3").CellStyle.Borders(ExcelBordersIndex.EdgeTop).LineStyle = ExcelLineStyle.Thin
			worksheet.Range("G2:H3").CellStyle.Borders(ExcelBordersIndex.EdgeLeft).LineStyle = ExcelLineStyle.Thin
			worksheet.Range("G2:H3").CellStyle.Borders(ExcelBordersIndex.EdgeRight).LineStyle = ExcelLineStyle.Thin
			worksheet.Range("G2").Value = GlobalVariables.selMethod.ETEQ


			'Grab sample to write out headings
			aSample = GlobalVariables.ReportSamList(0)
			intCountCol = 3
			'Weight
			For Each aCompound In aSample.CompoundNameList
				If InStr(aCompound.Name, "13C") = 0 Then
					If aCompound.WriteToReport Then
						worksheet.Range(5, intCountCol).Value = aCompound.Name & " [ng/kg d.w.]"
						workbook.ActiveSheet.SetColumnWidth(intCountCol, 12)
						worksheet.Range(5, intCountCol + 1).Value = "flag"
						workbook.ActiveSheet.SetColumnWidth(intCountCol + 1, 5)
						intCountCol = intCountCol + 2
					End If
				End If
			Next
			worksheet.Range(5, intCountCol).Value = "Spl. Flag"
			intCountCol = intCountCol + 1
			'Weight TEQ
			For Each aCompound In aSample.CompoundNameList
				If InStr(aCompound.Name, "13C") = 0 Then
					If aCompound.WriteToReport Then
						worksheet.Range(5, intCountCol).Value = aCompound.Name & " [ng TEQ/kg d.w.]"
						workbook.ActiveSheet.SetColumnWidth(intCountCol, 15)
						worksheet.Range(5, intCountCol + 1).Value = "flag"
						workbook.ActiveSheet.SetColumnWidth(intCountCol + 1, 5)
						intCountCol = intCountCol + 2
					End If
				End If
			Next
			'ETEQ
			worksheet.Range(5, intCountCol).Value = "ETEQ (ND = 0)"
			worksheet.Range(5, intCountCol).CellStyle.Font.Bold = True
			workbook.ActiveSheet.SetColumnWidth(intCountCol, 15)
			intCountCol = intCountCol + 1
			worksheet.Range(5, intCountCol).Value = "ETEQ (ND = 0.5*LOD)"
			worksheet.Range(5, intCountCol).CellStyle.Font.Bold = True
			workbook.ActiveSheet.SetColumnWidth(intCountCol, 20)
			intCountCol = intCountCol + 1
			worksheet.Range(5, intCountCol).Value = "ETEQ (ND = LOD)"
			worksheet.Range(5, intCountCol).CellStyle.Font.Bold = True
			workbook.ActiveSheet.SetColumnWidth(intCountCol, 15)
			intCountCol = intCountCol + 1
			'13c Recovery
			For Each aStandard In aSample.InternalStdList
				If aStandard.WriteToReport Then
					worksheet.Range(5, intCountCol).Value = "Recovery " & aStandard.Name & " [%]"
					worksheet.Range(5, intCountCol).CellStyle.Font.Bold = True
					worksheet.Range(5, intCountCol).CellStyle.Font.Color = ExcelKnownColors.Blue_grey
					workbook.ActiveSheet.SetColumnWidth(intCountCol, 20)
					worksheet.Range(5, intCountCol + 1).Value = "flag"
					workbook.ActiveSheet.SetColumnWidth(intCountCol + 1, 5)
					intCountCol = intCountCol + 2
				End If
			Next
			worksheet.Range(5, intCountCol).Value = "Sample Date"
			workbook.ActiveSheet.SetColumnWidth(intCountCol, 20)
			'Formatting
			rng = "A5:" & worksheet.Range(5, intCountCol).AddressLocal
			worksheet.Range(rng).CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thin
			worksheet.Range(rng).CellStyle.Borders(ExcelBordersIndex.EdgeTop).LineStyle = ExcelLineStyle.Thin
			worksheet.Range(rng).CellStyle.WrapText = True
			workbook.ActiveSheet.SetRowHeight(5, 45)
			worksheet.Range(rng).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
			worksheet.Range(rng).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter

			intCountRow = 6
			'Gather SIS and Samples 
			For Each aSIS In GlobalVariables.SISList
				'Report #
				worksheet.Range("B2").Value = aSIS.ProjNum
				strEARLNum = aSIS.ProjNum
				worksheet.Range("B3").Value = aSIS.Reviewer
				For Each aSample In GlobalVariables.ReportSamList
					If aSample.Include Then
						For Each aSISSample In aSIS.SampleList
							'Match samples
							If aSample.LimsID = aSISSample.SISLabNum Then
								intCountCol = 1
								worksheet.Range(intCountRow, intCountCol).Value = aSISSample.SISLabNum
								worksheet.Range(intCountRow, intCountCol).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(intCountRow, intCountCol).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								intCountCol = intCountCol + 1
								worksheet.Range(intCountRow, intCountCol).Value = aSISSample.SISClientSampID
								worksheet.Range(intCountRow, intCountCol).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(intCountRow, intCountCol).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								intCountCol = intCountCol + 1
								For Each aCompound In aSample.CompoundNameList
									If aCompound.WriteToReport Then
										If InStr(aCompound.Name, "13C") = 0 Then
											If aCompound.MidFNonDetect Then
												worksheet.Range(intCountRow, intCountCol).Text = "< " & GlobalVariables.Calculations.FormatSF(aCompound.MidFFinalWeight)
												worksheet.Range(intCountRow, intCountCol).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
												worksheet.Range(intCountRow, intCountCol).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
											Else
												worksheet.Range(intCountRow, intCountCol).Text = GlobalVariables.Calculations.FormatSF(aCompound.MidFFinalWeight)
												worksheet.Range(intCountRow, intCountCol).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
												worksheet.Range(intCountRow, intCountCol).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
											End If
											'Flags
											worksheet.Range(intCountRow, intCountCol + 1).Value = aCompound.MidFFlags
											worksheet.Range(intCountRow, intCountCol + 1).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
											worksheet.Range(intCountRow, intCountCol + 1).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
											intCountCol = intCountCol + 2
										End If
									End If
								Next
								worksheet.Range(intCountRow, intCountCol).Value = ""
								intCountCol = intCountCol + 1
								For Each aCompound In aSample.CompoundNameList
									If aCompound.WriteToReport Then
										If InStr(aCompound.Name, "13C") = 0 Then
											If aCompound.MidFNonDetect Then
												worksheet.Range(intCountRow, intCountCol).Text = "< " & GlobalVariables.Calculations.FormatSF(aCompound.MidFTEQFinalWeight)
												worksheet.Range(intCountRow, intCountCol).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
												worksheet.Range(intCountRow, intCountCol).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
											Else
												worksheet.Range(intCountRow, intCountCol).Text = GlobalVariables.Calculations.FormatSF(aCompound.MidFTEQFinalWeight)
												worksheet.Range(intCountRow, intCountCol).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
												worksheet.Range(intCountRow, intCountCol).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
											End If
											worksheet.Range(intCountRow, intCountCol + 1).Value = ""
											intCountCol = intCountCol + 2
										End If
									End If
								Next
								'ETEQ to Integer only
								worksheet.Range(intCountRow, intCountCol).Text = GlobalVariables.Calculations.FormatSF(aSample.MidFETEQ0)
								worksheet.Range(intCountRow, intCountCol).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(intCountRow, intCountCol).CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent
								worksheet.Range(intCountRow, intCountCol).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								intCountCol = intCountCol + 1
								worksheet.Range(intCountRow, intCountCol).Text = GlobalVariables.Calculations.FormatSF(aSample.MidFETEQ05)
								worksheet.Range(intCountRow, intCountCol).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(intCountRow, intCountCol).CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent
								worksheet.Range(intCountRow, intCountCol).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								intCountCol = intCountCol + 1
								worksheet.Range(intCountRow, intCountCol).Text = GlobalVariables.Calculations.FormatSF(aSample.MidFETEQLOD)
								worksheet.Range(intCountRow, intCountCol).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(intCountRow, intCountCol).CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent
								worksheet.Range(intCountRow, intCountCol).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								intCountCol = intCountCol + 1
								For Each aStandard In aSample.InternalStdList
									If aStandard.WriteToReport Then
										If aStandard.MidF13CAmt = -1 Or aStandard.MidF13CRecovery = -1 Then
											worksheet.Range(intCountRow, intCountCol).Text = "NR"
										Else
											worksheet.Range(intCountRow, intCountCol).Text = GlobalVariables.Calculations.FormatSF(CStr(aStandard.MidF13CRecovery))
										End If
										worksheet.Range(intCountRow, intCountCol).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
										worksheet.Range(intCountRow, intCountCol).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
										worksheet.Range(intCountRow, intCountCol).CellStyle.ColorIndex = ExcelKnownColors.Pale_blue
										worksheet.Range(intCountRow, intCountCol + 1).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
										worksheet.Range(intCountRow, intCountCol + 1).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
										worksheet.Range(intCountRow, intCountCol + 1).Value = aStandard.MidFFlags
										intCountCol = intCountCol + 2
									End If
								Next
								If aSISSample.SISSampDateEnd <> CDate("1/1/1970") Then
									worksheet.Range(intCountRow, intCountCol).Value = aSISSample.SISSampDate & " - " & aSISSample.SISSampDateEnd
								Else
									worksheet.Range(intCountRow, intCountCol).Value = aSISSample.SISSampDate
								End If
								worksheet.Range(intCountRow, intCountCol).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(intCountRow, intCountCol).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
							End If
						Next
						intCountRow = intCountRow + 1
					End If

				Next
			Next
			workbook.ActiveSheet.SetColumnWidth(4, 10)
			workbook.ActiveSheet.SetColumnWidth(6, 12)

			'Copy QC sheet
			MidlandFASTQCSheet(workbook)

			'Copy SIS sheet
			MidlandFASTSISSheet(workbook, strSISLoc)

			'Copy Chains
			MidlandFASTCOCCopy(workbook, strSISLoc)

			'Start setting up save
			workbook.Version = ExcelVersion.Excel2010
			workbook.SaveAs(GlobalVariables.Report.SavLoc & "\" & strEARLNum & "_" & GlobalVariables.Report.RName & ".xlsx")
			workbook.Close()
			exEngine.Dispose()

		Catch ex As Exception
			MsgBox("Error creating Final Data Report" & vbCrLf &
			   "Sub Procedure: MidlandFASTFinalDataReport()" & vbCrLf &
			   "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
		End Try


	End Sub

	'SampleRPT.xls for Midland FAST
	Sub MidlandFASTSampleReport(ByVal strEMethod As String)
		Dim exEngine As New ExcelEngine
		Dim exApp As IApplication
		Dim aSample As Sample
		Dim aCS1Sample As Sample
		Dim aCompound1 As Compound
		Dim aCompound2 As Compound
		Dim aCS1Compound1 As Compound
		Dim aCS1Compound2 As Compound
		Dim aTCompound As Compound
		Dim workbook As IWorkbook
		Dim worksheet As IWorksheet
		Dim fDialog As New FolderBrowserDialog
		Dim flgs As String
		Dim strSaveLoc As String
		Dim arrSpl() As String
		Dim intCmpdCount As Integer

		Try
			exApp = exEngine.Excel

			'Get CS1(LOQ)
			aCS1Sample = Nothing
			For Each aSample In GlobalVariables.ReportSamList
				If InStr(aSample.DataFile, "CS1") Then
					aCS1Sample = aSample
				End If
			Next
			aSample = Nothing

			For Each aSample In GlobalVariables.ReportSamList
				intCmpdCount = 0
				strSaveLoc = fDialog.SelectedPath

				'reset written flag
				For Each aCompound1 In aSample.CompoundList
					aCompound1.Written = False
				Next

				workbook = exApp.Workbooks.Create(1)
				'Final sheet
				worksheet = workbook.Worksheets(0)
				worksheet.Name = "Final"
				worksheet.Range("B1").Value = "Sample Report"
				worksheet.Range("A3").Value = "Data Path: " & aSample.DataPath
				worksheet.Range("A4").Value = "Data File: " & aSample.DataFile
				worksheet.Range("A5").Value = "Acq On: " & aSample.AcqDate
				worksheet.Range("A6").Value = "Operator: " & aSample.Analyst
				worksheet.Range("A7").Value = "Sample: " & aSample.Name
				worksheet.Range("A8").Value = "Misc: " & aSample.Misc
				worksheet.Range("A9").Value = "ALS Vial: " & aSample.Vial & " Spike Multiplier: " & aSample.Multiplier
				worksheet.Range("A11").Value = "Quant Time: " & aSample.QuantTime
				worksheet.Range("A12").Value = "Quant Method: " & aSample.QuantMethod
				worksheet.Range("A13").Value = "eTrain Method: " & strEMethod
				worksheet.Range("A20").Value = "Analyte Name"
				worksheet.Range("B20").Value = "Amount (ppt)"
				worksheet.Range("C20").Value = "QC Flag"
				worksheet.Range("D20").Value = "LOQ (ppt)"
				'Compound data
				For Each aCompound1 In aSample.CompoundList
					If Not aCompound1.Written Then
						For Each aCompound2 In aSample.CompoundList
							If Trim(aCompound2.Name.Substring(0, InStrRev(aCompound2.Name, "(") - 1)) = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) And aCompound1.Name <> aCompound2.Name Then
								'Find CS1 compounds 
								For Each aCS1Compound1 In aCS1Sample.CompoundList
									If aCS1Compound1.Name = aCompound1.Name Then
										For Each aCS1Compound2 In aCS1Sample.CompoundList
											If aCS1Compound2.Name = aCompound2.Name Then
												'QC1, report average of two amounts
												worksheet.Range("A" & CStr(intCmpdCount + 21)).Value = aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)
												worksheet.Range("B" & CStr(intCmpdCount + 21)).Value = aCompound1.MidFReportedAmt

												'Flags flown
												flgs = ""
												If aCompound1.MidFQC1 Then
													If flgs = "" Then
														flgs = flgs & "1"
													Else
														flgs = flgs & ",1"
													End If
												End If
												If aCompound1.MidFQC2 Then
													If flgs = "" Then
														flgs = flgs & "2"
													Else
														flgs = flgs & ",2"
													End If
												End If
												If aCompound1.MidFQC3 Then
													If flgs = "" Then
														flgs = flgs & "3"
													Else
														flgs = flgs & ",3"
													End If
												End If
												If aCompound1.MidFIsTarget Then
													If aCompound1.MidFQC4 Then
														If flgs = "" Then
															flgs = flgs & "4"
														Else
															flgs = flgs & ",4"
														End If
													End If
													If aCompound2.MidFQC5 Then
														If flgs = "" Then
															flgs = flgs & "5"
														Else
															flgs = flgs & ",5"
														End If
													End If
												Else
													If aCompound2.MidFQC4 Then
														If flgs = "" Then
															flgs = flgs & "4"
														Else
															flgs = flgs & ",4"
														End If
													End If
													If aCompound1.MidFQC5 Then
														If flgs = "" Then
															flgs = flgs & "5"
														Else
															flgs = flgs & ",5"
														End If
													End If
												End If
												If aCompound1.MidFQC6 Then
													If flgs = "" Then
														flgs = flgs & "6"
													Else
														flgs = flgs & ",6"
													End If
												End If
												If aCompound1.MidFQC7 Then
													If flgs = "" Then
														flgs = flgs & "7"
													Else
														flgs = flgs & ",7"
													End If
												End If
												If aCompound1.MidFIsTarget Then
													If aCompound1.MidFQC8 Then
														If flgs = "" Then
															flgs = flgs & "8"
														Else
															flgs = flgs & ",8"
														End If
													End If
													If aCompound2.MidFQC9 Then
														If flgs = "" Then
															flgs = flgs & "9"
														Else
															flgs = flgs & ",9"
														End If
													End If
												Else
													If aCompound2.MidFQC8 Then
														If flgs = "" Then
															flgs = flgs & "8"
														Else
															flgs = flgs & ",8"
														End If
													End If
													If aCompound1.MidFQC9 Then
														If flgs = "" Then
															flgs = flgs & "9"
														Else
															flgs = flgs & ",9"
														End If
													End If
												End If
												If aCompound1.MidFQC10 Then
													If flgs = "" Then
														flgs = flgs & "10"
													Else
														flgs = flgs & ",10"
													End If
												End If
												If aCompound1.MidFQC11 Then
													If flgs = "" Then
														flgs = flgs & "11"
													Else
														flgs = flgs & ",11"
													End If
												End If
												If aCompound1.MidFQC12 Then
													If flgs = "" Then
														flgs = flgs & "12"
													Else
														flgs = flgs & ",12"
													End If
												End If
												worksheet.Range("C" & CStr(intCmpdCount + 21)).Text = flgs

												'Always report average of LOQ
												worksheet.Range("D" & CStr(intCmpdCount + 21)).Value = aCompound1.MidFReportedLOQAmt
												aCompound1.Written = True
												aCompound2.Written = True
												intCmpdCount = intCmpdCount + 1
											End If
										Next
									End If
								Next
							End If
						Next
					End If
				Next
				intCmpdCount = intCmpdCount + 2
				worksheet.Range("A" & CStr(intCmpdCount + 21)).Value = "QC Notes:"
				worksheet.Range("A" & CStr(intCmpdCount + 22)).Value = "1 - Isotope ion ratio is within EPA range(15%).   Reported value is the average conc of both ions."
				worksheet.Range("A" & CStr(intCmpdCount + 23)).Value = "2 - Isotope ion ratio is below or above the EPA range (15%).  Reported value is the lowest conc. of the two ions."
				worksheet.Range("A" & CStr(intCmpdCount + 24)).Value = "3 - Reported average value is less than the average LOD."
				worksheet.Range("A" & CStr(intCmpdCount + 25)).Value = "4 - Reported ion 1 concentration is below the ion 1 LOD."
				worksheet.Range("A" & CStr(intCmpdCount + 26)).Value = "5 - Reported ion 2 concentration is below the ion 2 LOD."
				worksheet.Range("A" & CStr(intCmpdCount + 27)).Value = "6 - Isotope ion ratio for the internal standard is outside the EPA range (+/-15%)."
				worksheet.Range("A" & CStr(intCmpdCount + 28)).Value = "7 - Internal standard peak manually integrated."
				worksheet.Range("A" & CStr(intCmpdCount + 29)).Value = "8 - Isotope ion 1 peak manually integrated."
				worksheet.Range("A" & CStr(intCmpdCount + 30)).Value = "9 - Isotope ion 2 peak manually integrated."
				'worksheet.Range("A" & CStr(intCmpdCount + 31)).Value = "10 - Internal standard recovery is outside of EPA range."
				worksheet.Range("A" & CStr(intCmpdCount + 31)).Value = "11 - Alternate 2378-TCDD Calculation used."
				worksheet.Range("A" & CStr(intCmpdCount + 32)).Value = "12 - One or more components were Non-Detect."

				'reset written flag
				For Each aCompound1 In aSample.CompoundList
					aCompound1.Written = False
				Next

				'Detail Sheet
				intCmpdCount = 0
				worksheet = workbook.Worksheets.Create()
				worksheet.Name = "Detail"
				worksheet.Range("B1").Value = "Sample Report"
				worksheet.Range("A20").Value = "Analyte Name"
				worksheet.Range("B20").Value = "Ion 1 Area"
				worksheet.Range("C20").Value = "Ion 1 Amount (ppt)"
				worksheet.Range("D20").Value = "Ion 1 LOQ (Area)"
				worksheet.Range("E20").Value = "Ion 1 LOQ (Amt)"
				worksheet.Range("F20").Value = "Ion 2 Area"
				worksheet.Range("G20").Value = "Ion 2 Amount (ppt)"
				worksheet.Range("H20").Value = "Ion 2 LOQ (Area)"
				worksheet.Range("I20").Value = "Ion 2 LOQ (Amt)"
				worksheet.Range("J20").Value = "Sample Ion Ratio"
				worksheet.Range("K20").Value = "Expected Ion Ratio"

				For Each aCompound1 In aSample.CompoundList
					If Not aCompound1.Written Then
						For Each aCompound2 In aSample.CompoundList
							If Trim(aCompound2.Name.Substring(0, InStrRev(aCompound2.Name, "(") - 1)) = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) And aCompound1.Name <> aCompound2.Name Then
								If aCompound1.MidFIsTarget Then
									worksheet.Range("A" & CStr(intCmpdCount + 21)).Value = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1))
									worksheet.Range("B" & CStr(intCmpdCount + 21)).Value = aCompound1.Response
									worksheet.Range("C" & CStr(intCmpdCount + 21)).Value = aCompound1.Conc
									For Each aCS1Compound In aCS1Sample.CompoundList
										If aCS1Compound.Name = aCompound1.Name Then
											worksheet.Range("D" & CStr(intCmpdCount + 21)).Value = aCS1Compound.Response
											worksheet.Range("E" & CStr(intCmpdCount + 21)).Value = aCompound1.MidFLoq
										End If
									Next
									worksheet.Range("F" & CStr(intCmpdCount + 21)).Value = aCompound2.Response
									worksheet.Range("G" & CStr(intCmpdCount + 21)).Value = aCompound2.Conc
									For Each aCS1Compound1 In aCS1Sample.CompoundList
										If aCS1Compound1.Name = aCompound2.Name Then
											worksheet.Range("H" & CStr(intCmpdCount + 21)).Value = aCS1Compound1.Response
											worksheet.Range("I" & CStr(intCmpdCount + 21)).Value = aCompound2.MidFLoq
										End If
									Next
									worksheet.Range("J" & CStr(intCmpdCount + 21)).Value = aCompound1.MidFIonRatio
									For Each aTCompound In GlobalVariables.TheoComps
										If aTCompound.Name = Trim(aCompound1.Name.Substring(0, InStrRev(aCompound1.Name, "(") - 1)) Then
											worksheet.Range("K" & CStr(intCmpdCount + 21)).Value = aTCompound.MidFIonRatio
										End If
									Next
									aCompound1.Written = True
									aCompound2.Written = True
									intCmpdCount = intCmpdCount + 1
								Else
									worksheet.Range("A" & CStr(intCmpdCount + 21)).Value = Trim(aCompound2.Name.Substring(0, InStrRev(aCompound2.Name, "(") - 1))
									worksheet.Range("B" & CStr(intCmpdCount + 21)).Value = aCompound2.Response
									worksheet.Range("C" & CStr(intCmpdCount + 21)).Value = aCompound2.Conc
									For Each aCS1Compound1 In aCS1Sample.CompoundList
										If aCS1Compound1.Name = aCompound2.Name Then
											worksheet.Range("D" & CStr(intCmpdCount + 21)).Value = aCS1Compound1.Response
											worksheet.Range("E" & CStr(intCmpdCount + 21)).Value = aCompound2.MidFLoq
										End If
									Next
									worksheet.Range("F" & CStr(intCmpdCount + 21)).Value = aCompound1.Response
									worksheet.Range("G" & CStr(intCmpdCount + 21)).Value = aCompound1.Conc
									For Each aCS1Compound In aCS1Sample.CompoundList
										If aCS1Compound.Name = aCompound1.Name Then
											worksheet.Range("H" & CStr(intCmpdCount + 21)).Value = aCS1Compound.Response
											worksheet.Range("I" & CStr(intCmpdCount + 21)).Value = aCompound1.MidFLoq
										End If
									Next
									worksheet.Range("J" & CStr(intCmpdCount + 21)).Value = aCompound2.MidFIonRatio
									For Each aTCompound In GlobalVariables.TheoComps
										If aTCompound.Name = Trim(aCompound2.Name.Substring(0, InStrRev(aCompound2.Name, "(") - 1)) Then
											worksheet.Range("K" & CStr(intCmpdCount + 21)).Value = aTCompound.MidFIonRatio
										End If
									Next
									aCompound1.Written = True
									aCompound2.Written = True
									intCmpdCount = intCmpdCount + 1
								End If
							End If
						Next
					End If
				Next

				arrSpl = aSample.DataFile.Split(".")

				'Start setting up save
				workbook.Version = ExcelVersion.Excel2010
				workbook.SaveAs(GlobalVariables.Report.SavLoc & "\" & arrSpl(0) & "_SampleReport.xlsx")
				workbook.Close()

				'If MsgBox("Would you like to open the newly generated report?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
				'    System.Diagnostics.Process.Start(GlobalVariables.Report.SavLoc & GlobalVariables.Report.RName& ".xlsx")
				'End If
			Next

			exEngine.Dispose()
		Catch ex As Exception
			MsgBox("Error creating Sample Report" & vbCrLf &
				"Sub Procedure: MidlandFASTSampleReport()" & vbCrLf &
				"Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
		End Try


	End Sub

	'Sub to copy QC sheet into a report for Midland FAST
	Sub MidlandFASTQCSheet(ByRef workbook As IWorkbook)
		Dim wksQC As IWorksheet

		Try
			wksQC = workbook.Worksheets.Create(workbook.Worksheets.Count - 1)
			wksQC.Name = "Qualifier Codes"
			wksQC.SetColumnWidth(1, 6)
			wksQC.SetColumnWidth(2, 152)
			wksQC.Range("B1").Value = "Dioxin/Furan Raw Data Annotation Codes"
			wksQC.Range("B1").CellStyle.Font.Bold = True
			wksQC.Range("B1").CellStyle.Font.Underline = True
			wksQC.Range("B1").CellStyle.Font.Size = 16
			wksQC.Range("A3").Value = "S/N"
			wksQC.Range("A5").Value = "A"
			wksQC.Range("A7").Value = "BA"
			wksQC.Range("A9").Value = "LR"
			wksQC.Range("A11").Value = "E"
			wksQC.Range("A13").Value = "EMPC"
			wksQC.Range("A15").Value = "F"
			wksQC.Range("A17").Value = "I"
			wksQC.Range("A19").Value = "HEX"
			wksQC.Range("A21").Value = "IS"
			wksQC.Range("A23").Value = "J"
			wksQC.Range("A25").Value = "M"
			wksQC.Range("A27").Value = "R"
			wksQC.Range("A29").Value = "S"
			wksQC.Range("A31").Value = "T"
			wksQC.Range("A33").Value = "X"
			wksQC.Range("A35").Value = "Y"
			wksQC.Range("A37").Value = "W"
			wksQC.Range("A39").Value = "Z"
			wksQC.Range("A41").Value = "NR"
			wksQC.Range("A43").Value = "P,T"
			wksQC.Range("A45").Value = "N"
			wksQC.Range("B3").Value = "Signal to Noise Ratio"
			wksQC.Range("B5").Value = "13C isotope recovery outside method criteria"
			wksQC.Range("B7").Value = "Manually integrated because initial baseline was determined incorrectly by the software"
			wksQC.Range("B9").Value = "Peak area outside the range of calibration curve but within the linear range of the instrument"
			wksQC.Range("B11").Value = "Excluded  - < 2.5 times signal to noise ratio"
			wksQC.Range("B13").Value = "Estimated Maximum Peak"
			wksQC.Range("B15").Value = "2378-TCDD quantified via 13C-2378-TCDF"
			wksQC.Range("B17").Value = "Excluded – interference"
			wksQC.Range("B19").Value = "Manually integrated to include both 1,2,3,4,7,8-HxCDF and 1,2,3,6,7,8-HxCDF as a single peak"
			wksQC.Range("B21").Value = "Excluded peak from measurement.  Isotope ratio does not match EPA Method 1613 ratios"
			wksQC.Range("B23").Value = "Reported value is below lower calibration limit"
			wksQC.Range("B25").Value = "Manually integrated because signal was above 2.5 S/N"
			wksQC.Range("B27").Value = "Excluded – retention time mismatch between internal standard and one or both native peaks"
			wksQC.Range("B29").Value = "Peak split during ITEF measurement – peak separation evident in unsmoothed data"
			wksQC.Range("B31").Value = "Manual integration of peak at retention time of 13C-standard (shift of elution time) "
			wksQC.Range("B33").Value = "Peaks annotated with X code were measured as part of " & Chr(34) & "Total" & Chr(34) & " calculation"
			wksQC.Range("B35").Value = "Isotope ion ratio is below or above the method range (15 %). Reported value is the lowest concentration of the two ions"
			wksQC.Range("B37").Value = "Isotope ion ratio for the internal standard is out of the method range (15 %)"
			wksQC.Range("B39").Value = "Reagent blank contains significant level of ETEQ"
			wksQC.Range("B41").Value = "Not Reported"
			wksQC.Range("B43").Value = "Potential bias due to PCP/TCP related PCDD/F, sample recommended for confirmation analysis with 1613b method"
			wksQC.Range("B45").Value = "Non-toxic congener distribution indicates potential bias due to PCP related PCDD/F, sample recommended for confirmation analysis with 1613b method"
		Catch ex As Exception
			MsgBox("Error creating QC sheet in Report" & vbCrLf &
				 "Sub Procedure: MidlandFASTQCSheet()" & vbCrLf &
				 "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
		End Try


	End Sub

	'Sub to copy SIS Sheet into a report for Midland FAST
	Sub MidlandFASTSISSheet(ByRef workbook As IWorkbook, ByVal strSISLoc As String)
		Dim exEngine As New ExcelEngine
		Dim exApp As IApplication
		Dim wkbSIS As IWorkbook
		Dim aSIS As New SIS
		Dim aSISwks As IWorksheet

		Try
			exApp = exEngine.Excel
			wkbSIS = exApp.Workbooks.Open(strSISLoc)

			wkbSIS.Version = ExcelVersion.Excel2013
			workbook.Version = ExcelVersion.Excel2013

			For Each aSISwks In wkbSIS.Worksheets
				If aSISwks.Name = "SIS" Then
					workbook.Worksheets.AddCopyAfter(wkbSIS.Worksheets(aSISwks.Index), workbook.Worksheets(0))
				End If
			Next

		Catch ex As Exception
			MsgBox("Error Copying SIS sheet from SIS to Report" & vbCrLf &
				 "Sub Procedure: MidlandFASTSISSheet()" & vbCrLf &
				 "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
		End Try

	End Sub

	'Sub to copy Chains from SIS into a report for Midland FAST
	Sub MidlandFASTCOCCopy(ByRef workbook As IWorkbook, ByVal strSISLoc As String)
		Dim exEngine As New ExcelEngine
		Dim exApp As IApplication
		Dim wkbSIS As IWorkbook
		Dim aSIS As New SIS
		Dim aSISwks As IWorksheet

		Try
			exApp = exEngine.Excel
			wkbSIS = exApp.Workbooks.Open(strSISLoc)

			wkbSIS.Version = ExcelVersion.Excel2013
			workbook.Version = ExcelVersion.Excel2013

			For Each aSISwks In wkbSIS.Worksheets
				If InStr(aSISwks.Name, "Coc") Then
					workbook.Worksheets.AddCopy(aSISwks)
				End If
			Next

		Catch ex As Exception
			MsgBox("Error Copying COC's from SIS to Report" & vbCrLf &
				 "Sub Procedure: MidlandFASTCOCCopy()" & vbCrLf &
				 "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
		End Try

	End Sub

	Function FreeportChromDUPReport(ByVal strLimitType As String) As Boolean
		Dim exEngine As New ExcelEngine
		Dim worksheet As IWorksheet
		Dim aSample As Sample
		Dim aSample2 As Sample
		Dim aCompound As Compound
		Dim blnGTG As Boolean
		Dim intCount As Integer
		Dim intCount2 As Integer
		Dim intSurrStart As Integer
		Dim exApp As IApplication
		Dim intWksCount As Integer
		exApp = exEngine.Excel

		Try
			'Grab samples for this report
			blnGTG = False
			intWksCount = 0
			GlobalVariables.TempReportSamList.Clear()
			For Each aSample In GlobalVariables.ReportSamList
				If aSample.Include Then
					If aSample.Type = "DUP" And Not aSample.Reported Then
						For Each aSample2 In GlobalVariables.ReportSamList
							If aSample2.Name = Trim(aSample.Name.Substring(0, aSample.Name.Length - 3)) And Not aSample.Reported And aSample.Methylated And aSample2.Methylated Then
								GlobalVariables.TempReportSamList.Add(aSample2)
								GlobalVariables.TempReportSamList.Add(aSample)
								intWksCount = intWksCount + 1
								blnGTG = True
								Exit For
							ElseIf aSample2.Name = Trim(aSample.Name.Substring(0, aSample.Name.Length - 3)) And Not aSample.Reported And Not aSample.Methylated And Not aSample2.Methylated Then
								GlobalVariables.TempReportSamList.Add(aSample2)
								GlobalVariables.TempReportSamList.Add(aSample)
								intWksCount = intWksCount + 1
								blnGTG = True
								Exit For
							End If
						Next
					End If
				End If
			Next

			'aSample is dup, aSample2 is original
			If blnGTG Then
				For u = 1 To intWksCount
					GlobalVariables.workbook.Worksheets.Create("DUP-" & CStr(u))
					For Each wks In GlobalVariables.workbook.Worksheets
						If wks.name = "DUP-" & CStr(u) Then
							worksheet = wks

							'Begin building sheet..
							worksheet.Range("B1:D1").Merge()
							worksheet.Range("B1").Value = "Sample Duplicate RPD Report"
							worksheet.Range("A3").Value = "LIMS #"
							worksheet.Range("A4").Value = "Sample Point"
							worksheet.Range("A5").Value = "Sample Date"
							worksheet.Range("A6").Value = "Analysis Date"
							worksheet.Range("A7").Value = "Analysis Time"
							worksheet.Range("A8").Value = "Data Folder Name"
							worksheet.Range("A9").Value = "Analyte/Parameter"
							worksheet.Range("A3:A9").BorderInside()
							worksheet.Range("A3:A9").BorderAround()
							worksheet.Range("A3:A9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							aSample = GlobalVariables.TempReportSamList.Item(0)

							'List Compounds
							intCount = 10
							For Each aCompound In aSample.CompoundList
								worksheet.Range(intCount, 1).Value = aCompound.Name
								intCount = intCount + 1
							Next
							intSurrStart = intCount
							worksheet.Range(10, 1, intCount, 1).BorderAround()

							intCount = 2 'column start

							worksheet.Range(3, intCount).Value = aSample.LimsID
							worksheet.Range(4, intCount).Value = aSample.Name
							worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
							worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
							worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
							worksheet.Range(8, intCount).Value = aSample.DataFile
							worksheet.Range(9, intCount).Value = "Amount (" & aSample.ReportedUnits & ") (From Chemstation Report)"
							worksheet.Range(9, intCount).CellStyle.WrapText = True
							worksheet.Range(3, intCount, 9, intCount).BorderAround()
							worksheet.Range(3, intCount, 9, intCount).BorderInside()
							worksheet.Range(3, intCount, 9, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							GlobalVariables.workbook.ActiveSheet.SetColumnWidthInPixels(2, 140)
							GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)

							'Begin analyte readout
							intCount2 = 10 'Row compounds start at
							For Each aCompound In aSample.CompoundList
								For i = 0 To aSample.CompoundList.Count
									If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
										worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
									End If
								Next
							Next
							worksheet.Range(10, intCount, intSurrStart, intCount).BorderAround()

							intCount = 3 'column start
							aSample = GlobalVariables.TempReportSamList.Item(1)
							worksheet.Range(3, intCount).Value = aSample.LimsID
							worksheet.Range(4, intCount).Value = aSample.Name
							worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
							worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
							worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
							worksheet.Range(8, intCount).Value = aSample.DataFile
							worksheet.Range(9, intCount).Value = "Amount (" & aSample.ReportedUnits & ") (From Chemstation Report)"
							worksheet.Range(9, intCount).CellStyle.WrapText = True
							worksheet.Range(3, intCount, 9, intCount).BorderAround()
							worksheet.Range(3, intCount, 9, intCount).BorderInside()
							worksheet.Range(3, intCount, 9, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)
							GlobalVariables.workbook.ActiveSheet.SetColumnWidthInPixels(3, 140)

							'Begin analyte readout
							intCount2 = 10 'Row compounds start at
							For Each aCompound In aSample.CompoundList
								For i = 0 To aSample.CompoundList.Count
									If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
										worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
									End If
								Next
							Next
							worksheet.Range(10, intCount, intSurrStart, intCount).BorderAround()

							'3 remaining columns

							intCount = 4

							worksheet.Range(9, intCount).Value = "% RPD"
							worksheet.Range(9, intCount + 1).Value = "RPD Limit"
							worksheet.Range(9, intCount + 2).Value = "Pass/Fail"
							worksheet.Range(3, intCount, 9, intCount + 2).BorderInside()
							worksheet.Range(3, intCount, 9, intCount + 2).BorderAround()
							worksheet.Range(3, intCount, 9, intCount + 2).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							aSample = GlobalVariables.TempReportSamList.Item(0)
							'Begin analyte readout
							intCount2 = 10 'Row compounds start at
							For Each aCompound In aSample.CompoundList
								For i = 0 To aSample.CompoundList.Count
									If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
										worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.ChromRPD)
										worksheet.Range(intCount2 + i, intCount + 1).Value = "30"
										If aCompound.ChromRPD <> "N/A" Then
											If CDbl(aCompound.ChromRPD) <= CDbl(worksheet.Range(intCount2 + i, intCount + 1).Value) Then
												worksheet.Range(intCount2 + i, intCount + 2).Value = "Pass"
											Else
												worksheet.Range(intCount2 + i, intCount + 2).Value = "Fail"
											End If
										Else
											worksheet.Range(intCount2 + i, intCount + 2).Value = "N/A"
										End If
									End If
								Next
							Next
							worksheet.Range(10, intCount, intSurrStart, intCount).BorderAround()
							worksheet.Range(10, intCount + 1, intSurrStart, intCount + 1).BorderAround()
							worksheet.Range(10, intCount + 1, intSurrStart, intCount + 2).BorderAround()
							worksheet.Range(1, 2, intSurrStart, intCount + 2).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
							worksheet.Range(1, 2, intSurrStart, intCount + 2).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
							worksheet.Range(1, 1, intSurrStart, intCount + 2).CellStyle.Font.Bold = True
							worksheet.Range(1, 1, intSurrStart, intCount + 2).CellStyle.Font.Size = 8
							worksheet.Range(1, 1, intSurrStart, intCount + 2).CellStyle.Font.FontName = "Arial"
							worksheet.Range(1, 1, intSurrStart, intCount + 2).AutofitColumns()

							'Clear out Samples so it is not reported twice
							aSample = GlobalVariables.TempReportSamList.Item(0)
							GlobalVariables.TempReportSamList.Remove(aSample)
							aSample = GlobalVariables.TempReportSamList.Item(0)
							aSample.Reported = True
							GlobalVariables.TempReportSamList.Remove(aSample)
						End If
					Next
				Next
				Return True
			Else
				Return False
			End If


		Catch ex As Exception
			MsgBox("Error creating DUP Report" & vbCrLf &
			   "Sub Procedure: FreeportChromDUPReport()" & vbCrLf &
			   "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
		End Try

	End Function

	Function FreeportChromMSReport(ByVal strLimitType As String) As Boolean
		Dim exEngine As New ExcelEngine
		Dim worksheet As IWorksheet
		Dim aSample As Sample
		Dim aSample2 As Sample
		Dim aSample3 As Sample
		Dim aCompound As Compound
		Dim aSurrogate As Surrogate
		Dim blnGTG As Boolean
		Dim intCount As Integer
		Dim intCount2 As Integer
		Dim intSurrStart As Integer
		Dim exApp As IApplication
		Dim intWksCount As Integer
		Dim blnMSD As Boolean
		Dim intMSDCounter As Integer
		exApp = exEngine.Excel

		Try
			'Grab samples for this report
			blnGTG = False
			intWksCount = 0
			GlobalVariables.TempReportSamList.Clear()
			For Each aSample In GlobalVariables.ReportSamList
				If aSample.Include Then
					If aSample.Type = "MS" Then
						For Each aSample2 In GlobalVariables.ReportSamList
							If aSample2.Name = Trim(aSample.Name.Substring(0, aSample.Name.Length - 2)) And Not aSample.Reported And aSample.Methylated And aSample2.Methylated Then
								For Each aSample3 In GlobalVariables.ReportSamList
									If aSample3.Type = "MSD" And InStr(aSample3.Name, aSample2.Name) And Not aSample3.Reported And aSample3.Methylated And aSample2.Methylated Then
										GlobalVariables.TempReportSamList.Add(aSample2)
										GlobalVariables.TempReportSamList.Add(aSample)
										GlobalVariables.TempReportSamList.Add(aSample3)
										blnMSD = True
										intWksCount = intWksCount + 1
										blnGTG = True
										Exit For
									End If
								Next
								If Not blnMSD Then
									GlobalVariables.TempReportSamList.Add(aSample2)
									GlobalVariables.TempReportSamList.Add(aSample)
									intWksCount = intWksCount + 1
									blnGTG = True
									Exit For
								Else
									Exit For
								End If
							ElseIf aSample2.Name = Trim(aSample.Name.Substring(0, aSample.Name.Length - 2)) And Not aSample.Reported And Not aSample.Methylated And Not aSample2.Methylated Then
								For Each aSample3 In GlobalVariables.ReportSamList
									If aSample3.Type = "MSD" And InStr(aSample3.Name, aSample2.Name) And Not aSample3.Reported And Not aSample3.Methylated And Not aSample2.Methylated Then
										GlobalVariables.TempReportSamList.Add(aSample2)
										GlobalVariables.TempReportSamList.Add(aSample)
										GlobalVariables.TempReportSamList.Add(aSample3)
										blnMSD = True
										intWksCount = intWksCount + 1
										blnGTG = True
										Exit For
									End If
								Next
								If Not blnMSD Then
									GlobalVariables.TempReportSamList.Add(aSample2)
									GlobalVariables.TempReportSamList.Add(aSample)
									intWksCount = intWksCount + 1
									blnGTG = True
									Exit For
								Else
									Exit For
								End If
							End If
						Next
					End If
				End If

			Next

			If blnGTG Then
				'Matrix switch
				For u = 1 To intWksCount
					blnMSD = False
					GlobalVariables.workbook.Worksheets.Create("MS-" & CStr(u))
					For Each wks In GlobalVariables.workbook.Worksheets
						If wks.name = "MS-" & CStr(u) Then
							worksheet = wks
							aSample = GlobalVariables.TempReportSamList.Item(0)
							If aSample.Matrix = "W" Then
								aSample = Nothing
								aSample2 = Nothing
								'Begin building sheet..
								worksheet.Range("D1:F1").Merge()
								worksheet.Range("D1").Value = "Spike Recovery Report"
								worksheet.Range("A3").Value = "LIMS #"
								worksheet.Range("A4").Value = "Sample Point"
								worksheet.Range("A5").Value = "Sample Date"
								worksheet.Range("A6").Value = "Analysis Date"
								worksheet.Range("A7").Value = "Analysis Time"
								worksheet.Range("A8").Value = "Data Folder Name"
								worksheet.Range("A9").Value = "Analyte/Parameter"
								worksheet.Range("A3:A9").BorderInside()
								worksheet.Range("A3:A9").BorderAround()
								worksheet.Range("A3:A9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								aSample = GlobalVariables.TempReportSamList.Item(0)

								'List Compounds
								intCount = 10
								For Each aCompound In aSample.CompoundList
									worksheet.Range(intCount, 1).Value = aCompound.Name
									intCount = intCount + 1
								Next
								worksheet.Range(10, 1, intCount, 1).BorderAround()

								'List Surrogates
								intCount = intCount + 1
								worksheet.Range(intCount, 1).Value = "Surrogate Recovery (%)"
								worksheet.Range(intCount, 2).BorderAround()
								worksheet.Range(intCount, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
								intCount = intCount + 1
								intSurrStart = intCount
								For Each aSurrogate In aSample.SurrogateList
									worksheet.Range(intCount, 1).Value = aSurrogate.Name
									intCount = intCount + 1
								Next
								worksheet.Range(intSurrStart, 1, intCount, 1).BorderAround()
								worksheet.Range(intSurrStart - 1, 1, intCount, 2).BorderAround()

								intCount = 2 'column start
								aSample = GlobalVariables.TempReportSamList.Item(0)
								worksheet.Range(3, intCount).Value = aSample.LimsID
								worksheet.Range(4, intCount).Value = aSample.Name
								worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
								worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
								worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
								worksheet.Range(8, intCount).Value = aSample.DataFile
								worksheet.Range(9, intCount).Value = "Amount (" & aSample.ReportedUnits & ") (From Chemstation Report) x Dil. Factor"
								'Begin analyte readout
								intCount2 = 10 'Row compounds start at
								For Each aCompound In aSample.CompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											If IsNumeric(aCompound.Conc) Then
												worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF(CStr(CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)))
											Else
												worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
											End If
										End If
									Next
								Next
								worksheet.Range(9, intCount).CellStyle.WrapText = True
								worksheet.Range(3, intCount, 9, intCount).BorderAround()
								worksheet.Range(3, intCount, 9, intCount).BorderInside()
								worksheet.Range(3, intCount, 9, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)
								'Surrogates
								For Each aSurrogate In aSample.SurrogateList
									For i = 0 To aSample.SurrogateList.Count
										If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
											worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery)
										End If
									Next
								Next
								worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
								worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 5).BorderAround()

								GlobalVariables.ReportSamList.Remove(aSample)

								intCount = 3 'column start

								'If MSD then run twice, if not, then just once
								aSample = GlobalVariables.TempReportSamList.Item(0)
								aSample2 = GlobalVariables.TempReportSamList.Item(1)
								If InStr(aSample2.Name, aSample.Name) And aSample.Type = "MS" And aSample2.Type = "MSD" Then
									blnMSD = True
								Else
									blnMSD = False
								End If
								aSample = Nothing
								aSample2 = Nothing
								If blnMSD Then
									intMSDCounter = 1
								Else
									intMSDCounter = 0
								End If
								For n = 0 To intMSDCounter
									aSample = GlobalVariables.TempReportSamList.Item(0 + n)
									worksheet.Range(3, intCount, 3, intCount + 4).Merge()
									worksheet.Range(3, intCount).Value = aSample.LimsID
									worksheet.Range(4, intCount, 4, intCount + 4).Merge()
									worksheet.Range(4, intCount).Value = aSample.Name
									worksheet.Range(5, intCount, 5, intCount + 4).Merge()
									worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
									worksheet.Range(6, intCount, 6, intCount + 4).Merge()
									worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
									worksheet.Range(7, intCount, 7, intCount + 4).Merge()
									worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
									worksheet.Range(8, intCount, 8, intCount + 4).Merge()
									worksheet.Range(8, intCount).Value = aSample.DataFile
									worksheet.Range(9, intCount).Value = "Amount " & aSample.Units & " (From Chemstation Report x Dil.Factor)"
									GlobalVariables.workbook.ActiveSheet.SetColumnWidthInPixels(intCount, 78)
									worksheet.Range(9, intCount).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 1).Value = "Recovered Spiked Amount (" & aSample.ReportedUnits & ")"
									worksheet.Range(9, intCount + 1).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 2).Value = "Corrected Spiked amount (" & aSample.ReportedUnits & ")"
									worksheet.Range(9, intCount + 2).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 3).Value = "% Recovery"
									worksheet.Range(9, intCount + 3).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 4).Value = "Recovery Limit Check"
									worksheet.Range(9, intCount + 4).CellStyle.WrapText = True
									worksheet.Range(3, intCount, 9, intCount + 4).BorderAround()
									worksheet.Range(3, intCount, 9, intCount + 4).BorderInside()
									worksheet.Range(3, intCount, 9, intCount + 4).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)

									'Begin analyte readout
									intCount2 = 10 'Row compounds start at
									For Each aCompound In aSample.CompoundList
										For i = 0 To aSample.CompoundList.Count
											If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
												If IsNumeric(aCompound.Conc) Then
													worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF(CStr(CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)))
												Else
													worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
												End If
												worksheet.Range(intCount2 + i, intCount + 1).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
												worksheet.Range(intCount2 + i, intCount + 2).Text = GlobalVariables.Calculations.FormatSF(aCompound.ChromCorrectedSpike)
												worksheet.Range(intCount2 + i, intCount + 3).Text = GlobalVariables.Calculations.FormatSF(aCompound.ChromSpikeRecovery)
												If aCompound.ChromSpikePass Then
													worksheet.Range(intCount2 + i, intCount + 4).Value = "Passed"
												Else
													worksheet.Range(intCount2 + i, intCount + 4).Value = "Failed"
												End If
												If Not aCompound.ChromSpikePass Then
													worksheet.Range(intCount2 + i, intCount + 4).CellStyle.Color = System.Drawing.Color.FromArgb(197, 217, 241)
												End If
											End If
										Next
									Next
									worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
									worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
									worksheet.Range(10, intCount + 2, intSurrStart - 1, intCount + 2).BorderAround()
									worksheet.Range(10, intCount + 3, intSurrStart - 1, intCount + 3).BorderAround()
									worksheet.Range(10, intCount + 4, intSurrStart - 1, intCount + 4).BorderAround()

									'Surrogate write out
									For Each aSurrogate In aSample.SurrogateList
										For i = 0 To aSample.SurrogateList.Count
											If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
												worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery)
											End If
										Next
									Next
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
									worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 4).BorderAround()
									intCount = intCount + 5
								Next
								worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount + 2).BorderAround()

								'3 remaining columns

								worksheet.Range(9, intCount).Value = "% RPD"
								worksheet.Range(9, intCount + 1).Value = "Recovery Limit (%)"
								worksheet.Range(9, intCount + 2).Value = "RPD Limit (%)"
								worksheet.Range(3, intCount, 9, intCount + 2).BorderInside()
								worksheet.Range(3, intCount, 9, intCount + 2).BorderAround()
								worksheet.Range(3, intCount, 9, intCount + 2).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								aSample = GlobalVariables.TempReportSamList.Item(0)
								'Begin analyte readout
								intCount2 = 10 'Row compounds start at
								For Each aCompound In aSample.CompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.ChromRPD)
											worksheet.Range(intCount2 + i, intCount + 1).Value = aCompound.ChromLowMSLim & "-" & aCompound.ChromUpMSLim
											worksheet.Range(intCount2 + i, intCount + 2).Value = aCompound.ChromRPDLimit
										End If
									Next
								Next
								worksheet.Range(9, intCount, intSurrStart - 1, intCount).BorderAround()
								worksheet.Range(9, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
								worksheet.Range(9, intCount + 1, intSurrStart - 1, intCount + 2).BorderAround()
								worksheet.Range(10, 4, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.WrapText = True
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.Bold = True
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.Size = 8
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.FontName = "Arial"
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).AutofitColumns()
								For n = 0 To intMSDCounter
									GlobalVariables.TempReportSamList.RemoveAt(0)
								Next
							ElseIf aSample.Matrix = "S" Then
								aSample = Nothing
								aSample2 = Nothing

								'Begin building sheet..
								worksheet.Range("D1:F1").Merge()
								worksheet.Range("D1").Value = "Spike Recovery Report"
								worksheet.Range("A3").Value = "LIMS #"
								worksheet.Range("A4").Value = "Sample Point"
								worksheet.Range("A5").Value = "Sample Date"
								worksheet.Range("A6").Value = "Analysis Date"
								worksheet.Range("A7").Value = "Analysis Time"
								worksheet.Range("A8").Value = "Data Folder Name"
								worksheet.Range("A9").Value = "Analyte/Parameter"
								worksheet.Range("A3:A9").BorderInside()
								worksheet.Range("A3:A9").BorderAround()
								worksheet.Range("A3:A9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								aSample = GlobalVariables.TempReportSamList.Item(0)

								'List Compounds
								intCount = 10
								For Each aCompound In aSample.CompoundList
									worksheet.Range(intCount, 1).Value = aCompound.Name
									intCount = intCount + 1
								Next
								worksheet.Range(10, 1, intCount + 1, 1).BorderAround()

								'List Surrogates
								intCount = intCount + 1
								worksheet.Range(intCount, 1).Value = "Surrogate Recovery (%)"
								worksheet.Range(intCount, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
								intCount = intCount + 1
								intSurrStart = intCount
								For Each aSurrogate In aSample.SurrogateList
									worksheet.Range(intCount, 1).Value = aSurrogate.Name
									intCount = intCount + 1
								Next
								worksheet.Range(intSurrStart - 1, 1, intCount, 1).BorderAround()
								worksheet.Range(intSurrStart - 1, 1, intCount, 4).BorderAround()

								intCount = 2 'column start
								aSample = GlobalVariables.TempReportSamList.Item(0)
								worksheet.Range(3, intCount, 3, intCount + 2).Merge()
								worksheet.Range(3, intCount).Value = aSample.LimsID
								worksheet.Range(4, intCount, 4, intCount + 2).Merge()
								worksheet.Range(4, intCount).Value = aSample.Name
								worksheet.Range(5, intCount, 5, intCount + 2).Merge()
								worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
								worksheet.Range(6, intCount, 6, intCount + 2).Merge()
								worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
								worksheet.Range(7, intCount, 7, intCount + 2).Merge()
								worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
								worksheet.Range(8, intCount, 8, intCount + 2).Merge()
								worksheet.Range(8, intCount).Value = aSample.DataFile
								worksheet.Range(9, intCount).Value = "Amount (" & aSample.ReportedUnits & ") (From Chemstation Report)"
								worksheet.Range(9, intCount).CellStyle.WrapText = True
								worksheet.Range(9, intCount + 1).Value = "Factor (" & aSample.ReportedUnits & " to (ug/Kg) of Sample)"
								worksheet.Range(9, intCount + 1).CellStyle.WrapText = True
								worksheet.Range(9, intCount + 2).Value = "Amount (ug/Kg)"
								worksheet.Range(9, intCount + 2).CellStyle.WrapText = True
								worksheet.Range(3, intCount, 9, intCount + 2).BorderAround()
								worksheet.Range(3, intCount, 9, intCount + 2).BorderInside()
								worksheet.Range(3, intCount, 9, intCount + 2).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(7, 106)

								'Begin analyte readout
								intCount2 = 10 'Row compounds start at
								For Each aCompound In aSample.CompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
											worksheet.Range(intCount2 + i, intCount + 1).Value = aSample.DilutionFactor
											If IsNumeric(aCompound.Conc) Then
												worksheet.Range(intCount2 + i, intCount + 2).Text = GlobalVariables.Calculations.FormatSF((CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)))
											Else
												worksheet.Range(intCount2 + i, intCount + 2).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
											End If
										End If
									Next
								Next
								worksheet.Range(9, intCount).CellStyle.WrapText = True
								worksheet.Range(3, intCount, 9, intCount).BorderAround()
								worksheet.Range(3, intCount, 9, intCount).BorderInside()
								worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
								worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
								worksheet.Range(10, intCount + 2, intSurrStart - 1, intCount + 2).BorderAround()
								worksheet.Range(3, intCount, 9, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)
								'Surrogates
								For Each aSurrogate In aSample.SurrogateList
									For i = 0 To aSample.SurrogateList.Count
										If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
											worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery)
										End If
									Next
								Next
								worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
								worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 3).BorderAround()

								GlobalVariables.TempReportSamList.Remove(aSample)

								intCount = 5 'column start

								'If MSD then run twice, if not, then just once
								aSample = GlobalVariables.TempReportSamList.Item(0)
								aSample2 = GlobalVariables.TempReportSamList.Item(1)
								If InStr(aSample2.Name, aSample.Name) And aSample.Type = "MS" And aSample2.Type = "MSD" Then
									blnMSD = True
								Else
									blnMSD = False
								End If
								aSample = Nothing
								aSample2 = Nothing
								If blnMSD Then
									intMSDCounter = 1
								Else
									intMSDCounter = 0
								End If
								For n = 0 To intMSDCounter
									aSample = GlobalVariables.TempReportSamList.Item(0 + n)
									worksheet.Range(3, intCount, 3, intCount + 7).Merge()
									worksheet.Range(3, intCount).Value = aSample.LimsID
									worksheet.Range(4, intCount, 4, intCount + 7).Merge()
									worksheet.Range(4, intCount).Value = aSample.Name
									worksheet.Range(5, intCount, 5, intCount + 7).Merge()
									worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
									worksheet.Range(6, intCount, 6, intCount + 7).Merge()
									worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
									worksheet.Range(7, intCount, 7, intCount + 7).Merge()
									worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
									worksheet.Range(8, intCount, 8, intCount + 7).Merge()
									worksheet.Range(8, intCount).Value = aSample.DataFile
									worksheet.Range(9, intCount).Value = "Amount " & aSample.Units & " (From Chemstation Report)"
									worksheet.Range(9, intCount).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 1).Value = "Factor (" & aSample.Units & " to (ug/Kg) of Sample)"
									worksheet.Range(9, intCount + 1).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 2).Value = "Amount (ug/Kg)"
									worksheet.Range(9, intCount + 2).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 3).Value = "Rec. Spiked Amount (ug/Kg) of Sample"
									worksheet.Range(9, intCount + 3).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 4).Value = "Rec. Spiked Amount (" & aSample.ReportedUnits & ")"
									worksheet.Range(9, intCount + 4).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 5).Value = "Corrected Spiked Amount (" & aSample.ReportedUnits & ")"
									worksheet.Range(9, intCount + 5).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 6).Value = "% Recovery"
									worksheet.Range(9, intCount + 6).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 7).Value = "Recovery Limit Check"
									worksheet.Range(9, intCount + 7).CellStyle.WrapText = True
									worksheet.Range(3, intCount, 9, intCount + 7).BorderAround()
									worksheet.Range(3, intCount, 9, intCount + 7).BorderInside()
									worksheet.Range(3, intCount, 9, intCount + 7).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)

									'Begin analyte readout
									intCount2 = 10 'Row compounds start at
									For Each aCompound In aSample.CompoundList
										For i = 0 To aSample.CompoundList.Count
											If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
												worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
												worksheet.Range(intCount2 + i, intCount + 1).Value = GlobalVariables.Calculations.FormatSF(aSample.DilutionFactor)
												If IsNumeric(aCompound.Conc) Then
													worksheet.Range(intCount2 + i, intCount + 2).Text = GlobalVariables.Calculations.FormatSF((CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)))
												Else
													worksheet.Range(intCount2 + i, intCount + 2).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
												End If
												worksheet.Range(intCount2 + i, intCount + 3).Text = GlobalVariables.Calculations.FormatSF(CStr(CDbl(worksheet.Range(intCount2 + i, intCount + 2).Value) - CDbl(worksheet.Range(intCount2 + i, 4).Value)))
												worksheet.Range(intCount2 + i, intCount + 4).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
												worksheet.Range(intCount2 + i, intCount + 5).Text = GlobalVariables.Calculations.FormatSF(aCompound.ChromCorrectedSpike)
												worksheet.Range(intCount2 + i, intCount + 6).Text = GlobalVariables.Calculations.FormatSF(aCompound.ChromSpikeRecovery)
												If aCompound.ChromSpikePass Then
													worksheet.Range(intCount2 + i, intCount + 7).Text = "Passed"
												Else
													worksheet.Range(intCount2 + i, intCount + 7).Text = "Failed"
												End If
												If Not aCompound.ChromSpikePass Then
													worksheet.Range(intCount2 + i, intCount + 7).CellStyle.Color = System.Drawing.Color.FromArgb(197, 217, 241)
												End If
											End If
										Next
									Next
									worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
									worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
									worksheet.Range(10, intCount + 2, intSurrStart - 1, intCount + 2).BorderAround()
									worksheet.Range(10, intCount + 3, intSurrStart - 1, intCount + 3).BorderAround()
									worksheet.Range(10, intCount + 4, intSurrStart - 1, intCount + 4).BorderAround()
									worksheet.Range(10, intCount + 5, intSurrStart - 1, intCount + 5).BorderAround()
									worksheet.Range(10, intCount + 6, intSurrStart - 1, intCount + 6).BorderAround()
									worksheet.Range(10, intCount + 7, intSurrStart - 1, intCount + 7).BorderAround()

									'Surrogate write out
									For Each aSurrogate In aSample.SurrogateList
										For i = 0 To aSample.SurrogateList.Count
											If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
												worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
											End If
										Next
									Next
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
									worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 7).BorderAround()
									intCount = intCount + 8
								Next
								worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount + 2).BorderAround()

								'3 remaining columns

								worksheet.Range(9, intCount).Value = "% RPD"
								worksheet.Range(9, intCount + 1).Value = "Recovery Limit (%)"
								worksheet.Range(9, intCount + 2).Value = "RPD Limit (%)"
								worksheet.Range(3, intCount, 9, intCount + 2).BorderInside()
								worksheet.Range(3, intCount, 9, intCount + 2).BorderAround()
								worksheet.Range(3, intCount, 9, intCount + 2).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								aSample = GlobalVariables.TempReportSamList.Item(0)
								'Begin analyte readout
								intCount2 = 10 'Row compounds start at
								For Each aCompound In aSample.CompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.ChromRPD)
											worksheet.Range(intCount2 + i, intCount + 1).Value = aCompound.ChromLowContLim & "-" & aCompound.ChromUpContLim
											worksheet.Range(intCount2 + i, intCount + 2).Value = aCompound.ChromRPDLimit
										End If
									Next
								Next
								worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
								worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
								worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 2).BorderAround()
								worksheet.Range(11, 4, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.WrapText = True
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.Bold = True
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.Size = 8
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.FontName = "Arial"
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).AutofitColumns()
								For n = 0 To intMSDCounter
									GlobalVariables.TempReportSamList.RemoveAt(0)
								Next
							End If
						End If
					Next
				Next
				Return True
			Else
				MsgBox("Error generating report!" & vbCrLf &
					  "Sub Procedure: FreeportChromMSReport()" & vbCrLf &
					  "Logic Error: Could not find MS/MSD pair to generate report.", "(╯°□°)╯︵ ┻━┻")
				Return False
			End If
		Catch ex As Exception
			MsgBox("Error generating report!" & vbCrLf &
					  "Sub Procedure: FreeportChromMSReport()" & vbCrLf &
					  "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
			Return False
		End Try
	End Function

	'Freeport LCS Report
	Function FreeportChromLCSReport() As Boolean
		Dim exEngine As New ExcelEngine
		Dim worksheet As IWorksheet
		Dim aSample As Sample
		Dim aSample2 As Sample
		Dim aCompound As Compound
		Dim aSurrogate As Surrogate
		Dim blnGTG As Boolean
		Dim intCount As Integer
		Dim intCount2 As Integer
		Dim intTotalCols As Integer
		Dim intSurrStart As Integer
		Dim intWksCount As Integer
		Dim blnLCSD As Boolean
		Dim intLCSDCounter As Integer
		Dim exApp As IApplication
		exApp = exEngine.Excel

		Try
			'Grab samples for this report
			'blnGTG = False
			'GlobalVariables.ReportSamList.Clear()
			'For Each aSample In GlobalVariables.SampleList
			'    If aSample.Type = "LCS" Then
			'        GlobalVariables.ReportSamList.Add(aSample)
			'        blnGTG = True
			'        Exit For
			'    End If
			'Next


			'Grab samples for this report
			blnGTG = False
			intWksCount = 0
			GlobalVariables.TempReportSamList.Clear()
			For Each aSample In GlobalVariables.ReportSamList
				If aSample.Include Then
					If aSample.Type = "LCS" Then
						'Look for LCSD
						blnLCSD = False
						For Each aSample2 In GlobalVariables.ReportSamList
							If aSample2.Type = "LCSD" And Not aSample2.Reported And aSample.Methylated And aSample2.Methylated Then
								aSample.Reported = True
								aSample2.Reported = True
								GlobalVariables.TempReportSamList.Add(aSample)
								GlobalVariables.TempReportSamList.Add(aSample2)
								blnLCSD = True
								intWksCount = intWksCount + 1
								blnGTG = True
								Exit For
							ElseIf aSample2.Type = "LCSD" And Not aSample2.Reported And Not aSample.Methylated And Not aSample2.Methylated Then
								aSample.Reported = True
								aSample2.Reported = True
								GlobalVariables.TempReportSamList.Add(aSample)
								GlobalVariables.TempReportSamList.Add(aSample2)
								blnLCSD = True
								intWksCount = intWksCount + 1
								blnGTG = True
								Exit For
							End If
						Next
						If Not blnLCSD And Not aSample.Reported Then
							aSample.Reported = True
							GlobalVariables.TempReportSamList.Add(aSample)
							intWksCount = intWksCount + 1
							blnGTG = True
						End If
					End If
				End If

			Next

			If blnGTG Then
				'Reset Reported 
				For Each aSample In GlobalVariables.TempReportSamList
					aSample.Reported = False
				Next
				For u = 1 To intWksCount
					GlobalVariables.workbook.Worksheets.Create("LCS-" & CStr(u))
					For Each wks In GlobalVariables.workbook.Worksheets
						If wks.name = "LCS-" & CStr(u) Then
							worksheet = wks
							'Begin building sheet..
							worksheet.Range("D1:F1").Merge()
							worksheet.Range("D1").Value = "LCS Recovery Report"
							worksheet.Range("A3").Value = "LIMS #"
							worksheet.Range("A4").Value = "Sample Point"
							worksheet.Range("A5").Value = "Sample Date"
							worksheet.Range("A6").Value = "Analysis Date"
							worksheet.Range("A7").Value = "Analysis Time"
							worksheet.Range("A8").Value = "Data Folder Name"
							worksheet.Range("A9").Value = "Analyte/Parameter"
							worksheet.Range("A3:A9").BorderInside()
							worksheet.Range("A3:A9").BorderAround()
							worksheet.Range("A3:A9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							aSample = GlobalVariables.TempReportSamList.Item(0)

							'List Compounds
							intTotalCols = 10
							For Each aCompound In aSample.CompoundList
								worksheet.Range(intTotalCols, 1).Value = aCompound.Name
								intTotalCols = intTotalCols + 1
							Next

							'List Surrogates
							If aSample.SurrogateList.Count > 0 Then
								intTotalCols = intTotalCols + 1
								worksheet.Range(intTotalCols, 1).Value = "Surrogate Recovery (%)"
								worksheet.Range(intTotalCols, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
								intTotalCols = intTotalCols + 1
								intSurrStart = intTotalCols
								For Each aSurrogate In aSample.SurrogateList
									worksheet.Range(intTotalCols, 1).Value = aSurrogate.Name
									intTotalCols = intTotalCols + 1
								Next
								worksheet.Range(intSurrStart, 1, intTotalCols, 1).BorderAround()
								worksheet.Range(10, 1, intSurrStart - 2, 1).BorderAround()
							Else
								intSurrStart = intTotalCols
								worksheet.Range(10, 1, intTotalCols - 1, 1).BorderAround()
							End If

							intCount = 2 'column start

							If GlobalVariables.TempReportSamList.Count > 1 Then
								aSample = GlobalVariables.TempReportSamList.Item(0)
								aSample2 = GlobalVariables.TempReportSamList.Item(1)
								If InStr(aSample2.Name, aSample.Name) And aSample.Type = "LCS" And aSample2.Type = "LCSD" Then
									blnLCSD = True
								Else
									blnLCSD = False
								End If
								aSample = Nothing
								aSample2 = Nothing
								If blnLCSD Then
									intLCSDCounter = 1
								Else
									intLCSDCounter = 0
								End If
							Else
								intLCSDCounter = 0
							End If


							For n = 0 To intLCSDCounter
								aSample = GlobalVariables.TempReportSamList.Item(n)
								worksheet.Range(3, intCount, 3, intCount + 4).Merge()
								worksheet.Range(3, intCount).Value = aSample.LimsID
								worksheet.Range(4, intCount, 4, intCount + 4).Merge()
								worksheet.Range(4, intCount).Value = aSample.Name
								worksheet.Range(5, intCount, 5, intCount + 4).Merge()
								worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
								worksheet.Range(6, intCount, 6, intCount + 4).Merge()
								worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
								worksheet.Range(7, intCount, 7, intCount + 4).Merge()
								worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
								worksheet.Range(8, intCount, 8, intCount + 4).Merge()
								worksheet.Range(8, intCount).Value = aSample.DataFile
								worksheet.Range(9, intCount).Value = "Amount " & aSample.Units & " (From Chemstation Report x Dil.Factor)"
								GlobalVariables.workbook.ActiveSheet.SetColumnWidthInPixels(intCount, 78)
								worksheet.Range(9, intCount).CellStyle.WrapText = True
								worksheet.Range(9, intCount + 1).Value = "Recovered Spiked Amount (" & aSample.ReportedUnits & ")"
								worksheet.Range(9, intCount + 1).CellStyle.WrapText = True
								worksheet.Range(9, intCount + 2).Value = "Corrected Spiked amount (" & aSample.ReportedUnits & ")"
								worksheet.Range(9, intCount + 2).CellStyle.WrapText = True
								worksheet.Range(9, intCount + 3).Value = "% Recovery"
								worksheet.Range(9, intCount + 3).CellStyle.WrapText = True
								worksheet.Range(9, intCount + 4).Value = "Recovery Limit Check"
								worksheet.Range(9, intCount + 4).CellStyle.WrapText = True
								worksheet.Range(3, intCount, 9, intCount + 4).BorderAround()
								worksheet.Range(3, intCount, 9, intCount + 4).BorderInside()
								worksheet.Range(3, intCount, 9, intCount + 4).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)

								'Begin analyte readout
								intCount2 = 10 'Row compounds start at
								For Each aCompound In aSample.CompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											If IsNumeric(aCompound.Conc) Then
												worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF((CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)))
											Else
												worksheet.Range(intCount2 + i, intCount).Text = "N.D."
											End If
											worksheet.Range(intCount2 + i, intCount + 1).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
											worksheet.Range(intCount2 + i, intCount + 2).Text = GlobalVariables.Calculations.FormatSF(aCompound.ChromCorrectedSpike)
											worksheet.Range(intCount2 + i, intCount + 3).Text = GlobalVariables.Calculations.FormatSF(aCompound.ChromSpikeRecovery)
											If aCompound.ChromSpikePass Then
												worksheet.Range(intCount2 + i, intCount + 4).Text = "Passed"
											Else
												worksheet.Range(intCount2 + i, intCount + 4).Text = "Failed"
											End If
											If Not aCompound.ChromSpikePass Then
												worksheet.Range(intCount2 + i, intCount + 4).CellStyle.Color = System.Drawing.Color.FromArgb(197, 217, 241)
											End If
										End If
									Next
								Next
								If aSample.SurrogateList.Count > 0 Then
									worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
									worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
									worksheet.Range(10, intCount + 2, intSurrStart - 1, intCount + 2).BorderAround()
									worksheet.Range(10, intCount + 3, intSurrStart - 1, intCount + 3).BorderAround()
									worksheet.Range(10, intCount + 4, intSurrStart - 1, intCount + 4).BorderAround()
								Else
									worksheet.Range(10, intCount, intTotalCols - 1, intCount).BorderAround()
									worksheet.Range(10, intCount + 1, intTotalCols - 1, intCount + 1).BorderAround()
									worksheet.Range(10, intCount + 2, intTotalCols - 1, intCount + 2).BorderAround()
									worksheet.Range(10, intCount + 3, intTotalCols - 1, intCount + 3).BorderAround()
									worksheet.Range(10, intCount + 4, intTotalCols - 1, intCount + 4).BorderAround()
								End If


								'Surrogate write out
								If aSample.SurrogateList.Count > 0 Then
									For Each aSurrogate In aSample.SurrogateList
										For i = 0 To aSample.SurrogateList.Count
											If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
												worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
											End If
										Next
									Next
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
									worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 4).BorderAround()
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
								End If
								intCount = intCount + 5
							Next


							'3 remaining columns

							worksheet.Range(9, intCount).Value = "Recovery Limit (%)"
							worksheet.Range(3, intCount, 9, intCount).BorderInside()
							worksheet.Range(3, intCount, 9, intCount).BorderAround()
							worksheet.Range(3, intCount, 9, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							aSample = GlobalVariables.TempReportSamList.Item(0)
							'Begin analyte readout
							intCount2 = 10 'Row compounds start at
							For Each aCompound In aSample.CompoundList
								For i = 0 To aSample.CompoundList.Count
									If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
										worksheet.Range(intCount2 + i, intCount).Value = aCompound.ChromLowLCSLim & "-" & aCompound.ChromUpLCSLim
									End If
								Next
							Next
							intCount2 = intSurrStart
							For Each aSurrogate In aSample.SurrogateList
								For i = 0 To aSample.SurrogateList.Count
									If worksheet.Range(intCount2 + i, 1).Value = aSurrogate.Name Then
										worksheet.Range(intCount2 + i, intCount).Value = aSurrogate.ChromLowLCSLim & "-" & aSurrogate.ChromUpLCSLim
									End If
								Next
							Next
							If aSample.SurrogateList.Count > 0 Then
								worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
								worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
								worksheet.Range(11, 4, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.WrapText = True
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Bold = True
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Size = 8
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.FontName = "Arial"
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).AutofitColumns()
							Else
								worksheet.Range(10, intCount, intTotalCols - 1, intCount).BorderAround()
								worksheet.Range(11, 4, intTotalCols, intCount).CellStyle.WrapText = True
								worksheet.Range(1, 2, intTotalCols, intCount).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(1, 2, intTotalCols, intCount).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								worksheet.Range(1, 1, intTotalCols, intCount).CellStyle.Font.Bold = True
								worksheet.Range(1, 1, intTotalCols, intCount).CellStyle.Font.Size = 8
								worksheet.Range(1, 1, intTotalCols, intCount).CellStyle.Font.FontName = "Arial"
								worksheet.Range(1, 1, intTotalCols, intCount).AutofitColumns()
							End If
							For n = 0 To intLCSDCounter
								GlobalVariables.TempReportSamList.RemoveAt(0)
							Next
						End If
					Next
				Next
				Return True

			Else
				MsgBox("Error generating report!" & vbCrLf &
					  "Sub Procedure: FreeportChromLCSReport()" & vbCrLf &
					  "Logic Error: Could not find LCS to generate report.", "(╯°□°)╯︵ ┻━┻")
				Return False
			End If
		Catch ex As Exception
			MsgBox("Error generating report!" & vbCrLf &
					  "Sub Procedure: FreeportChromLCSReport()" & vbCrLf &
					  "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
			Return False
		End Try


	End Function

	Function FreeportChromMBReport(ByVal strPath As String, ByVal strLimitType As String) As Boolean
		Dim worksheet As IWorksheet
		Dim aSample As Sample
		Dim aCompound As Compound
		Dim intCount As Integer
		Dim intCount2 As Integer
		Dim intTotalRows As Integer
		Dim blnGTG As Boolean
		Dim exEngine As New ExcelEngine
		Dim exApp As IApplication
		Dim intWksCount As Integer
		exApp = exEngine.Excel

		Try
			'Import limits for MB
			intWksCount = 0
			If GlobalVariables.Import.FreeportChromBuildMBCompoundList(strPath) Then
				'Grab sample for this report
				blnGTG = False
				GlobalVariables.TempReportSamList.Clear()
				For Each aSample In GlobalVariables.ReportSamList
					If aSample.Include Then
						If aSample.Type = "MB" Then
							GlobalVariables.TempReportSamList.Add(aSample)
							intWksCount = intWksCount + 1
							blnGTG = True
						End If
					End If

				Next

				If blnGTG Then
					For u = 1 To intWksCount
						GlobalVariables.workbook.Worksheets.Create("MB-" & CStr(u))
						For Each wks In GlobalVariables.workbook.Worksheets
							If wks.name = "MB-" & CStr(u) Then
								worksheet = wks
								'Begin building sheet..
								worksheet.Range("B1:C1").Merge()
								worksheet.Range("B1").Value = GlobalVariables.strFreeportAnalysis & " Daily Blank Report"
								worksheet.Range("A3").Value = "LIMS #"
								worksheet.Range("A4").Value = "Sample Point"
								worksheet.Range("A5").Value = "Sample Date"
								worksheet.Range("A6").Value = "Analysis Date"
								worksheet.Range("A7").Value = "Analysis Time"
								worksheet.Range("A8").Value = "Data Folder Name"
								worksheet.Range("A9").Value = "Dilution Factor"
								worksheet.Range("A10").Value = "Analyte/Parameter"
								worksheet.Range("A3:A10").BorderInside()
								worksheet.Range("A3:A10").BorderAround()
								worksheet.Range("A3:A10").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								aSample = GlobalVariables.TempReportSamList.Item(0)

								'List Compounds
								intCount = 11
								For Each aCompound In GlobalVariables.FreeportMBCompoundList
									worksheet.Range(intCount, 1).Value = aCompound.Name
									worksheet.Range(intCount, 3).Value = aCompound.ChromMBLim
									intCount = intCount + 1
								Next
								intTotalRows = intCount - 1
								worksheet.Range(11, 1, intTotalRows, 1).BorderAround()

								intCount = 2 'column start

								worksheet.Range(3, intCount).Value = aSample.LimsID
								worksheet.Range(4, intCount).Value = aSample.Name
								worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
								worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
								worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
								worksheet.Range(8, intCount).Value = aSample.DataFile
								worksheet.Range(9, intCount).Value = aSample.DilutionFactor
								worksheet.Range(10, intCount).Value = "Amount (" & aSample.ReportedUnits & ") (From Chemstation Report)"
								worksheet.Range(10, intCount).CellStyle.WrapText = True
								worksheet.Range(3, intCount, 10, intCount).BorderAround()
								worksheet.Range(3, intCount, 10, intCount).BorderInside()
								worksheet.Range(3, intCount, 10, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(10, 106)

								'Begin analyte readout
								intCount2 = 11 'Row compounds start at
								For Each aCompound In aSample.CompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
										End If
									Next
								Next
								worksheet.Range(11, intCount, intTotalRows, intCount).BorderAround()

								'3 remaining columns

								intCount = 3

								worksheet.Range(10, intCount).Value = strLimitType & "Limit UCL " & aSample.Units
								worksheet.Range(10, intCount + 1).Value = "Result"
								worksheet.Range(3, intCount, 10, intCount + 1).BorderInside()
								worksheet.Range(3, intCount, 10, intCount + 1).BorderAround()
								worksheet.Range(3, intCount, 10, intCount + 1).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								aSample = GlobalVariables.ReportSamList.Item(0)
								'Begin analyte readout
								intCount2 = 11 'Row compounds start at
								For Each aCompound In GlobalVariables.FreeportMBCompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											If worksheet.Range(intCount2 + i, 2).Value <> "N.D." And worksheet.Range(intCount2 + i, 2).Value <> "" Then
												If CDbl(worksheet.Range(intCount2 + i, 2).Value) <= CDbl(worksheet.Range(intCount2 + i, intCount).Value) Then
													worksheet.Range(intCount2 + i, intCount + 1).Value = "Pass"
												Else
													worksheet.Range(intCount2 + i, intCount + 1).Value = "Fail"
												End If
											Else
												worksheet.Range(intCount2 + i, intCount + 1).Value = "N/A"
											End If
										End If
									Next
								Next
								worksheet.Range(11, intCount, intTotalRows, intCount).BorderAround()
								worksheet.Range(11, intCount + 1, intTotalRows, intCount + 1).BorderAround()
								worksheet.Range(1, 2, intTotalRows, intCount + 1).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(1, 2, intTotalRows, intCount + 1).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								worksheet.Range(1, 1, intTotalRows, intCount + 1).CellStyle.Font.Bold = True
								worksheet.Range(1, 1, intTotalRows, intCount + 1).CellStyle.Font.Size = 8
								worksheet.Range(1, 1, intTotalRows, intCount + 1).CellStyle.Font.FontName = "Arial"
								worksheet.Range(1, 1, intTotalRows, intCount + 1).AutofitColumns()
								'Clear out Samples so it is not reported twice
								aSample = GlobalVariables.ReportSamList.Item(0)
								GlobalVariables.ReportSamList.Remove(aSample)
							End If
						Next
					Next
					Return True
				Else
					Return False
				End If
			Else
				MsgBox("Error generating report!" & vbCrLf &
						"Sub Procedure: FreeportChromMBReport()" & vbCrLf &
						"Logic Error: Could not import Method Blank compound list", MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
				Return False
			End If
		Catch ex As Exception
			MsgBox("Error generating report!" & vbCrLf &
					  "Sub Procedure: FreeportChromMBReport()" & vbCrLf &
					  "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
			Return False
		End Try

	End Function

	Function FreeportChromCVSReport() As Boolean
		Dim exEngine As New ExcelEngine
		Dim worksheet As IWorksheet
		Dim aSample As Sample
		Dim aSample2 As Sample
		Dim aCompound As Compound
		Dim aSurrogate As Surrogate
		Dim blnGTG As Boolean
		Dim intCount As Integer
		Dim intCount2 As Integer
		Dim intTotalCols As Integer
		Dim intSurrStart As Integer
		Dim exApp As IApplication
		exApp = exEngine.Excel

		Try
			'Grab samples for this report
			blnGTG = False
			GlobalVariables.TempReportSamList.Clear()
			For Each aSample In GlobalVariables.ReportSamList
				If aSample.Include Then
					If aSample.Type = "CVS" Then
						GlobalVariables.TempReportSamList.Add(aSample)
						blnGTG = True
						Exit For
					End If
				End If

			Next
			If blnGTG Then

				'Check for LCS
				For Each aSample In GlobalVariables.ReportSamList
					If aSample.Include Then
						If aSample.Type = "LCS" Then
							GlobalVariables.TempReportSamList.Add(aSample)
						End If
					End If

				Next
				aSample = Nothing
				aSample2 = Nothing
				GlobalVariables.workbook.Worksheets.Create("CVS Report")
				For Each wks In GlobalVariables.workbook.Worksheets
					If wks.name = "CVS Report" Then
						worksheet = wks

						'Begin building sheet..
						worksheet.Range("B1:F1").Merge()
						worksheet.Range("B1").Value = "CVS Report"
						worksheet.Range("A3").Value = "LIMS #"
						worksheet.Range("A4").Value = "Sample Point"
						worksheet.Range("A5").Value = "Sample Date"
						worksheet.Range("A6").Value = "Analysis Date"
						worksheet.Range("A7").Value = "Analysis Time"
						worksheet.Range("A8").Value = "Data Folder Name"
						worksheet.Range("A9").Value = "Analyte/Parameter"
						worksheet.Range("A3:A9").BorderInside()
						worksheet.Range("A3:A9").BorderAround()
						worksheet.Range("A3:A9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
						aSample = GlobalVariables.TempReportSamList.Item(0)

						'List Compounds
						intTotalCols = 10
						For Each aCompound In aSample.CompoundList
							worksheet.Range(intTotalCols, 1).Value = aCompound.Name
							intTotalCols = intTotalCols + 1
						Next

						'List Surrogates
						If aSample.SurrogateList.Count > 0 Then
							intTotalCols = intTotalCols + 1
							worksheet.Range(intTotalCols, 1).Value = "Surrogate Recovery (%)"
							worksheet.Range(intTotalCols, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
							intTotalCols = intTotalCols + 1
							intSurrStart = intTotalCols
							For Each aSurrogate In aSample.SurrogateList
								worksheet.Range(intTotalCols, 1).Value = aSurrogate.Name
								intTotalCols = intTotalCols + 1
							Next
							worksheet.Range(intSurrStart, 1, intTotalCols, 1).BorderAround()
							worksheet.Range(8, 1, intSurrStart - 2, 1).BorderAround()
						Else
							intSurrStart = intTotalCols
							worksheet.Range(8, 1, intTotalCols - 1, 1).BorderAround()
						End If

						intCount = 2 'column start
						worksheet.Range(3, intCount, 3, intCount + 3).Merge()
						worksheet.Range(3, intCount).Value = aSample.LimsID
						worksheet.Range(4, intCount, 4, intCount + 3).Merge()
						worksheet.Range(4, intCount).Value = aSample.Name
						worksheet.Range(5, intCount, 5, intCount + 3).Merge()
						worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
						worksheet.Range(6, intCount, 6, intCount + 3).Merge()
						worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
						worksheet.Range(7, intCount, 7, intCount + 3).Merge()
						worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
						worksheet.Range(8, intCount, 8, intCount + 3).Merge()
						worksheet.Range(8, intCount).Value = aSample.DataFile
						worksheet.Range(9, intCount).Value = "CVS Conc"
						GlobalVariables.workbook.ActiveSheet.SetColumnWidthInPixels(intCount, 78)
						worksheet.Range(9, intCount + 1).Value = "LCL (" & aSample.ReportedUnits & ")"
						worksheet.Range(9, intCount + 1).CellStyle.WrapText = True
						worksheet.Range(9, intCount + 2).Value = "UCL (" & aSample.ReportedUnits & ")"
						worksheet.Range(9, intCount + 2).CellStyle.WrapText = True
						worksheet.Range(9, intCount + 3).Value = "Status Check"
						worksheet.Range(9, intCount + 3).CellStyle.WrapText = True
						worksheet.Range(3, intCount, 9, intCount + 3).BorderAround()
						worksheet.Range(3, intCount, 9, intCount + 3).BorderInside()
						worksheet.Range(3, intCount, 9, intCount + 3).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
						GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)

						'Begin analyte readout
						intCount2 = 10 'Row compounds start at
						For Each aCompound In aSample.CompoundList
							For i = 0 To aSample.CompoundList.Count
								If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
									If IsNumeric(aCompound.Conc) Then
										worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
									Else
										worksheet.Range(intCount2 + i, intCount).Value = "N.D."
									End If
									If aCompound.ChromLowCVSLim = "" Then
										worksheet.Range(intCount2 + i, intCount + 1).Value = aCompound.ChromLowContLim
										worksheet.Range(intCount2 + i, intCount + 2).Value = aCompound.ChromUpContLim
									Else
										worksheet.Range(intCount2 + i, intCount + 1).Value = aCompound.ChromLowCVSLim
										worksheet.Range(intCount2 + i, intCount + 2).Value = aCompound.ChromUpCVSLim
									End If
									If aCompound.ChromCVSPass Then
										worksheet.Range(intCount2 + i, intCount + 3).Value = "Passed"
									Else
										worksheet.Range(intCount2 + i, intCount + 3).Value = "Failed"
									End If
									If Not aCompound.ChromCVSPass Then
										worksheet.Range(intCount2 + i, intCount + 3).CellStyle.Color = System.Drawing.Color.FromArgb(197, 217, 241)
									End If
								End If
							Next
						Next
						If aSample.SurrogateList.Count > 0 Then
							worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
							worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
							worksheet.Range(10, intCount + 2, intSurrStart - 1, intCount + 2).BorderAround()
							worksheet.Range(10, intCount + 3, intSurrStart - 1, intCount + 3).BorderAround()
						Else
							worksheet.Range(10, intCount, intTotalCols - 1, intCount).BorderAround()
							worksheet.Range(10, intCount + 1, intTotalCols - 1, intCount + 1).BorderAround()
							worksheet.Range(10, intCount + 2, intTotalCols - 1, intCount + 2).BorderAround()
							worksheet.Range(10, intCount + 3, intTotalCols - 1, intCount + 3).BorderAround()
						End If


						'Surrogate write out
						If aSample.SurrogateList.Count > 0 Then
							For Each aSurrogate In aSample.SurrogateList
								For i = 0 To aSample.SurrogateList.Count
									If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
										worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
									End If
								Next
							Next
							worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
							worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 3).BorderAround()
							worksheet.Range(10, intCount, intTotalCols, intCount + 3).BorderAround()
						Else
							worksheet.Range(10, intCount, intTotalCols - 1, intCount + 3).BorderAround()
						End If

						'LCS if exists
						If GlobalVariables.TempReportSamList.Count > 1 Then
							aSample = GlobalVariables.TempReportSamList.Item(1)
							intCount = 6 'column start
							worksheet.Range(3, intCount, 3, intCount + 3).Merge()
							worksheet.Range(3, intCount).Value = aSample.LimsID
							worksheet.Range(4, intCount, 4, intCount + 3).Merge()
							worksheet.Range(4, intCount).Value = aSample.Name
							worksheet.Range(5, intCount, 5, intCount + 3).Merge()
							worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
							worksheet.Range(6, intCount, 6, intCount + 3).Merge()
							worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
							worksheet.Range(7, intCount, 7, intCount + 3).Merge()
							worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
							worksheet.Range(8, intCount, 8, intCount + 3).Merge()
							worksheet.Range(8, intCount).Value = aSample.DataFile
							worksheet.Range(9, intCount).Value = "LCS Conc"
							GlobalVariables.workbook.ActiveSheet.SetColumnWidthInPixels(intCount, 78)
							worksheet.Range(9, intCount + 1).Value = "LCL " & aSample.Units
							worksheet.Range(9, intCount + 1).CellStyle.WrapText = True
							worksheet.Range(9, intCount + 2).Value = "UCL" & aSample.Units
							worksheet.Range(9, intCount + 2).CellStyle.WrapText = True
							worksheet.Range(9, intCount + 3).Value = "% Recovery"
							worksheet.Range(9, intCount + 3).CellStyle.WrapText = True
							worksheet.Range(9, intCount + 4).Value = "Status Check"
							worksheet.Range(9, intCount + 4).CellStyle.WrapText = True
							worksheet.Range(3, intCount, 9, intCount + 4).BorderAround()
							worksheet.Range(3, intCount, 9, intCount + 4).BorderInside()
							worksheet.Range(3, intCount, 9, intCount + 4).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)

							'Begin analyte readout
							intCount2 = 10 'Row compounds start at
							For Each aCompound In aSample.CompoundList
								For i = 0 To aSample.CompoundList.Count
									If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
										If IsNumeric(aCompound.Conc) Then
											worksheet.Range(intCount2 + i, intCount).Value = aCompound.Conc
										Else
											worksheet.Range(intCount2 + i, intCount).Value = "N.D."
										End If
										worksheet.Range(intCount2 + i, intCount + 1).Value = aCompound.ChromLowLCSLim
										worksheet.Range(intCount2 + i, intCount + 2).Value = aCompound.ChromUpLCSLim
										worksheet.Range(intCount2 + i, intCount + 3).Value = GlobalVariables.Calculations.FormatSF(aCompound.ChromSpikeRecovery)
										If aCompound.ChromSpikePass Then
											worksheet.Range(intCount2 + i, intCount + 4).Value = "Passed"
										Else
											worksheet.Range(intCount2 + i, intCount + 4).Value = "Failed"
										End If
										If Not aCompound.ChromSpikePass Then
											worksheet.Range(intCount2 + i, intCount + 4).CellStyle.Color = System.Drawing.Color.FromArgb(197, 217, 241)
										End If
									End If
								Next
							Next
							If aSample.SurrogateList.Count > 0 Then
								worksheet.Range(8, intCount, intSurrStart - 1, intCount).BorderAround()
								worksheet.Range(8, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
								worksheet.Range(8, intCount + 2, intSurrStart - 1, intCount + 2).BorderAround()
								worksheet.Range(8, intCount + 3, intSurrStart - 1, intCount + 3).BorderAround()
								worksheet.Range(8, intCount + 4, intSurrStart - 1, intCount + 4).BorderAround()
							Else
								worksheet.Range(8, intCount, intTotalCols - 1, intCount).BorderAround()
								worksheet.Range(8, intCount + 1, intTotalCols - 1, intCount + 1).BorderAround()
								worksheet.Range(8, intCount + 2, intTotalCols - 1, intCount + 2).BorderAround()
								worksheet.Range(8, intCount + 3, intTotalCols - 1, intCount + 3).BorderAround()
								worksheet.Range(8, intCount + 4, intSurrStart - 1, intCount + 4).BorderAround()
							End If


							'Surrogate write out
							If aSample.SurrogateList.Count > 0 Then
								For Each aSurrogate In aSample.SurrogateList
									For i = 0 To aSample.SurrogateList.Count
										If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
											worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
										End If
									Next
								Next
								worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
								worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 4).BorderAround()
								worksheet.Range(10, intCount, intTotalCols, intCount).BorderAround()
							Else
								worksheet.Range(10, intCount, intTotalCols - 1, intCount).BorderAround()
							End If
							worksheet.Range(9, intCount, intTotalCols, intCount + 4).CellStyle.WrapText = True
							worksheet.Range(1, 2, intTotalCols, intCount + 4).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
							worksheet.Range(1, 2, intTotalCols, intCount + 4).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
							worksheet.Range(1, 1, intTotalCols, intCount + 4).CellStyle.Font.Bold = True
							worksheet.Range(1, 1, intTotalCols, intCount + 4).CellStyle.Font.Size = 8
							worksheet.Range(1, 1, intTotalCols, intCount + 4).CellStyle.Font.FontName = "Arial"
							worksheet.Range(1, 1, intTotalCols, intCount + 4).AutofitColumns()
						Else
							worksheet.Range(9, 4, intTotalCols, intCount + 3).CellStyle.WrapText = True
							worksheet.Range(1, 2, intTotalCols, intCount + 3).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
							worksheet.Range(1, 2, intTotalCols, intCount + 3).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
							worksheet.Range(1, 1, intTotalCols, intCount + 3).CellStyle.Font.Bold = True
							worksheet.Range(1, 1, intTotalCols, intCount + 3).CellStyle.Font.Size = 8
							worksheet.Range(1, 1, intTotalCols, intCount + 3).CellStyle.Font.FontName = "Arial"
							worksheet.Range(1, 1, intTotalCols, intCount + 3).AutofitColumns()
						End If
						Return True
					End If
				Next
			Else
				MsgBox("Error generating report!" & vbCrLf &
					  "Sub Procedure: FreeportChromCVSReport()" & vbCrLf &
					  "Logic Error: Could not find CVS to generate report.", "(╯°□°)╯︵ ┻━┻")
				Return False
			End If
		Catch ex As Exception
			MsgBox("Error generating report!" & vbCrLf &
					  "Sub Procedure: FreeportChromCVSReport()" & vbCrLf &
					  "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
			Return False
		End Try

	End Function

	Function FreeportChromICVReport() As Boolean
		Dim exEngine As New ExcelEngine
		Dim worksheet As IWorksheet
		Dim aSample As Sample
		Dim aSample2 As Sample
		Dim aCompound As Compound
		Dim aSurrogate As Surrogate
		Dim blnGTG As Boolean
		Dim intCount As Integer
		Dim intCount2 As Integer
		Dim intTotalCols As Integer
		Dim intSurrStart As Integer
		Dim exApp As IApplication
		exApp = exEngine.Excel

		Try
			'Grab samples for this report
			blnGTG = False
			GlobalVariables.TempReportSamList.Clear()
			For Each aSample In GlobalVariables.ReportSamList
				If aSample.Include Then
					If aSample.Type = "ICV" Then
						GlobalVariables.TempReportSamList.Add(aSample)
						blnGTG = True
						Exit For
					End If
				End If

			Next
			If blnGTG Then

				'Check for LCS
				For Each aSample In GlobalVariables.ReportSamList
					If aSample.Include Then
						If aSample.Type = "LCS" Then
							GlobalVariables.TempReportSamList.Add(aSample)
						End If
					End If

				Next
				aSample = Nothing
				aSample2 = Nothing
				GlobalVariables.workbook.Worksheets.Create("ICV Report")
				For Each wks In GlobalVariables.workbook.Worksheets
					If wks.name = "ICV Report" Then
						worksheet = wks

						'Begin building sheet..
						worksheet.Range("B1:F1").Merge()
						worksheet.Range("B1").Value = "ICV Report"
						worksheet.Range("A3").Value = "LIMS #"
						worksheet.Range("A4").Value = "Sample Point"
						worksheet.Range("A5").Value = "Sample Date"
						worksheet.Range("A6").Value = "Analysis Date"
						worksheet.Range("A7").Value = "Analysis Time"
						worksheet.Range("A8").Value = "Data Folder Name"
						worksheet.Range("A9").Value = "Analyte/Parameter"
						worksheet.Range("A3:A9").BorderInside()
						worksheet.Range("A3:A9").BorderAround()
						worksheet.Range("A3:A9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
						aSample = GlobalVariables.TempReportSamList.Item(0)

						'List Compounds
						intTotalCols = 10
						For Each aCompound In aSample.CompoundList
							worksheet.Range(intTotalCols, 1).Value = aCompound.Name
							intTotalCols = intTotalCols + 1
						Next

						'List Surrogates
						If aSample.SurrogateList.Count > 0 Then
							intTotalCols = intTotalCols + 1
							worksheet.Range(intTotalCols, 1).Value = "Surrogate Recovery (%)"
							worksheet.Range(intTotalCols, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
							intTotalCols = intTotalCols + 1
							intSurrStart = intTotalCols
							For Each aSurrogate In aSample.SurrogateList
								worksheet.Range(intTotalCols, 1).Value = aSurrogate.Name
								intTotalCols = intTotalCols + 1
							Next
							worksheet.Range(intSurrStart, 1, intTotalCols, 1).BorderAround()
							worksheet.Range(8, 1, intSurrStart - 2, 1).BorderAround()
						Else
							intSurrStart = intTotalCols
							worksheet.Range(8, 1, intTotalCols - 1, 1).BorderAround()
						End If

						intCount = 2 'column start
						worksheet.Range(3, intCount, 3, intCount + 3).Merge()
						worksheet.Range(3, intCount).Value = aSample.LimsID
						worksheet.Range(4, intCount, 4, intCount + 3).Merge()
						worksheet.Range(4, intCount).Value = aSample.Name
						worksheet.Range(5, intCount, 5, intCount + 3).Merge()
						worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
						worksheet.Range(6, intCount, 6, intCount + 3).Merge()
						worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
						worksheet.Range(7, intCount, 7, intCount + 3).Merge()
						worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
						worksheet.Range(8, intCount, 8, intCount + 3).Merge()
						worksheet.Range(8, intCount).Value = aSample.DataFile
						worksheet.Range(9, intCount).Value = "ICV Conc"
						GlobalVariables.workbook.ActiveSheet.SetColumnWidthInPixels(intCount, 78)
						worksheet.Range(9, intCount + 1).Value = "LCL (" & aSample.ReportedUnits & ")"
						worksheet.Range(9, intCount + 1).CellStyle.WrapText = True
						worksheet.Range(9, intCount + 2).Value = "UCL (" & aSample.ReportedUnits & ")"
						worksheet.Range(9, intCount + 2).CellStyle.WrapText = True
						worksheet.Range(9, intCount + 3).Value = "Status Check"
						worksheet.Range(9, intCount + 3).CellStyle.WrapText = True
						worksheet.Range(3, intCount, 9, intCount + 3).BorderAround()
						worksheet.Range(3, intCount, 9, intCount + 3).BorderInside()
						worksheet.Range(3, intCount, 9, intCount + 3).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
						GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)

						'Begin analyte readout
						intCount2 = 10 'Row compounds start at
						For Each aCompound In aSample.CompoundList
							For i = 0 To aSample.CompoundList.Count
								If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
									If IsNumeric(aCompound.Conc) Then
										worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
									Else
										worksheet.Range(intCount2 + i, intCount).Value = "N.D."
									End If
									If aCompound.ChromLowICVLim = "" Then
										worksheet.Range(intCount2 + i, intCount + 1).Value = aCompound.ChromLowContLim
										worksheet.Range(intCount2 + i, intCount + 2).Value = aCompound.ChromUpContLim
									Else
										worksheet.Range(intCount2 + i, intCount + 1).Value = aCompound.ChromLowICVLim
										worksheet.Range(intCount2 + i, intCount + 2).Value = aCompound.ChromUpICVLim
									End If
									If aCompound.ChromICVPass Then
										worksheet.Range(intCount2 + i, intCount + 3).Value = "Passed"
									Else
										worksheet.Range(intCount2 + i, intCount + 3).Value = "Failed"
									End If
									If Not aCompound.ChromICVPass Then
										worksheet.Range(intCount2 + i, intCount + 3).CellStyle.Color = System.Drawing.Color.FromArgb(197, 217, 241)
									End If
								End If
							Next
						Next
						If aSample.SurrogateList.Count > 0 Then
							worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
							worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
							worksheet.Range(10, intCount + 2, intSurrStart - 1, intCount + 2).BorderAround()
							worksheet.Range(10, intCount + 3, intSurrStart - 1, intCount + 3).BorderAround()
						Else
							worksheet.Range(10, intCount, intTotalCols - 1, intCount).BorderAround()
							worksheet.Range(10, intCount + 1, intTotalCols - 1, intCount + 1).BorderAround()
							worksheet.Range(10, intCount + 2, intTotalCols - 1, intCount + 2).BorderAround()
							worksheet.Range(10, intCount + 3, intTotalCols - 1, intCount + 3).BorderAround()
						End If


						'Surrogate write out
						If aSample.SurrogateList.Count > 0 Then
							For Each aSurrogate In aSample.SurrogateList
								For i = 0 To aSample.SurrogateList.Count
									If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
										worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
									End If
								Next
							Next
							worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
							worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 3).BorderAround()
							worksheet.Range(10, intCount, intTotalCols, intCount + 3).BorderAround()
						Else
							worksheet.Range(10, intCount, intTotalCols - 1, intCount + 3).BorderAround()
						End If

						'LCS if exists
						If GlobalVariables.TempReportSamList.Count > 1 Then
							aSample = GlobalVariables.TempReportSamList.Item(1)
							intCount = 6 'column start
							worksheet.Range(3, intCount, 3, intCount + 3).Merge()
							worksheet.Range(3, intCount).Value = aSample.LimsID
							worksheet.Range(4, intCount, 4, intCount + 3).Merge()
							worksheet.Range(4, intCount).Value = aSample.Name
							worksheet.Range(5, intCount, 5, intCount + 3).Merge()
							worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
							worksheet.Range(6, intCount, 6, intCount + 3).Merge()
							worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
							worksheet.Range(7, intCount, 7, intCount + 3).Merge()
							worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
							worksheet.Range(8, intCount, 8, intCount + 3).Merge()
							worksheet.Range(8, intCount).Value = aSample.DataFile
							worksheet.Range(9, intCount).Value = "LCS Conc"
							GlobalVariables.workbook.ActiveSheet.SetColumnWidthInPixels(intCount, 78)
							worksheet.Range(9, intCount + 1).Value = "LCL (" & aSample.ReportedUnits & ")"
							worksheet.Range(9, intCount + 1).CellStyle.WrapText = True
							worksheet.Range(9, intCount + 2).Value = "UCL (" & aSample.ReportedUnits & ")"
							worksheet.Range(9, intCount + 2).CellStyle.WrapText = True
							worksheet.Range(9, intCount + 3).Value = "% Recovery"
							worksheet.Range(9, intCount + 3).CellStyle.WrapText = True
							worksheet.Range(9, intCount + 4).Value = "Status Check"
							worksheet.Range(9, intCount + 4).CellStyle.WrapText = True
							worksheet.Range(3, intCount, 9, intCount + 4).BorderAround()
							worksheet.Range(3, intCount, 9, intCount + 4).BorderInside()
							worksheet.Range(3, intCount, 9, intCount + 4).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)

							'Begin analyte readout
							intCount2 = 10 'Row compounds start at
							For Each aCompound In aSample.CompoundList
								For i = 0 To aSample.CompoundList.Count
									If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
										If IsNumeric(aCompound.Conc) Then
											worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
										Else
											worksheet.Range(intCount2 + i, intCount).Value = "N.D."
										End If
										worksheet.Range(intCount2 + i, intCount + 1).Value = aCompound.ChromLowLCSLim
										worksheet.Range(intCount2 + i, intCount + 2).Value = aCompound.ChromUpLCSLim
										worksheet.Range(intCount2 + i, intCount + 3).Value = GlobalVariables.Calculations.FormatSF(aCompound.ChromSpikeRecovery)
										If aCompound.ChromSpikePass Then
											worksheet.Range(intCount2 + i, intCount + 4).Value = "Passed"
										Else
											worksheet.Range(intCount2 + i, intCount + 4).Value = "Failed"
										End If
										If Not aCompound.ChromSpikePass Then
											worksheet.Range(intCount2 + i, intCount + 4).CellStyle.Color = System.Drawing.Color.FromArgb(197, 217, 241)
										End If
									End If
								Next
							Next
							If aSample.SurrogateList.Count > 0 Then
								worksheet.Range(8, intCount, intSurrStart - 1, intCount).BorderAround()
								worksheet.Range(8, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
								worksheet.Range(8, intCount + 2, intSurrStart - 1, intCount + 2).BorderAround()
								worksheet.Range(8, intCount + 3, intSurrStart - 1, intCount + 3).BorderAround()
								worksheet.Range(8, intCount + 4, intSurrStart - 1, intCount + 4).BorderAround()
							Else
								worksheet.Range(8, intCount, intTotalCols - 1, intCount).BorderAround()
								worksheet.Range(8, intCount + 1, intTotalCols - 1, intCount + 1).BorderAround()
								worksheet.Range(8, intCount + 2, intTotalCols - 1, intCount + 2).BorderAround()
								worksheet.Range(8, intCount + 3, intTotalCols - 1, intCount + 3).BorderAround()
								worksheet.Range(8, intCount + 4, intSurrStart - 1, intCount + 4).BorderAround()
							End If


							'Surrogate write out
							If aSample.SurrogateList.Count > 0 Then
								For Each aSurrogate In aSample.SurrogateList
									For i = 0 To aSample.SurrogateList.Count
										If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
											worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
										End If
									Next
								Next
								worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
								worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 4).BorderAround()
								worksheet.Range(10, intCount, intTotalCols, intCount).BorderAround()
							Else
								worksheet.Range(10, intCount, intTotalCols - 1, intCount).BorderAround()
							End If
							worksheet.Range(9, intCount, intTotalCols, intCount + 4).CellStyle.WrapText = True
							worksheet.Range(1, 2, intTotalCols, intCount + 4).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
							worksheet.Range(1, 2, intTotalCols, intCount + 4).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
							worksheet.Range(1, 1, intTotalCols, intCount + 4).CellStyle.Font.Bold = True
							worksheet.Range(1, 1, intTotalCols, intCount + 4).CellStyle.Font.Size = 8
							worksheet.Range(1, 1, intTotalCols, intCount + 4).CellStyle.Font.FontName = "Arial"
							worksheet.Range(1, 1, intTotalCols, intCount + 4).AutofitColumns()
						Else
							worksheet.Range(9, 4, intTotalCols, intCount + 3).CellStyle.WrapText = True
							worksheet.Range(1, 2, intTotalCols, intCount + 3).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
							worksheet.Range(1, 2, intTotalCols, intCount + 3).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
							worksheet.Range(1, 1, intTotalCols, intCount + 3).CellStyle.Font.Bold = True
							worksheet.Range(1, 1, intTotalCols, intCount + 3).CellStyle.Font.Size = 8
							worksheet.Range(1, 1, intTotalCols, intCount + 3).CellStyle.Font.FontName = "Arial"
							worksheet.Range(1, 1, intTotalCols, intCount + 3).AutofitColumns()
						End If
						Return True
					End If
				Next
			Else
				MsgBox("Error generating report!" & vbCrLf &
					  "Sub Procedure: FreeportChromICVReport()" & vbCrLf &
					  "Logic Error: Could not find ICV to generate report.", "(╯°□°)╯︵ ┻━┻")
				Return False
			End If
		Catch ex As Exception
			MsgBox("Error generating report!" & vbCrLf &
					  "Sub Procedure: FreeportChromICVReport()" & vbCrLf &
					  "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
			Return False
		End Try

	End Function

	'Freeport Summary Report Methylated
	Function FreeportChromSummaryMethReport(ByVal strLimitType As String) As Boolean
		Dim worksheet As IWorksheet
		Dim aSample As Sample
		Dim aSample2 As Sample
		Dim aPermit As Permit
		Dim aProject As Project
		Dim aInstrument As mInstrument
		Dim amCompound As mCompound
		Dim aCompound As Compound
		Dim aSurrogate As Surrogate
		Dim intCount As Integer
		Dim intCount2 As Integer
		Dim intSurrStart As Integer
		Dim exEngine As New ExcelEngine
		Dim exApp As IApplication
		exApp = exEngine.Excel

		Try
			GlobalVariables.workbook.Worksheets.Create("Summary MED Report")
			For Each wks In GlobalVariables.workbook.Worksheets
				If wks.name = "Summary MED Report" Then
					worksheet = wks
					aSample = GlobalVariables.ReportSamList.Item(0)
					If strLimitType <> "RL" Then
						'Begin building sheet..
						worksheet.Range("D1:E1").Merge()
						worksheet.Range("D1").Value = "Summary Report"
						worksheet.Range("A3").Value = "LIMS #"
						worksheet.Range("A4").Value = "Sample Point"
						worksheet.Range("A5").Value = "Sample Date"
						worksheet.Range("A6").Value = "Analysis Date"
						worksheet.Range("A7").Value = "Analysis Time"
						worksheet.Range("A8").Value = "Data Folder Name"
						worksheet.Range("A9").Value = "Analyte/Parameter"
						worksheet.Range("B9").Value = "CAS #"
						worksheet.Range("C9").Value = strLimitType & " (" & aSample.ReportedUnits & ")"
						worksheet.Range("A3:C9").BorderInside()
						worksheet.Range("A3:C9").BorderAround()
						worksheet.Range("A3:C9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
						For Each aSample2 In GlobalVariables.ReportSamList
							If aSample2.Methylated Then
								aSample = aSample2
								Exit For
							End If
						Next

						'List Compounds
						intCount = 10
						For Each aPermit In GlobalVariables.PermitList
							If aPermit.Name = GlobalVariables.selPermit.Name Then
								For Each aProject In aPermit.ProjectList
									If aProject.Name = GlobalVariables.selProject Then
										For Each aInstrument In aProject.mInstrumentList
											If aInstrument.Name = GlobalVariables.selInstrument Then
												For Each amCompound In aInstrument.mCompoundList
													For Each aCompound In aSample.CompoundList
														If amCompound.Name = aCompound.Name And aCompound.Methylated Then
															worksheet.Range(intCount, 1).Value = aCompound.Name
															worksheet.Range(intCount, 2).Value = amCompound.CAS
															If strLimitType = "MDL" Then
																worksheet.Range(intCount, 3).Value = amCompound.MDL
															ElseIf strLimitType = "PQL" Then
																worksheet.Range(intCount, 3).Value = amCompound.PQL
															End If
															intCount = intCount + 1
														End If
													Next
												Next
											End If
										Next
									End If
								Next
							End If
						Next
						worksheet.Range(10, 1, intCount, 1).BorderAround()
						worksheet.Range(10, 2, intCount, 2).BorderAround()
						worksheet.Range(10, 3, intCount, 3).BorderAround()

						'List Surrogates
						intCount = intCount + 1
						If aSample.SurrogateList.Count > 0 Then
							worksheet.Range(intCount, 1).Value = "Surrogate Recovery (%)"
							worksheet.Range(intCount, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
							worksheet.Range(intCount, 3).Value = "Recovery Limit (%)"
							worksheet.Range(intCount, 3).CellStyle.Font.Color = ExcelKnownColors.Blue
							intCount = intCount + 1
							intSurrStart = intCount
							For Each aSurrogate In aSample.SurrogateList
								If aSurrogate.Methylated Then
									worksheet.Range(intCount, 1).Value = aSurrogate.Name
									worksheet.Range(intCount, 3).Value = aSurrogate.ChromLowContLim & "-" & aSurrogate.ChromUpContLim
									intCount = intCount + 1
								End If
							Next
							worksheet.Range(intSurrStart, 1, intCount, 1).BorderAround()
							worksheet.Range(intSurrStart, 2, intCount, 2).BorderAround()
							worksheet.Range(intSurrStart, 3, intCount, 3).BorderAround()
							worksheet.Range(intSurrStart - 1, 1).BorderAround()
							worksheet.Range(intSurrStart - 1, 2).BorderAround()
							worksheet.Range(intSurrStart - 1, 3).BorderAround()
						Else
							intSurrStart = intCount
						End If


						intCount = 4
						For Each aSample In GlobalVariables.ReportSamList
							If aSample.Include Then
								If aSample.Methylated Then
									worksheet.Range(3, intCount, 3, intCount + 1).Merge()
									worksheet.Range(3, intCount).Value = aSample.LimsID
									worksheet.Range(4, intCount, 4, intCount + 1).Merge()
									worksheet.Range(4, intCount).Value = aSample.Name
									worksheet.Range(5, intCount, 5, intCount + 1).Merge()
									worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
									worksheet.Range(6, intCount, 6, intCount + 1).Merge()
									worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
									worksheet.Range(7, intCount, 7, intCount + 1).Merge()
									worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
									worksheet.Range(8, intCount, 8, intCount + 1).Merge()
									worksheet.Range(8, intCount).Value = aSample.DataFile
									worksheet.Range(3, intCount, 8, intCount + 1).BorderAround()
									worksheet.Range(3, intCount, 8, intCount + 1).BorderInside()
									worksheet.Range(3, intCount, 8, intCount + 1).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
									worksheet.Range(9, intCount).Value = "Adjusted Limit (" & aSample.ReportedUnits & ")"
									worksheet.Range(9, intCount + 1).Value = "Amount (" & aSample.ReportedUnits & ")"
									worksheet.Range(9, intCount).BorderAround()
									worksheet.Range(9, intCount + 1).BorderAround()
									worksheet.Range(9, intCount, 9, intCount + 1).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
									GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 45)

									'Begin analyte readout
									intCount2 = 10 'Row compounds start at
									For Each aCompound In aSample.CompoundList
										For i = 0 To aSample.CompoundList.Count
											If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
												worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.ChromAdjustedLimit)
												worksheet.Range(intCount2 + i, intCount + 1).Value = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
											End If
										Next
									Next
									worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
									worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()

									'Surrogate write out
									If aSample.SurrogateList.Count > 0 Then
										For Each aSurrogate In aSample.SurrogateList
											For i = 0 To aSample.SurrogateList.Count
												If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
													worksheet.Range(intSurrStart + i, intCount + 1).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
												End If
											Next
										Next
										worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
										worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).BorderAround()
									End If

									worksheet.Range(9, 4, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.WrapText = True
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.Font.Bold = True
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.Font.Size = 8
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.Font.FontName = "Arial"
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).AutofitColumns()

									intCount = intCount + 2
								End If
							End If

						Next
					Else
						'RL Limit or N/A selected
						aSample = GlobalVariables.ReportSamList.Item(0)
						'Begin building sheet..
						worksheet.Range("D1:E1").Merge()
						worksheet.Range("D1").Value = "Summary Report"
						worksheet.Range("A3").Value = "LIMS #"
						worksheet.Range("A4").Value = "Sample Point"
						worksheet.Range("A5").Value = "Sample Date"
						worksheet.Range("A6").Value = "Analysis Date"
						worksheet.Range("A7").Value = "Analysis Time"
						worksheet.Range("A8").Value = "Data Folder Name"
						worksheet.Range("A9").Value = "Analyte/Parameter"
						worksheet.Range("B9").Value = "CAS #"
						If strLimitType = "RL" Then
							worksheet.Range("C9").Value = strLimitType & " (" & aSample.ReportedUnits & ")"
						Else
							worksheet.Range("C9").Value = strLimitType
						End If

						worksheet.Range("A3:C9").BorderInside()
						worksheet.Range("A3:C9").BorderAround()
						worksheet.Range("A3:C9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
						aSample = GlobalVariables.ReportSamList.Item(0)

						'List Compounds
						intCount = 10
						For Each aPermit In GlobalVariables.PermitList
							If aPermit.Name = GlobalVariables.selPermit.Name Then
								For Each aProject In aPermit.ProjectList
									If aProject.Name = GlobalVariables.selProject Then
										For Each aInstrument In aProject.mInstrumentList
											If aInstrument.Name = GlobalVariables.selInstrument Then
												For Each amCompound In aInstrument.mCompoundList
													For Each aCompound In aSample.CompoundList
														If amCompound.Name = aCompound.Name And aCompound.Methylated Then
															worksheet.Range(intCount, 1).Value = aCompound.Name
															worksheet.Range(intCount, 2).Value = amCompound.CAS
															If strLimitType = "RL" Then
																worksheet.Range(intCount, 3).Value = amCompound.RL
															Else
																worksheet.Range(intCount, 3).Value = "N/A"
															End If

															intCount = intCount + 1
														End If
													Next
												Next
											End If
										Next
									End If
								Next
							End If
						Next
						worksheet.Range(10, 1, intCount, 1).BorderAround()
						worksheet.Range(10, 2, intCount, 2).BorderAround()
						worksheet.Range(10, 3, intCount, 3).BorderAround()

						'List Surrogates
						intCount = intCount + 1
						If aSample.SurrogateList.Count > 0 Then
							worksheet.Range(intCount, 1).Value = "Surrogate Recovery (%)"
							worksheet.Range(intCount, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
							worksheet.Range(intCount, 3).Value = "Recovery Limit (%)"
							worksheet.Range(intCount, 3).CellStyle.Font.Color = ExcelKnownColors.Blue
							intCount = intCount + 1
							intSurrStart = intCount
							For Each aSurrogate In aSample.SurrogateList
								If aSurrogate.Methylated Then
									worksheet.Range(intCount, 1).Value = aSurrogate.Name
									worksheet.Range(intCount, 3).Value = aSurrogate.ChromLowContLim & "-" & aSurrogate.ChromUpContLim
									intCount = intCount + 1
								End If
							Next
							worksheet.Range(intSurrStart, 1, intCount, 1).BorderAround()
							worksheet.Range(intSurrStart, 2, intCount, 2).BorderAround()
							worksheet.Range(intSurrStart, 3, intCount, 3).BorderAround()
							worksheet.Range(intSurrStart - 1, 1).BorderAround()
							worksheet.Range(intSurrStart - 1, 2).BorderAround()
							worksheet.Range(intSurrStart - 1, 3).BorderAround()
						Else
							intSurrStart = intCount
						End If


						intCount = 4
						For Each aSample In GlobalVariables.ReportSamList
							If aSample.Include Then
								If aSample.Methylated Then
									worksheet.Range(3, intCount).Value = aSample.LimsID
									worksheet.Range(4, intCount).Value = aSample.Name
									worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
									worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
									worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
									worksheet.Range(8, intCount).Value = aSample.DataFile
									worksheet.Range(3, intCount, 9, intCount).BorderAround()
									worksheet.Range(3, intCount, 9, intCount).BorderInside()
									worksheet.Range(3, intCount, 9, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
									worksheet.Range(9, intCount).Value = "Amount (" & aSample.ReportedUnits & ")"
									GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 45)

									'Begin analyte readout
									intCount2 = 10 'Row compounds start at
									For Each aCompound In aSample.CompoundList
										For i = 0 To aSample.CompoundList.Count
											If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
												worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
											End If
										Next
									Next
									worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()

									'Surrogate write out
									If aSample.SurrogateList.Count > 0 Then
										For Each aSurrogate In aSample.SurrogateList
											For i = 0 To aSample.SurrogateList.Count
												If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
													worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
												End If
											Next
										Next
										worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
									End If
									worksheet.Range(9, 4, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.WrapText = True
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Bold = True
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Size = 8
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.FontName = "Arial"
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).AutofitColumns()

									intCount = intCount + 1
								End If
							End If

						Next
					End If

					Return True
				End If
			Next
		Catch ex As Exception
			MsgBox("Error generating report!" & vbCrLf &
						"Sub Procedure: FreeportChromSummaryMethReport()" & vbCrLf &
						"Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
			Return False
		End Try

	End Function

	'Freeport Summary Report
	Function FreeportChromSummaryReport(ByVal strLimitType As String) As Boolean
		Dim worksheet As IWorksheet
		Dim aSample As Sample
		Dim aPermit As Permit
		Dim aProject As Project
		Dim aInstrument As mInstrument
		Dim amCompound As mCompound
		Dim aCompound As Compound
		Dim aSurrogate As Surrogate
		Dim intCount As Integer
		Dim intCount2 As Integer
		Dim intSurrStart As Integer
		Dim exEngine As New ExcelEngine
		Dim exApp As IApplication
		Dim blnMethylated As Boolean
		exApp = exEngine.Excel

		blnMethylated = False
		For Each aSample In GlobalVariables.ReportSamList
			If aSample.Methylated Then
				blnMethylated = True
			End If
		Next

		'If methylated, means we need to make a methylated sheet as well
		If blnMethylated Then
			FreeportChromSummaryMethReport(strLimitType)
		End If


		Try
			GlobalVariables.workbook.Worksheets.Create("Summary Report")
			For Each wks In GlobalVariables.workbook.Worksheets
				If wks.name = "Summary Report" Then
					worksheet = wks
					aSample = GlobalVariables.ReportSamList.Item(0)
					If strLimitType <> "RL" Then
						'Begin building sheet..
						worksheet.Range("D1:E1").Merge()
						worksheet.Range("D1").Value = "Summary Report"
						worksheet.Range("A3").Value = "LIMS #"
						worksheet.Range("A4").Value = "Sample Point"
						worksheet.Range("A5").Value = "Sample Date"
						worksheet.Range("A6").Value = "Analysis Date"
						worksheet.Range("A7").Value = "Analysis Time"
						worksheet.Range("A8").Value = "Data Folder Name"
						worksheet.Range("A9").Value = "Analyte/Parameter"
						worksheet.Range("B9").Value = "CAS #"
						worksheet.Range("C9").Value = strLimitType & " (" & aSample.ReportedUnits & ")"
						worksheet.Range("A3:C9").BorderInside()
						worksheet.Range("A3:C9").BorderAround()
						worksheet.Range("A3:C9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
						aSample = GlobalVariables.ReportSamList.Item(0)

						'List Compounds
						intCount = 10
						For Each aPermit In GlobalVariables.PermitList
							If aPermit.Name = GlobalVariables.selPermit.Name Then
								For Each aProject In aPermit.ProjectList
									If aProject.Name = GlobalVariables.selProject Then
										For Each aInstrument In aProject.mInstrumentList
											If aInstrument.Name = GlobalVariables.selInstrument Then
												For Each amCompound In aInstrument.mCompoundList
													For Each aCompound In aSample.CompoundList
														If amCompound.Name = aCompound.Name Then
															worksheet.Range(intCount, 1).Value = aCompound.Name
															worksheet.Range(intCount, 2).Value = amCompound.CAS
															If strLimitType = "MDL" Then
																worksheet.Range(intCount, 3).Value = amCompound.MDL
															ElseIf strLimitType = "PQL" Then
																worksheet.Range(intCount, 3).Value = amCompound.PQL
															End If
															intCount = intCount + 1
														End If
													Next
												Next
											End If
										Next
									End If
								Next
							End If
						Next
						worksheet.Range(10, 1, intCount, 1).BorderAround()
						worksheet.Range(10, 2, intCount, 2).BorderAround()
						worksheet.Range(10, 3, intCount, 3).BorderAround()

						'List Surrogates
						intCount = intCount + 1
						If aSample.SurrogateList.Count > 0 Then
							worksheet.Range(intCount, 1).Value = "Surrogate Recovery (%)"
							worksheet.Range(intCount, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
							worksheet.Range(intCount, 3).Value = "Recovery Limit (%)"
							worksheet.Range(intCount, 3).CellStyle.Font.Color = ExcelKnownColors.Blue
							intCount = intCount + 1
							intSurrStart = intCount
							For Each aSurrogate In aSample.SurrogateList
								worksheet.Range(intCount, 1).Value = aSurrogate.Name
								worksheet.Range(intCount, 3).Value = aSurrogate.ChromLowContLim & "-" & aSurrogate.ChromUpContLim
								intCount = intCount + 1
							Next
							worksheet.Range(intSurrStart, 1, intCount, 1).BorderAround()
							worksheet.Range(intSurrStart, 2, intCount, 2).BorderAround()
							worksheet.Range(intSurrStart, 3, intCount, 3).BorderAround()
							worksheet.Range(intSurrStart - 1, 1).BorderAround()
							worksheet.Range(intSurrStart - 1, 2).BorderAround()
							worksheet.Range(intSurrStart - 1, 3).BorderAround()
						Else
							intSurrStart = intCount
						End If


						intCount = 4
						For Each aSample In GlobalVariables.ReportSamList
							If aSample.Include Then
								worksheet.Range(3, intCount, 3, intCount + 1).Merge()
								worksheet.Range(3, intCount).Value = aSample.LimsID
								worksheet.Range(4, intCount, 4, intCount + 1).Merge()
								worksheet.Range(4, intCount).Value = aSample.Name
								worksheet.Range(5, intCount, 5, intCount + 1).Merge()
								worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
								worksheet.Range(6, intCount, 6, intCount + 1).Merge()
								worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
								worksheet.Range(7, intCount, 7, intCount + 1).Merge()
								worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
								worksheet.Range(8, intCount, 8, intCount + 1).Merge()
								worksheet.Range(8, intCount).Value = aSample.DataFile
								worksheet.Range(3, intCount, 8, intCount + 1).BorderAround()
								worksheet.Range(3, intCount, 8, intCount + 1).BorderInside()
								worksheet.Range(3, intCount, 8, intCount + 1).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								worksheet.Range(9, intCount).Value = "Adjusted Limit (" & aSample.ReportedUnits & ")"
								worksheet.Range(9, intCount + 1).Value = "Amount (" & aSample.ReportedUnits & ")"
								worksheet.Range(9, intCount).BorderAround()
								worksheet.Range(9, intCount + 1).BorderAround()
								worksheet.Range(9, intCount, 9, intCount + 1).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 45)

								'Begin analyte readout
								intCount2 = 10 'Row compounds start at
								For Each aCompound In aSample.CompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.ChromAdjustedLimit)
											worksheet.Range(intCount2 + i, intCount + 1).Value = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
										End If
									Next
								Next
								worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
								worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()

								'Surrogate write out
								If aSample.SurrogateList.Count > 0 Then
									For Each aSurrogate In aSample.SurrogateList
										For i = 0 To aSample.SurrogateList.Count
											If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
												worksheet.Range(intSurrStart + i, intCount + 1).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
											End If
										Next
									Next
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
									worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).BorderAround()
								End If

								worksheet.Range(9, 4, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.WrapText = True
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.Font.Bold = True
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.Font.Size = 8
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.Font.FontName = "Arial"
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).AutofitColumns()

								intCount = intCount + 2
							End If

						Next
					Else
						'RL Limit selected
						aSample = GlobalVariables.ReportSamList.Item(0)
						'Begin building sheet..
						worksheet.Range("D1:E1").Merge()
						worksheet.Range("D1").Value = "Summary Report"
						worksheet.Range("A3").Value = "LIMS #"
						worksheet.Range("A4").Value = "Sample Point"
						worksheet.Range("A5").Value = "Sample Date"
						worksheet.Range("A6").Value = "Analysis Date"
						worksheet.Range("A7").Value = "Analysis Time"
						worksheet.Range("A8").Value = "Data Folder Name"
						worksheet.Range("A9").Value = "Analyte/Parameter"
						worksheet.Range("B9").Value = "CAS #"
						worksheet.Range("C9").Value = strLimitType & " (" & aSample.ReportedUnits & ")"
						worksheet.Range("A3:C9").BorderInside()
						worksheet.Range("A3:C9").BorderAround()
						worksheet.Range("A3:C9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
						aSample = GlobalVariables.ReportSamList.Item(0)

						'List Compounds
						intCount = 10
						For Each aPermit In GlobalVariables.PermitList
							If aPermit.Name = GlobalVariables.selPermit.Name Then
								For Each aProject In aPermit.ProjectList
									If aProject.Name = GlobalVariables.selProject Then
										For Each aInstrument In aProject.mInstrumentList
											If aInstrument.Name = GlobalVariables.selInstrument Then
												For Each amCompound In aInstrument.mCompoundList
													For Each aCompound In aSample.CompoundList
														If amCompound.Name = aCompound.Name Then
															worksheet.Range(intCount, 1).Value = aCompound.Name
															worksheet.Range(intCount, 2).Value = amCompound.CAS
															If strLimitType = "RL" Then
																worksheet.Range(intCount, 3).Value = amCompound.RL
															Else
																worksheet.Range(intCount, 3).Value = "N/A"
															End If
															intCount = intCount + 1
														End If
													Next
												Next
											End If
										Next
									End If
								Next
							End If
						Next
						worksheet.Range(10, 1, intCount, 1).BorderAround()
						worksheet.Range(10, 2, intCount, 2).BorderAround()
						worksheet.Range(10, 3, intCount, 3).BorderAround()

						'List Surrogates
						intCount = intCount + 1
						If aSample.SurrogateList.Count > 0 Then
							worksheet.Range(intCount, 1).Value = "Surrogate Recovery (%)"
							worksheet.Range(intCount, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
							worksheet.Range(intCount, 3).Value = "Recovery Limit (%)"
							worksheet.Range(intCount, 3).CellStyle.Font.Color = ExcelKnownColors.Blue
							intCount = intCount + 1
							intSurrStart = intCount
							For Each aSurrogate In aSample.SurrogateList
								worksheet.Range(intCount, 1).Value = aSurrogate.Name
								worksheet.Range(intCount, 3).Value = aSurrogate.ChromLowContLim & "-" & aSurrogate.ChromUpContLim
								intCount = intCount + 1
							Next
							worksheet.Range(intSurrStart, 1, intCount, 1).BorderAround()
							worksheet.Range(intSurrStart, 2, intCount, 2).BorderAround()
							worksheet.Range(intSurrStart, 3, intCount, 3).BorderAround()
							worksheet.Range(intSurrStart - 1, 1).BorderAround()
							worksheet.Range(intSurrStart - 1, 2).BorderAround()
							worksheet.Range(intSurrStart - 1, 3).BorderAround()
						Else
							intSurrStart = intCount
						End If


						intCount = 4
						For Each aSample In GlobalVariables.ReportSamList
							If aSample.Include Then
								worksheet.Range(3, intCount).Value = aSample.LimsID
								worksheet.Range(4, intCount).Value = aSample.Name
								worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
								worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
								worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
								worksheet.Range(8, intCount).Value = aSample.DataFile
								worksheet.Range(3, intCount, 9, intCount).BorderAround()
								worksheet.Range(3, intCount, 9, intCount).BorderInside()
								worksheet.Range(3, intCount, 9, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								worksheet.Range(9, intCount).Value = "Amount (" & aSample.ReportedUnits & ")"
								GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 45)

								'Begin analyte readout
								intCount2 = 10 'Row compounds start at
								For Each aCompound In aSample.CompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
										End If
									Next
								Next
								worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()

								'Surrogate write out
								If aSample.SurrogateList.Count > 0 Then
									For Each aSurrogate In aSample.SurrogateList
										For i = 0 To aSample.SurrogateList.Count
											If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
												worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
											End If
										Next
									Next
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
								End If
								worksheet.Range(9, 4, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.WrapText = True
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Bold = True
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Size = 8
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.FontName = "Arial"
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).AutofitColumns()

								intCount = intCount + 1
							End If

						Next
					End If
				End If
			Next
			Return True
		Catch ex As Exception
			MsgBox("Error generating report!" & vbCrLf &
						"Sub Procedure: FreeportChromSummaryReport()" & vbCrLf &
						"Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
			Return False
		End Try

	End Function

	Function MidlandChromDUPReport(ByVal strLimitType As String) As Boolean
		Dim exEngine As New ExcelEngine
		Dim worksheet As IWorksheet
		Dim aSample As Sample
		Dim aSample2 As Sample
		Dim aCompound As Compound
		Dim blnGTG As Boolean
		Dim intCount As Integer
		Dim intCount2 As Integer
		Dim intSurrStart As Integer
		Dim exApp As IApplication
		Dim intWksCount As Integer
		exApp = exEngine.Excel

		Try
			'Grab samples for this report
			blnGTG = False
			intWksCount = 0
			GlobalVariables.TempReportSamList.Clear()
			For Each aSample In GlobalVariables.ReportSamList
				If aSample.Include Then
					If aSample.Type = "DUP" And Not aSample.Reported Then
						For Each aSample2 In GlobalVariables.ReportSamList
							If aSample2.Name = Trim(aSample.Name.Substring(0, aSample.Name.Length - 3)) And Not aSample.Reported And aSample.Methylated And aSample2.Methylated Then
								GlobalVariables.TempReportSamList.Add(aSample2)
								GlobalVariables.TempReportSamList.Add(aSample)
								intWksCount = intWksCount + 1
								blnGTG = True
								Exit For
							ElseIf aSample2.Name = Trim(aSample.Name.Substring(0, aSample.Name.Length - 3)) And Not aSample.Reported And Not aSample.Methylated And Not aSample2.Methylated Then
								GlobalVariables.TempReportSamList.Add(aSample2)
								GlobalVariables.TempReportSamList.Add(aSample)
								intWksCount = intWksCount + 1
								blnGTG = True
								Exit For
							End If
						Next
					End If
				End If

			Next

			'aSample is dup, aSample2 is original
			If blnGTG Then
				For u = 1 To intWksCount
					GlobalVariables.workbook.Worksheets.Create("DUP-" & CStr(u))
					For Each wks In GlobalVariables.workbook.Worksheets
						If wks.name = "DUP-" & CStr(u) Then
							worksheet = wks

							'Begin building sheet..
							worksheet.Range("B1:D1").Merge()
							worksheet.Range("B1").Value = "Sample Duplicate RPD Report"
							worksheet.Range("A3").Value = "LIMS #"
							worksheet.Range("A4").Value = "Sample Point"
							worksheet.Range("A5").Value = "Sample Date"
							worksheet.Range("A6").Value = "Analysis Date"
							worksheet.Range("A7").Value = "Analysis Time"
							worksheet.Range("A8").Value = "Data Folder Name"
							worksheet.Range("A9").Value = "Analyte/Parameter"
							worksheet.Range("A3:A9").BorderInside()
							worksheet.Range(10, 1).FreezePanes()
							worksheet.Range("A3:A9").BorderAround()
							worksheet.Range("A3:A9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							aSample = GlobalVariables.TempReportSamList.Item(0)

							'List Compounds
							intCount = 10
							For Each aCompound In aSample.CompoundList
								worksheet.Range(intCount, 1).Value = aCompound.Name
								intCount = intCount + 1
							Next
							intSurrStart = intCount
							worksheet.Range(10, 1, intCount, 1).BorderAround()

							intCount = 2 'column start

							worksheet.Range(3, intCount).Value = aSample.LimsID
							worksheet.Range(4, intCount).Value = aSample.Name
							worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
							worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
							worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
							worksheet.Range(8, intCount).Value = aSample.DataFile
							worksheet.Range(9, intCount).Value = "Amount (" & aSample.ReportedUnits & ") (From Chemstation Report)"
							worksheet.Range(9, intCount).CellStyle.WrapText = True
							worksheet.Range(3, intCount, 9, intCount).BorderAround()
							worksheet.Range(3, intCount, 9, intCount).BorderInside()
							worksheet.Range(3, intCount, 9, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							GlobalVariables.workbook.ActiveSheet.SetColumnWidthInPixels(2, 140)
							GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)

							'Begin analyte readout
							intCount2 = 10 'Row compounds start at
							For Each aCompound In aSample.CompoundList
								For i = 0 To aSample.CompoundList.Count
									If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
										worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
									End If
								Next
							Next
							worksheet.Range(10, intCount, intSurrStart, intCount).BorderAround()

							intCount = 3 'column start
							aSample = GlobalVariables.TempReportSamList.Item(1)
							worksheet.Range(3, intCount).Value = aSample.LimsID
							worksheet.Range(4, intCount).Value = aSample.Name
							worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
							worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
							worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
							worksheet.Range(8, intCount).Value = aSample.DataFile
							worksheet.Range(9, intCount).Value = "Amount (" & aSample.ReportedUnits & ") (From Chemstation Report)"
							worksheet.Range(9, intCount).CellStyle.WrapText = True
							worksheet.Range(3, intCount, 9, intCount).BorderAround()
							worksheet.Range(3, intCount, 9, intCount).BorderInside()
							worksheet.Range(3, intCount, 9, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)
							GlobalVariables.workbook.ActiveSheet.SetColumnWidthInPixels(3, 140)

							'Begin analyte readout
							intCount2 = 10 'Row compounds start at
							For Each aCompound In aSample.CompoundList
								For i = 0 To aSample.CompoundList.Count
									If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
										worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
									End If
								Next
							Next
							worksheet.Range(10, intCount, intSurrStart, intCount).BorderAround()

							'3 remaining columns

							intCount = 4

							worksheet.Range(9, intCount).Value = "% RPD"
							worksheet.Range(9, intCount + 1).Value = "RPD Limit"
							worksheet.Range(9, intCount + 2).Value = "Pass/Fail"
							worksheet.Range(3, intCount, 9, intCount + 2).BorderInside()
							worksheet.Range(3, intCount, 9, intCount + 2).BorderAround()
							worksheet.Range(3, intCount, 9, intCount + 2).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							aSample = GlobalVariables.TempReportSamList.Item(0)
							'Begin analyte readout
							intCount2 = 10 'Row compounds start at
							For Each aCompound In aSample.CompoundList
								For i = 0 To aSample.CompoundList.Count
									If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
										worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF(aCompound.ChromRPD)
										worksheet.Range(intCount2 + i, intCount + 1).Value = "30"
										If aCompound.ChromRPD <> "N/A" Then
											If CDbl(aCompound.ChromRPD) <= CDbl(worksheet.Range(intCount2 + i, intCount + 1).Value) Then
												worksheet.Range(intCount2 + i, intCount + 2).Value = "Pass"
											Else
												worksheet.Range(intCount2 + i, intCount + 2).Value = "Fail"
											End If
										Else
											worksheet.Range(intCount2 + i, intCount + 2).Value = "N/A"
										End If
									End If
								Next
							Next
							worksheet.Range(10, intCount, intSurrStart, intCount).BorderAround()
							worksheet.Range(10, intCount + 1, intSurrStart, intCount + 1).BorderAround()
							worksheet.Range(10, intCount + 1, intSurrStart, intCount + 2).BorderAround()
							worksheet.Range(1, 2, intSurrStart, intCount + 2).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
							worksheet.Range(1, 2, intSurrStart, intCount + 2).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
							worksheet.Range(1, 1, intSurrStart, intCount + 2).CellStyle.Font.Bold = True
							worksheet.Range(1, 1, intSurrStart, intCount + 2).CellStyle.Font.Size = 8
							worksheet.Range(1, 1, intSurrStart, intCount + 2).CellStyle.Font.FontName = "Arial"
							worksheet.Range(1, 1, intSurrStart, intCount + 2).AutofitColumns()

							'Clear out Samples so it is not reported twice
							aSample = GlobalVariables.TempReportSamList.Item(0)
							GlobalVariables.TempReportSamList.Remove(aSample)
							aSample = GlobalVariables.TempReportSamList.Item(0)
							aSample.Reported = True
							GlobalVariables.TempReportSamList.Remove(aSample)
						End If
					Next
				Next
				Return True
			Else
				Return False
			End If


		Catch ex As Exception
			MsgBox("Error creating DUP Report" & vbCrLf &
			   "Sub Procedure: MidlandChromDUPReport()" & vbCrLf &
			   "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
		End Try

	End Function

	'Midland LCS Report
	Function MidlandChromLCSReport() As Boolean
		Dim exEngine As New ExcelEngine
		Dim worksheet As IWorksheet
		Dim aSample As Sample
		Dim aSample2 As Sample
		Dim aCompound As Compound
		Dim aSurrogate As Surrogate
		Dim blnGTG As Boolean
		Dim intCount As Integer
		Dim intCount2 As Integer
		Dim intTotalCols As Integer
		Dim intSurrStart As Integer
		Dim intWksCount As Integer
		Dim blnLCSD As Boolean
		Dim intLCSDCounter As Integer
		Dim exApp As IApplication
		exApp = exEngine.Excel

		Try

			'Grab samples for this report
			blnGTG = False
			intWksCount = 0
			GlobalVariables.TempReportSamList.Clear()
			For Each aSample In GlobalVariables.ReportSamList
				If aSample.Include Then
					If aSample.Type = "LCS" Then
						'Look for LCSD
						blnLCSD = False
						For Each aSample2 In GlobalVariables.ReportSamList
							If aSample2.Type = "LCSD" And Not aSample2.Reported And aSample.Methylated And aSample2.Methylated Then
								aSample.Reported = True
								aSample2.Reported = True
								GlobalVariables.TempReportSamList.Add(aSample)
								GlobalVariables.TempReportSamList.Add(aSample2)
								blnLCSD = True
								intWksCount = intWksCount + 1
								blnGTG = True
								Exit For
							ElseIf aSample2.Type = "LCSD" And Not aSample2.Reported And Not aSample.Methylated And Not aSample2.Methylated Then
								aSample.Reported = True
								aSample2.Reported = True
								GlobalVariables.TempReportSamList.Add(aSample)
								GlobalVariables.TempReportSamList.Add(aSample2)
								blnLCSD = True
								intWksCount = intWksCount + 1
								blnGTG = True
								Exit For
							End If
						Next
						If Not blnLCSD And Not aSample.Reported Then
							aSample.Reported = True
							GlobalVariables.TempReportSamList.Add(aSample)
							intWksCount = intWksCount + 1
							blnGTG = True
						End If
					End If
				End If

			Next

			If blnGTG Then
				'Reset Reported 
				For Each aSample In GlobalVariables.TempReportSamList
					aSample.Reported = False
				Next
				For u = 1 To intWksCount
					GlobalVariables.workbook.Worksheets.Create("LCS-" & CStr(u))
					For Each wks In GlobalVariables.workbook.Worksheets
						If wks.name = "LCS-" & CStr(u) Then
							worksheet = wks
							'Begin building sheet..
							worksheet.Range("D1:F1").Merge()
							worksheet.Range("D1").Value = "LCS Recovery Report"
							worksheet.Range("A3").Value = "LIMS #"
							worksheet.Range("A4").Value = "Sample Point"
							worksheet.Range("A5").Value = "Sample Date"
							worksheet.Range("A6").Value = "Analysis Date"
							worksheet.Range("A7").Value = "Analysis Time"
							worksheet.Range("A8").Value = "Data Folder Name"
							worksheet.Range("A9").Value = "Analyte/Parameter"
							worksheet.Range("A3:A9").BorderInside()
							worksheet.Range("A3:A9").BorderAround()
							worksheet.Range(10, 1).FreezePanes()
							worksheet.Range("A3:A9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							aSample = GlobalVariables.TempReportSamList.Item(0)

							'List Compounds
							intTotalCols = 10
							For Each aCompound In aSample.CompoundList
								worksheet.Range(intTotalCols, 1).Value = aCompound.Name
								intTotalCols = intTotalCols + 1
							Next

							'List Surrogates
							If aSample.SurrogateList.Count > 0 Then
								intTotalCols = intTotalCols + 1
								worksheet.Range(intTotalCols, 1).Value = "Surrogate Recovery (%)"
								worksheet.Range(intTotalCols, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
								intTotalCols = intTotalCols + 1
								intSurrStart = intTotalCols
								For Each aSurrogate In aSample.SurrogateList
									worksheet.Range(intTotalCols, 1).Value = aSurrogate.Name
									intTotalCols = intTotalCols + 1
								Next
								worksheet.Range(intSurrStart, 1, intTotalCols, 1).BorderAround()
								worksheet.Range(10, 1, intSurrStart - 2, 1).BorderAround()
							Else
								intSurrStart = intTotalCols
								worksheet.Range(10, 1, intTotalCols - 1, 1).BorderAround()
							End If

							intCount = 2 'column start

							If GlobalVariables.TempReportSamList.Count > 1 Then
								aSample = GlobalVariables.TempReportSamList.Item(0)
								aSample2 = GlobalVariables.TempReportSamList.Item(1)
								If InStr(aSample2.Name, aSample.Name) And aSample.Type = "LCS" And aSample2.Type = "LCSD" Then
									blnLCSD = True
								Else
									blnLCSD = False
								End If
								aSample = Nothing
								aSample2 = Nothing
								If blnLCSD Then
									intLCSDCounter = 1
								Else
									intLCSDCounter = 0
								End If
							Else
								intLCSDCounter = 0
							End If


							For n = 0 To intLCSDCounter
								aSample = GlobalVariables.TempReportSamList.Item(n)
								worksheet.Range(3, intCount, 3, intCount + 4).Merge()
								worksheet.Range(3, intCount).Value = aSample.LimsID
								worksheet.Range(4, intCount, 4, intCount + 4).Merge()
								worksheet.Range(4, intCount).Value = aSample.Name
								worksheet.Range(5, intCount, 5, intCount + 4).Merge()
								worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
								worksheet.Range(6, intCount, 6, intCount + 4).Merge()
								worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
								worksheet.Range(7, intCount, 7, intCount + 4).Merge()
								worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
								worksheet.Range(8, intCount, 8, intCount + 4).Merge()
								worksheet.Range(8, intCount).Value = aSample.DataFile
								worksheet.Range(9, intCount).Value = "Amount " & aSample.Units & " (From Chemstation Report x Dil.Factor)"
								GlobalVariables.workbook.ActiveSheet.SetColumnWidthInPixels(intCount, 78)
								worksheet.Range(9, intCount).CellStyle.WrapText = True
								worksheet.Range(9, intCount + 1).Value = "Recovered Spiked Amount (" & aSample.ReportedUnits & ")"
								worksheet.Range(9, intCount + 1).CellStyle.WrapText = True
								worksheet.Range(9, intCount + 2).Value = "Corrected Spiked amount (" & aSample.ReportedUnits & ")"
								worksheet.Range(9, intCount + 2).CellStyle.WrapText = True
								worksheet.Range(9, intCount + 3).Value = "% Recovery"
								worksheet.Range(9, intCount + 3).CellStyle.WrapText = True
								worksheet.Range(9, intCount + 4).Value = "Recovery Limit Check"
								worksheet.Range(9, intCount + 4).CellStyle.WrapText = True
								worksheet.Range(3, intCount, 9, intCount + 4).BorderAround()
								worksheet.Range(3, intCount, 9, intCount + 4).BorderInside()
								worksheet.Range(3, intCount, 9, intCount + 4).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)

								'Begin analyte readout
								intCount2 = 10 'Row compounds start at
								For Each aCompound In aSample.CompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											If IsNumeric(aCompound.Conc) Then
												worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF((CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)))
											Else
												worksheet.Range(intCount2 + i, intCount).Text = "N.D."
											End If
											worksheet.Range(intCount2 + i, intCount + 1).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
											worksheet.Range(intCount2 + i, intCount + 2).Text = GlobalVariables.Calculations.FormatSF(aCompound.ChromCorrectedSpike)
											worksheet.Range(intCount2 + i, intCount + 3).Text = GlobalVariables.Calculations.FormatSF(aCompound.ChromSpikeRecovery)
											If aCompound.ChromSpikePass Then
												worksheet.Range(intCount2 + i, intCount + 4).Text = "Passed"
											Else
												worksheet.Range(intCount2 + i, intCount + 4).Text = "Failed"
											End If
											If Not aCompound.ChromSpikePass Then
												worksheet.Range(intCount2 + i, intCount + 4).CellStyle.Color = System.Drawing.Color.FromArgb(197, 217, 241)
											End If
										End If
									Next
								Next
								If aSample.SurrogateList.Count > 0 Then
									worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
									worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
									worksheet.Range(10, intCount + 2, intSurrStart - 1, intCount + 2).BorderAround()
									worksheet.Range(10, intCount + 3, intSurrStart - 1, intCount + 3).BorderAround()
									worksheet.Range(10, intCount + 4, intSurrStart - 1, intCount + 4).BorderAround()
								Else
									worksheet.Range(10, intCount, intTotalCols - 1, intCount).BorderAround()
									worksheet.Range(10, intCount + 1, intTotalCols - 1, intCount + 1).BorderAround()
									worksheet.Range(10, intCount + 2, intTotalCols - 1, intCount + 2).BorderAround()
									worksheet.Range(10, intCount + 3, intTotalCols - 1, intCount + 3).BorderAround()
									worksheet.Range(10, intCount + 4, intTotalCols - 1, intCount + 4).BorderAround()
								End If


								'Surrogate write out
								If aSample.SurrogateList.Count > 0 Then
									For Each aSurrogate In aSample.SurrogateList
										For i = 0 To aSample.SurrogateList.Count
											If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
												worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
											End If
										Next
									Next
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
									worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 4).BorderAround()
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
								End If
								intCount = intCount + 5
							Next

							'3 remaining columns

							worksheet.Range(9, intCount).Value = "% RPD"
							worksheet.Range(9, intCount + 1).Value = "Recovery Limit (%)"
							worksheet.Range(9, intCount + 2).Value = "RPD Limit (%)"
							worksheet.Range(3, intCount, 9, intCount + 2).BorderInside()
							worksheet.Range(3, intCount, 9, intCount + 2).BorderAround()
							worksheet.Range(3, intCount, 9, intCount + 2).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
							aSample = GlobalVariables.TempReportSamList.Item(0)
							'Begin analyte readout
							intCount2 = 10 'Row compounds start at
							For Each aCompound In aSample.CompoundList
								For i = 0 To aSample.CompoundList.Count
									If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
										worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.ChromRPD)
										worksheet.Range(intCount2 + i, intCount + 1).Value = aCompound.ChromLowLCSLim & "-" & aCompound.ChromUpLCSLim
										worksheet.Range(intCount2 + i, intCount + 2).Value = aCompound.ChromRPDLimit
									End If
								Next
							Next
							intCount2 = intSurrStart
							For Each aSurrogate In aSample.SurrogateList
								For i = 0 To aSample.SurrogateList.Count
									If worksheet.Range(intCount2 + i, 1).Value = aSurrogate.Name Then
										worksheet.Range(intCount2 + i, intCount).Value = aSurrogate.ChromLowLCSLim & "-" & aSurrogate.ChromUpLCSLim
									End If
								Next
							Next
							If aSample.SurrogateList.Count > 0 Then
								worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
								worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
								worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 2).BorderAround()
								worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
								worksheet.Range(11, 4, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.WrapText = True
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.Bold = True
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.Size = 8
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.FontName = "Arial"
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).AutofitColumns()
							Else
								worksheet.Range(10, intCount, intTotalCols - 1, intCount).BorderAround()
								worksheet.Range(10, intCount + 1, intTotalCols - 1, intCount + 1).BorderAround()
								worksheet.Range(10, intCount + 1, intTotalCols - 1, intCount + 2).BorderAround()
								worksheet.Range(11, 4, intTotalCols, intCount + 2).CellStyle.WrapText = True
								worksheet.Range(1, 2, intTotalCols, intCount + 2).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(1, 2, intTotalCols, intCount + 2).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								worksheet.Range(1, 1, intTotalCols, intCount + 2).CellStyle.Font.Bold = True
								worksheet.Range(1, 1, intTotalCols, intCount + 2).CellStyle.Font.Size = 8
								worksheet.Range(1, 1, intTotalCols, intCount + 2).CellStyle.Font.FontName = "Arial"
								worksheet.Range(1, 1, intTotalCols, intCount + 2).AutofitColumns()
							End If
							For n = 0 To intLCSDCounter
								GlobalVariables.TempReportSamList.RemoveAt(0)
							Next
						End If
					Next
				Next
				Return True

			Else
				MsgBox("Error generating report!" & vbCrLf &
					  "Sub Procedure: MidlandChromLCSReport()" & vbCrLf &
					  "Logic Error: Could not find LCS to generate report.", "(╯°□°)╯︵ ┻━┻")
				Return False
			End If
		Catch ex As Exception
			MsgBox("Error generating report!" & vbCrLf &
					  "Sub Procedure: MidlandChromLCSReport()" & vbCrLf &
					  "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
			Return False
		End Try


	End Function

	'Function MidlandChromMBReport(ByVal strPath As String) As Boolean
	'    Dim worksheet As IWorksheet
	'    Dim aSample As Sample
	'    Dim aCompound As Compound
	'    Dim intCount As Integer
	'    Dim intCount2 As Integer
	'    Dim intTotalRows As Integer
	'    Dim blnGTG As Boolean
	'    Dim exEngine As New ExcelEngine
	'    Dim exApp As IApplication
	'    Dim intWksCount As Integer
	'    exApp = exEngine.Excel

	'    Try
	'        'Import limits for MB
	'        intWksCount = 0
	'        If GlobalVariables.Import.MidlandChromBuildMBCompoundList(strPath) Then
	'            'Grab sample for this report
	'            blnGTG = False
	'            GlobalVariables.TempReportSamList.Clear()
	'            For Each aSample In GlobalVariables.ReportSamList
	'                If aSample.Include Then
	'                    If aSample.Type = "MB" Then
	'                        GlobalVariables.TempReportSamList.Add(aSample)
	'                        intWksCount = intWksCount + 1
	'                        blnGTG = True
	'                    End If
	'                End If

	'            Next
	'            'test changes
	'            If blnGTG Then
	'                For u = 1 To intWksCount
	'                    GlobalVariables.workbook.Worksheets.Create("MB-" & CStr(u))
	'                    For Each wks In GlobalVariables.workbook.Worksheets
	'                        If wks.name = "MB-" & CStr(u) Then
	'                            worksheet = wks
	'                            'Begin building sheet..
	'                            worksheet.Range("B1:C1").Merge()
	'                            worksheet.Range("B1").Value = GlobalVariables.strFreeportAnalysis & " Daily Blank Report"
	'                            worksheet.Range("A3").Value = "LIMS #"
	'                            worksheet.Range("A4").Value = "Sample Point"
	'                            worksheet.Range("A5").Value = "Sample Date"
	'                            worksheet.Range("A6").Value = "Analysis Date"
	'                            worksheet.Range("A7").Value = "Analysis Time"
	'                            worksheet.Range("A8").Value = "Data Folder Name"
	'                            worksheet.Range("A9").Value = "Dilution Factor"
	'                            worksheet.Range("A10").Value = "Analyte/Parameter"
	'                            worksheet.Range("A3:A10").BorderInside()
	'                            worksheet.Range("A3:A10").BorderAround()
	'                            worksheet.Range(10, 1).FreezePanes()
	'                            worksheet.Range("A3:A10").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
	'                            aSample = GlobalVariables.TempReportSamList.Item(0)

	'                            'List Compounds
	'                            intCount = 11
	'                            For Each aCompound In GlobalVariables.MidlandMBCompoundList
	'                                worksheet.Range(intCount, 1).Value = aCompound.Name
	'                                worksheet.Range(intCount, 3).Value = aCompound.ChromMBLim
	'                                intCount = intCount + 1
	'                            Next
	'                            intTotalRows = intCount - 1
	'                            worksheet.Range(11, 1, intTotalRows, 1).BorderAround()

	'                            intCount = 2 'column start

	'                            worksheet.Range(3, intCount).Value = aSample.LimsID
	'                            worksheet.Range(4, intCount).Value = aSample.Name
	'                            worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
	'                            worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
	'                            worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
	'                            worksheet.Range(8, intCount).Value = aSample.DataFile
	'                            worksheet.Range(9, intCount).Value = aSample.DilutionFactor
	'                            worksheet.Range(10, intCount).Value = "Amount (" & aSample.ReportedUnits & ") (From Chemstation Report)"
	'                            worksheet.Range(10, intCount).CellStyle.WrapText = True
	'                            worksheet.Range(3, intCount, 10, intCount).BorderAround()
	'                            worksheet.Range(3, intCount, 10, intCount).BorderInside()
	'                            worksheet.Range(3, intCount, 10, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
	'                            GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(10, 106)

	'                            'Begin analyte readout
	'                            intCount2 = 11 'Row compounds start at
	'                            For Each aCompound In aSample.CompoundList
	'                                For i = 0 To aSample.CompoundList.Count
	'                                    If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
	'                                        worksheet.Range(intCount2 + i, intCount).Text = aCompound.Conc
	'                                    End If
	'                                Next
	'                            Next
	'                            worksheet.Range(11, intCount, intTotalRows, intCount).BorderAround()

	'                            '3 remaining columns

	'                            intCount = 3

	'                            worksheet.Range(10, intCount).Value = "MAL/MDL Limit ucl (" & aSample.ReportedUnits & ")"
	'                            worksheet.Range(10, intCount + 1).Value = "Result"
	'                            worksheet.Range(3, intCount, 10, intCount + 1).BorderInside()
	'                            worksheet.Range(3, intCount, 10, intCount + 1).BorderAround()
	'                            worksheet.Range(3, intCount, 10, intCount + 1).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
	'                            aSample = GlobalVariables.TempReportSamList.Item(0)
	'                            'Begin analyte readout
	'                            intCount2 = 11 'Row compounds start at
	'                            For Each aCompound In GlobalVariables.MidlandMBCompoundList
	'                                For i = 0 To aSample.CompoundList.Count
	'                                    If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
	'                                        If worksheet.Range(intCount2 + i, 2).Value <> "N.D." And worksheet.Range(intCount2 + i, 2).Value <> "" Then
	'                                            If CDbl(worksheet.Range(intCount2 + i, 2).Value) <= CDbl(worksheet.Range(intCount2 + i, intCount).Value) Then
	'                                                worksheet.Range(intCount2 + i, intCount + 1).Value = "Pass"
	'                                            Else
	'                                                worksheet.Range(intCount2 + i, intCount + 1).Value = "Fail"
	'                                            End If
	'                                        Else
	'                                            worksheet.Range(intCount2 + i, intCount + 1).Value = "N/A"
	'                                        End If
	'                                    End If
	'                                Next
	'                            Next
	'                            worksheet.Range(11, intCount, intTotalRows, intCount).BorderAround()
	'                            worksheet.Range(11, intCount + 1, intTotalRows, intCount + 1).BorderAround()
	'                            worksheet.Range(1, 2, intTotalRows, intCount + 1).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
	'                            worksheet.Range(1, 2, intTotalRows, intCount + 1).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
	'                            worksheet.Range(1, 1, intTotalRows, intCount + 1).CellStyle.Font.Bold = True
	'                            worksheet.Range(1, 1, intTotalRows, intCount + 1).CellStyle.Font.Size = 8
	'                            worksheet.Range(1, 1, intTotalRows, intCount + 1).CellStyle.Font.FontName = "Arial"
	'                            worksheet.Range(1, 1, intTotalRows, intCount + 1).AutofitColumns()
	'                            'Clear out Samples so it is not reported twice
	'                            aSample = GlobalVariables.ReportSamList.Item(0)
	'                            GlobalVariables.ReportSamList.Remove(aSample)
	'                        End If
	'                    Next
	'                Next
	'                Return True
	'            Else
	'                Return False
	'            End If
	'        Else
	'            MsgBox("Error generating report!" & vbCrLf & _
	'                    "Sub Procedure: MidlandChromMBReport()" & vbCrLf & _
	'                    "Logic Error: Could not import Method Blank compound list", MsgBoxStyle.Critical)
	'            Return False
	'        End If
	'    Catch ex As Exception
	'        MsgBox("Error generating report!" & vbCrLf & _
	'                  "Sub Procedure: MidlandChromMBReport()" & vbCrLf & _
	'                  "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
	'        Return False
	'    End Try

	'End Function

	Function MidlandChromMSReport(ByVal strLimitType As String) As Boolean
		Dim exEngine As New ExcelEngine
		Dim worksheet As IWorksheet
		Dim aSample As Sample
		Dim aSample2 As Sample
		Dim aSample3 As Sample
		Dim aCompound As Compound
		Dim aSurrogate As Surrogate
		Dim blnGTG As Boolean
		Dim intCount As Integer
		Dim intCount2 As Integer
		Dim intSurrStart As Integer
		Dim exApp As IApplication
		Dim intWksCount As Integer
		Dim blnMSD As Boolean
		Dim intMSDCounter As Integer
		exApp = exEngine.Excel

		Try
			'Grab samples for this report
			blnGTG = False
			intWksCount = 0
			GlobalVariables.TempReportSamList.Clear()
			For Each aSample In GlobalVariables.ReportSamList
				If blnGTG Then
					Exit For
				End If
				If aSample.Include Then
					If aSample.Type = "MS" Then
						For Each aSample2 In GlobalVariables.ReportSamList
							If aSample2.Name = Trim(aSample.Name.Substring(0, aSample.Name.Length - 2)) And Not aSample.Reported And aSample.Methylated And aSample2.Methylated Then
								For Each aSample3 In GlobalVariables.ReportSamList
									If aSample3.Type = "MSD" And InStr(aSample3.Name, aSample2.Name) And Not aSample3.Reported And aSample3.Methylated And aSample2.Methylated Then
										GlobalVariables.TempReportSamList.Add(aSample2)
										GlobalVariables.TempReportSamList.Add(aSample)
										GlobalVariables.TempReportSamList.Add(aSample3)
										blnMSD = True
										intWksCount = intWksCount + 1
										blnGTG = True
										Exit For
									End If
								Next
								If Not blnMSD Then
									GlobalVariables.TempReportSamList.Add(aSample2)
									GlobalVariables.TempReportSamList.Add(aSample)
									intWksCount = intWksCount + 1
									blnGTG = True
									Exit For
								Else
									Exit For
								End If
							ElseIf aSample2.Name = Trim(aSample.Name.Substring(0, aSample.Name.Length - 2)) And Not aSample.Reported And Not aSample.Methylated And Not aSample2.Methylated Then
								For Each aSample3 In GlobalVariables.ReportSamList
									If aSample3.Type = "MSD" And InStr(aSample3.Name, aSample2.Name) And Not aSample3.Reported And Not aSample3.Methylated And Not aSample2.Methylated Then
										GlobalVariables.TempReportSamList.Add(aSample2)
										GlobalVariables.TempReportSamList.Add(aSample)
										GlobalVariables.TempReportSamList.Add(aSample3)
										blnMSD = True
										intWksCount = intWksCount + 1
										blnGTG = True
										Exit For
									End If
								Next
								If Not blnMSD Then
									GlobalVariables.TempReportSamList.Add(aSample2)
									GlobalVariables.TempReportSamList.Add(aSample)
									intWksCount = intWksCount + 1
									blnGTG = True
									Exit For
								Else
									Exit For
								End If
							End If
						Next
					End If
				End If

			Next

			If blnGTG Then
				'Matrix switch
				For u = 1 To intWksCount
					'blnMSD = False
					GlobalVariables.workbook.Worksheets.Create("MS-" & CStr(u))
					For Each wks In GlobalVariables.workbook.Worksheets
						If wks.name = "MS-" & CStr(u) Then
							worksheet = wks
							aSample = GlobalVariables.TempReportSamList.Item(0)
							If aSample.Matrix = "W" Then
								aSample = Nothing
								aSample2 = Nothing
								'Begin building sheet..
								worksheet.Range("D1:F1").Merge()
								worksheet.Range("D1").Value = "Spike Recovery Report"
								worksheet.Range("A3").Value = "LIMS #"
								worksheet.Range("A4").Value = "Sample Point"
								worksheet.Range("A5").Value = "Sample Date"
								worksheet.Range("A6").Value = "Analysis Date"
								worksheet.Range("A7").Value = "Analysis Time"
								worksheet.Range("A8").Value = "Data Folder Name"
								worksheet.Range("A9").Value = "Analyte/Parameter"
								worksheet.Range("A3:A9").BorderInside()
								worksheet.Range("A3:A9").BorderAround()
								worksheet.Range("A3:A9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								worksheet.Range(10, 1).FreezePanes()
								aSample = GlobalVariables.TempReportSamList.Item(0)

								'List Compounds
								intCount = 10
								For Each aCompound In aSample.CompoundList
									worksheet.Range(intCount, 1).Value = aCompound.Name
									intCount = intCount + 1
								Next
								worksheet.Range(10, 1, intCount, 1).BorderAround()

								'List Surrogates
								intCount = intCount + 1
								worksheet.Range(intCount, 1).Value = "Surrogate Recovery (%)"
								worksheet.Range(intCount, 2).BorderAround()
								worksheet.Range(intCount, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
								intCount = intCount + 1
								intSurrStart = intCount
								For Each aSurrogate In aSample.SurrogateList
									worksheet.Range(intCount, 1).Value = aSurrogate.Name
									intCount = intCount + 1
								Next
								worksheet.Range(intSurrStart, 1, intCount, 1).BorderAround()
								worksheet.Range(intSurrStart - 1, 1, intCount, 2).BorderAround()

								intCount = 2 'column start
								aSample = GlobalVariables.TempReportSamList.Item(0)
								worksheet.Range(3, intCount).Value = aSample.LimsID
								worksheet.Range(4, intCount).Value = aSample.Name
								worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
								worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
								worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
								worksheet.Range(8, intCount).Value = aSample.DataFile
								worksheet.Range(9, intCount).Value = "Amount (" & aSample.ReportedUnits & ") (From Chemstation Report) x Dil. Factor"
								'Begin analyte readout
								intCount2 = 10 'Row compounds start at
								For Each aCompound In aSample.CompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											If IsNumeric(aCompound.Conc) Then
												worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF(CStr(CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)))
											Else
												worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
											End If
										End If
									Next
								Next
								worksheet.Range(9, intCount).CellStyle.WrapText = True
								worksheet.Range(3, intCount, 9, intCount).BorderAround()
								worksheet.Range(3, intCount, 9, intCount).BorderInside()
								worksheet.Range(3, intCount, 9, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)
								'Surrogates
								For Each aSurrogate In aSample.SurrogateList
									For i = 0 To aSample.SurrogateList.Count
										If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
											worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery)
										End If
									Next
								Next
								worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
								worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 5).BorderAround()

								GlobalVariables.TempReportSamList.Remove(aSample)

								intCount = 3 'column start

								'If MSD then run twice, if not, then just once
								If blnMSD Then
									aSample = GlobalVariables.TempReportSamList.Item(0)
									aSample2 = GlobalVariables.TempReportSamList.Item(1)

									If InStr(aSample2.Name, aSample.Name) And aSample.Type = "MS" And aSample2.Type = "MSD" Then
										blnMSD = True
									Else
										blnMSD = False
									End If
									aSample = Nothing
									aSample2 = Nothing
									intMSDCounter = 1
								Else
									intMSDCounter = 0
								End If


								For n = 0 To intMSDCounter
									aSample = GlobalVariables.TempReportSamList.Item(0 + n)
									worksheet.Range(3, intCount, 3, intCount + 4).Merge()
									worksheet.Range(3, intCount).Value = aSample.LimsID
									worksheet.Range(4, intCount, 4, intCount + 4).Merge()
									worksheet.Range(4, intCount).Value = aSample.Name
									worksheet.Range(5, intCount, 5, intCount + 4).Merge()
									worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
									worksheet.Range(6, intCount, 6, intCount + 4).Merge()
									worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
									worksheet.Range(7, intCount, 7, intCount + 4).Merge()
									worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
									worksheet.Range(8, intCount, 8, intCount + 4).Merge()
									worksheet.Range(8, intCount).Value = aSample.DataFile
									worksheet.Range(9, intCount).Value = "Amount " & aSample.Units & " (From Chemstation Report x Dil.Factor)"
									GlobalVariables.workbook.ActiveSheet.SetColumnWidthInPixels(intCount, 78)
									worksheet.Range(9, intCount).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 1).Value = "Recovered Spiked Amount (" & aSample.ReportedUnits & ")"
									worksheet.Range(9, intCount + 1).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 2).Value = "Corrected Spiked amount (" & aSample.ReportedUnits & ")"
									worksheet.Range(9, intCount + 2).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 3).Value = "% Recovery"
									worksheet.Range(9, intCount + 3).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 4).Value = "Recovery Limit Check"
									worksheet.Range(9, intCount + 4).CellStyle.WrapText = True
									worksheet.Range(3, intCount, 9, intCount + 4).BorderAround()
									worksheet.Range(3, intCount, 9, intCount + 4).BorderInside()
									worksheet.Range(3, intCount, 9, intCount + 4).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)

									'Begin analyte readout
									intCount2 = 10 'Row compounds start at
									For Each aCompound In aSample.CompoundList
										For i = 0 To aSample.CompoundList.Count
											If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
												If IsNumeric(aCompound.Conc) Then
													worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF(CStr(CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)))
												Else
													worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
												End If
												worksheet.Range(intCount2 + i, intCount + 1).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
												worksheet.Range(intCount2 + i, intCount + 2).Text = GlobalVariables.Calculations.FormatSF(aCompound.ChromCorrectedSpike)
												worksheet.Range(intCount2 + i, intCount + 3).Text = GlobalVariables.Calculations.FormatSF(aCompound.ChromSpikeRecovery)
												If aCompound.ChromSpikePass Then
													worksheet.Range(intCount2 + i, intCount + 4).Text = "Passed"
												Else
													worksheet.Range(intCount2 + i, intCount + 4).Text = "Failed"
												End If
												If Not aCompound.ChromSpikePass Then
													worksheet.Range(intCount2 + i, intCount + 4).CellStyle.Color = System.Drawing.Color.FromArgb(197, 217, 241)
												End If
											End If
										Next
									Next
									worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
									worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
									worksheet.Range(10, intCount + 2, intSurrStart - 1, intCount + 2).BorderAround()
									worksheet.Range(10, intCount + 3, intSurrStart - 1, intCount + 3).BorderAround()
									worksheet.Range(10, intCount + 4, intSurrStart - 1, intCount + 4).BorderAround()

									'Surrogate write out
									For Each aSurrogate In aSample.SurrogateList
										For i = 0 To aSample.SurrogateList.Count
											If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
												worksheet.Range(intSurrStart + i, intCount).Value = aSurrogate.Recovery
											End If
										Next
									Next
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
									worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 4).BorderAround()
									intCount = intCount + 5
								Next


								'3 remaining columns
								If blnMSD Then
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount + 2).BorderAround()
									worksheet.Range(9, intCount).Value = "% RPD"
									worksheet.Range(9, intCount + 1).Value = "Recovery Limit (%)"
									worksheet.Range(9, intCount + 2).Value = "RPD Limit (%)"
									worksheet.Range(3, intCount, 9, intCount + 2).BorderInside()
									worksheet.Range(3, intCount, 9, intCount + 2).BorderAround()
									worksheet.Range(3, intCount, 9, intCount + 2).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
									aSample = GlobalVariables.TempReportSamList.Item(0)
									'Begin analyte readout
									intCount2 = 10 'Row compounds start at
									For Each aCompound In aSample.CompoundList
										For i = 0 To aSample.CompoundList.Count
											If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
												worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.ChromRPD)
												worksheet.Range(intCount2 + i, intCount + 1).Value = aCompound.ChromLowMSLim & "-" & aCompound.ChromUpMSLim
												worksheet.Range(intCount2 + i, intCount + 2).Value = aCompound.ChromRPDLimit
											End If
										Next
									Next
									worksheet.Range(9, intCount, intSurrStart - 1, intCount).BorderAround()
									worksheet.Range(9, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
									worksheet.Range(9, intCount + 1, intSurrStart - 1, intCount + 2).BorderAround()
									worksheet.Range(10, 4, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.WrapText = True
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.Bold = True
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.Size = 8
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.FontName = "Arial"
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).AutofitColumns()
								Else
									'worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
									'worksheet.Range(9, intCount, intSurrStart - 1, intCount).BorderAround()
									'worksheet.Range(9, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
									'worksheet.Range(9, intCount + 1, intSurrStart - 1, intCount + 2).BorderAround()
									worksheet.Range(10, 4, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.WrapText = True
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Bold = True
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Size = 8
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.FontName = "Arial"
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).AutofitColumns()
								End If

								For n = 0 To intMSDCounter
									GlobalVariables.TempReportSamList.RemoveAt(0)
								Next

							ElseIf aSample.Matrix = "S" Then
								aSample = Nothing
								aSample2 = Nothing

								'Begin building sheet..
								worksheet.Range("D1:F1").Merge()
								worksheet.Range("D1").Value = "Spike Recovery Report"
								worksheet.Range("A3").Value = "LIMS #"
								worksheet.Range("A4").Value = "Sample Point"
								worksheet.Range("A5").Value = "Sample Date"
								worksheet.Range("A6").Value = "Analysis Date"
								worksheet.Range("A7").Value = "Analysis Time"
								worksheet.Range("A8").Value = "Data Folder Name"
								worksheet.Range("A9").Value = "Analyte/Parameter"
								worksheet.Range("A3:A9").BorderInside()
								worksheet.Range("A3:A9").BorderAround()
								worksheet.Range(10, 1).FreezePanes()
								worksheet.Range("A3:A9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								aSample = GlobalVariables.TempReportSamList.Item(0)

								'List Compounds
								intCount = 10
								For Each aCompound In aSample.CompoundList
									worksheet.Range(intCount, 1).Value = aCompound.Name
									intCount = intCount + 1
								Next
								worksheet.Range(10, 1, intCount + 1, 1).BorderAround()

								'List Surrogates
								intCount = intCount + 1
								worksheet.Range(intCount, 1).Value = "Surrogate Recovery (%)"
								worksheet.Range(intCount, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
								intCount = intCount + 1
								intSurrStart = intCount
								For Each aSurrogate In aSample.SurrogateList
									worksheet.Range(intCount, 1).Value = aSurrogate.Name
									intCount = intCount + 1
								Next
								worksheet.Range(intSurrStart - 1, 1, intCount, 1).BorderAround()
								worksheet.Range(intSurrStart - 1, 1, intCount, 4).BorderAround()

								intCount = 2 'column start
								aSample = GlobalVariables.TempReportSamList.Item(0)
								worksheet.Range(3, intCount, 3, intCount + 2).Merge()
								worksheet.Range(3, intCount).Value = aSample.LimsID
								worksheet.Range(4, intCount, 4, intCount + 2).Merge()
								worksheet.Range(4, intCount).Value = aSample.Name
								worksheet.Range(5, intCount, 5, intCount + 2).Merge()
								worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
								worksheet.Range(6, intCount, 6, intCount + 2).Merge()
								worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
								worksheet.Range(7, intCount, 7, intCount + 2).Merge()
								worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
								worksheet.Range(8, intCount, 8, intCount + 2).Merge()
								worksheet.Range(8, intCount).Value = aSample.DataFile
								worksheet.Range(9, intCount).Value = "Amount " & aSample.Units & " (From Chemstation Report)"
								worksheet.Range(9, intCount).CellStyle.WrapText = True
								worksheet.Range(9, intCount + 1).Value = "Factor (" & aSample.Units & " to (ug/Kg) of Sample)"
								worksheet.Range(9, intCount + 1).CellStyle.WrapText = True
								worksheet.Range(9, intCount + 2).Value = "Amount (ug/Kg)"
								worksheet.Range(9, intCount + 2).CellStyle.WrapText = True
								worksheet.Range(3, intCount, 9, intCount + 2).BorderAround()
								worksheet.Range(3, intCount, 9, intCount + 2).BorderInside()
								worksheet.Range(3, intCount, 9, intCount + 2).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(7, 106)

								'Begin analyte readout
								intCount2 = 10 'Row compounds start at
								For Each aCompound In aSample.CompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
											worksheet.Range(intCount2 + i, intCount + 1).Value = GlobalVariables.Calculations.FormatSF(aSample.DilutionFactor)
											If IsNumeric(aCompound.Conc) Then
												worksheet.Range(intCount2 + i, intCount + 2).Text = GlobalVariables.Calculations.FormatSF((CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)))
											Else
												worksheet.Range(intCount2 + i, intCount + 2).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
											End If
										End If
									Next
								Next
								worksheet.Range(9, intCount).CellStyle.WrapText = True
								worksheet.Range(3, intCount, 9, intCount).BorderAround()
								worksheet.Range(3, intCount, 9, intCount).BorderInside()
								worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
								worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
								worksheet.Range(10, intCount + 2, intSurrStart - 1, intCount + 2).BorderAround()
								worksheet.Range(3, intCount, 9, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 106)
								'Surrogates
								For Each aSurrogate In aSample.SurrogateList
									For i = 0 To aSample.SurrogateList.Count
										If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
											worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery)
										End If
									Next
								Next
								worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
								worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 3).BorderAround()

								GlobalVariables.TempReportSamList.Remove(aSample)

								intCount = 5 'column start

								'If MSD then run twice, if not, then just once

								If blnMSD Then
									aSample = GlobalVariables.TempReportSamList.Item(0)
									aSample2 = GlobalVariables.TempReportSamList.Item(1)

									If InStr(aSample2.Name, aSample.Name) And aSample.Type = "MS" And aSample2.Type = "MSD" Then
										blnMSD = True
									Else
										blnMSD = False
									End If
									aSample = Nothing
									aSample2 = Nothing
									intMSDCounter = 1
								Else
									intMSDCounter = 0
								End If

								For n = 0 To intMSDCounter
									aSample = GlobalVariables.TempReportSamList.Item(0 + n)
									worksheet.Range(3, intCount, 3, intCount + 7).Merge()
									worksheet.Range(3, intCount).Value = aSample.LimsID
									worksheet.Range(4, intCount, 4, intCount + 7).Merge()
									worksheet.Range(4, intCount).Value = aSample.Name
									worksheet.Range(5, intCount, 5, intCount + 7).Merge()
									worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
									worksheet.Range(6, intCount, 6, intCount + 7).Merge()
									worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
									worksheet.Range(7, intCount, 7, intCount + 7).Merge()
									worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
									worksheet.Range(8, intCount, 8, intCount + 7).Merge()
									worksheet.Range(8, intCount).Value = aSample.DataFile
									worksheet.Range(9, intCount).Value = "Amount " & aSample.Units & " (From Chemstation Report)"
									worksheet.Range(9, intCount).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 1).Value = "Factor (" & aSample.Units & " to (ug/Kg) of Sample)"
									worksheet.Range(9, intCount + 1).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 2).Value = "Amount (ug/Kg)"
									worksheet.Range(9, intCount + 2).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 3).Value = "Rec. Spiked Amount (ug/Kg) of Sample"
									worksheet.Range(9, intCount + 3).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 4).Value = "Rec. Spiked Amount " & aSample.Units & ""
									worksheet.Range(9, intCount + 4).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 5).Value = "Corrected Spiked Amount " & aSample.Units & ""
									worksheet.Range(9, intCount + 5).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 6).Value = "% Recovery"
									worksheet.Range(9, intCount + 6).CellStyle.WrapText = True
									worksheet.Range(9, intCount + 7).Value = "Recovery Limit Check"
									worksheet.Range(9, intCount + 7).CellStyle.WrapText = True
									worksheet.Range(3, intCount, 9, intCount + 7).BorderAround()
									worksheet.Range(3, intCount, 9, intCount + 7).BorderInside()
									worksheet.Range(3, intCount, 9, intCount + 7).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)

									'Begin analyte readout
									intCount2 = 10 'Row compounds start at
									For Each aCompound In aSample.CompoundList
										For i = 0 To aSample.CompoundList.Count
											If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
												worksheet.Range(intCount2 + i, intCount).Value = aCompound.Conc
												worksheet.Range(intCount2 + i, intCount + 1).Value = aSample.DilutionFactor
												If IsNumeric(aCompound.Conc) Then
													worksheet.Range(intCount2 + i, intCount + 2).Text = GlobalVariables.Calculations.FormatSF((CDbl(aCompound.Conc) * CDbl(aSample.DilutionFactor)))
												Else
													worksheet.Range(intCount2 + i, intCount + 2).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
												End If
												worksheet.Range(intCount2 + i, intCount + 3).Text = GlobalVariables.Calculations.FormatSF(CStr(CDbl(worksheet.Range(intCount2 + i, intCount + 2).Value) - CDbl(worksheet.Range(intCount2 + i, 4).Value)))
												worksheet.Range(intCount2 + i, intCount + 4).Text = GlobalVariables.Calculations.FormatSF(aCompound.Conc)
												worksheet.Range(intCount2 + i, intCount + 5).Text = GlobalVariables.Calculations.FormatSF(aCompound.ChromCorrectedSpike)
												worksheet.Range(intCount2 + i, intCount + 6).Text = GlobalVariables.Calculations.FormatSF(aCompound.ChromSpikeRecovery)
												If aCompound.ChromSpikePass Then
													worksheet.Range(intCount2 + i, intCount + 7).Text = "Passed"
												Else
													worksheet.Range(intCount2 + i, intCount + 7).Text = "Failed"
												End If
												If Not aCompound.ChromSpikePass Then
													worksheet.Range(intCount2 + i, intCount + 7).CellStyle.Color = System.Drawing.Color.FromArgb(197, 217, 241)
												End If
											End If
										Next
									Next
									worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
									worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
									worksheet.Range(10, intCount + 2, intSurrStart - 1, intCount + 2).BorderAround()
									worksheet.Range(10, intCount + 3, intSurrStart - 1, intCount + 3).BorderAround()
									worksheet.Range(10, intCount + 4, intSurrStart - 1, intCount + 4).BorderAround()
									worksheet.Range(10, intCount + 5, intSurrStart - 1, intCount + 5).BorderAround()
									worksheet.Range(10, intCount + 6, intSurrStart - 1, intCount + 6).BorderAround()
									worksheet.Range(10, intCount + 7, intSurrStart - 1, intCount + 7).BorderAround()

									'Surrogate write out
									For Each aSurrogate In aSample.SurrogateList
										For i = 0 To aSample.SurrogateList.Count
											If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
												worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
											End If
										Next
									Next
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
									worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 7).BorderAround()
									intCount = intCount + 8
								Next

								If blnMSD Then
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount + 2).BorderAround()

									'3 remaining columns

									worksheet.Range(9, intCount).Value = "% RPD"
									worksheet.Range(9, intCount + 1).Value = "Recovery Limit (%)"
									worksheet.Range(9, intCount + 2).Value = "RPD Limit (%)"
									worksheet.Range(3, intCount, 9, intCount + 2).BorderInside()
									worksheet.Range(3, intCount, 9, intCount + 2).BorderAround()
									worksheet.Range(3, intCount, 9, intCount + 2).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
									aSample = GlobalVariables.TempReportSamList.Item(0)
									'Begin analyte readout
									intCount2 = 10 'Row compounds start at
									For Each aCompound In aSample.CompoundList
										For i = 0 To aSample.CompoundList.Count
											If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
												worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.ChromRPD)
												worksheet.Range(intCount2 + i, intCount + 1).Value = aCompound.ChromLowContLim & "-" & aCompound.ChromUpContLim
												worksheet.Range(intCount2 + i, intCount + 2).Value = aCompound.ChromRPDLimit
											End If
										Next
									Next
									worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
									worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()
									worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 2).BorderAround()
									worksheet.Range(11, 4, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.WrapText = True
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.Bold = True
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.Size = 8
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).CellStyle.Font.FontName = "Arial"
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 2).AutofitColumns()
								Else
									worksheet.Range(11, 4, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.WrapText = True
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Bold = True
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Size = 8
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.FontName = "Arial"
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).AutofitColumns()
								End If

								For n = 0 To intMSDCounter
									GlobalVariables.TempReportSamList.RemoveAt(0)
								Next
							End If
						End If
					Next
				Next
				Return True
			Else
				MsgBox("Error generating report!" & vbCrLf &
					  "Sub Procedure: MidlandChromMSReport()" & vbCrLf &
					  "Logic Error: Could not find MS/MSD pair to generate report.", "(╯°□°)╯︵ ┻━┻")
				Return False
			End If
		Catch ex As Exception
			MsgBox("Error generating report!" & vbCrLf &
					  "Sub Procedure: MidlandChromMSReport()" & vbCrLf &
					  "Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
			Return False
		End Try
	End Function

	'Midland Summary Report Methylated
	Function MidlandChromSummaryMethReport(ByVal strLimitType As String) As Boolean
		Dim worksheet As IWorksheet
		Dim aSample As Sample
		Dim aSample2 As Sample
		Dim aPermit As Permit
		Dim aProject As Project
		Dim aInstrument As mInstrument
		Dim amCompound As mCompound
		Dim aCompound As Compound
		Dim aSurrogate As Surrogate
		Dim intCount As Integer
		Dim intCount2 As Integer
		Dim intSurrStart As Integer
		Dim exEngine As New ExcelEngine
		Dim exApp As IApplication
		exApp = exEngine.Excel

		Try
			GlobalVariables.workbook.Worksheets.Create("Summary MED Report")
			For Each wks In GlobalVariables.workbook.Worksheets
				If wks.name = "Summary MED Report" Then
					worksheet = wks
					aSample = GlobalVariables.ReportSamList.Item(0)
					If strLimitType <> "RL" Then
						'Begin building sheet..
						worksheet.Range("D1:E1").Merge()
						worksheet.Range("D1").Value = "Summary Report"
						worksheet.Range("A3").Value = "LIMS #"
						worksheet.Range("A4").Value = "Sample Point"
						worksheet.Range("A5").Value = "Sample Date"
						worksheet.Range("A6").Value = "Analysis Date"
						worksheet.Range("A7").Value = "Analysis Time"
						worksheet.Range("A8").Value = "Data Folder Name"
						worksheet.Range("A9").Value = "Analyte/Parameter"
						worksheet.Range("B9").Value = "CAS #"
						worksheet.Range("C9").Value = strLimitType & " (" & aSample.ReportedUnits & ")"
						worksheet.Range("A3:C9").BorderInside()
						worksheet.Range("A3:C9").BorderAround()
						worksheet.Range(10, 1).FreezePanes()
						worksheet.Range("A3:C9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
						For Each aSample2 In GlobalVariables.ReportSamList
							If aSample2.Methylated Then
								aSample = aSample2
								Exit For
							End If
						Next

						'List Compounds
						intCount = 10
						For Each aPermit In GlobalVariables.PermitList
							If aPermit.Name = GlobalVariables.selPermit.Name Then
								For Each aProject In aPermit.ProjectList
									If aProject.Name = GlobalVariables.selProject Then
										For Each aInstrument In aProject.mInstrumentList
											If aInstrument.Name = GlobalVariables.selInstrument Then
												For Each amCompound In aInstrument.mCompoundList
													For Each aCompound In aSample.CompoundList
														If amCompound.Name = aCompound.Name And aCompound.Methylated Then
															worksheet.Range(intCount, 1).Value = aCompound.Name
															worksheet.Range(intCount, 2).Value = amCompound.CAS
															If strLimitType = "MDL" Then
																worksheet.Range(intCount, 3).Value = amCompound.MDL
															ElseIf strLimitType = "PQL" Then
																worksheet.Range(intCount, 3).Value = amCompound.PQL
															End If
															intCount = intCount + 1
														End If
													Next
												Next
											End If
										Next
									End If
								Next
							End If
						Next
						worksheet.Range(10, 1, intCount, 1).BorderAround()
						worksheet.Range(10, 2, intCount, 2).BorderAround()
						worksheet.Range(10, 3, intCount, 3).BorderAround()

						'List Surrogates
						intCount = intCount + 1
						If aSample.SurrogateList.Count > 0 Then
							worksheet.Range(intCount, 1).Value = "Surrogate Recovery (%)"
							worksheet.Range(intCount, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
							worksheet.Range(intCount, 3).Value = "Recovery Limit (%)"
							worksheet.Range(intCount, 3).CellStyle.Font.Color = ExcelKnownColors.Blue
							intCount = intCount + 1
							intSurrStart = intCount
							For Each aSurrogate In aSample.SurrogateList
								If aSurrogate.Methylated Then
									worksheet.Range(intCount, 1).Value = aSurrogate.Name
									worksheet.Range(intCount, 3).Value = aSurrogate.ChromLowContLim & "-" & aSurrogate.ChromUpContLim
									intCount = intCount + 1
								End If
							Next
							worksheet.Range(intSurrStart, 1, intCount, 1).BorderAround()
							worksheet.Range(intSurrStart, 2, intCount, 2).BorderAround()
							worksheet.Range(intSurrStart, 3, intCount, 3).BorderAround()
							worksheet.Range(intSurrStart - 1, 1).BorderAround()
							worksheet.Range(intSurrStart - 1, 2).BorderAround()
							worksheet.Range(intSurrStart - 1, 3).BorderAround()
						Else
							intSurrStart = intCount
						End If


						intCount = 4
						For Each aSample In GlobalVariables.ReportSamList
							If aSample.Include Then
								If aSample.Methylated Then
									worksheet.Range(3, intCount, 3, intCount + 1).Merge()
									worksheet.Range(3, intCount).Value = aSample.LimsID
									worksheet.Range(4, intCount, 4, intCount + 1).Merge()
									worksheet.Range(4, intCount).Value = aSample.Name
									worksheet.Range(5, intCount, 5, intCount + 1).Merge()
									worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
									worksheet.Range(6, intCount, 6, intCount + 1).Merge()
									worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
									worksheet.Range(7, intCount, 7, intCount + 1).Merge()
									worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
									worksheet.Range(8, intCount, 8, intCount + 1).Merge()
									worksheet.Range(8, intCount).Value = aSample.DataFile
									worksheet.Range(3, intCount, 8, intCount + 1).BorderAround()
									worksheet.Range(3, intCount, 8, intCount + 1).BorderInside()
									worksheet.Range(3, intCount, 8, intCount + 1).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
									worksheet.Range(9, intCount).Value = "Adjusted Limit (" & aSample.ReportedUnits & ")"
									worksheet.Range(9, intCount + 1).Value = "Amount (" & aSample.ReportedUnits & ")"
									worksheet.Range(9, intCount).BorderAround()
									worksheet.Range(9, intCount + 1).BorderAround()
									worksheet.Range(9, intCount, 9, intCount + 1).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
									GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 45)

									'Begin analyte readout
									intCount2 = 10 'Row compounds start at
									For Each aCompound In aSample.CompoundList
										For i = 0 To aSample.CompoundList.Count
											If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
												worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.ChromAdjustedLimit)
												worksheet.Range(intCount2 + i, intCount + 1).Text = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
											End If
										Next
									Next
									worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
									worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()

									'Surrogate write out
									If aSample.SurrogateList.Count > 0 Then
										For Each aSurrogate In aSample.SurrogateList
											For i = 0 To aSample.SurrogateList.Count
												If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
													worksheet.Range(intSurrStart + i, intCount + 1).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
												End If
											Next
										Next
										worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
										worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).BorderAround()
									End If

									worksheet.Range(9, 4, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.WrapText = True
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.Font.Bold = True
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.Font.Size = 8
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.Font.FontName = "Arial"
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).AutofitColumns()

									intCount = intCount + 2
								End If
							End If

						Next
					Else
						'RL Limit or N/A selected
						aSample = GlobalVariables.ReportSamList.Item(0)
						'Begin building sheet..
						worksheet.Range("D1:E1").Merge()
						worksheet.Range("D1").Value = "Summary Report"
						worksheet.Range("A3").Value = "LIMS #"
						worksheet.Range("A4").Value = "Sample Point"
						worksheet.Range("A5").Value = "Sample Date"
						worksheet.Range("A6").Value = "Analysis Date"
						worksheet.Range("A7").Value = "Analysis Time"
						worksheet.Range("A8").Value = "Data Folder Name"
						worksheet.Range("A9").Value = "Analyte/Parameter"
						worksheet.Range("B9").Value = "CAS #"

						If strLimitType = "RL" Then
							worksheet.Range("C9").Value = strLimitType & " " & aSample.ReportedUnits
						Else
							worksheet.Range("C9").Value = strLimitType
						End If

						worksheet.Range("A3:C9").BorderInside()
						worksheet.Range("A3:C9").BorderAround()
						worksheet.Range(10, 1).FreezePanes()
						worksheet.Range("A3:C9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
						aSample = GlobalVariables.ReportSamList.Item(0)

						'List Compounds
						intCount = 10
						For Each aPermit In GlobalVariables.PermitList
							If aPermit.Name = GlobalVariables.selPermit.Name Then
								For Each aProject In aPermit.ProjectList
									If aProject.Name = GlobalVariables.selProject Then
										For Each aInstrument In aProject.mInstrumentList
											If aInstrument.Name = GlobalVariables.selInstrument Then
												For Each amCompound In aInstrument.mCompoundList
													For Each aCompound In aSample.CompoundList
														If amCompound.Name = aCompound.Name And aCompound.Methylated Then
															worksheet.Range(intCount, 1).Value = aCompound.Name
															worksheet.Range(intCount, 2).Value = amCompound.CAS
															If strLimitType = "RL" Then
																worksheet.Range(intCount, 3).Value = amCompound.RL
															Else
																worksheet.Range(intCount, 3).Value = "N/A"
															End If

															intCount = intCount + 1
														End If
													Next
												Next
											End If
										Next
									End If
								Next
							End If
						Next
						worksheet.Range(10, 1, intCount, 1).BorderAround()
						worksheet.Range(10, 2, intCount, 2).BorderAround()
						worksheet.Range(10, 3, intCount, 3).BorderAround()

						'List Surrogates
						intCount = intCount + 1
						If aSample.SurrogateList.Count > 0 Then
							worksheet.Range(intCount, 1).Value = "Surrogate Recovery (%)"
							worksheet.Range(intCount, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
							worksheet.Range(intCount, 3).Value = "Recovery Limit (%)"
							worksheet.Range(intCount, 3).CellStyle.Font.Color = ExcelKnownColors.Blue
							intCount = intCount + 1
							intSurrStart = intCount
							For Each aSurrogate In aSample.SurrogateList
								If aSurrogate.Methylated Then
									worksheet.Range(intCount, 1).Value = aSurrogate.Name
									worksheet.Range(intCount, 3).Value = aSurrogate.ChromLowContLim & "-" & aSurrogate.ChromUpContLim
									intCount = intCount + 1
								End If
							Next
							worksheet.Range(intSurrStart, 1, intCount, 1).BorderAround()
							worksheet.Range(intSurrStart, 2, intCount, 2).BorderAround()
							worksheet.Range(intSurrStart, 3, intCount, 3).BorderAround()
							worksheet.Range(intSurrStart - 1, 1).BorderAround()
							worksheet.Range(intSurrStart - 1, 2).BorderAround()
							worksheet.Range(intSurrStart - 1, 3).BorderAround()
						Else
							intSurrStart = intCount
						End If


						intCount = 4
						For Each aSample In GlobalVariables.ReportSamList
							If aSample.Include Then
								If aSample.Methylated Then
									worksheet.Range(3, intCount).Value = aSample.LimsID
									worksheet.Range(4, intCount).Value = aSample.Name
									worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
									worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
									worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
									worksheet.Range(8, intCount).Value = aSample.DataFile
									worksheet.Range(3, intCount, 9, intCount).BorderAround()
									worksheet.Range(3, intCount, 9, intCount).BorderInside()
									worksheet.Range(3, intCount, 9, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
									worksheet.Range(9, intCount).Value = "Amount " & aSample.ReportedUnits
									GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 45)

									'Begin analyte readout
									intCount2 = 10 'Row compounds start at
									For Each aCompound In aSample.CompoundList
										For i = 0 To aSample.CompoundList.Count
											If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
												worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
											End If
										Next
									Next
									worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()

									'Surrogate write out
									If aSample.SurrogateList.Count > 0 Then
										For Each aSurrogate In aSample.SurrogateList
											For i = 0 To aSample.SurrogateList.Count
												If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
													worksheet.Range(intSurrStart + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
												End If
											Next
										Next
										worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
									End If
									worksheet.Range(9, 4, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.WrapText = True
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
									worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Bold = True
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Size = 8
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.FontName = "Arial"
									worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).AutofitColumns()

									intCount = intCount + 1
								End If
							End If

						Next
					End If

					Return True
				End If
			Next
		Catch ex As Exception
			MsgBox("Error generating report!" & vbCrLf &
						"Sub Procedure: MidlandChromSummaryMethReport()" & vbCrLf &
						"Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
			Return False
		End Try

	End Function

	'Midland Summary Report
	Function MidlandChromSummaryReport(ByVal strLimitType As String) As Boolean
		Dim worksheet As IWorksheet
		Dim aSample As Sample
		Dim aPermit As Permit
		Dim aProject As Project
		Dim aInstrument As mInstrument
		Dim amCompound As mCompound
		Dim aCompound As Compound
		Dim aSurrogate As Surrogate
		Dim intCount As Integer
		Dim intCount2 As Integer
		Dim intSurrStart As Integer
		Dim exEngine As New ExcelEngine
		Dim exApp As IApplication
		Dim blnMethylated As Boolean
		exApp = exEngine.Excel

		blnMethylated = False
		For Each aSample In GlobalVariables.ReportSamList
			If aSample.Methylated Then
				blnMethylated = True
			End If
		Next

		'If methylated, means we need to make a methylated sheet as well
		If blnMethylated Then
			MidlandChromSummaryMethReport(strLimitType)
		End If


		Try
			GlobalVariables.workbook.Worksheets.Create("Summary Report")
			For Each wks In GlobalVariables.workbook.Worksheets
				If wks.name = "Summary Report" Then
					worksheet = wks
					aSample = GlobalVariables.ReportSamList.Item(0)
					If strLimitType <> "RL" Then
						'Begin building sheet..
						worksheet.Range("D1:E1").Merge()
						worksheet.Range("D1").Value = "Summary Report"
						worksheet.Range("A3").Value = "LIMS #"
						worksheet.Range("A4").Value = "Sample Point"
						worksheet.Range("A5").Value = "Sample Date"
						worksheet.Range("A6").Value = "Analysis Date"
						worksheet.Range("A7").Value = "Analysis Time"
						worksheet.Range("A8").Value = "Data Folder Name"
						worksheet.Range("A9").Value = "Analyte/Parameter"
						worksheet.Range("B9").Value = "CAS #"
						worksheet.Range("C9").Value = strLimitType & " (" & aSample.ReportedUnits & ")"
						worksheet.Range("A3:C9").BorderInside()
						worksheet.Range("A3:C9").BorderAround()
						worksheet.Range("A3:C9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
						worksheet.Range(10, 1).FreezePanes()

						aSample = GlobalVariables.ReportSamList.Item(0)

						'List components
						'intCount = 10
						'For Each aCompound In aSample.CompoundList

						'Next


						'List Compounds
						intCount = 10
						For Each aPermit In GlobalVariables.PermitList
							If aPermit.Name = GlobalVariables.selPermit.Name Then
								For Each aProject In aPermit.ProjectList
									If aProject.Name = GlobalVariables.selProject Then
										For Each aInstrument In aProject.mInstrumentList
											If aInstrument.Name = GlobalVariables.selInstrument Then
												For Each aCompound In aSample.CompoundList
													For Each amCompound In aInstrument.mCompoundList
														If amCompound.Name = aCompound.Name Then
															worksheet.Range(intCount, 1).Value = aCompound.Name
															worksheet.Range(intCount, 2).Value = amCompound.CAS
															If strLimitType = "MDL" Then
																worksheet.Range(intCount, 3).Value = amCompound.MDL
															ElseIf strLimitType = "PQL" Then
																worksheet.Range(intCount, 3).Value = amCompound.PQL
															End If
															intCount = intCount + 1
														End If
													Next
												Next
											End If
										Next
									End If
								Next
							End If
						Next
						worksheet.Range(10, 1, intCount, 1).BorderAround()
						worksheet.Range(10, 2, intCount, 2).BorderAround()
						worksheet.Range(10, 3, intCount, 3).BorderAround()

						'List Surrogates
						intCount = intCount + 1
						If aSample.SurrogateList.Count > 0 Then
							worksheet.Range(intCount, 1).Value = "Surrogate Recovery (%)"
							worksheet.Range(intCount, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
							worksheet.Range(intCount, 3).Value = "Recovery Limit (%)"
							worksheet.Range(intCount, 3).CellStyle.Font.Color = ExcelKnownColors.Blue
							intCount = intCount + 1
							intSurrStart = intCount
							For Each aSurrogate In aSample.SurrogateList
								worksheet.Range(intCount, 1).Value = aSurrogate.Name
								worksheet.Range(intCount, 3).Value = aSurrogate.ChromLowContLim & "-" & aSurrogate.ChromUpContLim
								intCount = intCount + 1
							Next
							worksheet.Range(intSurrStart, 1, intCount, 1).BorderAround()
							worksheet.Range(intSurrStart, 2, intCount, 2).BorderAround()
							worksheet.Range(intSurrStart, 3, intCount, 3).BorderAround()
							worksheet.Range(intSurrStart - 1, 1).BorderAround()
							worksheet.Range(intSurrStart - 1, 2).BorderAround()
							worksheet.Range(intSurrStart - 1, 3).BorderAround()
						Else
							intSurrStart = intCount
						End If


						intCount = 4
						For Each aSample In GlobalVariables.ReportSamList
							If aSample.Include Then
								worksheet.Range(3, intCount, 3, intCount + 1).Merge()
								worksheet.Range(3, intCount).Value = aSample.LimsID
								worksheet.Range(4, intCount, 4, intCount + 1).Merge()
								worksheet.Range(4, intCount).Value = aSample.Name
								worksheet.Range(5, intCount, 5, intCount + 1).Merge()
								worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
								worksheet.Range(6, intCount, 6, intCount + 1).Merge()
								worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
								worksheet.Range(7, intCount, 7, intCount + 1).Merge()
								worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
								worksheet.Range(8, intCount, 8, intCount + 1).Merge()
								worksheet.Range(8, intCount).Value = aSample.DataFile
								worksheet.Range(3, intCount, 8, intCount + 1).BorderAround()
								worksheet.Range(3, intCount, 8, intCount + 1).BorderInside()
								worksheet.Range(3, intCount, 8, intCount + 1).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								worksheet.Range(9, intCount).Value = "Adjusted Limit (" & aSample.ReportedUnits & ")"
								worksheet.Range(9, intCount + 1).Value = "Amount (" & aSample.ReportedUnits & ")"
								worksheet.Range(9, intCount).BorderAround()
								worksheet.Range(9, intCount + 1).BorderAround()
								worksheet.Range(9, intCount, 9, intCount + 1).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 45)

								'Begin analyte readout
								intCount2 = 10 'Row compounds start at
								For Each aCompound In aSample.CompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											worksheet.Range(intCount2 + i, intCount).Value = GlobalVariables.Calculations.FormatSF(aCompound.ChromAdjustedLimit)
											worksheet.Range(intCount2 + i, intCount + 1).Text = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
										End If
									Next
								Next
								worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()
								worksheet.Range(10, intCount + 1, intSurrStart - 1, intCount + 1).BorderAround()

								'Surrogate write out
								If aSample.SurrogateList.Count > 0 Then
									For Each aSurrogate In aSample.SurrogateList
										For i = 0 To aSample.SurrogateList.Count
											If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
												worksheet.Range(intSurrStart + i, intCount + 1).Value = GlobalVariables.Calculations.FormatSF(aSurrogate.Recovery) & "%"
											End If
										Next
									Next
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
									worksheet.Range(intSurrStart - 1, intCount + 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).BorderAround()
								End If
								worksheet.Range(4, 4, 4, intCount + 1).CellStyle.WrapText = True
								worksheet.Range(9, 4, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.WrapText = True
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.Font.Bold = True
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.Font.Size = 8
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).CellStyle.Font.FontName = "Arial"
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount + 1).AutofitColumns()
								intCount = intCount + 2
							End If

						Next
					Else
						'RL Limit selected
						aSample = GlobalVariables.ReportSamList.Item(0)
						'Begin building sheet..
						worksheet.Range("D1:E1").Merge()
						worksheet.Range("D1").Value = "Summary Report"
						worksheet.Range("A3").Value = "LIMS #"
						worksheet.Range("A4").Value = "Sample Point"
						worksheet.Range("A5").Value = "Sample Date"
						worksheet.Range("A6").Value = "Analysis Date"
						worksheet.Range("A7").Value = "Analysis Time"
						worksheet.Range("A8").Value = "Data Folder Name"
						worksheet.Range("A9").Value = "Analyte/Parameter"
						worksheet.Range("B9").Value = "CAS #"
						worksheet.Range("C9").Value = strLimitType & " " & aSample.ReportedUnits
						worksheet.Range("A3:C9").BorderInside()
						worksheet.Range("A3:C9").BorderAround()
						worksheet.Range("A3:C9").CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
						worksheet.Range(10, 1).FreezePanes()
						aSample = GlobalVariables.ReportSamList.Item(0)

						'List Compounds
						intCount = 10
						For Each aPermit In GlobalVariables.PermitList
							If aPermit.Name = GlobalVariables.selPermit.Name Then
								For Each aProject In aPermit.ProjectList
									If aProject.Name = GlobalVariables.selProject Then
										For Each aInstrument In aProject.mInstrumentList
											If aInstrument.Name = GlobalVariables.selInstrument Then
												For Each amCompound In aInstrument.mCompoundList
													For Each aCompound In aSample.CompoundList
														If amCompound.Name = aCompound.Name Then
															worksheet.Range(intCount, 1).Value = aCompound.Name
															worksheet.Range(intCount, 2).Value = amCompound.CAS
															If strLimitType = "RL" Then
																worksheet.Range(intCount, 3).Value = amCompound.RL
															Else
																worksheet.Range(intCount, 3).Value = "N/A"
															End If
															intCount = intCount + 1
														End If
													Next
												Next
											End If
										Next
									End If
								Next
							End If
						Next
						worksheet.Range(10, 1, intCount, 1).BorderAround()
						worksheet.Range(10, 2, intCount, 2).BorderAround()
						worksheet.Range(10, 3, intCount, 3).BorderAround()

						'List Surrogates
						intCount = intCount + 1
						If aSample.SurrogateList.Count > 0 Then
							worksheet.Range(intCount, 1).Value = "Surrogate Recovery (%)"
							worksheet.Range(intCount, 1).CellStyle.Font.Color = ExcelKnownColors.Blue
							worksheet.Range(intCount, 3).Value = "Recovery Limit (%)"
							worksheet.Range(intCount, 3).CellStyle.Font.Color = ExcelKnownColors.Blue
							intCount = intCount + 1
							intSurrStart = intCount
							For Each aSurrogate In aSample.SurrogateList
								worksheet.Range(intCount, 1).Value = aSurrogate.Name
								worksheet.Range(intCount, 3).Value = aSurrogate.ChromLowContLim & "-" & aSurrogate.ChromUpContLim
								intCount = intCount + 1
							Next
							worksheet.Range(intSurrStart, 1, intCount, 1).BorderAround()
							worksheet.Range(intSurrStart, 2, intCount, 2).BorderAround()
							worksheet.Range(intSurrStart, 3, intCount, 3).BorderAround()
							worksheet.Range(intSurrStart - 1, 1).BorderAround()
							worksheet.Range(intSurrStart - 1, 2).BorderAround()
							worksheet.Range(intSurrStart - 1, 3).BorderAround()
						Else
							intSurrStart = intCount
						End If


						intCount = 4
						For Each aSample In GlobalVariables.ReportSamList
							If aSample.Include Then
								worksheet.Range(3, intCount).Value = aSample.LimsID
								worksheet.Range(4, intCount).Value = aSample.Name
								worksheet.Range(5, intCount).Value = CStr(aSample.SampleDate)
								worksheet.Range(6, intCount).Value = aSample.QuantTime.ToString("MM/dd/yyyy")
								worksheet.Range(7, intCount).Value = aSample.QuantTime.ToString("hh:mm tt")
								worksheet.Range(8, intCount).Value = aSample.DataFile
								worksheet.Range(3, intCount, 9, intCount).BorderAround()
								worksheet.Range(3, intCount, 9, intCount).BorderInside()
								worksheet.Range(3, intCount, 9, intCount).CellStyle.Color = System.Drawing.Color.FromArgb(255, 255, 153)
								worksheet.Range(9, intCount).Value = "Amount " & aSample.ReportedUnits
								GlobalVariables.workbook.ActiveSheet.SetRowHeightInPixels(9, 45)

								'Begin analyte readout
								intCount2 = 10 'Row compounds start at
								For Each aCompound In aSample.CompoundList
									For i = 0 To aSample.CompoundList.Count
										If worksheet.Range(intCount2 + i, 1).Value = aCompound.Name Then
											worksheet.Range(intCount2 + i, intCount).Text = GlobalVariables.Calculations.FormatSF(aCompound.ReportedAmt)
										End If
									Next
								Next
								worksheet.Range(10, intCount, intSurrStart - 1, intCount).BorderAround()

								'Surrogate write out
								If aSample.SurrogateList.Count > 0 Then
									For Each aSurrogate In aSample.SurrogateList
										For i = 0 To aSample.SurrogateList.Count
											If worksheet.Range(intSurrStart + i, 1).Value = aSurrogate.Name Then
												worksheet.Range(intSurrStart + i, intCount).Value = aSurrogate.Recovery & "%"
											End If
										Next
									Next
									worksheet.Range(intSurrStart - 1, intCount, intSurrStart + aSample.SurrogateList.Count, intCount).BorderAround()
								End If
								worksheet.Range(9, 4, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.WrapText = True
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
								worksheet.Range(1, 2, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Bold = True
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.Size = 8
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).CellStyle.Font.FontName = "Arial"
								worksheet.Range(1, 1, intSurrStart + aSample.SurrogateList.Count, intCount).AutofitColumns()

								intCount = intCount + 1
							End If

						Next
					End If
				End If
			Next
			Return True
		Catch ex As Exception
			MsgBox("Error generating report!" & vbCrLf &
						"Sub Procedure: MidlandChromSummaryReport()" & vbCrLf &
						"Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
			Return False
		End Try

	End Function

	'Midland Customer Report
	Function MidlandChromCustomerReport() As Boolean
		Dim worksheet As IWorksheet
		Dim exEngine As New ExcelEngine
		Dim exApp As IApplication
		exApp = exEngine.Excel

		Try
			GlobalVariables.workbook.Worksheets.Create("Cover Pg")
			GlobalVariables.workbook.Worksheets.Create("Case Narrative")
			GlobalVariables.workbook.Worksheets.Create("Flag Sheet")
			For Each wks In GlobalVariables.workbook.Worksheets
				If wks.name = "Cover Pg" Then
					worksheet = wks
					worksheet.Range(1, 1, 32, 10).CellStyle.Font.FontName = "Arial"
					worksheet.Range("A1:J1").Merge()
					worksheet.Range("A1").Value = "DOW CONFIDENTIAL INFORMATION"
					worksheet.Range("A2:J2").Merge()
					worksheet.Range("A2").Value = "Do not share without permission"
					worksheet.Range("A4:J4").CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thin
					worksheet.Range("A1").CellStyle.Font.Size = 16
					worksheet.Range("A2").CellStyle.Font.Size = 12
					worksheet.Range(1, 1, 2, 10).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
					worksheet.Range(1, 1, 2, 10).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
					worksheet.Range("A5:C5").Merge()
					worksheet.Range("A5").Value = "EH&S ANALYTICAL REPORT"
					worksheet.Range("D5:F5").Merge()
					worksheet.Range("D5").Value = "DOW CHEMICAL, USA"
					worksheet.Range("A5").CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft
					worksheet.Range("A5:J5").CellStyle.HorizontalAlignment = ExcelVAlign.VAlignCenter
					worksheet.Range("J5").CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight
					worksheet.Range("J5").Value = "MIDLAND, MI"
					worksheet.Range("A7:B7").Merge()
					worksheet.Range("A7").Value = "Date Issued:"
					worksheet.Range("A7").CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft
					worksheet.Range("G7:H7").Merge()
					worksheet.Range("G7").Value = "Study Number:"
					worksheet.Range("G7").CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight
					worksheet.Range("A7:J7").CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thin
					worksheet.Range("A7:J7").CellStyle.Borders(ExcelBordersIndex.EdgeTop).LineStyle = ExcelLineStyle.Thin
					worksheet.Range("C9:H9").Merge()
					worksheet.Range("C9").Value = "Title:"
					worksheet.Range("C9").CellStyle.Font.Size = 11
					worksheet.Range("C9").CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft
					worksheet.Range("A10:J10").CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thin
					worksheet.Range("B13:C13").Merge()
					worksheet.Range("B13").Value = "FULL REPORT"
					worksheet.Range("A14").Value = "To:"
					worksheet.Range("A14").CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight
					worksheet.Range("A23:C23").Merge()
					worksheet.Range("A23").Value = "Pages in complete report:"
					worksheet.Range("A23:J23").CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thin
					worksheet.Range("A23:J23").CellStyle.Borders(ExcelBordersIndex.EdgeTop).LineStyle = ExcelLineStyle.Thin
					worksheet.Range("A26:B26").Merge()
					worksheet.Range("A26").Value = "Sample Date:"
					worksheet.Range("G26:I26").Merge()
					worksheet.Range("H26").Value = "Analysis Completion Date:"
					worksheet.Range("A28:B28").Merge()
					worksheet.Range("A28:J28").CellStyle.Borders(ExcelBordersIndex.EdgeTop).LineStyle = ExcelLineStyle.Thin
					worksheet.Range("A28").Value = "Signature/Date:"
					worksheet.Range("D28").Value = "Bldg:"
					worksheet.Range("E28").Value = "Phone:"
					worksheet.Range("G28:H28").Merge()
					worksheet.Range("G28").Value = "Reviewer/Date:"
					worksheet.Range("I28").Value = "Bldg:"
					worksheet.Range("J28").Value = "Phone:"
					worksheet.Range("A33:J33").CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thin
					worksheet.Range(1, 1, 33, 10).CellStyle.Font.Bold = True
				ElseIf wks.name = "Case Narrative" Then
					worksheet = wks
					worksheet.Range(1, 1, 34, 10).CellStyle.Font.FontName = "Arial"
					worksheet.Range(1, 1, 34, 10).CellStyle.Font.Bold = True
					worksheet.Range("A1").CellStyle.Font.Size = 16
					worksheet.Range("A1:J1").Merge()
					worksheet.Range("A1").CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
					worksheet.Range("A1").Value = "Case Narrative"
					worksheet.Range("A3").Value = "Project:"
					worksheet.Range("A4").Value = "Sample Date:"
					worksheet.Range("A5").Value = "Analysis Date:"
					worksheet.Range("A6").Value = "Analysis:"
					worksheet.Range("A7").Value = "Matrix:"
					worksheet.Range("A8").Value = "Samples Affected:"
					worksheet.Range("A10").Value = "Samples were received in acceptable condition."
					worksheet.Range("A12:J12").Merge()
					worksheet.Range("A12").Value = "All quality exceptions or observations that impacted data quality or indicate data bias/uncertainty are " &
					"documented below in this narrative and flagged in the data report.  All exceptions to method specified quality requirements that do not impact data quality " &
					"are documented in a Quality Control Exception Report filed with the raw data."
					worksheet.Range("A12").CellStyle.WrapText = True
					worksheet.SetRowHeightInPixels(12, 90)
					worksheet.Range("A14").Value = "Exceptions:"
					worksheet.Range("A30").Value = "Analyst:"
					worksheet.Range("A32").Value = "Quality Reviewer:"
					worksheet.Range("F30").Value = "Date:"
					worksheet.Range("F32").Value = "Date:"
					worksheet.Range("A30:F32").CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight
					worksheet.Range("B30:E30").CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thin
					worksheet.Range("G30:J30").CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thin
					worksheet.Range("B32:E32").CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thin
					worksheet.Range("G32:J32").CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thin
					worksheet.Range("B34:I34").Merge()
					worksheet.Range("B34").Value = "Case Narrative signed in hardcopy only."
					worksheet.Range("B34").CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
					worksheet.Range("B34:I34").CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thick
					worksheet.Range("B34:I34").CellStyle.Borders(ExcelBordersIndex.EdgeLeft).LineStyle = ExcelLineStyle.Thick
					worksheet.Range("B34:I34").CellStyle.Borders(ExcelBordersIndex.EdgeRight).LineStyle = ExcelLineStyle.Thick
					worksheet.Range("B34:I34").CellStyle.Borders(ExcelBordersIndex.EdgeTop).LineStyle = ExcelLineStyle.Thick
					worksheet.SetRowHeightInPixels(12, 70)
					worksheet.SetColumnWidthInPixels(1, 125)
				ElseIf wks.name = "Flag Sheet" Then
					worksheet = wks
					worksheet.Range(1, 1, 128, 10).CellStyle.Font.FontName = "Arial"
					worksheet.Range("F2:H2").Merge()
					worksheet.Range("F3:H3").Merge()
					worksheet.Range("F4:H4").Merge()
					worksheet.Range("F2").Value = "Organic Group"
					worksheet.Range("F3").Value = "Report Qualifiers (Flags)"
					worksheet.Range("F4").Value = "Revised: 5/9/13"
					worksheet.Range("A4").Value = "Qualifier"
					worksheet.Range("A4:O4").CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thin
					worksheet.Range("A1:H4").CellStyle.Font.Bold = True
					worksheet.Range("A6").Value = "B"
					worksheet.Range("B6:O6").Merge()
					worksheet.Range("B6").Value = "Analyte was detected in the associated Lab Blank/Method Blank/Reagent Blank, a high bias to data indicated."
					worksheet.Range("A7").Value = "B1"
					worksheet.Range("B7:O7").Merge()
					worksheet.Range("B7").Value = "Non-target analyte detected in associated Lab Blank/Method Blank/Reagent Blank and sample, producing interference."
					worksheet.Range("A8").Value = "B2"
					worksheet.Range("B8:O8").Merge()
					worksheet.Range("B8").Value = "Analyte was detected in the associated Lab Blank/Method Blank/Reagent Blank, but was not detected in Sample, sample data not affected."
					worksheet.Range("A9").Value = "B3"
					worksheet.Range("B9:O9").Merge()
					worksheet.Range("B9").Value = "Analyte detected in associated Lab Blank/Method Blank/Reagent Blank at level less than 5% that found in sample, data quality not affected."
					worksheet.Range("A10").Value = "B4"
					worksheet.Range("B10:O10").Merge()
					worksheet.Range("B10").Value = "Analyte detected in associated Lab Blank/Method Blank/Reagent Blank at or above MDL but at level less than 10% of the MAL, data quality not affected for TPH reporting."
					worksheet.SetRowHeightInPixels(10, 45)
					worksheet.Range("A13").Value = "C"
					worksheet.Range("B13:O13").Merge()
					worksheet.Range("B13").Value = "Calibration Verification recovery was below the method control limit for this analyte. Low bias to data indicated."
					worksheet.Range("A14").Value = "C1"
					worksheet.Range("B14:O14").Merge()
					worksheet.Range("B14").Value = "The sample was originally analyzed with a positive result, however the reanalysis did not confirm the presence of the analyte."
					worksheet.Range("A15").Value = "C2"
					worksheet.Range("B15:O15").Merge()
					worksheet.Range("B15").Value = "Results confirmed by reanalysis."
					worksheet.Range("A16").Value = "C3"
					worksheet.Range("B16:O16").Merge()
					worksheet.Range("B16").Value = "Quantitation curve failure: R^2 >15% or coefficient of determination <0.99."
					worksheet.Range("A17").Value = "C4"
					worksheet.Range("B17:O17").Merge()
					worksheet.Range("B17").Value = "Calibration Verification recovery was above the method control limit for this analyte. High bias to data indicated."
					worksheet.Range("A18").Value = "C5"
					worksheet.Range("B18:O18").Merge()
					worksheet.Range("B18").Value = "Calibration Verification recovery was above the method control limit for this analyte. Analyte not found in samples, data quality not affected."
					worksheet.Range("A20").Value = "D"
					worksheet.Range("B20:O20").Merge()
					worksheet.Range("B20").Value = "Compound quantitated on a diluted sample."
					worksheet.Range("A22").Value = "E"
					worksheet.Range("B22:O22").Merge()
					worksheet.Range("B22").Value = "Reported value is EMPC (estimated maximum possible concentration)"
					worksheet.Range("A23").Value = "E1"
					worksheet.Range("B23:O23").Merge()
					worksheet.Range("B23").Value = "Concentration exceeds the tested calibration range of the instrument."
					worksheet.Range("A24").Value = "E2"
					worksheet.Range("B24:O24").Merge()
					worksheet.Range("B24").Value = "Estimated due to interference."
					worksheet.Range("A26").Value = "H"
					worksheet.Range("B26:O26").Merge()
					worksheet.Range("B26").Value = "Sample analysis performed past the method-specified holding time per client's approval."
					worksheet.Range("A27").Value = "H1"
					worksheet.Range("B27:O27").Merge()
					worksheet.Range("B27").Value = "Initial analysis within holding time.  Reanalysis for the required dilution was past holding time."
					worksheet.Range("A29").Value = "I"
					worksheet.Range("B29:O29").Merge()
					worksheet.Range("B29").Value = "Internal Standard recovery was outside of method limits."
					worksheet.Range("A30").Value = "I1"
					worksheet.Range("B30:O30").Merge()
					worksheet.Range("B30").Value = "Estimated result. The associated internal standard did not meet the specified method criteria."
					worksheet.Range("A32").Value = "J"
					worksheet.Range("B32:O32").Merge()
					worksheet.Range("B32").Value = "Estimated value.  Analyte detected at a level less than the Reporting Limit (RL) and greater than or equal to the Method Detection Limit (MDL). The user of " &
					"this data should be aware that this data is of limited reliability."
					worksheet.SetRowHeightInPixels(32, 45)
					worksheet.Range("A33").Value = "J1"
					worksheet.Range("B33:O33").Merge()
					worksheet.Range("B33").Value = "Analyte detected at a level less then the Reporting Limit (RL) and greater than, or equal to the Method Detection Limit (MDL). On the Final Summary Report Analyte " &
						"reported Not Detected (ND) because analyzed concentration is below the RL. For Lab Control Spike (LCS)  & Matrix Spike (MS) recovery purposes the analyzed concentration below the RL and above " &
						"or equal to the MDL are shown on the LCS & MS recovery reports."
					worksheet.SetRowHeightInPixels(33, 67)
					worksheet.Range("A34").Value = "J2"
					worksheet.Range("B34:O34").Merge()
					worksheet.Range("B34").Value = "Analyte detected at a level less then the Reporting Limit (RL) and is less than Method Detection Limit (MDL). On the Final Summary Report Analyte reported Not " &
						"Detected (ND) because analyzed concentration is below the RL. For Lab Control Spike (LCS) & Matrix Spike (MS) recovery purposes the analyzed concentration below the RL and less then " &
						"the MDL are shown on the LCS & MS recovery reports."
					worksheet.SetRowHeightInPixels(34, 67)
					worksheet.Range("A35").Value = "J3"
					worksheet.Range("B35:O35").Merge()
					worksheet.Range("B35").Value = "Analyte detected above MDL but below lowest calibration point. Estimated value."
					worksheet.Range("A36").Value = "J4"
					worksheet.Range("B36:O36").Merge()
					worksheet.Range("B36").Value = "Analyte detected above detection level but below quantification level. Estimated value."
					worksheet.Range("A38").Value = "L"
					worksheet.Range("B38:O38").Merge()
					worksheet.Range("B38").Value = "Laboratory Control Spike and/or Laboratory Control Spike Duplicate recovery was above the acceptance limits. Analyte not detected at reporting limit " &
						"in samples, therefore data quality not affected. "
					worksheet.SetRowHeightInPixels(38, 45)
					worksheet.Range("A39").Value = "L1"
					worksheet.Range("B39:O39").Merge()
					worksheet.Range("B39").Value = "Laboratory Control Spike and/or Laboratory Control Spike Duplicate recovery was above acceptance limits. High bias to sample results indicated."
					worksheet.Range("A40").Value = "L2"
					worksheet.Range("B40:O40").Merge()
					worksheet.Range("B40").Value = "Laboratory Control Spike and/or Laboratory Control Spike Duplicate recovery was below the acceptance limits. A low bias to sample results is indicated."
					worksheet.Range("A41").Value = "L3"
					worksheet.Range("B41:O41").Merge()
					worksheet.Range("B41").Value = "The LCS and/or LCSD were above the acceptance limits.  Passing matrix spike (MS) satisfies method requirements. Data quality not affected."
					worksheet.Range("A42").Value = "L4"
					worksheet.Range("B42:O42").Merge()
					worksheet.Range("B42").Value = "The LCS and/or LCSD were below the acceptance limits.  Passing matrix spike (MS) satisfies method requirements. Data quality not affected."
					worksheet.Range("A44").Value = "M"
					worksheet.Range("B44:O44").Merge()
					worksheet.Range("B44").Value = "Duplicate sample precision not met."
					worksheet.Range("A45").Value = "M1"
					worksheet.Range("B45:O45").Merge()
					worksheet.Range("B45").Value = "The MS and/or MSD were above the acceptance limits.  Passing Lab Control Spike (LCS) satisfies method requirements.  Data quality not affected."
					worksheet.Range("A46").Value = "M2"
					worksheet.Range("B46:O46").Merge()
					worksheet.Range("B46").Value = "The MS and/or MSD were below the acceptance limits.  Passing Lab Control Spike (LCS) satisfies method requirements. "
					worksheet.Range("A47").Value = "M3"
					worksheet.Range("B47:O47").Merge()
					worksheet.Range("B47").Value = "The sample spiked had a pH of less than 2.  Analyte degrades under acidic conditions."
					worksheet.Range("A48").Value = "M4"
					worksheet.Range("B48:O48").Merge()
					worksheet.Range("B48").Value = "Due to high levels of analyte in the sample, the MS/MSD calculation does not provide useful spike recovery information. See Laboratory Control Spike (LCS)."
					worksheet.SetRowHeightInPixels(48, 45)
					worksheet.Range("A49").Value = "M5"
					worksheet.Range("B49:O49").Merge()
					worksheet.Range("B49").Value = "No results were reported for the MS/MSD.  The sample used for the MS/MSD required dilution due to the sample matrix.  Because of this, the spike " &
						"compounds were diluted below the detection limit."
					worksheet.SetRowHeightInPixels(49, 45)
					worksheet.Range("A50").Value = "M6"
					worksheet.Range("B50:O50").Merge()
					worksheet.Range("B50").Value = "There was no MS/MSD analyzed with this batch due to insufficient sample volume.  See Lab Control Spike/Lab Control Spike Duplicate."
					worksheet.Range("A51").Value = "M7"
					worksheet.Range("B51:O51").Merge()
					worksheet.Range("B51").Value = "No recovery range given in method for analyte in LCS/MS/MSD"
					worksheet.Range("A52").Value = "M8"
					worksheet.Range("B52:O52").Merge()
					worksheet.Range("B52").Value = "Matrix spike and/or Matrix Spike Duplicate recovery was above acceptance limits. Analyte not found in samples, data quality not affected."
					worksheet.Range("A53").Value = "M9"
					worksheet.Range("B53:O53").Merge()
					worksheet.Range("B53").Value = "Matrix Spike and/or Matrix Spike Duplicate recovery was above acceptance limits. High bias to sample results indicated."
					worksheet.Range("A54").Value = "M10"
					worksheet.Range("B54:O54").Merge()
					worksheet.Range("B54").Value = "Matrix Spike and/or Matrix Spike Duplicate recovery was below the acceptance limits.   A low bias to sample results is indicated."
					worksheet.Range("A55").Value = "M11"
					worksheet.Range("B55:O55").Merge()
					worksheet.Range("B55").Value = "Matrix spike results reported from diluted sample, due to target analyte recovery above calibration limit in undiluted sample."
					worksheet.Range("A57").Value = "N"
					worksheet.Range("B57:O57").Merge()
					worksheet.Range("B57").Value = "Spike recovery not within control limits."
					worksheet.Range("A58").Value = "N1"
					worksheet.Range("B58:O58").Merge()
					worksheet.Range("B58").Value = "See case narrative."
					worksheet.Range("A59").Value = "N2"
					worksheet.Range("B59:O59").Merge()
					worksheet.Range("B59").Value = "Sample point not sampled, see case narrative."
					worksheet.Range("A61").Value = "P"
					worksheet.Range("B61:O61").Merge()
					worksheet.Range("B61").Value = "The sample, as received, was not preserved in accordance to the referenced analytical method."
					worksheet.Range("A62").Value = "P2"
					worksheet.Range("B62:O62").Merge()
					worksheet.Range("B62").Value = "Sample was not sufficiently preserved at time of collection.  Sample pH is >2"
					worksheet.Range("A63").Value = "P3"
					worksheet.Range("B63:O63").Merge()
					worksheet.Range("B63").Value = "Sample received without chemical preservation, but preserved by the laboratory."
					worksheet.Range("A64").Value = "P4"
					worksheet.Range("B64:O64").Merge()
					worksheet.Range("B64").Value = "Sample was received above recommended temperature."
					worksheet.Range("A65").Value = "P5"
					worksheet.Range("B65:O65").Merge()
					worksheet.Range("B65").Value = "Sample received in inappropriate sample container."
					worksheet.Range("A66").Value = "P6"
					worksheet.Range("B66:O66").Merge()
					worksheet.Range("B66").Value = "Insufficient sample received to meet method QC requirements."
					worksheet.Range("A67").Value = "P7"
					worksheet.Range("B67:O67").Merge()
					worksheet.Range("B67").Value = "Sample taken from VOA vial with air bubble > 6mm diameter."
					worksheet.Range("A68").Value = "P8"
					worksheet.Range("B68:O68").Merge()
					worksheet.Range("B68").Value = "Sample pH  < 2 when received, target analyte(s) is acid sensitive, thererfore analysis for target analyte(s) is not valid."
					worksheet.Range("A70").Value = "R"
					worksheet.Range("B70:O70").Merge()
					worksheet.Range("B70").Value = "The RPD exceeded the method control limit due to sample matrix effects.  The individual analyte QA/QC recoveries, however,  pass the method control limits. "
					worksheet.SetRowHeightInPixels(70, 45)
					worksheet.Range("A71").Value = "R1"
					worksheet.Range("B71:O71").Merge()
					worksheet.Range("B71").Value = "The RPD exceeded the acceptance limit due to sample matrix effects."
					worksheet.Range("A72").Value = "R2"
					worksheet.Range("B72:O72").Merge()
					worksheet.Range("B72").Value = "Due to the low levels of analyte in the sample, the duplicate RPD calculation does not provide useful information."
					worksheet.Range("A73").Value = "R3"
					worksheet.Range("B73:O73").Merge()
					worksheet.Range("B73").Value = "Reporting limit raised due to high concentrations of non-target analytes."
					worksheet.Range("A74").Value = "R4"
					worksheet.Range("B74:O74").Merge()
					worksheet.Range("B74").Value = "Reporting limit raised due to insufficient sample volume."
					worksheet.Range("A75").Value = "R5"
					worksheet.Range("B75:O75").Merge()
					worksheet.Range("B75").Value = "Sample required dilution due to high concentrations of target analyte."
					worksheet.Range("A76").Value = "R6"
					worksheet.Range("B76:O76").Merge()
					worksheet.Range("B76").Value = "RPD exceeded method control limits due to a lack of sample homogeneity"
					worksheet.Range("A77").Value = "R7"
					worksheet.Range("B77:O77").Merge()
					worksheet.Range("B77").Value = "Reporting limit raised due to dilution required for sample matrix.  See case narrative for details."
					worksheet.Range("A78").Value = "R8"
					worksheet.Range("B78:O78").Merge()
					worksheet.Range("B78").Value = "Due to the high levels of analyte in the sample, the duplicate RPD calculation does not provide useful information."
					worksheet.Range("A79").Value = "R9"
					worksheet.Range("B79:O79").Merge()
					worksheet.Range("B79").Value = "RPD exceeds the acceptance limit.  Analyte not detected at reporting limit in samples, therefore data quality not affected."
					worksheet.Range("A81").Value = "S"
					worksheet.Range("B81:O81").Merge()
					worksheet.Range("B81").Value = "Sediment present."
					worksheet.Range("A82").Value = "S1"
					worksheet.Range("B82:O82").Merge()
					worksheet.Range("B82").Value = "Insufficient sample available for reanalysis."
					worksheet.Range("A83").Value = "S2"
					worksheet.Range("B83:O83").Merge()
					worksheet.Range("B83").Value = "Analytical results not reliable due to potential sample container contamination."
					worksheet.Range("A85").Value = "T1"
					worksheet.Range("B85:O85").Merge()
					worksheet.Range("B85").Value = "Tentatively identified compound.  Concentration is estimated based on the closest internal standard."
					worksheet.Range("A86").Value = "T2"
					worksheet.Range("B86:O86").Merge()
					worksheet.Range("B86").Value = "Not to be reported for compliance purposes."
					worksheet.Range("A88").Value = "Z"
					worksheet.Range("B88:O88").Merge()
					worksheet.Range("B88").Value = "The surrogate/Isotopic recovery was below the acceptance limits, results may be biased low."
					worksheet.Range("A89").Value = "Z1"
					worksheet.Range("B89:O89").Merge()
					worksheet.Range("B89").Value = "Surrogate/Isotopic standard recovery was outside the acceptance limits.  Data not impacted."
					worksheet.Range("A90").Value = "Z2"
					worksheet.Range("B90:O90").Merge()
					worksheet.Range("B90").Value = "The sample required a dilution due to the nature of the sample matrix.  Because of this dilution, the surrogate/Isotopic spike concentration " &
						"in the sample was reduced to a level where the recovery calculation does not provide useful information."
					worksheet.SetRowHeightInPixels(90, 45)
					worksheet.Range("A91").Value = "Z3"
					worksheet.Range("B91:O91").Merge()
					worksheet.Range("B91").Value = "The surrogate/isotopic recovery was above the acceptance limits, results may be biased high."
					worksheet.Range("A93:O93").Merge()
					worksheet.Range("A93").Value = "Organic Group Abbreviations"
					worksheet.Range("A95:O95").CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Thin
					worksheet.Range("B97:G97").Merge()
					worksheet.Range("I97:O97").Merge()
					worksheet.Range("A97").Value = "AF"
					worksheet.Range("B97").Value = "Antifoam"
					worksheet.Range("H97").Value = "MS"
					worksheet.Range("I97").Value = "Matrix Spike"
					worksheet.Range("B98:G98").Merge()
					worksheet.Range("I98:O98").Merge()
					worksheet.Range("A98").Value = "<"
					worksheet.Range("B98").Value = "Less Than"
					worksheet.Range("H98").Value = "MSD"
					worksheet.Range("I98").Value = "Matrix Spike Duplicate"
					worksheet.Range("B99:G99").Merge()
					worksheet.Range("I99:O99").Merge()
					worksheet.Range("A99").Value = ">"
					worksheet.Range("B99").Value = "Greater Than"
					worksheet.Range("H99").Value = "MW"
					worksheet.Range("I99").Value = "Monitoring Well"
					worksheet.Range("B100:G100").Merge()
					worksheet.Range("I100:O100").Merge()
					worksheet.Range("A100").Value = "Bldg"
					worksheet.Range("B100").Value = "Building"
					worksheet.Range("H100").Value = "m3"
					worksheet.Range("I100").Value = "Cubic Meters"
					worksheet.Range("B101:G101").Merge()
					worksheet.Range("I101:O101").Merge()
					worksheet.Range("A101").Value = "C"
					worksheet.Range("B101").Value = "Degrees Celsius"
					worksheet.Range("H101").Value = "N/A"
					worksheet.Range("I101").Value = "Not Applicable"
					worksheet.Range("B102:G102").Merge()
					worksheet.Range("I102:O102").Merge()
					worksheet.Range("A102").Value = "CAS #"
					worksheet.Range("B102").Value = "Chemical Abstracts Service Number"
					worksheet.Range("H102").Value = "ng/dscm"
					worksheet.Range("I102").Value = "Nanograms per Dry Standard Cubic Meter"
					worksheet.Range("B103:G103").Merge()
					worksheet.Range("I103:O103").Merge()
					worksheet.Range("A103").Value = "CFR"
					worksheet.Range("B103").Value = "Code of Federal Regulations"
					worksheet.Range("H103").Value = "ND"
					worksheet.Range("I103").Value = "Not Detected at the Reporting Limit"
					worksheet.Range("B104:G104").Merge()
					worksheet.Range("I104:O104").Merge()
					worksheet.Range("A104").Value = "Dil."
					worksheet.Range("B104").Value = "Dilution"
					worksheet.Range("H104").Value = "OF"
					worksheet.Range("I104").Value = "Outfall"
					worksheet.Range("B105:G105").Merge()
					worksheet.Range("I105:O105").Merge()
					worksheet.Range("A105").Value = "Dup."
					worksheet.Range("B105").Value = "Duplicate"
					worksheet.Range("H105").Value = "ppb"
					worksheet.Range("I105").Value = "Parts Per Billion"
					worksheet.Range("B106:G106").Merge()
					worksheet.Range("I106:O106").Merge()
					worksheet.Range("A106").Value = "F"
					worksheet.Range("B106").Value = "Degrees Fahrenheit"
					worksheet.Range("H106").Value = "ppm"
					worksheet.Range("I106").Value = "Parts Per Million"
					worksheet.Range("B107:G107").Merge()
					worksheet.Range("I107:O107").Merge()
					worksheet.Range("A107").Value = "g"
					worksheet.Range("B107").Value = "gram(s)"
					worksheet.Range("H107").Value = "ppt"
					worksheet.Range("I107").Value = "Parts Per Trillion"
					worksheet.Range("B108:G108").Merge()
					worksheet.Range("I108:O108").Merge()
					worksheet.Range("A108").Value = "GC/MSD"
					worksheet.Range("B108").Value = "Gas Chromatography/Mass Spectrometry Detection"
					worksheet.Range("H108").Value = "ppq"
					worksheet.Range("I108").Value = "Parts Per Quadrillion"
					worksheet.Range("B109:G109").Merge()
					worksheet.Range("I109:O109").Merge()
					worksheet.Range("A109").Value = "GWW"
					worksheet.Range("B109").Value = "Ground Water Well"
					worksheet.Range("H109").Value = "pg/dscm"
					worksheet.Range("I109").Value = "Picograms Per Dry Standard Cubic Meter"
					worksheet.Range("B110:G110").Merge()
					worksheet.Range("I110:O110").Merge()
					worksheet.Range("A110").Value = "kg"
					worksheet.Range("B110").Value = "Kilogram(s)"
					worksheet.Range("H110").Value = "PQL"
					worksheet.Range("I110").Value = "Practical Quantitation Limit"
					worksheet.Range("B111:G111").Merge()
					worksheet.Range("I111:O111").Merge()
					worksheet.Range("A111").Value = "l"
					worksheet.Range("B111").Value = "Liter(s)"
					worksheet.Range("H111").Value = "QA/QC"
					worksheet.Range("I111").Value = "Quality Assurance / Quality Control"
					worksheet.Range("B112:G112").Merge()
					worksheet.Range("I112:O112").Merge()
					worksheet.Range("A112").Value = "lb."
					worksheet.Range("B112").Value = "Pound(s)"
					worksheet.Range("H112").Value = "Rec"
					worksheet.Range("I112").Value = "Recovered"
					worksheet.Range("B113:G113").Merge()
					worksheet.Range("I113:O113").Merge()
					worksheet.Range("A113").Value = "LCS"
					worksheet.Range("B113").Value = "Lab Control Spike"
					worksheet.Range("H113").Value = "RL"
					worksheet.Range("I113").Value = "Reporting Limit"
					worksheet.Range("B114:G114").Merge()
					worksheet.Range("I114:O114").Merge()
					worksheet.Range("A114").Value = "LIMS"
					worksheet.Range("B114").Value = "Laboratory Information Management System"
					worksheet.Range("H114").Value = "RPD"
					worksheet.Range("I114").Value = "Relative Percent Difference"
					worksheet.Range("B115:G115").Merge()
					worksheet.Range("I115:O115").Merge()
					worksheet.Range("A115").Value = "LS"
					worksheet.Range("B115").Value = "Lift Station"
					worksheet.Range("H115").Value = "SS"
					worksheet.Range("I115").Value = "Surrogate Standard"
					worksheet.Range("B116:G116").Merge()
					worksheet.Range("I116:O116").Merge()
					worksheet.Range("A116").Value = "MAL"
					worksheet.Range("B116").Value = "Minimum Analytical Limit"
					worksheet.Range("H116").Value = "TDL"
					worksheet.Range("I116").Value = "Target Detection Limit"
					worksheet.Range("B117:G117").Merge()
					worksheet.Range("I117:O117").Merge()
					worksheet.Range("A117").Value = "MDL"
					worksheet.Range("B117").Value = "Method Detection Limit"
					worksheet.Range("H117").Value = "TPH"
					worksheet.Range("I117").Value = "Total Purgable Halocarbons"
					worksheet.Range("B118:G118").Merge()
					worksheet.Range("I118:O118").Merge()
					worksheet.Range("A118").Value = "Med"
					worksheet.Range("B118").Value = "Methylated"
					worksheet.Range("H118").Value = "ug"
					worksheet.Range("I118").Value = "Microgram(s)"
					worksheet.Range("B119:G119").Merge()
					worksheet.Range("I119:O119").Merge()
					worksheet.Range("A119").Value = "mg"
					worksheet.Range("B119").Value = "Milligram(s)"
					worksheet.Range("H119").Value = "ul"
					worksheet.Range("I119").Value = "Microliter(s)"
					worksheet.Range("B120:G120").Merge()
					worksheet.Range("I120:O120").Merge()
					worksheet.Range("A120").Value = "mg/L"
					worksheet.Range("B120").Value = "Milligrams Per Liter"
					worksheet.Range("H120").Value = "ug/L"
					worksheet.Range("I120").Value = "Microgram(s) per Liter"
					worksheet.Range("B121:G121").Merge()
					worksheet.Range("I121:O121").Merge()
					worksheet.Range("A121").Value = "ml"
					worksheet.Range("B121").Value = "Milliliter(s)"
					worksheet.Range("H121").Value = "WWTP"
					worksheet.Range("I121").Value = "Wastewater Treatment Plant"
					worksheet.Range("B122:G122").Merge()
					worksheet.Range("A122").Value = "mm"
					worksheet.Range("B122").Value = "Millimeters(s)"
					worksheet.Range("A124:C124").Merge()
					worksheet.Range("A125:C125").Merge()
					worksheet.Range("A126:C126").Merge()
					worksheet.Range("A127:C127").Merge()
					worksheet.Range("A128:C128").Merge()
					worksheet.Range("D124:O124").Merge()
					worksheet.Range("D125:O125").Merge()
					worksheet.Range("D126:O126").Merge()
					worksheet.Range("D127:O127").Merge()
					worksheet.Range("D128:O128").Merge()
					worksheet.Range("A124").Value = "Dry Weight Basis:"
					worksheet.Range("A125").Value = "Dilution Factor:"
					worksheet.Range("A126").Value = "Elevated Limit:"
					worksheet.Range("A127").Value = "Adjusted Limit:"
					worksheet.Range("A128").Value = "Calculation Factor:"
					worksheet.Range("D124").Value = "Results based on dry weight"
					worksheet.Range("D125").Value = "Elevated Limit / Original Limit"
					worksheet.Range("D126").Value = "Original Limit multiplied by Dilution Factor"
					worksheet.Range("D127").Value = "Original Limit multiplied by Calculation Factor based on Dilution and Sample Size"
					worksheet.Range("D128").Value = "Adjusted Limit / Original Limit"
					worksheet.Range("A5:A128").CellStyle.Font.Bold = True
					worksheet.Range("H97:H121").CellStyle.Font.Bold = True
					worksheet.Range(1, 1, 128, 10).CellStyle.WrapText = True

				End If
			Next
			Return True

		Catch ex As Exception
			MsgBox("Error generating report!" & vbCrLf &
						"Sub Procedure: MidlandChromCustomerReport()" & vbCrLf &
						"Logic Error: " & ex.Message, MsgBoxStyle.Critical, "(╯°□°)╯︵ ┻━┻")
			Return False
		End Try

	End Function


End Class

