Call rptReportLog("Start Execution","","")
Webutil.DeleteCookies
Wait(2)
Call ExecTestSuite(sTestSuitePath)

''##################################################################################################################################
''Function Name : ExecTestSuite
''Description : Reads the Test Suite Information and Executes the test cases/ scenarios accordingly
''Parameters : sTestSuitePath, sTestSuiteSheet
''Author : Gallop Automation Team	
''##################################################################################################################################
Function ExecTestSuite(sTestSuiteFile)
	ExecTestSuite = false
	XLObj.DisplayAlerts = False
	if Not ifFileExist(sTestSuiteFile) Then
		Call rptReportLog("Start Execution","Open the test suite file",sTestSuiteFile & " file not exist")
		ExitTest
	End if 
	
	Set XLWB = XLObj.WorkBooks.Open(sTestSuiteFile)
	if Err.number<>0 then 
		Call rptReportLog("Start Execution","Open the test suite file","Error occur while opening the test suite file ::: " & Err.description)
	Else
		Call rptReportLog("Start Execution","Open the test suite file","open the test suite file successfully")
	End if 	
	Set XLSheet = XLWB.WorkSheets(1)
    RowCnt = XLSheet.Usedrange.rows.count
	ColCnt = XLSheet.Usedrange.columns.count
    ''**********Make sure the location can be shifted, however the header name is pre-defined.'**********
	For colcount = 1 To ColCnt
		Select Case Lcase(XLSheet.Cells (4,colcount).Value)
			Case "portfolio" PortfolioCol = colcount
			Case "application" ApplicationCol = colcount
			Case "execute" ExecuteCol = colcount
			Case "module" ModuleCol = colcount
			Case "test_case_name" Test_Case_NameCol = colcount
			Case "object_repositories" Object_RepositoriesCol = colcount
			Case "onfail" OnFailCol = colcount
		End Select
	Next
	TCPortfolioCol = ""
	CompPortfolioCol = ""
	For rc = 5 to RowCnt
		If StrComp(trim(XLSheet.Cells (rc,ExecuteCol).Value),"yes",1) = 0 Then
			If trim(XLSheet.Cells (rc,ModuleCol).Value) <> "" Then
				Environment.Value("Test_Name") = XLSheet.Cells (rc,Test_Case_NameCol).Value 
				sTSModulePath = sTestCasesPath & XLSheet.Cells (rc,ModuleCol).Value & "\" & XLSheet.Cells (rc,Test_Case_NameCol).Value & ".xlsm"
				strTestCaseName = XLSheet.Cells (rc,Test_Case_NameCol).Value
				strTestCaseID = Split(strTestCaseName,"_")(0)
				strSheetName = Split(strTestCaseName,"_")(1)
				strTestDataPath = sTestDataPath & XLSheet.Cells (rc,ModuleCol).Value & ".xls" 
				if Not ifFileExist(sTSModulePath) Then
					Call rptReportLog("Start test case execution::"&strTestCaseName,"Open the "& strTestCaseName &" test case file",sTSModulePath & " file not exist")
					rptWriteReport "FAIL", "Open the test case excel file" , sTSModulePath & " file not exist"
				End if 
				Set XLTCWB = XLObj.WorkBooks.Open(sTSModulePath)
					if Err.number<>0 then 
						Call rptReportLog("Start Execution","Open the " & strTestCaseName & " file","Error occur while opening the test case file ::: " & Err.description)
					Else
						Call rptReportLog("Start Execution","Open the " & strTestCaseName & " file","Open the " & strTestCaseName & " test case file successfully")
					End if 
				Set XLTCSheet = XLTCWB.WorkSheets(1)
				TCRowCnt = XLTCSheet.Usedrange.rows.count
				TCColCnt = XLTCSheet.Usedrange.columns.count
				sTCORs = XLSheet.Cells (rc,Object_RepositoriesCol).Value
				If RepositoriesCollection.Count > 0 Then
					RepositoriesCollection.RemoveAll
				End If
				If instr(sTCORs,",") > 0 Then
					ORs = Split(sTCORs,chr(44))
					For ORRep = 0 to UBound(ORs)
						RepositoriesCollection.Add sORPath & ORs(ORRep) & ".tsr", ORRep+1
						Call rptReportLog("Start Execution","Associate the " & ORs(ORRep) & " object repository to test"&vbtab,"Associate the " & ORs(ORRep) & " object repository to test successfully")
					Next
				Else
					RepositoriesCollection.Add sORPath & sTCORs & ".tsr", 1
					Call rptReportLog("Start Execution","Associate the " & sTCORs & " object repository to test"&vbtab,"Associate the " & sTCORs & " object repository to test successfully")
				End If
				If TCPortfolioCol = "" Then 
					For TCcolcount = 1 To TCColCnt
						Select Case Lcase(XLTCSheet.Cells (4,TCcolcount).Value)
							Case "portfolio" TCPortfolioCol = TCcolcount
							Case "application" TCApplicationCol = TCcolcount
							Case "execute" TCExecuteCol = TCcolcount
							Case "module" TCModuleCol = TCcolcount
							Case "component" TCComponentCol = TCcolcount
							Case "onfail" TCOnFailCol = TCcolcount
						End Select
					Next
				End If 
				if Not ifFileExist(sTSModulePath) Then
					rptWriteReport "FAIL", "Read the test data" , strTestDataPath & " file is not exist"
				End if
				dtScenario.ImportData strTestDataPath,strSheetName,strTestCaseID
				if dtScenario.GetRowCount > 0 Then
					Call rptReportLog("Start Execution","Read the test data from " & strTestDataPath & " File ","Read the test data from " & strTestDataPath & " File successfully")
				Else
					rptWriteReport "FAIL", "Read the test data from " & strTestDataPath & " File ","No rows in test data with "&strTestCaseID
				End If
				Dim i
				gErrorFlag = False
				Err.Clear
				For i  = 1 To dtScenario.GetRowCount
					Call rptReportLog("Start Execution","start execute of  " & Environment.Value("Test_Name"),"")
					Call rptReportLog(vbtab&vbtab&"Step Name",vbtab&vbtab&"Expected",btab&vbtab&"Status")
					On Error Resume Next
					dtScenario.SetCurrentRow i
					For tcrc = 5 To TCRowCnt
						If StrComp(trim(XLTCSheet.Cells (tcrc,TCExecuteCol).Value),"yes",1) = 0 Then
							If trim(XLTCSheet.Cells (tcrc,TCModuleCol).Value) <> "" Then
								sResult = eval(Replace(XLTCSheet.Cells (tcrc,TCComponentCol).Value," ","_"))
								If Err.Number <> 0 then
									rptWriteReport "Fail", Environment.Value("StepName"), "Error occur while execution:: " & Err.description
									gErrorFlag = True
									Err.Clear
								End If	
								If gErrorFlag = true and Ucase(Trim(XLTCSheet.Cells (tcrc,TCOnFailCol).Value)) = "EXIT" Then									
									Exit For 
								Else
									gErrorFlag = False											
								End If
							End If 
						End If 
					Next 
				Next
			End If
		End If
	Next
	XLObj.DisplayAlerts = False
	XLWB.Close
	XLTCWB.Close
	XLTDWB.Close
	XLObj.Quit
	Set dtScenario = Nothing 
	''Nullify all the objects
	Set XLTCWB = nothing
	Set XLTCSheet = nothing 
	Set XLTDWB =  nothing
	Set XLTDSheet = nothing
	Set XLWB = nothing
	Set XLSheet = nothing
	Set XLObj = nothing	
	Call rptWriteResultsSummary()
	call KillReqdProcess("EXCEL.EXE")
End Function




