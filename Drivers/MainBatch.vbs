
''Script			: Main Batch driver
''Script Desc		: Reads the controlFile and executes the Scripts/test cases
''Parameters		: N/A
''****************************************************************************************************************************************************************
'Option Explicit
Wscript.Echo " Killing all the Processes before execution starts..."
Call KillProcessesBeforeExec()
Call includeFile ( GetAbsolutePath() & "\Environment Files\GlobalDef.vbs")
Call includeFile (GetAbsolutePath() & "\Environment Files\SetDirs.vbs")
slogFile = vBaseDir &"\Batch Execution Logs\"& GetCurrUserLoggedIn() &"_Log_"& ReplaceFunc(Date) &"_"& ReplaceFunc(time) &".txt"
Call includeFile (GetAbsolutePath()  & "\Environment Files\Welcome.txt")
Call EchoAndLog(fileData,slogFile,true)

Call EchoAndLog(" Test Suite Execution Start Date and Time: "& Now,slogFile,false)
Call EchoAndLog(" Current User: " & GetCurrUserLoggedIn(),slogFile,true)
Call EchoAndLog(" Current folder: " & vBaseDir,slogFile,true)

''If the Test Suite file does not exist then execution will exit
If Not ifFileExist(vLogDir + filename) Then
	Call EchoAndLog(" Specified Log file " & filename & " not found. Please check and re-execute. Quiting...",slogFile,true)
	Wscript.quit
End If

Call ConvertXMLToVbs(vLibLocation & "AUTLibraries\AUTLibrary.xml")

Call EchoAndLog(" Test suite execution started...",slogFile,true)  
call BatchExecution()


'***************************************************************************************************************************************************************
'Function Desc	: Adds all the .vbs files from a specific folder to the Test Resources.
'Function		: AddLibraries(FolderPath)
'Parameters		: FolderPath - Path to the folder containing the .vbs files to be included in the Test Resources
'***************************************************************************************************************************************************************
Public Function ConvertXMLToVbs(XMLPath)
	Set xmlDoc = CreateObject( "Microsoft.XMLDOM" )
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	xmlDoc.Async = "False"
	xmlDoc.Load(XMLPath)
	Wscript.Echo XMLPath
	strVbsPath = mid(XMLPath,1,len(XMLPath)-4)&".vbs"
	strQuery = "/AutLibrary/BusinessFunction"
	Set objFile = objFSO.CreateTextFile(strVbsPath,True)
	Set colNodes = xmlDoc.selectNodes(strQuery)
	wscript.echo colNodes.length
	For Each objNode in colNodes
	objFile.WriteLine objNode.text
	Next
	objFile.close
	Set xmlDoc = Nothing
	Set objFSO = Nothing
End Function


'***************************************************************************************************************************************************************
'Sub Desc	: Includes and executes the VBScript from an external file.
'Sub			: includeFile (fSpec)
'Parameters : fSpec - Path of the file to be included
'***************************************************************************************************************************************************************
Public Sub includeFile (fSpec)
    Dim fileSys, file, fileData
    Set fileSys = CreateObject("Scripting.FileSystemObject")
    Set file = fileSys.openTextFile (fSpec)
    fileData = file.readAll()
    file.close
    ExecuteGlobal fileData
    Set file = Nothing
    Set fileSys = Nothing
End Sub
'***************************************************************************************************************************************************************
'Function Desc	: Displays the current execution progress and also saves it to log file.
'Function		: EchoAndLog(strText,LogFile,append)
'Parameters		: strText - Text string to be written to log
					'logFile - Path of the logFile
					'append - Appends to the existing file if set True
'***************************************************************************************************************************************************************
Sub EchoAndLog(strText,logFile,append)
	WScript.Echo(strText)
	strText = Time & " -" & strText
	Call WriteTextFile(strText,logFile,append)
End Sub

'***************************************************************************************************************************************************************
'Function Desc	: Adds all the .vbs files from a specific folder to the Test Resources.
'Function		: AddLibraries(FolderPath)
'Parameters		: FolderPath - Path to the folder containing the .vbs files to be included in the Test Resources
'***************************************************************************************************************************************************************
Public Function AddLibraries(FolderPath)
	Dim fso, folder, files, NewsFile,sFolder    
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set folder = fso.GetFolder(FolderPath)  
	Set files = folder.Files    
	For each fileIdx In files
		If fso.GetExtensionName(fileIdx.Name) = "vbs" Then
			qtApp.Test.Settings.Resources.Libraries.Add FolderPath & "\" & fileIdx.Name
		End if
	Next
	Set fso = Nothing
	Set folder = Nothing
	Set files = Nothing
End Function

'***********************************************************************************************************************************************
'Function Desc  : Reads the QTP variables and writes into the control file excel
'Function		: GetOutputValues()
'Parameters		: N/A
'************************************************************************************************************************************************
Function GetOutputValues()
	Dim fileSys, file, fileData,lines,words,i,j
	Set fileSys = CreateObject("Scripting.FileSystemObject")
	Set file = fileSys.openTextFile(tempFile,1,False)
	fileData = file.ReadAll()
	file.close
	lines = Split(fileData,vbcrlf)
	For i = LBound(lines) To UBound(lines)
		If lines(i) <> "" Then
			words = Split(lines(i),"|")
			objWorksheet.Cells(RowNum,CInt(words(0))).Value = words(1)
		End if
	Next
	Set file = Nothing
	Set fileSys = Nothing
End Function

'***********************************************************************************************************************************************
'Function Desc  : Writes the sText value to the file at fPath
'Function		: WriteTextFile(sText,fPath,createNew)
'Parameters		: sText - Text value
				  'fPath - Path of the file
				  'createNew - Creates a new file if True
'************************************************************************************************************************************************
Public Function WriteTextFile(sText,fPath,append)
	CONST ForReading = 1, ForWriting = 2, ForAppending = 8 
	Dim fso, myFile, fileData
	Set fso = CreateObject("Scripting.FileSystemObject")
	If append Then
		Set myFile = fso.OpenTextFile(fPath,ForAppending,True)
	Else
		Set myFile = fso.CreateTextFile(fPath)
	End If
	myFile.WriteLine(sText)
	myFile.close
	Set myFile = Nothing
	Set fso = Nothing
End Function

'***************************************************************************************************************************************************************
'Function			: KillProcessesWhileExec()
'Function desc	: Kills all the processes listed
'Parameters		: N/A
'***************************************************************************************************************************************************************
Function KillProcessesWhileExec()
	ProcessNames = "QTPro.exe|QTAutomationAgent.exe"	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objNetwork = CreateObject("Wscript.Network")
	currUser = objNetwork.UserName
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	KillProcess = Split(ProcessNames,"|")
	For x = 0 to  Ubound(KillProcess)
		Set colProcessList = objWMIService.ExecQuery _
		("SELECT * FROM Win32_Process WHERE Name ='"& KillProcess(x) &"'")
		For Each objProcess in colProcessList
			If objProcess.GetOwner ( User, Domain ) = 0 Then
				If UCase(User) = UCase(currUser) Then
					Call EchoAndLog(" '" & KillProcess(x) & "' Process Killed while execution",slogFile,true)
					objProcess.Terminate()
					KillProcessesWhileExec = true
				end if 
			end if
		Next
	Next 
	Set objFSO = Nothing
	Set objNetwork = Nothing 
	Set objWMIService = Nothing 
End Function

'***********************************************************************************************************************************************
'Function Desc   : Reads the Test suite execution and proceeds with execution
'Function			: BatchExecution()
'Parameters		  : N/A
'************************************************************************************************************************************************
Function BatchExecution()
	''Get Current User Logged In
	Set objNetwork1 = CreateObject("Wscript.Network")
	currUser1 = objNetwork1.UserName
	Call EchoAndLog(" Preparing to execute the QTP Script...",slogFile,true)
	''Script Execution
	call ExecuteQTPScript()
	Call EchoAndLog(" Script execution completed: " & time,slogFile,true)
	''Call EchoAndLog(" Closing QTP and BCI, while execution",slogFile,true)
	''qtApp.Quit
	''Call KillProcessesWhileExec()			
	Call EchoAndLog(" Test Suite Execution End Date and Time: "& Now,slogFile,true)
	Call EchoAndLog(" Execution completed",slogFile,true)
End Function

'***********************************************************************************************************************************************
'Function Desc   : Gets the excel rows, columns, and sets all the rows under Control column to "N"
'Function			: GetExcelRowsAndCols(strFileP, sSheetName, ControlCol, byref RowCnt, byref ColCnt)
'Parameters		  : strFileP, sSheetName, byref RowCnt, byref ColCnt
'************************************************************************************************************************************************
Function GetExcelRowsAndCols(strFileP, sSheetName, ControlCol, byref RowCnt, byref ColCnt)
	Set XLObj = CreateObject ("Excel.Application")
	Set XLWorkBook = XLObj.Workbooks.open (strFileP)
	XLObj.DisplayAlerts = False
	Set XLSheet = XLWorkBook.Worksheets (sSheetName)
	RowCnt = XLSheet.Usedrange.rows.count
	ColCnt = XLSheet.Usedrange.columns.count
	For rc = 2 to RowCnt
		XLSheet.Cells (rc,ControlCol).Value = "N"
	Next
	XLWorkBook.SaveAs strFileP
	XLWorkBook.Close
	XLObj.Quit
	Set XLObj = nothing
	Set XLWorkBook = nothing
	Set XLSheet = nothing
End Function

'***********************************************************************************************************************************************
'Function Desc   : Excel cells formating
'Function			: lCellFormating(objTestSetSheet,RowCnt,ColLbl)
'Parameters		  : objTestSetSheet,RowCnt,ColLbl
'************************************************************************************************************************************************
Function XlCellFormating(objTestSetSheet,RowCnt,ColLbl)
 	Lbl = Split(ColLbl,",")
	For k = 1 to Ubound(Lbl)
		objTestSetSheet.Cells(RowCnt,k) = Lbl(k)
		objTestSetSheet.Cells(RowCnt,k).interior.colorindex = 45
		objTestSetSheet.Cells(RowCnt,k).Font.Bold = True
		objTestSetSheet.Cells(RowCnt,k).HorizontalAlignment = -4108
	Next
End Function

'***********************************************************************************************************************************************
'Function Desc   : Opens desired QTP Script, associates required libraries, other files, manages settings and continues with the script execution
'Function			: ExecuteQTPScript(objTestSetSheet,TestSetRow, ScriptNameCol, ScriptLogCol, LibraryCol, ExecNum, CurrExecSetNo)
'Parameters		  : objTestSetSheet,TestSetRow, ScriptNameCol, ScriptLogCol, LibraryCol, ExecNum, CurrExecSetNo
'***********************************************************************************************************************************************
Function ExecuteQTPScript()	
	Call EchoAndLog(" Launching QTP...",slogFile,true)
	call LaunchQTP()
	Set qtApp = CreateObject("QuickTest.Application")	 	
	''**********Test Script Execution start'**********	
	Call EchoAndLog(" Opening the driver script...",slogFile,true)
	'wscript.sleep(1000)
	qtApp.Open vDriverScriptDir & "SuiteDriver"
	Call EchoAndLog(" Script opened, associating all the libraries to the current script...",slogFile,true) 
	qtApp.visible = True
	Set qtTest = qtApp.Test 
	Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions") ''Create the Run Results Options object
	qtResultsLoc = TestExecResultsPath 
	qtResultsOpt.ResultsLocation = qtResultsLoc ''Set the results location
	''Loading Environment Variables file
	qtTest.Environment.LoadFromFile vBaseDir & "\Environment Files\EnvironVariables.XML", True
	''Setting Environmanet Variables for Script Log file, 
	qtTest.Environment.Value("vBaseDir") = vBaseDir & "\"
	qtTest.Environment.Value("TestExecResultsPath") = TestExecResultsPath
	
	'Loading necessary Datatable Sheets
'	qtApp.Test.DataTable.DeleteSheet ("Global")
	qtApp.Test.DataTable.AddSheet ("TestSuite")
'	qtApp.Test.DataTable.AddSheet ("TestCase")
'	qtApp.Test.DataTable.AddSheet ("TestData")
'	qtApp.Test.DataTable.AddSheet ("Component")
'	qtApp.Test.DataTable.AddSheet ("Parameters")
'	qtApp.Test.DataTable.DeleteSheet ("Action1")
	'Loading Necessary Libraries for the test
	Set qtLibraries = qtTest.Settings.Resources.Libraries 
	qtLibraries.RemoveAll
	qtLibraries.Add vLibLocation & "GeneralLibraries\" + "GlobalVariables.vbs", 1 ' Add the library to the collection
	qtLibraries.Add vLibLocation & "GeneralLibraries\" + "ConfigFile.vbs", 2
	qtLibraries.Add vLibLocation & "GeneralLibraries\" + "CommonLibrary.vbs", 3
	qtLibraries.Add vLibLocation & "GeneralLibraries\" + "RegisteredFunctions.vbs", 4
	qtLibraries.Add vLibLocation & "GeneralLibraries\" + "CommonLibrary.vbs", 5
	qtLibraries.Add vLibLocation & "GeneralLibraries\" + "BusinessFunction.vbs", 6
	qtLibraries.Add vLibLocation & "GeneralLibraries\" + "ReportingLib.vbs", 7
	qtLibraries.Add vLibLocation & "GeneralLibraries\" + "Utilities.vbs",8
	qtLibraries.Add vLibLocation & "AUTLibraries\" + "AUTLibrary.vbs", 9
	
	''Adding Recovery Sceenarios
'	Set qtTestRecovery = qtApp.Test.Settings.Recovery	
	
'	If qtTestRecovery.Count > 0 Then 	'' If there are any default scenarios specified for the test
'	    qtTestRecovery.RemoveAll ' Remove them
'	End If
'	qtTestRecovery.Add vRecoveryScn1 & "BluesEnrollRecovery.qrs", "BluesEnrollPopUps", 1	
'	qtTestRecovery.Item(1).Enabled = True
'	For intIndex = 1 To qtTestRecovery.Count
'		qtTestRecovery.Item(intIndex).Enabled = True
'	Next
'	qtTestRecovery.Enabled = True
'	qtTestRecovery.SetActivationMode "OnEveryStep"	
'	Set qtTestRecovery = Nothing
		
	qtTest.Settings.Run.IterationMode = "oneIteration"
	qtTest.Settings.Run.ObjectSyncTimeOut = 100
	qtApp.Options.Run.RunMode = "Normal"
	qtApp.Options.Run.StepExecutionDelay = 8
'	If RecordingMode Then
'		qtApp.Options.Run.MovieCaptureForTestResults = "Always" 	'Set QuickTest to capture movie segments of each error and warning
'		qtApp.Options.Run.SaveMovieOfEntireRun = "TRUE" 	'Save only the movie segements that lead up to an error or warning	
'		Set qtMovieOpts = qtApp.Options.Run.ScreenRecorder 	'Create a Screen Recorder Options object
'		qtMovieOpts.DeactivateShowWindowContents = True
'		qtMovieOpts.RecordSound = False
'		qtMovieOpts.SetPlainWallpaper = True
'		Call EchoAndLog(" Please wait, the QTP script exec started at: '" & Time & "' and is in progress...",slogFile,true)
		'Wscript.Echo " Please wait, while the QTP script execution started at: '" & Time & "' is in progress..."
'		qtTest.Run qtResultsOpt
'	Else
		Call EchoAndLog(" Please wait, the QTP script exec started at: '" & Time & "' and is in progress...",slogFile,true)
		'Wscript.Echo " Please wait, while the QTP script execution started at: '" & Time & "' is in progress..."
		qtTest.Run
'	End If
'Call EchoAndLog(" Making QTP visible",slogFile,true)
	qtApp.Visible = true
	qtApp.Quit
	Set qtApp = Nothing
End Function

'***********************************************************************************************************************************************
'Function Desc   : Launches QTP 
'Function			: LaunchQTP()
'Parameters		  : NA
'***********************************************************************************************************************************************
Function LaunchQTP()
	'On Error Resume Next
	Dim qtApp, Ret_Value
	Set qtApp = CreateObject("QuickTest.Application")	
	''Storing all the necessary Addins in any array
	Dim QTPAddin(1)
	QTPAddin(0) = "ActiveX"	
	QTPAddin(1) = "Web"
	''Checking if the QTP is already opened. If opened closes the QTP app and reopens with the defined Addins
	If qtApp.Launched Then
		qtApp.Quit
	End If
	Ret_Value = qtApp.SetActiveAddins(QTPAddin)
	qtApp.Launch
	qtApp.Visible = false
	If (Ret_Value = True) Then
		'Nothing
	Else
		Call EchoAndLog(" Specified Addins does not exist",slogFile,true)
		'WScript.echo "Specified Addins does not exist" 
		qtApp.Quit 
		WScript.Quit
	End If
	Set qtApp = Nothing
	Erase QTPAddin
End Function

'***********************************************************************************************************************************************
'Function Desc   : Gets the Absolute Path
'Function			: GetAbsolutePath()
'Parameters		  : NA
'***********************************************************************************************************************************************
Function GetAbsolutePath()
	Set WshShell = CreateObject("WScript.Shell") 
	Set WshSysEnv = WshShell.Environment("System") 
    Set objFSO = CreateObject("Scripting.FileSystemObject")
	vBaseDir = objFSO.GetAbsolutePathName(".")
	GetAbsolutePath = vBaseDir
	Set objFSO = Nothing
End Function

'***********************************************************************************************************************************************
'Function Desc : Checks if the file exists in the path mentioned 
'Function : ifFileExist(searchFilePath)
'Parameters : searchFilePath
'***********************************************************************************************************************************************
Function ifFileExist(searchFilePath)
	ifFileExist = false
	Dim fObj
	Set fObj = CreateObject("Scripting.FileSystemObject")
	if fObj.FileExists(searchFilePath) Then
		ifFileExist = true
	Else
		ifFileExist = false
	End If
	Set fObj = nothing	
End Function

'***********************************************************************************************************************************************
'Function Desc   : Checks if the folder exists in the path mentioned 
'Function			: ifFolderExist(searchFolderPath)
'Parameters		  : searchFolderPath
'***********************************************************************************************************************************************
Function ifFolderExist(searchFolderPath)
	ifFolderExist = false
	Dim fObj	
	Set fObj = CreateObject("Scripting.FileSystemObject")
	if fObj.FolderExists(searchFolderPath) Then
		ifFolderExist = true
	Else
		ifFolderExist = false
	End If
	Set fObj = nothing
End Function

Function ifNotCreateFolder(sFolders)
	'Dim objFSO
	'Set objFSO = CreateObject("Scripting.FileSystemObject")
	Call GenerateFolderPath(sFolders)
	'Set objFSO = nothing
End Function

Function GenerateFolderPath(pFolderPath)
	GenerateFolderPath = False
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FolderExists(pFolderPath) Then
		If GenerateFolderPath(objFSO.GetParentFolderName(pFolderPath)) Then
			GenerateFolderPath = True
			Call objFSO.CreateFolder(pFolderPath)
		End If
	Else
		GenerateFolderPath = True
	End If
	Set objFSO = nothing
End Function

'***********************************************************************************************************************************************
'Function Desc   : returns the test duration based on the execution start and end time
'Function			: TestDuration(a,b)
'Parameters		  : Start time, End time
'***********************************************************************************************************************************************
Function TestDuration(a,b)
	s=Second(b)-Second(a)
	m=Minute(b)-Minute(a)
	h=Hour(b)-Hour(a)  
	If a<>b Then
	If (s<0) Then
	   s=s+60
	   m=m-1
	End if
	if m<0 Then
	   m=m+60
	   h=h-1
	End If
	If h<0 Then
	   h=h+24
	End If
	End if
	If m<10 Then
	m=0&m
	End If
	If s<10 Then
	s=0&s
	End If
	TestDuration = h&":"&m&":"&s
End Function

'***********************************************************************************************************************************************
'Function Desc   : replaces all the special characters with "_"
'Function			: ReplaceFunc(l_strVariable)
'Parameters		  : string
'***********************************************************************************************************************************************
Function ReplaceFunc(l_strVariable)
	Dim tempVariable
	tempVariable = replace(l_strVariable,":", "_")
	tempVariable = replace(tempVariable,".","_")
	tempVariable = replace(tempVariable,"(","_")
	tempVariable = replace(tempVariable,"/","_")
	tempVariable = replace(tempVariable,")","_")
	tempVariable = replace(tempVariable,"&","_")
	tempVariable = replace(tempVariable,"-","_")
	tempVariable = replace(tempVariable,"#","_")
	tempVariable = replace(tempVariable,"""","_")
	tempVariable = replace(tempVariable," ","_",1,-1)
	tempVariable = replace(tempVariable,"'","_")
	tempVariable = replace(tempVariable,",","_")
	tempVariable = replace(tempVariable,">","_")
	tempVariable = replace(tempVariable,"<","_")
	ReplaceFunc = replace(tempVariable,"*","_")
End Function

Function GetCurrUserLoggedIn()
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objNetwork = CreateObject("Wscript.Network")
	GetCurrUserLoggedIn = objNetwork.UserName
	Set objFSO = NOTHING
	Set objNetwork = NOTHING	
End Function 

''****************************************************************************************************************************************************************
''Function Desc   : Sends Email with on the execution status inputs
''Function			: ExecStatusMail(sFrom,sTo,sCc,sBcc,sSubject,sBody)
''Parameters		  : sFrom,sTo,sCc,sBcc,sSubject,sBody
''****************************************************************************************************************************************************************
Function ExecStatusMail(sFrom,sTo,sCc,sBcc,sSubject,sBody,ESMAttachments)
	On Error Resume Next
	Const cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing"
	Const cdoSendUsingPort = 2
	Const cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
	Const cdoSMTPConnectionTimeout = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
	Const cdoSMTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
	Const cdoBasic = 1
	Const cdoSendUserName = "http://schemas.microsoft.com/cdo/configuration/sendusername"
	Const cdoSendPassword = "http://schemas.microsoft.com/cdo/configuration/sendpassword"

	Set objConfig = CreateObject("CDO.Configuration")									
	Set Fields = objConfig.Fields
									
	With Fields
		.Item(cdoSendUsingMethod)       = cdoSendUsingPort
		.Item(cdoSMTPServer)            = "smtp.gmail.com"		''Mention Smtp details here
		.Item(cdoSMTPServerPort)        = 25
		.Item(cdoSMTPConnectionTimeout) = 10
		.Item(cdoSMTPAuthenticate)      = cdoBasic
		.Item(cdoSendUserName)          = sFrom
		.Item(cdoSendPassword)          = sPassword		''Mention password here
		.Update()
	End With 
	Set cdoMessage = CreateObject("CDO.Message")
	Set cdoMessage.Configuration = objConfig     
	cdoMessage.From = sFrom 
	cdoMessage.To = sTo 
	cdoMessage.Cc = sCc
	cdoMessage.Bcc = sBcc
	cdoMessage.Subject = sSubject 
    cdoMessage.TextBody = sBody
    If ESMAttachments <> "" Or ESMAttachments <> Empty Then
	    If InStr(ESMAttachments,"||") > 0 Then
			ESMAttach = Split(ESMAttachments,"||")
			For EsmAtt = 0 To UBound(ESMAttach)
				If ESMAttach(EsmAtt) <> "" Or ESMAttach(EsmAtt) <> Empty Then 
					cdoMessage.AddAttachment ESMAttach(EsmAtt)
				End If  
			Next
		Else
			cdoMessage.AddAttachment ESMAttachments
		End If
	End If 
	cdoMessage.Send()  
    Set cdoMessage = Nothing 
    Set cdoConfig = Nothing 
end Function

'***************************************************************************************************************************************************************
'Function		: KillProcessesBeforeExec()
'Function desc	: Kills all the processes listed
'Parameters		: N/A
'***************************************************************************************************************************************************************
Function KillProcessesBeforeExec()
	ProcessNames = "QTPro.exe|QTAutomationAgent.exe|EXCEL.EXE|wscript.exe" 	''|cscript.exe|
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objNetwork = CreateObject("Wscript.Network")
	currUser = objNetwork.UserName
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	KillProcess = Split(ProcessNames,"|")
	For x = 0 to  Ubound(KillProcess)
		Set colProcessList = objWMIService.ExecQuery _
		("SELECT * FROM Win32_Process WHERE Name ='"& KillProcess(x) &"'")
		For Each objProcess in colProcessList
			If objProcess.GetOwner ( User, Domain ) = 0 Then
				If UCase(User) = UCase(currUser) Then
					Wscript.Echo " '" & KillProcess(x) & "' Process Killed while execution"
					'Call EchoAndLog(" '" & KillProcess(x) & "' Process Killed while execution",slogFile,true)
					objProcess.Terminate()
					KillProcessesBeforeExec = true
				end if 
			end if
		Next
	Next 
	Set objFSO = Nothing
	Set objNetwork = Nothing 
	Set objWMIService = Nothing 
End Function