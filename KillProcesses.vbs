ProcessNames = "QTPro.exe|QTAutomationAgent.exe|EXCEL.EXE|wscript.exe" 	''|cscript.exe|
Call KillReqdProcess(ProcessNames)
MsgBox "Kill Process Complete"

'***************************************************************************************************************************************************************
'Script			: KillReqdProcess
'Script Desc	: Kills all the processes listed
'Parameters		: N/A
'***************************************************************************************************************************************************************
Function KillReqdProcess(ProcessNames)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	ParentFolder = objFSO.GetAbsolutePathName(".")

	'Get Current User Logged In
	Set objNetwork = Wscript.CreateObject("Wscript.Network")
	currUser = objNetwork.UserName
	'WScript.echo "Current User of the machine: "& currUser 
	'WScript.echo "	" 
	'WScript.echo "Test Suite Execution Start Date and Time: "& now
	'Kill Process(es) based on *.exe and current user 
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	KillProcess = Split(ProcessNames,"|")
	For x = 0 to  Ubound(KillProcess)
		Set colProcessList = objWMIService.ExecQuery _
		("SELECT * FROM Win32_Process WHERE Name ='"& KillProcess(x) &"'")
		Call WriteTextFile(time& " - Trying to find '"& KillProcess(x) &"' process in the Task Manager.",ParentFolder & "\ProcessesKilledLog.txt",true)
		'WScript.echo "Trying to find '"& KillProcess(x) &"' process in the Task Manager." 
		For Each objProcess in colProcessList
			If objProcess.GetOwner ( User, Domain ) = 0 Then
				'MsgBox User & " " & Domain
				If UCase(User) = UCase(currUser) then
					objProcess.Terminate()
					Call WriteTextFile(time& " - "& KillProcess(x) &" - terminated successfully.",ParentFolder & "\ProcessesKilledLog.txt",true)
					WScript.echo KillProcess(x) &" - terminated successfully."
					'WScript.echo "	" 
					'MsgBox KillProcess(x) &" - terminated successfully."
				end if 
			end if
		Next
	Next	
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