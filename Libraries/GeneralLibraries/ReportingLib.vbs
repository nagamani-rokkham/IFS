'Option Explicit

'@Function Name <rpt_WriteReport>
'@CreationDate <02-mar-2014>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description	this method write the test case results in html file   
'@Documentation <param> and <param> will do….  

Public gFolderName
Public strResultsFolder1
Dim iSNO, iTestCaseNumber

Public Function rptWriteReport(ByVal strResult, ByVal strStepName , ByVal strExpected)
'***********************************************Initial Setup******************************************
  Dim objFso,objFolder,strResultsFolder,strTestcasesPath,objFile,status, html, link,arrPath,strResources  
  strResources=gTestDir & "Resources"
  Set objFso = CreateObject("Scripting.FileSystemObject")
 
   strResultsFolder = strResultsFolder1
   strLogos=strResultsFolder&"\Logos"
   If not objFso.FolderExists(strLogos) Then
	   Set objFolder=objFso.CreateFolder(strLogos)
	   objFso.CopyFile strResources&"\Pass.png",strLogos&"\"
	   objFso.CopyFile strResources&"\Fail.png",strLogos&"\"
   End If
   
   strTestcasesPath=strResultsFolder&"\Testcases"
   If not objFso.FolderExists(strTestcasesPath) Then
	   Set objFolder=objFso.CreateFolder(strTestcasesPath)
   End If
	If not objFso.FileExists(strTestcasesPath&"\"&Environment.Value("Test_Name")&".html") Then
		iSNO = 1
		Set  objFile = objFso.CreateTextFile(strTestcasesPath&"\"&Environment.Value("Test_Name")&".html",true, false)				
		objFile.writeline  "<html><head> <meta charset='UTF-8'> <title>Detailed Results Report</title><style type='text/css'>" 
		objFile.writeline  "body { background-color: #FFFFFF; font-family: Verdana, Geneva, sans-serif; text-align: center; } small { font-size: 0.7em; } "
		objFile.writeline  "table { box-shadow: 9px 9px 10px 4px #BDBDBD;border: 0px solid #4D7C7B; border-collapse: collapse; "
		objFile.writeline  "border-spacing: 0px; width: 1000px; margin-left: auto; margin-right: auto; } tr.heading { background-color: #041944;"
		objFile.writeline  "color: #FFFFFF; font-size: 0.7em; font-weight: bold; background:-o-linear-gradient(bottom, #999999 5%, #000000 100%);"
		objFile.writeline  "background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #999999), color-stop(1, #000000)"
		objFile.writeline  ");background:-moz-linear-gradient( center top, #999999 5%, #000000 100%);"
		objFile.writeline  "filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#999999, endColorstr=#000000);	"
		objFile.writeline  "background: -o-linear-gradient(top,#999999,000000);} tr.subheading { background-color: #FFFFFF; color: #000000; "
		objFile.writeline  "font-weight: bold; font-size: 0.7em; text-align: justify; } tr.section { background-color: #A4A4A4; color: #333300; "
		objFile.writeline  "cursor: pointer; font-weight: bold; font-size: 0.7em; text-align: justify; "
		objFile.writeline  "background:-o-linear-gradient(bottom, #56aaff 5%, #e5e5e5 100%);	"
		objFile.writeline  "background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #56aaff), color-stop(1, #e5e5e5));"
		objFile.writeline  "background:-moz-linear-gradient( center top, #56aaff 5%, #e5e5e5 100%);"
		objFile.writeline  "filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#56aaff, endColorstr=#e5e5e5);	"
		objFile.writeline  "background: -o-linear-gradient(top,#56aaff,e5e5e5);} tr.subsection { cursor: pointer; } tr.content { background-color: #FFFFFF;"
		objFile.writeline  "color: #000000; font-size: 0.7em; display: table-row; } tr.content2 { background-color: #E1E1E1; border: 1px solid #4D7C7B;"
		objFile.writeline  "color: #000000; font-size: 0.75em; display: table-row; } td, th { padding: 5px; border: 1px solid #4D7C7B; "
		objFile.writeline  "text-align: inherit /; } th.Logos { padding: 5px; border: 0px solid #4D7C7B; text-align: inherit /;} "
		objFile.writeline  "td.justified { text-align: Left; } td.pass { font-weight: bold; color: green; } td.fail { font-weight: bold; color: red; }"
		objFile.writeline  "td.done, td.screenshot { font-weight: bold; color: black; } td.debug { font-weight: bold;color: blue; } "
		objFile.writeline  "td.warning { font-weight: bold; color: orange; } </style> </head><body></br><table id='Logos'> "
		objFile.writeline  "<colgroup><col style='width: 25%' /><col style='width: 25%' /><col style='width: 25%' /><col style='width: 25%' /></colgroup> "
		objFile.writeline  "<thead> <tr class='content'><th class ='Logos' colspan='2' ><img align ='left' src="
		objFile.writeline  "	http://apps.ifsautomation.com/RIADemo/(S(s0slutdzgk41tgmifxbtjxhc))/images/IFSLogoForLogin.png></img></th>"
		objFile.writeline  "<th class = 'Logos' colspan='2' > <img align ='right' src= http://www.gallop.net/images/logo.png></img></th> </tr> "
		objFile.writeline  "</thead> </table><table id='header'> <colgroup> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> "
		objFile.writeline  "<col style='width: 25%' /> </colgroup>  <thead> <tr class='heading'> <th colspan='4' style='font-family:Copperplate Gothic Bold;"
		objFile.writeline  "font-size:1.4em;'> **"& Environment.Value("Test_Name") & "- Execution Results **</th> </tr> <tr class='subheading'>"
		iCurrentTime = Now()
		objFile.writeline  "<th>&nbsp;Date&nbsp;&&nbsp;Time&nbsp;:&nbsp;</th> <th>"& DatePart("d", iCurrentTime) & "-" & MonthName(Month(iCurrentTime), True) & "-" & DatePart("yyyy", iCurrentTime) & Space(1) & Hour(iCurrentTime) & ":" & Minute(iCurrentTime) & "</th> <th>&nbsp;Operating&nbsp;System&nbsp;:&nbsp;</th>"
		objFile.writeline  "<th> "& Environment.Value("OS") & "</th> </tr> <tr class='subheading'> <th>&nbsp;Browser&nbsp;:&nbsp;</th>"
		If ucase(gBrowserIE) = ucase("yes")	Then
			BrowserName="Internet Explorer"
        ElseIf ucase(gBrowserChrome) =ucase("yes") Then
			BrowserName="CHROME"
		ElseIf ucase(gBrowserFF) =ucase("yes") Then
			BrowserName="FIREFOX"
        End If
		objFile.writeline  "<th> "& BrowserName & "</th> <th>&nbsp;Executed&nbsp;on&nbsp;:&nbsp;</th> <th>" & Environment.Value("LocalHostName") & "</th> </tr> </thead> </table><table id='main'> <colgroup>" 
		objFile.writeline  "<col style='width: 5%' /> <col style='width: 26%' /> <col style='width: 51%' /> "
		objFile.writeline  "<col style='width: 8%' /> <col style='width: 10%' />"
		objFile.writeline  "</colgroup> <thead> <tr class='heading'> <th>S.NO</th> <th>Steps</th> <th>Details</th> <th>Status</th> <th>Time</th> </tr> "
	Else
		Set objFile=objFso.OpenTextFile(strTestcasesPath&"\"&Environment.Value("Test_Name")&".html", 8,TRUE)    
   End If

   Select Case ucase(strResult)
		Case "PASS" 
				Reporter.ReportEvent micPass , strStepName , strActual
				objFile.WriteLine "<tr class='content2' ><td>" & iSNO & "</td> "
				objFile.WriteLine "<td class='justified'>" & strStepName &"</td>"
				objFile.WriteLine "<td class='justified'>" & strExpected & "</td>"
				objFile.WriteLine "<td class='Pass' align='center'><img  src='" & "..\Logos\Pass.png' width='18' height='18'/></td> "
				iCurrentTime = Now()
				objFile.WriteLine "<td><small>" & DatePart("d", iCurrentTime) & "-" & MonthName(Month(iCurrentTime), True) & "-" & DatePart("yyyy", iCurrentTime) & Space(1) & Hour(iCurrentTime) & ":" & Minute(iCurrentTime) & ":" & Second(iCurrentTime)& "</small></td> </tr>"
				rptReportLog strStepName,strExpected,"Pass"						
		Case "FAIL"	
				Reporter.ReportEvent micFail , strStepName , strActual
				objFile.WriteLine "<tr class='content' ><td>" & iSNO & "</td> "
				objFile.WriteLine "<td class='justified'>" & strStepName &"</td>"
				objFile.WriteLine "<td class='justified'>" & strExpected & "</td> "
				link = rptScreenCapture()
				objFile.WriteLine "<td class='Fail' align='center'><a href="& link &">"
				iCurrentTime = Now
				objFile.WriteLine "<img  src='" & "..\Logos\Fail.png' width='18' height='18'/></td> <td><small>" & DatePart("d", iCurrentTime) & "-" & MonthName(Month(iCurrentTime), True) & "-" & DatePart("yyyy", iCurrentTime) & Space(1) & Hour(iCurrentTime) & ":" & Minute(iCurrentTime) & ":" & Second(iCurrentTime)& "</small></td> </tr>"
				rptReportLog strStepName,strExpected,"Fail"
	End Select
  iSNO = iSNO+1
End Function

'*****************************************************************
'@CreationDate <02-mar-2014>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description	it capture screen and send path of image fie to called function     
'@Documentation <param> and <param> will do….
'******************************************************************
Function rptScreenCapture()
	 Dim objFso,strResultsPath,strScreenshotPath,objFolder,strImagePath,strFilePath,strImagelinkPath,objDesktop
	 Set objFso = CreateObject("Scripting.FileSystemObject")
	 'strResultsPath=Environment.Value("RESULTSPATH")&"\"&gFolderName
	 strScreenshotPath=strResultsFolder1 &"\screenshot"
	  If not  objFso.FolderExists(strScreenshotPath) Then
		   Set objFolder=objFso.CreateFolder(strScreenshotPath)
	  End If
	
	 strImagePath="\screenshot"&Replace(Replace(Replace(now(),":","_"),"/","_")," ","_") &".png"
	 strFilePath=strScreenshotPath&strImagePath
	 strImagelinkPath="..\screenshot"&strImagePath
	 Set objDesktop = Desktop
	 objDesktop.capturebitmap strFilePath ,  true
	
	 rptScreenCapture=strImagelinkPath
End Function 


'*********************************************************
'@CreationDate <02-mar-2014>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description	it creates summary report of executed  test cases.      
'@Documentation <param> and <param> will do….
  
'***********************************************************
Function rptWriteResultsSummary()
		
		Dim strResultsPath, objSummary,objFilesummary,strResultsFolder,objFSO,objFolder,objFiles,iCount,iFailCount,iPassCount,objFile,FailedScriptPercentage,PassedSrciptPercentage,SummaryChart,html,SummaryChart1,wshShell

		strResources=gTestDir &"Resources"
		SummaryChart =strResultsFolder1 &"\SummaryChart.html"
				
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFolder = objFSO.GetFolder(strResultsFolder1&"\Testcases")
		Set objFiles = objFolder.Files
		

		
		Set objFile = objFSO.CreateTextFile(SummaryChart, true, false)
		
		objFile.writeline "<table id='Logos'> <colgroup> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> </colgroup> "
		objFile.writeline "<thead>  <tr class='content'> <th class ='Logos' colspan='2' > <img align ='left' src='http://apps.ifsautomation.com/RIADemo/(S(s0slutdzgk41tgmifxbtjxhc))/images/IFSLogoForLogin.png'></img> </th>"
		objFile.writeline "<th class = 'Logos' colspan='2' > <img align ='right' src='http://www.gallop.net/images/logo.png'></img> </th> </tr> </thead> </table> "		
		objFile.writeline "<html> <head> <script src='http://ajax.googleapis.com/ajax/libs/jquery/1.7.1/jquery.min.js' type='text/javascript'></script>"
		objFile.writeline "<script src='/js/highcharts.js' type='text/javascript'></script><script src='http://code.highcharts.com/highcharts.js'></script>"
		objFile.writeline "<script src='http://code.highcharts.com/highcharts-3d.js'></script><script src='http://code.highcharts.com/modules/exporting.js'></script>"
		objFile.writeline "<meta charset='UTF-8'> <title>IFS - Automation Execution Results Summary</title><style type='text/css'>body {background-color: #FFFFFF; "
		objFile.writeline "font-family: Verdana, Geneva, sans-serif; text-align: center; } small { font-size: 0.7em; } table { box-shadow: 9px 9px 10px 4px #BDBDBD;"
		objFile.writeline "border: 0px solid #4D7C7B;border-collapse: collapse; border-spacing: 0px; width: 1000px; margin-left: auto; margin-right: auto; } "
		objFile.writeline "tr.heading { background-color: #041944;color: #FFFFFF; font-size: 0.7em; font-weight: bold; "
		objFile.writeline "background:-o-linear-gradient(bottom, #999999 5%, #000000 100%); "
		objFile.writeline "background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #999999), color-stop(1, #000000) );"
		objFile.writeline "background:-moz-linear-gradient( center top, #999999 5%, #000000 100% );"
		objFile.writeline "filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#999999, endColorstr=#000000); "
		objFile.writeline "background:-o-linear-"
		objFile.writeline "gradient(top,#999999,000000);} tr.subheading { background-color: #6A90B6;color: #000000; font-weight: bold; font-size: 0.7em; "
		objFile.writeline "text-align:justify; } tr.section { background-color: #A4A4A4; color: #333300; cursor: pointer; font-weight: bold;font-size: 0.8em; "
		objFile.writeline "text-align: justify;"
		objFile.writeline "background:-o-linear-gradient(bottom, #56aaff 5%, #e5e5e5 100%); "
		objFile.writeline "background:-webkit-gradient( linear, left top, left bottom,color-stop(0.05, #56aaff), color-stop(1, #e5e5e5) );"
		objFile.writeline "background:-moz-linear-gradient( center top, #56aaff 5%, #e5e5e5 100% );"
		objFile.writeline "filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#56aaff, endColorstr=#e5e5e5);"
		objFile.writeline "background:-o-linear-gradient(top,#56aaff,e5e5e5);} tr.subsection { cursor: pointer; } "
		objFile.writeline "tr.content { background-color: #FFFFFF; color:#000000; font-size: 0.7em; display: table-row; } "
		objFile.writeline "tr.content2 { background-color:#;E1E1E1border: 1px solid #4D7C7B;color: #000000; "
		objFile.writeline "font-size: 0.7em; display: table-row; } td, th { padding: 5px; border: 1px solid #4D7C7B; text-align: inherit/; } th.Logos {" 
		objFile.writeline "padding: 5px; "
		objFile.writeline "border: 0px solid #4D7C7B; text-align: inherit /;} td.justified { text-align: justify; } td.pass {font-weight: bold; color: green;"
		objFile.writeline "}" 
		objFile.writeline "td.fail { font-weight: bold; color: red; } td.done, td.screenshot { font-weight: bold; color: black; } "
		objFile.writeline "td.debug { font-weight: bold;color: blue; } td.warning { font-weight: bold; color: orange; } </style> </head> "
		objFile.writeline "<body> </br><table id='header'> "
		objFile.writeline "<colgroup> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> " 
		objFile.writeline "<col style='width: 25%' /> </colgroup> <thead> <tr class='heading'> <th colspan='4' style='font-family:Copperplate Gothic Bold;" 
		objFile.writeline "font-size:1.4em;'> IFS  -  Automation Execution Result Summary </th> </tr> <tr class='subheading'>   "
		objFile.writeline "<th>&nbsp;Date&nbsp;&&nbsp;Time&nbsp;India Standard Time</th> <th>&nbsp;&nbsp;"& Now &"</th>"
		If ucase(gBrowserIE) = ucase("yes")	Then
			BrowserName="Internet Explorer"
        ElseIf ucase(gBrowserChrome) =ucase("yes") Then
			BrowserName="CHROME"
		ElseIf ucase(gBrowserFF) =ucase("yes") Then
			BrowserName="FIREFOX"
        End If
		objFile.writeline "<th>&nbsp;Browser</th> <th>&nbsp;&nbsp;"&BrowserName&"</th> </tr> <tr class='subheading'> <th>&nbsp;Suite Executed</th> <th>&nbsp;&nbsp;Regression</th> <th>&nbsp;Host Name</th>"
		objFile.writeline "<th>&nbsp;&nbsp;"& Environment.Value("LocalHostName") & "</th></tr></thead></table>"
		
		objFile.writeline "<table id='main'> <colgroup> <col style='width: 10%' /> <col style='width: 40%' /> <col style='width: 20%' /> <col style='width:" 
		objFile.writeline "30%' /> </colgroup> "
		objFile.writeline "<thead> <tr class='heading'> <th>S.NO</th> <th>Test Case</th> <th>Status</th> <th>Time</th> </tr> </thead> <tbody>"
		Set objFile = Nothing
			

		iCount=0
		iFailCount=0
		iPassCount=0
		iTestCaseNumber = 0  
		iToltalExecutionTime = 0
		For Each Item In objFiles
		   If LCase(Right(Item.Name, 5)) = ".html" Or LCase(Right(Item.Name, 4)) = ".htm" Then
			  Set objFileDetailedReport = objFSO.OpenTextFile(Item.Path, 1, False)
				 strText = objFileDetailedReport.readAll()
				 
				 Set objReg = New RegExp
				 objReg.Pattern = "[\d]+-[a-zA-Z]+-[\d]+ [\d]+:[\d]+:[\d]+"			 
				 objReg.Global = True
				 Set objMatches =  objReg.Execute(strText)
				 iStepCount = objMatches.Count
				 
				 iStartTime = objMatches(0).Value
				 iEndTime = objMatches(iStepCount-1).Value
				 iExecutionTime = Round((CDbl(DateDiff("s",CDate(iStartTime),CDate(iEndTime)))/60), 2) &" Minutes"	
				iToltalExecutionTime = iToltalExecutionTime + (Round((CDbl(DateDiff("s",CDate(iStartTime),CDate(iEndTime)))/60), 2))
				 
				 If Instr(strText,"Fail.png") > 0 Then
					rptwritesum Item.Name, "FAIL", iExecutionTime
					iFailCount = iFailCount +1
				 Else
					rptwritesum Item.Name,"PASS" ,iExecutionTime
					iPassCount = iPassCount +1 
				 End If
		   End If
		   Set objFileDetailedReport = Nothing
		Next 
			
		Set objFile=objFSO.openTextFile(SummaryChart, 8, True)		
		objFile.writeline "</table> <table id='footer'> <colgroup> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> </colgroup> "
		objFile.writeline "<tfoot> <tr class='heading'>	<th colspan='4'>Total Duration (Including Report Creation) :" & iToltalExecutionTime &"Minutes</th> </tr> <tr class='content'>"
		objFile.writeline "<td class='pass'>&nbsp;Tests passed</td>	<td class='pass'>&nbsp;"&iPassCount&"</td> <td class='fail'>&nbsp;Tests failed</td>	<td class='fail'>&nbsp; "&iFailCount&"</td> </tr> </tfoot> </table>"
		  
		  set objFile=nothing
		  set objFSO=nothing
		  set objFolder=nothing
		  set objFiles=nothing 
End Function






'******************************************

'@CreationDate <02-mar-2014>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description	this method writes testcase results to summary file       
'@Documentation <param> and <param> will do….
'*********************************************

Function rptWriteSum(tname, tresult, iExecutionTime)		
		Dim strResultsPath,strResultsFolder,objFileSummary,objSummary
		strResources=gTestDir &"Resources"
		SummaryChart =strResultsFolder1 &"\SummaryChart.html"
		Set objFileSummary=CreateObject("scripting.filesystemobject")
		Set objSummary=objFileSummary.openTextFile(SummaryChart, 8, True)
		iTestCaseNumber = iTestCaseNumber+1	
		If  StrComp(tresult,0,1) = 0 or StrComp(tresult,"PASS",1) = 0  then		 
			
			objSummary.writeline "<tr class='content2'><td class='justified'><font color='#153e7e' size='1' face='arial'><b>"&iTestCaseNumber&"</b>"
			objSummary.writeline "</font></td><td class='justified'> <a href=.\TestCases\" & tname & ">"& Split(tName,".")(0) &"</a></td>"
			objSummary.writeline "<td class='justified'>Pass</td><td class='justified'>"&iExecutionTime&"</td></tr></tbody>" 
			
				
		Elseif  StrComp(tresult,1,1) = 0 or StrComp(tresult,"FAIL",1) = 0  then									
	
			objSummary.writeline "<tr class='content2'><td class='justified'><font color='#153e7e' size='1' face='arial'><b>"&iTestCaseNumber&"</b>"
			objSummary.writeline "</font></td><td class='justified'> <a href=.\TestCases\" & tname & ">"& Split(tName,".")(0) &"</a></td>"
			objSummary.writeline "<td class='justified'>Fail</td><td class='justified'>"&iExecutionTime&"</td></tr></tbody>" 
						
		End If
		
		Set objSummary = Nothing
		Set objFileSummary = Nothing 
End Function
'''#####################################################################################################################										--    
'@Function Name <rpt_ReportLog>
'@CreationDate <02-mar-2014>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description	-- This Function  useds to log the step details in log file  
'@Documentation <param> and <param> will do….

''#####################################################################################################################

Public Function rptReportLog (ByVal strStepName, ByVal strExpected,ByVal strStatus)
   Dim objFilesys,StrLogFloder,objFile
   set objFilesys = CreateObject("Scripting.FileSystemObject")
	If not objFilesys.FolderExists(strResultsPath) Then
		Set objFolder=objFilesys.CreateFolder(strResultsPath)
	End If
	
	If gFolderName=Empty Then
		rptFoldername
	End If
	 
   strResultsFolder=strResultsPath&"\"& gFolderName 
   strResultsFolder1 = strResultsFolder
   
   If not objFilesys.FolderExists(strResultsFolder) Then
	   Set objFolder=objFilesys.CreateFolder(strResultsFolder)
   End If
   	
	StrLogFloder=strResultsFolder1 & "\Logs\"
	
	If objFilesys.FolderExists(StrLogFloder)= False Then
			objFilesys.CreateFolder(StrLogFloder)
	End If
	
    If objFilesys.FileExists(StrLogFloder&"Log.txt")= false Then
		Set objFile=objFilesys.CreateTextFile(Trim(StrLogFloder)&"Log.txt")
		Set objFile=Nothing
    End if
     Set objFile = objFilesys.OpenTextFile(StrLogFloder&"Log.txt",8,True)
     objFile.WriteLine Now&vbtab&strStepName & vbTab & strExpected &  vbTab & Ucase(strStatus)
	Set objFile = Nothing
	Set objFilesys = Nothing

End Function


'''#####################################################################################################################										--    
'@Function Name <rptFilename>
'@CreationDate <15-mar-2014>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description	--   
'@Documentation <param> and <param> will do….

''#####################################################################################################################

sub rptFoldername 
Dim dDate,strdate,Filename
dDate=Now()
Foldername="Report_"&Month(dDate)&"-"&Day(dDate)&"-"&Year(dDate)&"-"&hour(dDate)&"-"&Minute(dDate)
gFolderName=Foldername	
End sub
