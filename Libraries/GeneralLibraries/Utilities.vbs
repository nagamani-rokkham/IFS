Option Explicit
''''#####################################################################################################################													
''@Class Name <DataManager>
''@CreationDate <03-Mar-2014>
''@Dependency  
''@Author 
''@ModifiedDate
''@ModifiedBy 
''@Description	 --   This class is used to retrieve data from excel sheet using ADODB object
''@Documentation <param> and <param> will do….
'''#####################################################################################################################
'@descriprion Manage all excel operation using the ole db model exposed my microsoft office provider
Class DataManager

	Private dContainer
	Private cRow, rCount
	
	'@description Default class variable values
	Sub Class_Initialize		
	    cRow = 0
	    rCount = 0
	End Sub

	'@description The function imports the data from a given excel sheet to be used by the test scripts
	'@documentation Import data from <file> [<sheet>] where combination_id field = <id>
	Function ImportData(ByVal file,ByVal sheet,ByVal id)
	
		'declarations
		Dim columnHeader, arrColumnValue, objDict
		Dim objExcel, objRS
		Dim i,max
		
		'constants
		Const adOpenForwardOnly = 0
		Const adOpenKeyset      = 1
		Const adOpenDynamic     = 2
		Const adOpenStatic      = 3
		
		'object declaration
		Set objDict = CreateObject("Scripting.Dictionary")
		Set objExcel = CreateObject("ADODB.Connection")		
		Set objRS = CreateObject("ADODB.Recordset")
		
		'Initiate connection
		objExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0;;Extended Properties=Excel 12.0;Data Source="&file&";"
		objRS.Open "Select * from ["&sheet&"$] where TestcaseId='"&id&"' And Execute = 'Y'", objExcel, adOpenStatic
		
		'read data and convert to 2D array
		i = 0
		max = objRS.Fields.count
		Do	
			columnHeader = objRS.Fields(i).Name
			Do	
				If objRS.RecordCount = 0 Then
					rptReportLog "Records not found", "Execute status should be 'Y'", "Execute status is 'N'", "FAIL"
					Exit Function
				End If
				arrColumnValue = objRS.GetRows(-1,1,i)
				objDict.Add columnHeader,arrColumnValue
				i = i + 1
			Loop While objRS.EOF <> true
			objRS.MoveFirst
		Loop While i < max
		
		Set dContainer = objDict
		rCount = objRS.RecordCount
		objRS.Close
		objExcel.Close
		
		Set objRS = Nothing
		Set objExcel = Nothing
		Set objDict = Nothing
		
	End Function
	

''''#####################################################################################################################													
''@Function Name <UpdateExcelSheet>
''@CreationDate <03-Mar-2014>
''@Dependency  
''@Author 
''@ModifiedDate
''@ModifiedBy 
''@Description	 --   This function is used to define test data path
''@Documentation <filepath,Sheetname,Scenariolist,updatecolumnslist in excel> and <> will do….
'''#####################################################################################################################

Function UpdateExcelSheet(ByVal file,ByVal sheet,ByVal ScenarioList,Byval UpDatedcolumnList)
	'declarations
	Dim values,ColumnName,ColumnVal,Scenarioids
	Dim columnHeader, objDict,UpdatedValue
	Dim objExcel, objRS
	Dim i,j
	'constants
	Const adOpenForwardOnly = 0
	Const adOpenKeyset      = 1
	Const adOpenDynamic     = 2
	Const adOpenStatic      = 3
	
	If Ucase(ScenarioList) = "ALL" Then''update all scenarios in sheet		
	
		'object declaration
		Set objDict = CreateObject("Scripting.Dictionary")
		Set objExcel = CreateObject("ADODB.Connection")		
		Set objRS = CreateObject("ADODB.Recordset")
		
		'Initiate connection
		objExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0;;Extended Properties=Excel 12.0;Data Source="&file&";"
		objRS.Open "Select * from ["&sheet&"$] where Execute='Y'", objExcel, adOpenStatic
		'split columnlist
		values = Split(UpDatedcolumnList,";;;")
		For i = 0 To UBound(values)			
		If Instr(1,Values(i),"|")>0 Then
				ColumnNameVal = Split(Values(i),"|")
				'columnname in Excel sheet
				ColumnName = ColumnNameVal(0)
				'value of global varible							
				ColumnVal = ColumnNameVal(1) 'Removed
				'update column
				objExcel.Execute "update ["&sheet&"$] Set " &ColumnName&"='"&ColumnVal&"'"  
		End If					
		Next 					
	
	ElseIf InStr(1,ScenarioList,"|")>0 Then		'more then one scenario update
		'object declaration
		Set objDict = CreateObject("Scripting.Dictionary")
		Set objExcel = CreateObject("ADODB.Connection")		
		Set objRS = CreateObject("ADODB.Recordset")
		'Initiate connection
		objExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0;;Extended Properties=Excel 12.0;Data Source="&file&";"
		objRS.Open "Select * from ["&sheet&"$] where ScenarioId='"&id&"' And Execute = 'Y'", objExcel, adOpenStatic
		Scenarioids = Split(ScenarioList,"|")
		For Each  Scenarioid  In Scenarioids
			'split columnlist			
			values = Split(UpDatedcolumnList,";;;")
			For i = 0 To UBound(values)			
				If Instr(1,Values(i),"|")>0 Then
						'split columnname and global varible name
						ColumnNameVal = Split(Values(i),"|")
						''column name of Excel sheet
						ColumnName = ColumnNameVal(0)
						 'value of global varible									
						ColumnVal = ColumnNameVal(1)	'Removed Eval
						'update the required scenarios
						objExcel.Execute "update ["&sheet&"$] Set " &ColumnName&"='"&ColumnVal&"' Where ScenarioId='"&Scenarioid&"'"
				End If					
			Next
		Next						
	Else
		'object declaration
		Set objDict = CreateObject("Scripting.Dictionary")
		Set objExcel = CreateObject("ADODB.Connection")		
		Set objRS = CreateObject("ADODB.Recordset")
		'Initiate connection
		objExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0;;Extended Properties=Excel 12.0;Data Source="&file&";"
		objRS.Open "Select * from ["&sheet&"$] where ScenarioId='"&id&"' And Execute = 'Y'", objExcel, adOpenStatic
		'split columnlist
		values = Split(UpDatedcolumnList,";;;")
		For i = 0 To UBound(values)			
			If Instr(1,Values(i),"|")>0 Then
					'split columnname and global varible name
					ColumnNameVal = Split(Values(i),"|")
					'column name of Excel Sheet
					ColumnName = ColumnNameVal(0)
					 'value of global varible									
					ColumnVal = ColumnNameVal(1)	'Removed Eval
					'update one scenario 
					objExcel.Execute "update ["&sheet&"$] Set " &ColumnName&"='"&ColumnVal&"' Where ScenarioId='"&ScenarioList&"'"
			End If					
		Next 
	End If	
	Set objRS = Nothing
	Set objExcel = Nothing
	Set objDict = Nothing
End Function


	'@description Set the row reference for the data extraction
	'@documentation Set <index> as the row reference
	Function SetCurrentRow(ByVal index)
		cRow = index - 1
	End Function
	
	'@description Returns the current row in the given data
	Function GetCurrentRow()
		GetCurrentRow = cRow + 1
	End Function
	
	'@description Returns the total rows in the given data
	Function GetRowCount()
		GetRowCount = rCount
	End Function
	
	'@description Return value from a specified column
	Function Value(ByVal columnName)
		Value=dContainer.Item(columnName)(0,cRow)
	End Function
	
	Private Sub Class_Terminate   ' Setup Terminate event.
      	Set dContainer = Nothing
    End Sub
    
End Class

Set dtScenario = New DataManager