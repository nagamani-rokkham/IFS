''**************************************************************************************************************************************************
''Function Name						:		 EnterText
''Function Description			    :		 Checks for the object existence, if enabled and sets the value in the object
''															 This function is generic for WebEdit, WebArea
''Function Parameters			    :		sValue is the value to be set in the object
''Author							:	    Gallop
''Date								:		3rd Aug 2014
''**************************************************************************************************************************************************
Public Function EnterText(ByRef test_object, ByRef sValue)
	On error resume next
	Call changeCreationTime (test_object)
	If test_object.exist(MaxWait) Then
		If test_object.GetROProperty("disabled") = False Then
			test_object.Set sValue
			If instr(1,test_object.GetROProperty("name"),"password",1) > 0 Then
				rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & "********" & chr(34) & " : value should be entered in " & chr(34) & test_object.ToString & chr(34) & " field."
			Else
				rptWriteReport "PASS", Environment.Value("StepName"),chr(34) & sValue & chr(34) & " : value should be entered in " & chr(34) & test_object.ToString & chr(34) &" field."
			End If
			EnterText = true
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"),test_object.ToString & " object should be enabled."
			EnterText = false
			gErrorFlag = True
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		EnterText = false
		gErrorFlag = True
	End If
	On error Goto 0
End Function
RegisterUserFunc "WebEdit", "EnterText", "EnterText"
RegisterUserFunc "WinEdit", "EnterText", "EnterText"

''**************************************************************************************************************************************************
''Function Name						 :	   ClickObject
''Function Description			     :	   Checks for the object existence, if enabled and Clicks the object
''										   This function is generic for Link, Image, WebElement, WebButton
''Function Parameters			     :	   NA
''Author					    	 :	   Gallop
''Date								 :	   4th Aug 2014
''**************************************************************************************************************************************************
Public Function ClickObject(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		If Instr(1,test_object.GetROProperty("Class Name"),"Win",1) > 0 Then
		
			If test_object.GetROProperty("enabled") = True Then
				Wait 1
				test_object.object.focus
				test_object.Click
				rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " object should be clicked."
				ClickObject = true
			Else
				rptWriteReport "FAIL", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " object should be enabled."
				ClickObject = false
				gErrorFlag = True
			End If
		ElseIf test_object.GetROProperty("disabled") = false Then
		Wait(3)
				test_object.Click
				rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " object Should be clicked."
				ClickObject = true
'			Else
'				rptWriteReport "FAIL", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " object should be enabled."
'				ClickObject = false
'				gErrorFlag = True
'			End If
		ElseIf Instr(1,test_object.GetROProperty("nativeclass"), "drop", 1) > 0 Then
		
			test_object.Click
			rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " object should be clicked."
			ClickObject = true			
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		ClickObject = false
		gErrorFlag = True
	End If
	On error Goto 0
End Function
RegisterUserFunc "Link", "ClickObject", "ClickObject"
RegisterUserFunc "Image", "ClickObject", "ClickObject"
RegisterUserFunc "WebButton", "ClickObject", "ClickObject"
RegisterUserFunc "WinButton", "ClickObject", "ClickObject"
RegisterUserFunc "WebElement", "ClickObject", "ClickObject"
RegisterUserFunc "WebEdit", "ClickObject", "ClickObject"
RegisterUserFunc "WebFile", "ClickObject", "ClickObject"
RegisterUserFunc "WebCheckBox", "ClickObject", "ClickObject"


''**************************************************************************************************************************************************
''Function Name						 :		VerifyExists
''Function Description			     :		Checks for the object existence
''											This function is generic for Browser, Page, WebEdit, WebElement, WebButton, WebList, WebRadioGroup, WebTable, WebCheckBox
''Function Parameters			     :		NA
''Author						   	 :	    Gallop
''Date								 :		4th Aug 2014
''**************************************************************************************************************************************************
Public Function VerifyExists(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " should be displayed."
		VerifyExists = true
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " should be displayed."
		VerifyExists = false
		gErrorFlag = True
	End If
	On error Goto 0
End Function
RegisterUserFunc "Browser", "VerifyExists", "VerifyExists"
RegisterUserFunc "Page", "VerifyExists", "VerifyExists"
RegisterUserFunc "Link", "VerifyExists", "VerifyExists"
RegisterUserFunc "Image", "VerifyExists", "VerifyExists"
RegisterUserFunc "WebList", "VerifyExists", "VerifyExists"
RegisterUserFunc "WebRadioGroup", "VerifyExists", "VerifyExists"
RegisterUserFunc "WebTable", "VerifyExists", "VerifyExists"
RegisterUserFunc "WebElement", "VerifyExists", "VerifyExists"
RegisterUserFunc "WebEdit", "VerifyExists", "VerifyExists"
RegisterUserFunc "WebButton", "VerifyExists", "VerifyExists"
RegisterUserFunc "WebCheckBox", "VerifyExists", "VerifyExists"
RegisterUserFunc "Static", "VerifyExists", "VerifyExists"


''**************************************************************************************************************************************************
''Function Name					:		WaitForObjectToLoad
''Function Description			:		Checks for the object existence, and then waits for the object to get loaded
''										This function is generic for Link,Image,Browser,Page,WebElement,WebEdit,WebButton,WebList,WebRadioGroup,WebTable
''Function Parameters			:		NA
''Author						:	   	Gallop
''Date							:		5th Aug 2014
''**************************************************************************************************************************************************
Public Function WaitForObjectToLoad(ByRef test_object)
	Call changeCreationTime (test_object)
	Wait 15
	On error resume Next
	If test_object.exist(MaxWait) Then
			rptWriteReport "PASS", Environment.Value("StepName"), test_object.ToString & "  should be loaded " 
			WaitForObjectToLoad = true 
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
	End If
	On error Goto 0
End Function
RegisterUserFunc "Link", "WaitForObjectToLoad", "WaitForObjectToLoad"
RegisterUserFunc "Image", "WaitForObjectToLoad", "WaitForObjectToLoad"
RegisterUserFunc "Browser", "WaitForObjectToLoad", "WaitForObjectToLoad"
RegisterUserFunc "Page", "WaitForObjectToLoad", "WaitForObjectToLoad"
RegisterUserFunc "WebElement", "WaitForObjectToLoad", "WaitForObjectToLoad"
RegisterUserFunc "WebEdit", "WaitForObjectToLoad", "WaitForObjectToLoad"
RegisterUserFunc "WebButton", "WaitForObjectToLoad", "WaitForObjectToLoad"
RegisterUserFunc "WebList", "WaitForObjectToLoad", "WaitForObjectToLoad"
RegisterUserFunc "WebRadioGroup", "WaitForObjectToLoad", "WaitForObjectToLoad"
RegisterUserFunc "WebTable", "WaitForObjectToLoad", "WaitForObjectToLoad"

''**************************************************************************************************************************************************
''Function Name				:		VerifyRadioButtonIsSelected
''Function Description		:		Checks for the object existence, if enabled and verify radiobutton is selected
''									This function is generic for WebradioGroup
''Function Parameters		:		strValue is the value of Radio Button
''Author					:	    Gallop
''Date						:		20th Feb 2015
''**************************************************************************************************************************************************
Public Function VerifyRadioButtonIsSelected(ByRef test_object,Byval strValue)

	If test_object.exist(MaxWait) Then
		Set odesc = Description.Create
		odesc("micclass").value ="WebRadioGroup"
		Set RadioGroup = test_object.ChildObjects(odesc)
		If RadioGroup(0).GetROProperty("checked") = 1 Then
			Select Case Ucase(strValue)
				Case "YES"
					If RadioGroup(0).GetROProperty("value") = "True" Then
						rptWriteReport "PASS", Environment.Value("StepName"), strValue & " radio button should be selected."
						VerifyRadioButtonIsSelected = True
					Else
						rptWriteReport "FAIL", Environment.Value("StepName"), strValue & " radio button should be selected."
						VerifyRadioButtonIsSelected = False
						gErrorFlag = True
					End If
				Case "NO"
					If RadioGroup(0).GetROProperty("value") = "False" Then
						rptWriteReport "PASS", Environment.Value("StepName"), strValue & " radio button should be selected."
						VerifyRadioButtonIsSelected = True
					Else
						rptWriteReport "FAIL", Environment.Value("StepName"), strValue & " radio button should be selected."
						VerifyRadioButtonIsSelected = False
						gErrorFlag = True
					End If
			End Select
		Else
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"), strValue & " radio button should be selected."
		End If			
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.GetROProperty("outertext") & " object should exist."
	End If
	On error Goto 0
End Function
RegisterUserFunc "WebElement", "VerifyRadioButtonIsSelected", "VerifyRadioButtonIsSelected"

''**************************************************************************************************************************************************
''Function Name				:		SelectListBox
''Function Description		:		Checks for the object existence, selects value from list box
''									This function is generic for WebElement
''Function Parameters		:		strValue is the value to selct froom list
''Author					:	    Gallop
''Date						:		20th Feb 2015
''**************************************************************************************************************************************************
Public Function SelectListBox(ByRef test_object, ByRef sValue)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		Set objDropdownDescription = Description.Create
		objDropdownDescription("micclass").value = "WebElement"
		objDropdownDescription("class").value = "k-input"
		set objDropdown = test_object.ChildObjects(objDropdownDescription)
		Wait(2)
		objDropdown(0).Click
		Wait(2)
		Set ObjDropdownListValueDescription = Description.Create
		ObjDropdownListValueDescription("micclass").value = "WebElement"
		ObjDropdownListValueDescription("html tag").value = "LI"
		ObjDropdownListValueDescription("visible").value = True
		ObjDropdownListValueDescription("innertext").value = sValue
		set objPage = test_object.GetTOProperty("parent")
		set objListValue = objPage.ChildObjects(ObjDropdownListValueDescription)
		If objListValue.Count = 1 Then
			objListValue(0).Click
			rptWriteReport "PASS", Environment.Value("StepName"), sValue & " should be selected from dropdown."			
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), sValue & " should be selected from dropdown."
			gErrorFlag = True
		End If	
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.GetROProperty("outertext") & " object should exist."
		gfSelectListBox = false
		gErrorFlag = True
	End If
	On error Goto 0
End Function

RegisterUserFunc "WebElement", "SelectListBox", "SelectListBox"

''**************************************************************************************************************************************************
''Function Name				:		VerifyMandatoryMessage
''Function Description		:		Checks for the object existence, verifies mandatory message displayed
''									This function is generic for WebElement
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		20th Feb 2015
''**************************************************************************************************************************************************

Function VerifyMandatoryMessage(ByRef test_object,Byval strMsg)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		If Instr(1,test_object.GetROProperty("outertext"),strMsg,1) <> 0 Then
			rptWriteReport "PASS", Environment.Value("StepName"), strMsg &  " message should be displayed at "& test_object.ToString
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), strMsg &  " message should be displayed at " & test_object.ToString
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function

RegisterUserFunc "WebElement", "VerifyMandatoryMessage", "VerifyMandatoryMessage"

''**************************************************************************************************************************************************
''Function Name				:		VerifyMessageNotExists
''Function Description		:		Checks for the object existence, verifies message is not displayed
''									This function is generic for WebElement
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		20th Feb 2015
''**************************************************************************************************************************************************

Function VerifyMessageNotExists(ByRef test_object,Byval strMsg)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		If Instr(1,test_object.GetROProperty("outertext"),strMsg,1) = 0 Then
			rptWriteReport "PASS", Environment.Value("StepName"), strMsg &  " message should disappear  at "& test_object.ToString
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), strMsg &  " message should disappear at "& test_object.ToString
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function

RegisterUserFunc "WebElement", "VerifyMessageNotExists", "VerifyMessageNotExists"
''**************************************************************************************************************************************************
''Function Name				:		VerifyDisplayedText
''Function Description		:		Checks for the object existence, verifies text displayed
''									This function is generic for WebElement
''Function Parameters		:		strValue is the text to be verified
''Author					:	    Gallop
''Date						:		20th Feb 2015
''**************************************************************************************************************************************************
Function VerifyDisplayedText(ByRef test_object,ByVal strValue)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		If Instr(1,test_object.GetROProperty("innertext"),strValue,1) <> 0 Then
			rptWriteReport "PASS", Environment.Value("StepName"), "Text: " & chr(34) & strValue & chr(34) & "  should be displayed at " & test_object.ToString
		Else
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"), "Text: " & chr(34) & strValue & chr(34) & "  should be displayed at " & test_object.ToString
		End If
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WebElement", "VerifyDisplayedText", "VerifyDisplayedText"

''**************************************************************************************************************************************************
''Function Name				:		ClickOnCalenderIcon
''Function Description		:		Checks for the object existence, Click on calender icon
''									This function is generic for WebElement
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		5th Mar 2015
''**************************************************************************************************************************************************

Public Function ClickOnCalenderIcon(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		Set objIcon = description.Create
		objIcon("micclass").value = "WebElement"
		objIcon("html tag").value = "SPAN"
		objIcon("class").value = ".*k-i-calendar"
		Set objCalendarIcon = test_object.childobjects(objIcon)
		If objCalendarIcon.Count = 1 Then
			objCalendarIcon(0).Click
			rptWriteReport "PASS", Environment.Value("StepName"), "cleandar icon should be clicked under " & test_object.ToString		
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "cleandar icon should be clicked under " & test_object.ToString
			gErrorFlag = True
		End If	
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		gErrorFlag = True
	End If
	On error Goto 0
End Function
RegisterUserFunc "WebElement", "ClickOnCalenderIcon", "ClickOnCalenderIcon"

''**************************************************************************************************************************************************
''Function Name				:		VerifyCalenderNavigation
''Function Description		:		Checks for the object existence, Check the calender Icons are disable/enable.
''									This function is generic for Link
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		5th Mar 2015
''**************************************************************************************************************************************************

Public Function VerifyCalenderNavigation(ByRef test_object,strCondation)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		strClass = test_object.GetROProperty("class")
		If ucase(strCondation) = "MOVABLE" Then
			If not Instr(1,strClass,"k-state-disabled",1)<>0 Then
				rptWriteReport "PASS", Environment.Value("StepName"), test_object.ToString & " should be enable"
			Else
				gErrorFlag = True
				rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " should be enable"	
			End If
		Else
			If Instr(1,strClass,"k-state-disabled",1)<>0 Then
				rptWriteReport "PASS", Environment.Value("StepName"), test_object.ToString & " should be disable"
			Else
				gErrorFlag = True
				rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " should be disable"	
			End If
		End If	
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.GetROProperty("outertext") & " object should exist."
		gErrorFlag = True
	End If
	On error Goto 0
End Function
RegisterUserFunc "Link", "VerifyCalenderNavigation", "VerifyCalenderNavigation"

''**************************************************************************************************************************************************
''Function Name				:		VerifyPreviousDateNotShownInCalendar
''Function Description		:		Checks for the object existence, Check Calendar pop-up widget should display only dates on or after the current date
''									This function is generic for WebTable
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		5th Mar 2015
''**************************************************************************************************************************************************

Public Function VerifyPreviousDateNotShownInCalendar(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		bflag = False
		bstartFlag = False
		bPassFlag = True
		ir = test_object.RowCount
		For iRow = 2 To ir Step 1
			ic = test_object.ColumnCount(iRow)
			For jCol = 1 To ic Step 1
				strcelldata = test_object.GetCellData(iRow,jCol)
				If len(Trim(strcelldata))>0 and bstartFlag = False Then
					strstartday = test_object.ChildItem(iRow,jCol,"Link",0).GetROProperty("title")
					strstartday = mid (strstartday,instr(strstartday,",")+1)
					If Day(CDate(Trim(strstartday))) = Day(Date) Then
						bstartFlag = True
					Else
						bPassFlag = False
						bflag = true
					    Exit for 
					End If
				End If
				If bstartFlag = True and len(Trim(strcelldata)) = 0 Then
					bPassFlag = False
					bflag = true
				    Exit for 
				End If				
			Next
			If bflag = true Then
				Exit For
			End If
		Next
		If bPassFlag Then
			rptWriteReport "Pass", Environment.Value("StepName"), "Calendar pop-up widget should display only dates on or after the current date"
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Calendar pop-up widget should display only dates on or after the current date"		
			gErrorFlag = True
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		gErrorFlag = True
	End If	
	On error Goto 0	
End Function
RegisterUserFunc "WebTable", "VerifyPreviousDateNotShownInCalendar", "VerifyPreviousDateNotShownInCalendar"


''**************************************************************************************************************************************************
''Function Name				:		VerifyToolTip
''Function Description		:		Checks for the object existence, verifies text displayed in tool tip
''									This function is generic for WebElement,WebEdit
''Function Parameters		:		strToolTip is the message displayed in tool tip
''Author					:	    Gallop
''Date						:		20th Feb 2015
''**************************************************************************************************************************************************
Function VerifyToolTip(ByRef test_object,strToolTip)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		If Instr(1,test_object.GetROProperty("title"),strToolTip,1) <> 0 Then
			rptWriteReport "PASS", Environment.Value("StepName"),chr(34) & strToolTip & chr(34) &  "should be displayed as tool tip for " & chr(34) & test_object.ToString & chr(34)
		Else
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"),chr(34) & strToolTip & chr(34) &  "should be displayed as tool tip for " & chr(34) & test_object.ToString & chr(34)
		End If
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WebEdit", "VerifyToolTip", "VerifyToolTip"
RegisterUserFunc "WebElement", "VerifyToolTip", "VerifyToolTip"


''**************************************************************************************************************************************************
''Function Name				:		VerifyMaxLength
''Function Description		:		Checks for the object , verifies maximum length allowed in text box
''									This function is generic for WebEdit
''Function Parameters		:		iLength,strType
''Author					:	    Gallop
''Date						:		9th Feb 2015
''**************************************************************************************************************************************************
Function VerifyMaxLength(ByRef test_object,Byval iLength,Byval strType)
	Call changeCreationTime (test_object)
	On error resume next
	Dim strValue
	Select Case Ucase(strType)
	Case "ALPHABETIC"
		strValue = GenerateRandomString(iLength+1,"ALPHABETIC")
	Case "NUMERIC"
		strValue = GenerateRandomString(iLength+1,"NUMERIC")
	Case "ALPHANUMERIC"
		strValue = GenerateRandomString(iLength+1,"ALPHANUMERIC")
	End Select
	If test_object.exist(MaxWait) Then
		test_object.Set strValue
		If Err.Number<>0 Then
			Err.Clear
			rptWriteReport "PASS", Environment.Value("StepName"), "Verify maximum length of " & test_object.ToString &",Try to enter "& iLength+1 &" " & strType &" characters. It should allow the " & iLength & " characters only"
		Else
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"), "Verify maximum length of " & test_object.ToString &",Try to enter "& iLength+1 &" " & strType &" characters. It should allow the " & iLength & " characters only"
		End If
	Else	
		gErrorFlag = True	
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WebEdit", "VerifyMaxLength", "VerifyMaxLength"



''**************************************************************************************************************************************************
''Function Name				:		ClickSpellCheckIcon
''Function Description		:		Checks for the object existence, clicks on spell check icon
''									This function is generic for WebElement
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		10th March 2015
''**************************************************************************************************************************************************
Function ClickSpellCheckIcon(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		Set objIcon = description.Create
		objIcon("micclass").value = "WebElement"
		objIcon("html tag").value = "SPAN"
		objIcon("class").value = ".*icon-check"
		Set objSpellCheckIcon = test_object.childobjects(objIcon)
		If objSpellCheckIcon.Count = 1 Then
			Wait 5
			objSpellCheckIcon(0).Click
			rptWriteReport "PASS", Environment.Value("StepName"), "Spell check icon should be clicked under " & test_object.ToString
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Spell check icon should be clicked under " & test_object.ToString
			gErrorFlag = True
		End If
	Else
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WebElement", "ClickSpellCheckIcon","ClickSpellCheckIcon"

''**************************************************************************************************************************************************
''Function Name				:		VerifySpellCheck
''Function Description		:		Checks for the object existence,verifies and selects correct word.
''									This function is generic for WebElement
''Function Parameters		:		strText
''Author					:	    Gallop
''Date						:		10th March 2015
''**************************************************************************************************************************************************
Function VerifySpellCheck(ByRef test_object,Byval strText)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		If Not(Len(Trim(test_object.getroproperty("all items")) > 0)) Then
			Wait 5
		End If	
			If Instr(1,test_object.getroproperty("all items"),strText,1) <> 0 Then
				rptWriteReport "PASS", Environment.Value("StepName"),chr(34) & strText&" should be displayed in suggestions list"
			Else
				rptWriteReport "FAIL", Environment.Value("StepName"),chr(34) & strText* " should be displayed in suggestions list"
				gErrorFlag = True
			End If
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WebList", "VerifySpellCheck", "VerifySpellCheck"

''**************************************************************************************************************************************************
''Function Name				:		VerifySpecialCharacterNotAllowed
''Function Description		:		Checks for the object existence, checks that special character is not allowed to enter
''									This function is generic for WebEdit
''Function Parameters		:		strText
''Author					:	    Gallop
''Date						:		9th march 2015
''**************************************************************************************************************************************************
Function VerifySpecialCharacterNotAllowed(ByRef test_object,strText)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		test_object.object.focus
		test_object.click
		Wait 2
		Set oshell = createobject("Wscript.shell")
		oshell.sendkeys strText
		oshell.sendkeys "{TAB}"
		wait(1)
		If Trim(test_object.GetROProperty("value")) <> strText And len(test_object.GetROProperty("value"))>0 Then
			rptWriteReport "PASS", Environment.Value("StepName"), "Verify special characters are not allowed in" & test_object.ToString &". Try to enter "& strText&" String.It should not allow special characters "
		Else
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"), "Verify special characters are not allowed in" & test_object.ToString &". Try to enter "& strText&" String.It should not allow special characters "
		End If
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WebEdit", "VerifySpecialCharacterNotAllowed", "VerifySpecialCharacterNotAllowed"

''**************************************************************************************************************************************************
''Function Name				:		VerifyAlphaCharacterNotAllowed
''Function Description		:		Checks for the object existence, checks that alpha characters are not allowed to enter
''									This function is generic for WebEdit
''Function Parameters		:		strText
''Author					:	    Gallop
''Date						:		9th march 2015
''**************************************************************************************************************************************************
Function VerifyAlphaCharacterNotAllowed(ByRef test_object,strText)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		test_object.focus
		test_object.click
		Wait 2
		Set oshell = createobject("Wscript.shell")
		oshell.sendkeys strText
		oshell.sendkeys "{TAB}"
		wait(1)
		If Trim(test_object.GetROProperty("value")) <> strText and Len(test_object.GetROProperty("value"))> 0 Then
			rptWriteReport "PASS", Environment.Value("StepName"), "Verify Alpha characters are not allowed in" & test_object.ToString &". Try to enter: "& strText&" .It should not allow alpha characters "
		Else
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"), "Verify Alpha characters are not allowed in" & test_object.ToString &". Try to enter: "& strText&" .It should not allow alpha characters "
		End If
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WebEdit", "VerifyAlphaCharacterNotAllowed", "VerifyAlphaCharacterNotAllowed"
 
''**************************************************************************************************************************************************
''Function Name				:		VerifyFieldIsReadOnly
''Function Description		:		Checks for the object existence, checks whether field is readonly
''									This function is generic for WebElement
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		9th March 2015
''**************************************************************************************************************************************************
Function VerifyFieldIsReadOnly(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		Set objState = description.Create
		objState("micclass").value = "WebElement"
		objState("class").value = ".*.*disabled"
		Set objStatic = description.Create
		'objStatic("micclass").value = "WebElement"
		objStatic("class").value = ".*readonly"
		Set objReadOnly = test_object.childobjects(objStatic)
		Set objDisabled = test_object.childobjects(objState)
		If objReadOnly.Count = 1 Then
			rptWriteReport "PASS", Environment.Value("StepName"), "Verify " & chr(34) & test_object.ToString & chr(34) & "is read only"
		Elseif objDisabled.Count = 1 then
			rptWriteReport "PASS", Environment.Value("StepName"), "Verify " & chr(34) & test_object.ToString & chr(34) & "is read only"
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Verify " & chr(34) & test_object.ToString & chr(34) & "is read only"
			gErrorFlag = True
		End If
	Else
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WebEdit", "VerifyFieldIsReadOnly", "VerifyFieldIsReadOnly"
RegisterUserFunc "WebElement", "VerifyFieldIsReadOnly","VerifyFieldIsReadOnly"

''**************************************************************************************************************************************************
''Function Name                      :        VerifyPageName
''Function Description               :        Checks for the page name
''                                            This function is generic for WebElement
''Function Parameters                :        NA
''Author                             :        Gallop
''Date                               :        09th Mar 2015
''**************************************************************************************************************************************************
Public Function VerifyPageName(ByRef test_object)
    Call changeCreationTime (test_object)
    On error resume next
	Wait 2
    set objPage = test_object.GetTOProperty("parent")
	objPage.sync
	Wait 2
    If test_object.exist(MaxWait) Then
        rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & objPage.GetTOProperty("title") & chr(34) & " page should be displayed."
    Else
        rptWriteReport "FAIL", Environment.Value("StepName"), chr(34) & objPage.GetTOProperty("title") & chr(34) & " page should be displayed."
        gErrorFlag = True
    End If
    On error Goto 0
End Function

RegisterUserFunc "WebElement", "VerifyPageName", "VerifyPageName"

''**************************************************************************************************************************************************
''Function Name				:		SelectRadioButton
''Function Description		:		Checks for the object existence, if enabled and selects the radiobutton
''									This function is generic for WebElement
''Function Parameters		:		strValue - Radio button value
''Author					:	    Gallop
''Date						:		12th March 2015
''**************************************************************************************************************************************************
Public Function SelectRadioButton(ByRef test_object,Byval strValue)
	Call changeCreationTime (test_object)
	If test_object.exist(MaxWait) Then
		Set odesc = Description.Create
		odesc("micclass").value ="WebRadioGroup"
		Set RadioGroup = test_object.ChildObjects(odesc)
		If RadioGroup.count = 0 Then
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"), test_object.GetROProperty("innertext") & " RadioGroup."
		Else
			Select Case Ucase(strValue)
				Case "YES"
					If (Instr(RadioGroup(0).GetROProperty("all items"),"Yes"))<>0 Then
						RadioGroup(0).Select "Yes"
						rptWriteReport "PASS", Environment.Value("StepName"), "select  " & strValue & " Radio button for " & chr(34) & test_object.ToString & chr(34) 
					Else
						If Instr(RadioGroup(0).GetROProperty("all items"),"True") Then
							RadioGroup(0).Select "True"
							rptWriteReport "PASS", Environment.Value("StepName"), "select  " & strValue & " Radio button for " & chr(34) & test_object.ToString & chr(34) 
						End If							
					End If
					
				Case "NO"
					If (Instr(RadioGroup(0).GetROProperty("all items"),"No")) <>0 Then
						RadioGroup(0).Select "No"
						rptWriteReport "PASS", Environment.Value("StepName"), "select  " & strValue & " Radio button for " & chr(34) & test_object.ToString & chr(34) 
					Else
						If Instr(RadioGroup(0).GetROProperty("all items"),"False") Then
							RadioGroup(0).Select "False"
							rptWriteReport "PASS", Environment.Value("StepName"), "select  " & strValue & " Radio button for " & chr(34) & test_object.ToString & chr(34) 
						End If
					End If	

				Case Else
					If (Instr(RadioGroup(0).GetROProperty("all items"),strValue)) <> 0 Then
						RadioGroup(0).Select strValue
						rptWriteReport "PASS", Environment.Value("StepName"), "select  " & strValue & " Radio button for " & chr(34) & test_object.ToString & chr(34) 
					Else
						gErrorFlag = True
						rptWriteReport "FAIL", Environment.Value("StepName"), chr(34) & test_object.ToString  &  chr(34)  & " object should exist."
					End If					
				End Select			
		End If
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), chr(34) & test_object.ToString  &  chr(34)  & " object should exist."
	End If
	
End Function
RegisterUserFunc "WebElement", "SelectRadioButton", "SelectRadioButton"
''**************************************************************************************************************************************************
''Function Name				:		SelectListBoxByIndex
''Function Description		:		Checks for the object existence, and select the value from drop down by using index number
''									This function is generic for WebElement
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		12th Mar 2015
''**************************************************************************************************************************************************
Public Function SelectListBoxByIndex(ByRef test_object, Byval iindex)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		Set objDropdownDescription = Description.Create
		objDropdownDescription("micclass").value = "WebElement"
		objDropdownDescription("class").value = "k-input"
		set objDropdown = test_object.ChildObjects(objDropdownDescription)
		Wait(2)
		If objDropdown.count = 1 Then
			objDropdown(0).Click
			Wait(2)			
			Set objlist = Description.Create
			objlist("micclass").value = "WebElement"
			objlist("html tag").value = "DIV"
			objlist("visible").value = True
			objlist("html id").value = ".*-list"
			set objPage = test_object.GetTOProperty("parent")
			Set list = objPage.ChildObjects(objlist)
			If list.count =1 Then
				Set ObjDropdownListValueDescription = Description.Create
				ObjDropdownListValueDescription("micclass").value = "WebElement"
				ObjDropdownListValueDescription("html tag").value = "LI"
				ObjDropdownListValueDescription("visible").value = True
				Set options = list(0).childobjects(ObjDropdownListValueDescription)		
				If options.count>0 Then
					If len(Trim(options(i).GetROproperty("innertext"))) > 0 Then
						iindex = iindex-1
					End If		
					If Not iindex > options.count-1 Then
						options(iindex).click
						rptWriteReport "PASS", Environment.Value("StepName"), "select the from " & chr(34) & test_object.ToString & chr(34)  & " dropdown. by using index ::" &iindex
						SelectListBoxByIndex = false
					Else
						rptWriteReport "FAIL", Environment.Value("StepName"), "Index value exceed the dropdown values in " &chr(34) & test_object.ToString & chr(34) & " dropdown."
						SelectListBoxByIndex = false
						gErrorFlag = True					
					End If			
				Else	
					rptWriteReport "FAIL", Environment.Value("StepName"), "Dropdown options should shown under " & chr(34) & test_object.ToString & chr(34)  & " dropdown."
					SelectListBoxByIndex = false
					gErrorFlag = True
				End If				
	
			Else
				rptWriteReport "FAIL", Environment.Value("StepName"), "Dropdown options should shown under " & chr(34) & test_object.ToString & chr(34)  &" dropdown."
				SelectListBoxByIndex = false
				gErrorFlag = True			
			End If
			
			
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Dropdown should shown under " & chr(34) & test_object.ToString & chr(34)  & " object."
			SelectListBoxByIndex = false
			gErrorFlag = True		
		End If
	
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"),  chr(34) & test_object.ToString & chr(34)  & " object should exist."
		SelectListBoxByIndex = false
		gErrorFlag = True
	End If
	On error Goto 0
End Function	

RegisterUserFunc "WebElement", "SelectListBoxByIndex", "SelectListBoxByIndex"

''**************************************************************************************************************************************************
''Function Name				:		VerifyValueSelectedInDropDown
''Function Description		:		Checks for the object existence, and verify the selected value in dropdown
''									This function is generic for WebElement
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		12th Mar 2015
''**************************************************************************************************************************************************

Public Function VerifyValueSelectedInDropDown(ByRef test_object,ByVal strValue)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		Set objdropdown = description.Create
		objdropdown("micclass").value = "WebElement"
		objdropdown("class").value = "k-input"
		Set objSelectedValue = test_object.childobjects(objdropdown)
		If objSelectedValue.Count = 1 Then
			If Trim(objSelectedValue(0).GetROProperty("innertext")) = strValue Then
				rptWriteReport "PASS", Environment.Value("StepName"),chr(34) & strValue & chr(34) & " should be displayed as selected value in dropdown : " & chr(34) & test_object.ToString & chr(34)		
			Else
				rptWriteReport "FAIL", Environment.Value("StepName"),chr(34) & strValue & chr(34) & " should be displayed as selected value in dropdown : " & chr(34) & test_object.ToString & chr(34)		
				gErrorFlag = True
			End If
		Else
			rptWriteReport "FAIl", Environment.Value("StepName"),"Dropdown should be displayed for " & chr(34) & test_object.ToString & chr(34)
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		gErrorFlag = True
	End If
	On error Goto 0
End Function
RegisterUserFunc "WebElement", "VerifyValueSelectedInDropDown", "VerifyValueSelectedInDropDown"

''**************************************************************************************************************************************************
''Function Name				:		VerifyYellowAlertNotExist
''Function Description		:		Checks for the object existence, checks whether field is not highlighted in yellow background
''									This function is generic for WebElement
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		13th March 2015
''**************************************************************************************************************************************************
Function VerifyYellowAlertNotExist(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		Set objAlert = description.Create
		objAlert("micclass").value = "WebElement"
		objAlert("class").value = ".*alert-warning"
		Set objYellowAlert = test_object.childobjects(objAlert)
		If objYellowAlert.Count = 0 Then
			rptWriteReport "PASS", Environment.Value("StepName"), "Verify yellow alert icon is not displayed for  " & chr(34) & test_object.ToString & chr(34)
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Verify yellow alert icon is not displayed for  " & chr(34) & test_object.ToString & chr(34)
			gErrorFlag = True
		End If
	Else
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WebElement", "VerifyYellowAlertNotExist","VerifyYellowAlertNotExist"


''**************************************************************************************************************************************************
''Function Name				:		VerifyYellowAlert
''Function Description		:		Checks for the object existence, checks whether field is highlighted in yellow background
''									This function is generic for WebElement
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		13th March 2015
''**************************************************************************************************************************************************
Function VerifyYellowAlert(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		Set objAlert = description.Create
		objAlert("micclass").value = "WebElement"
		objAlert("class").value = ".*alert-warning"
		Set objYellowAlert = test_object.childobjects(objAlert)
		If objYellowAlert.Count = 1 Then
			rptWriteReport "PASS", Environment.Value("StepName"), "Verify yellow alert icon is displayed for  " & chr(34) & test_object.ToString & chr(34)
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Verify yellow alert icon is displayed for  " & chr(34) & test_object.ToString & chr(34)
			gErrorFlag = True
		End If
	Else
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WebElement", "VerifyYellowAlert","VerifyYellowAlert"

''**************************************************************************************************************************************************
''Function Name				:		VerifyValuesInRow
''Function Description		:		Checks for the object existence, verifies the values in row
''									
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		5th Mar 2015
''**************************************************************************************************************************************************
Public Function VerifyValuesInRow(ByRef test_object,ByVal strValue)
	Call changeCreationTime (test_object)
	On error resume next
	Dim iFoundRow,bResultFlag
	bResultFlag = False
	iFoundRow = 0
	arrvalues = split(strvalue,"|")
	If test_object.exist(MaxWait) Then
		ir = test_object.RowCount
		For iRow = 1 To ir Step 1
			ic = test_object.ColumnCount(iRow)
			For jCol = 1 To ic Step 1
				strcelldata = test_object.GetCellData(iRow,jCol)
				If Trim(Ucase(arrvalues(0))) = Trim(Ucase(strcelldata)) Then
					iFoundRow = iRow
					bResultFlag = True
					Exit for	
				ElseIf InStr(1,Trim(Ucase(strcelldata)),Trim(Ucase(arrvalues(0))),1) > 0 Then
					iFoundRow = iRow
					bResultFlag = True
					Exit for		
				End If
			Next
			If iFoundRow > 0 Then
				Exit For
			End If
		Next
		If ubound(arrvalues)>0 and iFoundRow > 0 Then
			For icols = 1 To ubound(arrvalues) Step 1
				bfound = False
				ic = test_object.ColumnCount(iFoundRow)
				For itablecols = 1 To ic Step 1
					strcelldata = test_object.GetCellData(iFoundRow,itablecols)
					If Ucase(Trim(strcelldata)) = UCase(Trim(arrvalues(icols))) Then
						bfound = True
						Exit For
					End If
				Next
				If bfound = True Then
					bResultFlag = True
				Else
					bResultFlag = False					
				End If
			Next
		End If
		If bResultFlag = True Then
			rptWriteReport "PASS", Environment.Value("StepName"), arrvalues(1) &" Value should be verified under " &arrvalues(0)& test_object.ToString		
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), arrvalues(1) &" Value is not verified under " & arrvalues(0)&test_object.ToString		
		End If
	Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Object doesnt Exist: " & test_object.ToString	
	End If
End Function
RegisterUserFunc "webtable", "VerifyValuesInRow", "VerifyValuesInRow"
	
''**************************************************************************************************************************************************
''Function Name				:		verifyNumberexistIncellByText
''Function Description		:		Checks for the object existence, checks whether the cell data is numeric
''									
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		5th Mar 2015
''**************************************************************************************************************************************************
Public Function verifyNumberexistIncellByText(ByRef test_object,ByVal strValue)
	Call changeCreationTime (test_object)
	On error resume next
	Dim iFoundRow,bResultFlag
	bResultFlag = False
	iFoundRow = 0
	If test_object.exist(MaxWait) Then
		ir = test_object.RowCount
		For iRow = 1 To ir Step 1
			ic = test_object.ColumnCount(iRow)
			For jCol = 1 To ic Step 1
				strcelldata = test_object.GetCellData(iRow,jCol)
				If Ucase(strValue) = Ucase(strcelldata) Then
					iFoundRow = iRow
					Exit for
				End If
			Next
			If iFoundRow > 0 Then
				Exit For
			End If
		Next
		If ubound(arrvalues)>0 and iFoundRow > 0 Then
			ic = test_object.ColumnCount(iFoundRow)
			For itablecols = 1 To ic Step 1
				strcelldata = test_object.GetCellData(iFoundRow,itablecols)
				If IsNumeric(strcelldata) Then
					bResultFlag = True
					Exit For
				End If
			Next
		End If
		If bResultFlag = True Then
			rptWriteReport "PASS", Environment.Value("StepName"), strcelldata&"Value should be verified under " & test_object.ToString		
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), strcelldata&"Value is not verified under " & test_object.ToString		
		End If
	Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Object doesnt Exist: " & test_object.ToString	
	End If		
End Function
RegisterUserFunc "webtable", "verifyNumberexistIncellByText", "verifyNumberexistIncellByText"

''**************************************************************************************************************************************************
''Function Name				:		VerifyClassName
''Function Description		:		Checks for the object existence, verifies the class 
''									This function is generic for WebElement
''Function Parameters		:		strValue is the text to be verified
''Author					:	    Gallop
''Date						:		20th Feb 2015
''**************************************************************************************************************************************************
Function VerifyClassName(ByRef test_object,ByVal strValue)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		If Instr(1,test_object.GetROProperty("class"),strValue,1) <> 0 Then
			rptWriteReport "PASS", Environment.Value("StepName"),  "  It is  displayed at " & test_object.ToString
		Else
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"), "  It is not displayed at " & test_object.ToString
		End If
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WebElement", "VerifyClassName", "VerifyClassName"

''**************************************************************************************************************************************************
''Function Name				:		WaitinSec
''Function Description		:	    Static Wait
''									This function is Static wait for WebElement
''Function Parameters		:		strValue is the text to be verified
''Author					:	    Gallop
''Date						:		20th Feb 2015
''**************************************************************************************************************************************************
Function WaitinSec(ByRef test_object,ByVal TimeSec)
	Call changeCreationTime (test_object)
	If test_object.Exist Then
		Wait(TimeSec)
	End If
End Function
RegisterUserFunc "WebEdit", "WaitinSec", "WaitinSec"
RegisterUserFunc "winbutton", "WaitinSec", "WaitinSec"

''**************************************************************************************************************************************************
''Function Name				:		VerifyandClickCellsInRow
''Function Description		:		Checks for the object existence, verifies the values in row and clicks on the attachment
''									
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		5th Mar 2015
''**************************************************************************************************************************************************
    Public Function VerifyandClickLinkInRow(ByRef test_object,ByVal strValue)
	Call changeCreationTime (test_object)
	On error resume next
	Dim iFoundRow,bResultFlag
	bResultFlag = False
	iFoundRow = 0
	arrvalues = split(strvalue,"|")
	If test_object.exist(MaxWait) Then
		ir = test_object.RowCount
		For iRow = 1 To ir Step 1
			ic = test_object.ColumnCount(iRow)
			For jCol = 1 To ic Step 1
				strcelldata = test_object.GetCellData(iRow,jCol)
				If Trim(UCase(arrvalues(0))) = Trim(Ucase(strcelldata)) Then
					iFoundRow = iRow
					bResultFlag = True
					Exit for
				End If
			Next
			If iFoundRow > 0 Then
				Exit For
			End If
		Next
		If iFoundRow > 0 Then
			For icols = 1 To test_object.ColumnCount(iFoundRow) Step 1
				bfound = False	
				itemcount = test_object.childitemcount(iFoundRow,icols,"WebElement")					
				If itemcount > 0 Then
					For i  = 0 To itemcount-1 Step 1
						Set objIcon = test_object.childitem(iFoundRow,icols,"WebElement",i)
						Val = objIcon.GetRoProperty("title")
						If Trim(Val) = Trim(arrvalues(1)) Then
							objIcon.click
							click = True
						End If
					Next	
				End If				
			Next
			If click Then
				rptWriteReport "PASS", Environment.Value("StepName"), Val& " icon should be clicked under " & test_object.ToString		
			Else
			    rptWriteReport "FAIL", Environment.Value("StepName"),Val& "icon should be clicked under " & test_object.ToString
				gErrorFlag = True
			End If
		End If
	End If
	End Function
	RegisterUserFunc "webtable", "VerifyandClickLinkInRow", "VerifyandClickLinkInRow"
	
	''**************************************************************************************************************************************************
''Function Name						:		 EnterFileName
''Function Description			    :		 Checks for the object existence, if enabled and sets the fileName in the object
''															 This function is generic for WebEdit, WebArea
''Function Parameters			    :		sValue is the value to be set in the object
''Author							:	    Gallop
''Date								:		3rd Aug 2014
''**************************************************************************************************************************************************
Public Function EnterFileName(ByRef test_object,ByVal strfilename)
	Call changeCreationTime (test_object)
	On error resume Next
	Filename = sBaseDir  & "Resources\" & strfilename
	If test_object.exist(MaxWait) Then
		If test_object.GetROProperty("disabled") = False Then
			test_object.Set Filename			
			rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & Filename & chr(34) & " : value should be entered in " & chr(34) & test_object.ToString & chr(34) & " field."
			EnterFileName = True
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"),test_object.ToString & " object should be enabled."
			EnterFileName = false
			gErrorFlag = True
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		EnterText = false
		gErrorFlag = True
	End If
	On error Goto 0
End Function
RegisterUserFunc "WebEdit", "EnterFileName", "EnterFileName"
RegisterUserFunc "WinEdit", "EnterFileName", "EnterFileName"



''**************************************************************************************************************************************************
''Function Name				:		SelectWebList
''Function Description		:		Checks for the object existence, selects value from weblist
''									This function is generic for WebList
''Function Parameters		:		strValue
''Author					:	    Gallop
''Date						:		20th Feb 2015
''**************************************************************************************************************************************************
Public Function SelectWebList(ByRef test_object, Byval strValue)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		If Instr(1,test_object.GetROProperty("all items"),strValue,1) <> 0 Then
			test_object.Select strValue
			rptWriteReport "PASS", Environment.Value("StepName"), strValue & " should be selected from dropdown."			
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), strValue & " should be selected from dropdown."
			gErrorFlag = True
		End If	
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.GetROProperty("outertext") & " object should exist."
		gErrorFlag = True
	End If
	On error Goto 0
End Function

RegisterUserFunc "WebList", "SelectWebList", "SelectWebList"


''**************************************************************************************************************************************************
''Function Name						 :		VerifyNotExists
''Function Description			     :		Checks for the object existence, it should not be displayed
''											This function is generic for Browser, Page, WebEdit, WebElement, WebButton, WebList, WebRadioGroup, WebTable, WebCheckBox
''Function Parameters			     :		NA
''Author						   	 :	    Gallop
''Date								 :		4th Aug 2014
''**************************************************************************************************************************************************
Public Function VerifyNotExists(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume Next
	visible = test_object.GetRoProperty("visible")
	If Not visible  Then
		rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " is not  displayed."
		VerifyNotExists = true
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " is displayed."
		VerifyNotExists = false
		gErrorFlag = True
	End If
	On error Goto 0
End Function

RegisterUserFunc "WebEdit", "VerifyNotExists", "VerifyNotExists"
RegisterUserFunc "WebElement", "VerifyNotExists", "VerifyNotExists"

''**************************************************************************************************************************************************
''Function Name				:		VerifyLinkInRow
''Function Description		:		Checks for the object existence, verifies the values in row and validates the attachment
''									
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		5th Mar 2015
''**************************************************************************************************************************************************
    Public Function VerifyLinkInRow(ByRef test_object,ByVal strValue)
	Call changeCreationTime (test_object)
	On error resume next
	Dim iFoundRow,bResultFlag
	bResultFlag = False
	iFoundRow = 0
	arrvalues = split(strvalue,"|")
	If test_object.exist(MaxWait) Then
		ir = test_object.RowCount
		For iRow = 1 To ir Step 1
			ic = test_object.ColumnCount(iRow)
			For jCol = 1 To ic Step 1
				strcelldata = test_object.GetCellData(iRow,jCol)
				If Trim(Ucase(arrvalues(0))) = Trim(Ucase(strcelldata)) Then
					iFoundRow = iRow
					bResultFlag = True
					Exit for
				End If
			Next
			If iFoundRow > 0 Then
				Exit For
			End If
		Next
		If iFoundRow > 0 Then
			For icols = 1 To test_object.ColumnCount(iFoundRow) Step 1
				bfound = False	
				itemcount = test_object.childitemcount(iFoundRow,icols,"WebElement")					
				If itemcount > 0 Then
					For i  = 0 To itemcount-1 Step 1
						Set objIcon = test_object.childitem(iFoundRow,icols,"WebElement",i)
						Val = objIcon.GetRoProperty("title")
						If Trim(Val) = Trim(arrvalues(1)) Then
							visible = True
						End If
					Next	
				End If				
			Next
			If visible Then
				rptWriteReport "PASS", Environment.Value("StepName"), strValue&" icon should be displayed under " & test_object.ToString		
			Else
			    rptWriteReport "FAIL", Environment.Value("StepName"),strValue & "icon should be displayed under " & test_object.ToString
				gErrorFlag = True
			End If
		End If
	End If
	End Function
	RegisterUserFunc "webtable", "VerifyLinkInRow", "VerifyLinkInRow"
	
'''#####################################################################################################################
'@Function Name <VerifyCurrentDate>
'@CreationDate <16/03/2015>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description  this function return the Day after Tomorrow 
'@Documentation No input parameter
'''#####################################################################################################################
Public Function VerifyCurrentDate(ByRef test_object,ByVal strValue)
    Call changeCreationTime (test_object)
    On error resume next
	Dim bFlag,bResultFlag
	bFlag = False
	bResultFlag = False
	If test_object.exist(MaxWait) Then
		ir = test_object.RowCount
		For iRow = 1 To ir Step 1
			ic = test_object.ColumnCount(iRow)
			For jCol = 1 To ic Step 1
				strcelldata = test_object.GetCellData(iRow,jCol)
				If Ucase(strValue) = Ucase(strcelldata) Then
					strcelldata = test_object.GetCellData(iRow,jCol+1)
					If CDate(strcelldata) = Date Then
						bResultFlag = True
						bFlag = True
						Exit For
					End If
				End If
			Next
			If bFlag Then
				Exit For
			End If
		Next
		If bResultFlag = True Then
			rptWriteReport "PASS", Environment.Value("StepName"), arrvalues(0)&"Value has displayed Today Date "&CDate(strcelldata) & test_object.ToString		
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), arrvalues(0)&"Value has not  displayed Today Date "&CDate(strcelldata) &test_object.ToString		
		End If
	Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Object doesnt Exist: " & test_object.ToString	
	End If
	
    End Function

RegisterUserFunc "webtable", "VerifyCurrentDate", "VerifyCurrentDate"
	   
'''#####################################################################################################################
'@Function Name <VerifyDueDate>
'@CreationDate <16/03/2015>
'@Dependency  
'@Author 
'@ModifiedDate
'@ModifiedBy 
'@Description  this function return the Day after Tomorrow 
'@Documentation No input parameter
'''#####################################################################################################################
Public Function VerifyDueDate(ByRef test_object,ByVal strValue)
    Call changeCreationTime (test_object)
    On error resume next
	Dim bFlag,bResultFlag
	bFlag = False
	bResultFlag = False
	If test_object.exist(MaxWait) Then
		ir = test_object.RowCount
		For iRow = 1 To ir Step 1
			ic = test_object.ColumnCount(iRow)
			For jCol = 1 To ic Step 1
				strcelldata = test_object.GetCellData(iRow,jCol)
				If Ucase(strValue) = Ucase(strcelldata) Then
					strcelldata = test_object.GetCellData(iRow,jCol+1)
					If (strcelldata) > (Date) Then
						bResultFlag = True
						bFlag = True
						Exit For
					End If
				End If
			Next
			If bFlag Then
				Exit For
			End If
		Next
		If bResultFlag = True Then
			rptWriteReport "PASS", Environment.Value("StepName"), strValue & "Value has displayed future Date "&(strcelldata) & test_object.ToString		
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), strValue & "Value has not  displayed future Date "&(strcelldata) &test_object.ToString		
		End If
	Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Object doesnt Exist: " & test_object.ToString	
	End If
	
    End Function

RegisterUserFunc "webtable", "VerifyDueDate", "VerifyDueDate"


''**************************************************************************************************************************************************
''Function Name				:		VerifyValueinWebList
''Function Description		:		Checks for the object existence,verifies and selects correct word.
''									This function is generic for WebElement
''Function Parameters		:		strText
''Author					:	    Gallop
''Date						:		10th March 2015
''**************************************************************************************************************************************************
Function VerifyValueinWebList(ByRef test_object,Byval strText)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		If Not(Len(Trim(test_object.getroproperty("all items")) > 0)) Then
			Wait 5
		End If	
			If Instr(1,test_object.getroproperty("all items"),strText,1) <> 0 Then
				rptWriteReport "PASS", Environment.Value("StepName"),chr(34) & strText&" should be displayed in suggestions list"
			Else
				rptWriteReport "FAIL", Environment.Value("StepName"),chr(34) & strText* " should be displayed in suggestions list"
				gErrorFlag = True
			End If
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WebList", "VerifyValueinWebList", "VerifyValueinWebList"


''**************************************************************************************************************************************************
''Function Name				:		VerifyStatusofAttachment
''Function Description		:		Checks for the object existence, verifies the values of attachment and its status
''									
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		20th Mar 2015
''**************************************************************************************************************************************************
Public Function VerifyStatusofAttachment(ByRef test_object,ByVal strValue)
	Call changeCreationTime (test_object)
	Dim iFoundRow,bResultFlag
                bResultFlag = False
                iFoundRow = 0
                arrvalues = split(strvalue,"|")
                If test_object.exist(MaxWait) Then
                                ir = test_object.RowCount
                                For iRow = 1 To ir Step 1
                                                ic = test_object.ColumnCount(iRow)
                                                For jCol = 1 To ic Step 1
                                                                strcelldata = test_object.GetCellData(iRow,jCol)
                                                                If Trim(Ucase(arrvalues(0))) = Trim(Ucase(strcelldata)) Then
                                                                                iFoundRow = iRow
                                                                                bResultFlag = True
                                                                                Exit for
                                                                End If
                                                Next
                                                If iFoundRow > 0 Then
                                                                Exit For
                                                End If
                                Next
                                If ubound(arrvalues)>0 and iFoundRow > 0 Then
                                                For icols = 1 To ubound(arrvalues) Step 1
                                                                bfound = False
                                                                ic = test_object.ColumnCount(iFoundRow-1)
                                                                For itablecols = 1 To ic Step 1
                                                                                For irow = iFoundRow To 1 Step -1
                                                                                                strcelldata = test_object.GetCellData(irow,itablecols)
                                                                                                If Trim(Ucase(strcelldata)) = Trim(UCase(arrvalues(1))) Then
                                                                                                                bfound = True
                                                                                                                Exit For  
																								ElseIf Instr(1,Trim(Ucase(strcelldata)),Trim(UCase(arrvalues(1))),1) > 0 Then 
																												bfound = True
                                                                                                                Exit For 
                                                                                                End If
                                                                                Next                                                                      
                                                                Next
                                                                If bfound = True Then
                                                                                bResultFlag = True
                                                                                Exit For
                                                                Else
                                                                                bResultFlag = False                                                                          
                                                                End If
                                                Next
                                End If
                                If bResultFlag = True Then
                                                rptWriteReport "PASS", Environment.Value("StepName"), arrvalues(0)&"Value should be verified under " &arrvalues(1)& test_object.ToString     

                                Else
                                                rptWriteReport "FAIL", Environment.Value("StepName"), arrvalues(0)&"Value is not verified under " & arrvalues(1)&test_object.ToString                      
                                End If
                Else
                                                rptWriteReport "FAIL", Environment.Value("StepName"), "Object doesnt Exist: " & test_object.ToString      
                End If
End Function
RegisterUserFunc "webtable", "VerifyStatusofAttachment", "VerifyStatusofAttachment"

''**************************************************************************************************************************************************
''Function Name				:		ClickContextMenu
''Function Description		:		Clicks in context menu to saveas from browser
''Function Parameters		:		strText
''Author					:	    Gallop
''Date						:		10th March 2015
''**************************************************************************************************************************************************
Function ClickContextMenu(ByRef test_object,Byval strText)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		test_object.Select strText	
		rptWriteReport "PASS", Environment.Value("StepName"),chr(34) & strText&" Contect menu option is selected"	
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WinMenu", "ClickContextMenu", "ClickContextMenu"


''**************************************************************************************************************************************************
''Function Name						:		 TypeText
''Function Description			    :		 Checks for the object existence, if enabled and sets the value in the object with event
''															 This function is generic for WebEdit, WebArea
''Function Parameters			    :		sValue is the value to be set in the object
''Author							:	    Gallop
''Date								:		20th March 2014
''**************************************************************************************************************************************************
Public Function TypeText(ByRef test_object, ByRef sValue)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		test_object.object.focus
		If test_object.GetROProperty("disabled") = False Then
'			Setting.WebPackage("ReplayType") = 2
'			test_object.Set sValue
'			Setting.WebPackage("ReplayType") = 1
			Set oshell = createobject("Wscript.shell")
		oshell.sendkeys sValue
		oshell.sendkeys "{RIGHT}"
		oshell = Nothing
			rptWriteReport "PASS", Environment.Value("StepName"),chr(34) & sValue & chr(34) & " : value should be entered in " & chr(34) & test_object.ToString & chr(34) &" field."
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"),test_object.ToString & " object should be enabled."
			TypeText = false
			gErrorFlag = True
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		TypeText = false
		gErrorFlag = True
	End If
	On error Goto 0
End Function
RegisterUserFunc "WebEdit", "TypeText", "TypeText"
RegisterUserFunc "WinEdit", "TypeText", "TypeText"


''**************************************************************************************************************************************************
''Function Name						 :		 VerifyTextInTable
''Function Description			  :		 Checks for the object , gets the Column count, Row Count of the table and then verifies each cell of the table with the Value passed
''															 This function is generic for WebTable
''Function Parameters			:		NA
''Author										:	   Gallop
''**************************************************************************************************************************************************
Public Function VerifyTextInTable(ByRef test_object, strText)
	Call changeCreationTime (test_object)
	Table_VerifyText = false
	On error resume next
	If test_object.exist Then
		'rptWriteReport 0, Environment.Value("StepName"), test_object.GetROProperty("name") & " object exists."
		For CC = 1 to test_object.ColumnCount(1)
			For RC = 1 to test_object.RowCount
				If StrComp(Trim(test_object.GetCellData(RC,CC)), Trim(strText), 1) = 0 Then
					rptWriteReport 0, Environment.Value("StepName"), chr(34) & strText & chr(34) & " text is found in " & test_object.GetROProperty("name") & " table."
					tempFlag = true
					Table_VerifyText = true
					Exit for
				End If
			Next
		Next
		If not tempFlag Then
			rptWriteReport 1, Environment.Value("StepName"), chr(34) & strText & chr(34) & " text is not found in " & test_object.GetROProperty("name") & " table."
		Else
			rptWriteReport 0, Environment.Value("StepName"), chr(34) & strText & chr(34) & " text is found in " & test_object.GetROProperty("name") & " table."
		End If		
	Else
		rptWriteReport 1, Environment.Value("StepName"), test_object.GetROProperty("name") & " object doesn't exist."
	End If	
	On error Goto 0
End Function
RegisterUserFunc "WebTable", "VerifyTextInTable", "VerifyTextInTable"

''**************************************************************************************************************************************************
''Function Name				:		VerifyandClickButtonInRow
''Function Description		:		Checks for the button existence of button in a table row and clicks on the button
''									
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		5th Mar 2015
''**************************************************************************************************************************************************
    Public Function VerifyandClickButtonInRow(ByRef test_object,ByVal strValue)
'    Call changeCreationTime(test_object)
	On error resume next
	Dim iFoundRow,bResultFlag
	bResultFlag = False
	bclickflag = false
	iFoundRow = 0
	arrvalues = split(strvalue,"|")
	If test_object.exist(MaxWait) Then
		ir = test_object.RowCount
		For iRow = 1 To ir Step 1
			ic = test_object.ColumnCount(iRow)
			For jCol = 1 To ic Step 1
				strcelldata = test_object.GetCellData(iRow,jCol)
				If Trim(Ucase(arrvalues(0))) = Trim(Ucase(strcelldata)) Then
					iFoundRow = iRow
					iFoundCol = jCol
					bResultFlag = True
					Exit for
				ElseIf Instr(1,Trim(Ucase(strcelldata)),Trim(Ucase(arrvalues(0))),1) > 0 Then
					iFoundRow = iRow
					iFoundCol = jCol
					bResultFlag = True
					Exit for
				End If
			Next
			If iFoundRow > 0 Then
				Exit For
			End If
		Next
		If iFoundRow > 0 Then
			bclickflag = False
			itemcount = test_object.childitemcount(iFoundRow,iFoundCol,"WebElement")					
			If itemcount > 0 Then
				For i  = 0 To itemcount-1 Step 1
					Wait 3
					Set objIcon = test_object.childitem(iFoundRow,iFoundCol,"WebElement",i)
					strclass = objIcon.getROProperty("class")
					if instr(1,strclass,"dropdown-toggle",1) <> 0 then 						
						Wait 3
						objIcon.Focus
						Setting.WebPackage("ReplayType") = 2
						objIcon.click
						Setting.WebPackage("ReplayType") = 1
						bclickflag = True
						Exit For	
					End If
				Next	
			End If
			If bclickflag = true Then
				rptWriteReport "PASS", Environment.Value("StepName"), Val& " button should be clicked under " & test_object.ToString		
			Else
			    rptWriteReport "FAIL", Environment.Value("StepName"),Val& "button should be clicked under " & test_object.ToString
				gErrorFlag = True
			End If
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), "Object doesnt Exist: " & test_object.ToString	
	End If
	End Function

	RegisterUserFunc "webtable", "VerifyandClickButtonInRow", "VerifyandClickButtonInRow"
	
''**************************************************************************************************************************************************
''Function Name				:		VerifyButtonInRow
''Function Description		:		Checks for the Button  in a row
''									
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		5th Mar 2015
''**************************************************************************************************************************************************
    Public Function VerifyButtonInRow(ByRef test_object,ByVal strValue)
	Call changeCreationTime (test_object)
	On error resume next
	Dim iFoundRow,bResultFlag
	bResultFlag = False
	iFoundRow = 0
	arrvalues = split(strvalue,"|")
	If test_object.exist(MaxWait) Then
		ir = test_object.RowCount
		For iRow = 1 To ir Step 1
			ic = test_object.ColumnCount(iRow)
			For jCol = 1 To ic Step 1
				strcelldata = test_object.GetCellData(iRow,jCol)
				If Trim(Ucase(arrvalues(0))) = Trim(Ucase(strcelldata)) Then
					iFoundRow = iRow
					bResultFlag = True
					Exit for
				End If
			Next
			If iFoundRow > 0 Then
				Exit For
			End If
		Next
		If iFoundRow > 0 Then
			For icols = 1 To test_object.ColumnCount(iFoundRow) Step 1
				bfound = False	
				itemcount = test_object.childitemcount(iFoundRow,icols,"WebButton")					
				If itemcount > 0 Then
					For i  = 0 To itemcount-1 Step 1
						Set objIcon = test_object.childitem(iFoundRow,icols,"WebButton",i)
						Val = objIcon.GetRoProperty("html tag")
						If Trim(Val) = Trim(arrvalues(1)) Then
							visible = True
						End If
					Next	
				End If				
			Next
			If visible Then
				rptWriteReport "PASS", Environment.Value("StepName"), strValue&" Button should be displayed under " & test_object.ToString		
			Else
			    rptWriteReport "FAIL", Environment.Value("StepName"),strValue & "Button should be displayed under " & test_object.ToString
				gErrorFlag = True
			End If
		End If
	End If
	End Function
	RegisterUserFunc "webtable", "VerifyButtonInRow", "VerifyButtonInRow"
	
	''**************************************************************************************************************************************************
''Function Name				:		VerifyButtonInRow
''Function Description		:		Checks for the Button  in a row
''									
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		5th Mar 2015
''**************************************************************************************************************************************************
	
	Public Function VerifyDataInCell(ByRef test_object,ByVal strValue)
	Call changeCreationTime (test_object)
	On error resume next
	Dim iFoundRow,bResultFlag
	bResultFlag = False
	iColNumber = 0
	iRowNumber = 0
	arrvalues = split(strvalue,"|")	
	
	If test_object.exist(MaxWait) Then
		ir = test_object.RowCount
		For iRow = 1 To ir Step 1
			ic = test_object.ColumnCount(iRow)
			For jCol = 1 To ic Step 1
				strData = Trim(test_object.GetCellData(iRow,jCol))
				
				If Trim(Ucase(arrvalues(1))) = Trim(Ucase(strData)) Then
					iColNumber = jCol	
				ElseIf 	Trim(Ucase(arrvalues(0))) = Trim(Ucase(strData)) Then	
					iRowNumber  = iRow	
					Exit For	
				End If				
			Next
			If iColNumber > 0 And iRowNumber > 0 Then
				Exit For
			End If
		Next
		If ubound(arrvalues)>0 and iColNumber > 0 and iRowNumber > 0  Then
			strcelldata = Trim(test_object.GetCellData(iRowNumber,iColNumber))			
			If strcelldata = Trim(arrvalues(2)) Then
				rptWriteReport "PASS", Environment.Value("StepName"), strcelldata &" Value should be verified under " &arrvalues(0)& test_object.ToString
				
			Else
				rptWriteReport "FAIL", Environment.Value("StepName"), strcelldata &" Value is not verified under " & arrvalues(0)&test_object.ToString
				
			End If
		End If		
	Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Object doesnt Exist: " & test_object.ToString	
	End If
End Function
RegisterUserFunc "webtable", "VerifyDataInCell", "VerifyDataInCell"


''**************************************************************************************************************************************************
''Function Name				:		ClickonWorkButton
''Function Description		:		Click on Work Button in Actions 
''Function Parameters		:		strText
''Author					:	    Gallop
''Date						:		10th March 2015
''**************************************************************************************************************************************************
Function ClickonWorkButton(ByVal test_object)
    Call changeCreationTime (test_object)
    If test_object.Exist(MaxWait) Then
	   Set objBtn = description.Create
       objBtn("micclass").value = "WebButton"
       objBtn("html tag").value = "BUTTON"
       objBtn("class").value = "btn"
       objBtn("type").value = "submit"
       objBtn("title").value = "Work"
       Set objButton =  test_object.childobjects(objBtn) 
       If objButton.count = 1 Then
       	 Wait 2
         objButton(0).Click
         rptWriteReport "PASS", Environment.Value("StepName"), " Work Button should be clicked under " & test_object.ToString	
	   Else
		 rptWriteReport "FAIL", Environment.Value("StepName"), " Work Button should be clicked under " & test_object.ToString	   
       End If      
     Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		gErrorFlag = True
     End If
End Function
RegisterUserFunc "WebElement", "ClickonWorkButton", "ClickonWorkButton"


''**************************************************************************************************************************************************
''Function Name				:		VerifyActionsareNotDisplayedInRow
''Function Description		:		Validated for the button  in a table row
''									
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		5th Mar 2015
''**************************************************************************************************************************************************
    Public Function VerifyActionsareNotDisplayedInRow(ByRef test_object,ByVal strValue)
	Call changeCreationTime (test_object)
	On error resume next
	Dim iFoundRow,bResultFlag
	bResultFlag = False
	iFoundRow = 0
	arrvalues = split(strvalue,"|")
	If test_object.exist(MaxWait) Then
		ir = test_object.RowCount
		For iRow = 1 To ir Step 1
			ic = test_object.ColumnCount(iRow)
			For jCol = 1 To ic Step 1
				strcelldata = test_object.GetCellData(iRow,jCol)
				If Trim(Ucase(arrvalues(0))) = Trim(Ucase(strcelldata)) Then
					itemcount = test_object.childitemcount(iRow,jCol,"WebButton")					
				        If itemcount = 0 Then
					   ButtonDoesnotExist = True
					   ifoundRow = iRow
					   Exit For
				        End If
				End If
			Next
		    If ifoundRow > 0 Then
			Exit For
		    End If
		Next		
		If ButtonDoesnotExist Then
		     rptWriteReport "PASS", Environment.Value("StepName"),  " button should not be displayed under " & test_object.ToString		
		Else
		    rptWriteReport "FAIL", Environment.Value("StepName"),"button should not be displayed under" & test_object.ToString
		    gErrorFlag = True
		End If	
	Else		
	     gErrorFlag = True
	     rptWriteReport "FAIL", Environment.Value("StepName"), test_object.GetROProperty("outertext") & " object should exist."
	End If
	End Function
	RegisterUserFunc "webtable", "VerifyActionsareNotDisplayedInRow", "VerifyActionsareNotDisplayedInRow"



	

''**************************************************************************************************************************************************
''Function Name				:		VerifyDatainAuditHistory
''Function Description		:		Checks for the object existence, checks whether the values are displayed in a column
''									
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		31th Mar 2015
''**************************************************************************************************************************************************
	Public Function VerifyDatainAuditHistory(ByRef test_object,ByVal strcolname, ByVal strValue)
		Call changeCreationTime (test_object)
		On error resume next
		Dim iFoundRow,bResultFlag
		bResultFlag = False
		iFoundcol = 0
		If test_object.exist(MaxWait) Then
			icolcount = Browser("Summary").Page("Summary").WebTable("TBL_AuditHistoryHeader").ColumnCount(1)
			For icol = 1 To icolcount Step 1
				strcellval = Browser("Summary").Page("Summary").WebTable("TBL_AuditHistoryHeader").GetCellData(1,icol)
				If Trim(ucase(strcolname)) = Trim(ucase(strcellval)) Then
					iFoundcol = icol
					Exit For
				End If			
			Next
			If iFoundcol > 0 Then
				iRowValue = test_object.GetRowWithCellText(strValue)
				If iRowValue > 0 Then
					StrRetValue = test_object.GetCellData(iRowValue,iFoundcol)
					If Trim(ucase(strValue)) = Trim(ucase(StrRetValue)) Then
						rptWriteReport "PASS", Environment.Value("StepName"), StrRetValue & " should be verified under " &strcolname &"in Audit Tab" & test_object.ToString
					ElseIf InStr(1,Trim(ucase(StrRetValue)),Trim(ucase(strValue)),1) > 0 then
						rptWriteReport "PASS", Environment.Value("StepName"), StrRetValue & " should be verified under " &strcolname &"in Audit Tab" & test_object.ToString
					Else
						rptWriteReport "Fail", Environment.Value("StepName"), strValue & " should be verified under " &strcolname &"in Audit Tab" & test_object.ToString
					End If	
				Else
					rptWriteReport "Fail", Environment.Value("StepName"), strValue & " should be verified under " &strcolname &"in Audit Tab" & test_object.ToString
				End If				
			Else
				rptWriteReport "Fail", Environment.Value("StepName"), strValue & " should be verified under " &strcolname &"in Audit Tab" & test_object.ToString
			End If		
		Else
				rptWriteReport "FAIL", Environment.Value("StepName"), "Object doesnt Exist: " & test_object.ToString	
		End If			
	End Function
		RegisterUserFunc "webtable", "VerifyDatainAuditHistory", "VerifyDatainAuditHistory"
		
		
		
''**************************************************************************************************************************************************
''Function Name				:		SelectCheckBox
''Function Description		:		Checks for the object existence, if enabled and selects the checkbox
''									This function is generic for WebElement
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		19th Feb 2015
''**************************************************************************************************************************************************
Public Function SelectCheckBox(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		Set odesc = Description.Create
		odesc("micclass").value ="WebCheckBox"
		Set CheckBox = test_object.ChildObjects(odesc)
		If CheckBox.count = 0 Then
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"), test_object.GetROProperty("outertext") & " checkbox."
		Else
			CheckBox(0).Set "ON"
			rptWriteReport "PASS", Environment.Value("StepName"), "selected" & test_object.GetROProperty("outertext") & " checkbox."
		End If
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.GetROProperty("outertext") & " object should exist."
	End If
	On error Goto 0
End Function
RegisterUserFunc "WebElement", "SelectCheckBox", "SelectCheckBox"

''**************************************************************************************************************************************************
''Function Name				:		PageSync
''Function Description		:		wait for page load
''									This function is generic for web element,web edit
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		25th March 2015
''**************************************************************************************************************************************************
Function PageSync (ByRef test_object)
	Call changeCreationTime (test_object)
	If test_object.exist(MaxWait) Then  
		set objPage = test_object.GetTOProperty("parent")
		strclass = objPage.getTOProperty("micclass")
		If strcomp(strclass,"Page",1)=0 Then
			objPage.sync
			wait(2)
		End If
	End If   				
End Function
RegisterUserFunc "WebElement", "PageSync","PageSync"
RegisterUserFunc "WebEdit", "PageSync","PageSync"
RegisterUserFunc "WebTable", "PageSync","PageSync"
RegisterUserFunc "webbutton", "PageSync","PageSync"


''**************************************************************************************************************************************************
''Function Name				:		VerifyYellowBanner
''Function Description		:		Checks for the object existence, checks yellow alert message(banner) is displayed
''									This function is generic for WebElement
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		26th March 2015
''**************************************************************************************************************************************************
Function VerifyYellowBanner(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		If Instr(1,test_object.GetROProperty("class"),"alert-warning",1) <> 0 Then
			rptWriteReport "PASS", Environment.Value("StepName"), "Verify yellow banner is displayed with text  " & chr(34) & test_object.GetROProperty("innertext") & chr(34)
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Verify yellow banner is displayed with text  " & chr(34) & test_object.GetROProperty("innertext") & chr(34)
			gErrorFlag = True
		End If
	Else
			gErrorFlag = True
			rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function
RegisterUserFunc "WebElement", "VerifyYellowBanner","VerifyYellowBanner"

''**************************************************************************************************************************************************
''Function Name						:		 VerifyValue
''Function Description			    :		 Checks for the object existence, if enabled and sets the value in the object
''															 This function is generic for WebEdit, WebArea
''Function Parameters			    :		sValue is the value to be set in the object
''Author							:	    Gallop
''Date								:		3rd Aug 2014
''**************************************************************************************************************************************************
Public Function VerifyValue(ByRef test_object, ByVal sValue)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then		
		If StrComp(test_object.GetROProperty("value"),sValue,1) = 0 Then
			rptWriteReport "PASS", Environment.Value("StepName"), sValue & " : value should be displayed in " & chr(34) & test_object.ToString & chr(34) & " field."
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), sValue & " : value should be displayed in " & chr(34) & test_object.ToString & chr(34) & " field."
		End If
		VerifyValue = true		
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		VerifyValue = false
		gErrorFlag = True
	End If
	On error Goto 0
End Function
RegisterUserFunc "WebEdit", "VerifyValue", "VerifyValue"

''**************************************************************************************************************************************************
''Function Name				:		ClickOnToggles
''Function Description		:		Checks for the object existence, Click on Toggles
''									This function is generic for WebElement
''Function Parameters		:		sValue - is toggle value
''Author					:	    Gallop
''Date						:		16th Mar 2015
''**************************************************************************************************************************************************

Public Function ClickOnToggles(ByRef test_object,ByRef sValue)
	Call changeCreationTime (test_object)
	If test_object.exist(MaxWait) Then
		Set objToggle = description.Create
		objToggle("micclass").value = "WebElement"
		objToggle("class").value = "btn btn-default.*"
		objToggle("innertext").value = sValue
		Set objToggleValue = test_object.childobjects(objToggle)
		If objToggleValue.Count = 1 Then
			objToggleValue(0).Click
			rptWriteReport "PASS", Environment.Value("StepName"), sValue&" Toggle should be clicked under " & test_object.ToString		
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), sValue&" Toggle should be Shown under " & test_object.ToString
			gErrorFlag = True
		End If	
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		gErrorFlag = True
	End If
	
End Function
RegisterUserFunc "WebElement", "ClickOnToggles", "ClickOnToggles"

''**************************************************************************************************************************************************
''Function Name				:		VerifyNoValuesSelectedForToggle
''Function Description		:		Checks for the object existence, verifies no value is selected for toggle
''									This function is generic for WebElement
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		30th Mar 2015
''**************************************************************************************************************************************************

Public Function VerifyNoValuesSelectedForToggle(ByRef test_object)
	Call changeCreationTime (test_object)
	Dim ItemsNotSelected: ItemsNotSelected = 0
	If test_object.exist(MaxWait) Then
		Set objToggle = description.Create
		objToggle("micclass").value = "WebElement"
		objToggle("class").value = "btn btn-default.*"
		Set objToggleValue = test_object.childobjects(objToggle)
		For i = 0 To objToggleValue.Count-1
			If(Instr(1,objToggleValue(i).getROProperty("class"),"active",1) <> 0) Then
				gErrorFlag = True
				rptWriteReport "FAIL", Environment.Value("StepName"),chr(34) & otest_object.ToString & chr(34) & " should not have default value selected"
				Exit For
			Else
				ItemsNotSelected = ItemsNotSelected+1
			End If
		Next
		If ItemsNotSelected = objToggleValue.Count Then
			rptWriteReport "PASS", Environment.Value("StepName"),chr(34) & test_object.ToString & chr(34) & " should not have default value selected"
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"),chr(34) & test_object.ToString & chr(34) & " should not have default value selected"
			gErrorFlag = True
		End If	
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		gErrorFlag = True
	End If
	
End Function
RegisterUserFunc "WebElement", "VerifyNoValuesSelectedForToggle", "VerifyNoValuesSelectedForToggle"

''**************************************************************************************************************************************************
''Function Name				:		VerifyTableIsEmpty
''Function Description		:		Checks for the object existence, verifies webtable has no rows or empty
''									
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		30th Mar 2015
''**************************************************************************************************************************************************
Public Function VerifyTableIsEmpty(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	If test_object.exist(MaxWait) Then
		ir = test_object.RowCount
		If ir = 1 Then
			rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) &" should be empty"
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) &" should be empty"		
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."	
	End If
End Function
RegisterUserFunc "WebTable", "VerifyTableIsEmpty", "VerifyTableIsEmpty"



Public Function clickProceed(ByRef test_object)
Call changeCreationTime (test_object)
On error resume next
wait(10)
bPrceedflag = false
		Set objBrowser = description.Create
		objBrowser("micclass").value = "Browser"
		Set objProceed = description.Create
		objProceed("micclass").value = "WebButton"
		objProceed("value").value = "Proceed"
		objProceed("type").value = "submit"
		Set objPage = description.Create
		objPage("micclass").value = "Page"
		Set BrowserList = desktop.ChildObjects(objBrowser)
		For ib = 0 To BrowserList.count-1 Step 1
			Set PageList = BrowserList(ib).ChildObjects(objPage)	
			For ip = 0 To PageList.count-1 Step 1
				Set Proceedobj = PageList(ip).childobjects(objProceed)
				If Proceedobj.count = 1 Then
					Proceedobj(0).click
					bPrceedflag = True
				End If
			Next
		Next
		If bPrceedflag = True Then
			rptWriteReport "PASS", Environment.Value("StepName"),"Should click on Proceed button"
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"),"Should click on Proceed button"		
		End If
End Function
RegisterUserFunc "WebButton", "clickProceed", "clickProceed"



Public Function clickAudit(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	wait(10)
	bAuditFlag = false
	Set objBrowser = description.Create
	objBrowser("micclass").value = "Browser"
	Set objAudit = description.Create
	objAudit("micclass").value = "WebElement"
	objAudit("innertext").value = "Audit"
	objAudit("class").value = "k-link"
	Set objPage = description.Create
	objPage("micclass").value = "Page"
	Set BrowserList = desktop.ChildObjects(objBrowser)
	For ib = 0 To BrowserList.count-1 Step 1
		Set PageList = BrowserList(ib).ChildObjects(objPage)	
		For ip = 0 To PageList.count-1 Step 1
			Set Auditobj = PageList(ip).childobjects(objAudit)
			If Auditobj.count = 1 Then
				Auditobj(0).click
				bAuditFlag = True
			End If
		Next
	Next
	If bAuditFlag = True Then
		rptWriteReport "PASS", Environment.Value("StepName"),"Should click on Audit button"
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"),"Should click on Audit button"		
	End If
End Function
RegisterUserFunc "WebElement", "clickAudit", "clickAudit"


''**************************************************************************************************************************************************
''Function Name					:		WaitForObjectToLoadWithTime
''Function Description			:		Checks for the object existence, and then waits for the object to get loaded
''										This function is generic for Link,Image,Browser,Page,WebElement,WebEdit,WebButton,WebList,WebRadioGroup,WebTable
''Function Parameters			:		NA
''Author						:	   	Gallop
''Date							:		5th Aug 2014
''**************************************************************************************************************************************************
Public Function WaitForObjectToLoadWithTime(ByRef test_object,ByVal sTime)
	Call changeCreationTime (test_object)
	On error resume Next
	If test_object.exist(MaxWait) Then
		If test_object.WaitProperty("attribute/readyState","complete",sTime * 1000) or test_object.WaitProperty("attribute/ReadyState",4,sTime * 1000) Then
			rptWriteReport "PASS", Environment.Value("StepName"), test_object.ToString & "  should be loaded " 
			WaitForObjectToLoadWithTime = true
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "  should be loaded "
			WaitForObjectToLoadWithTime = false
			gErrorFlag = True
		End if 
	End If
	On error Goto 0
End Function
RegisterUserFunc "Link", "WaitForObjectToLoadWithTime", "WaitForObjectToLoadWithTime"
RegisterUserFunc "Image", "WaitForObjectToLoadWithTime", "WaitForObjectToLoadWithTime"
RegisterUserFunc "Browser", "WaitForObjectToLoadWithTime", "WaitForObjectToLoadWithTime"
RegisterUserFunc "Page", "WaitForObjectToLoadWithTime", "WaitForObjectToLoadWithTime"
RegisterUserFunc "WebElement", "WaitForObjectToLoadWithTime", "WaitForObjectToLoadWithTime"
RegisterUserFunc "WebEdit", "WaitForObjectToLoadWithTime", "WaitForObjectToLoadWithTime"
RegisterUserFunc "WebButton", "WaitForObjectToLoadWithTime", "WaitForObjectToLoadWithTime"
RegisterUserFunc "WebList", "WaitForObjectToLoadWithTime", "WaitForObjectToLoadWithTime"
RegisterUserFunc "WebRadioGroup", "WaitForObjectToLoadWithTime", "WaitForObjectToLoadWithTime"
RegisterUserFunc "WebTable", "WaitForObjectToLoadWithTime", "WaitForObjectToLoadWithTime"

	''**************************************************************************************************************************************************
''Function Name						:		 DownloadFile
''Function Description			    :		 Checks for the object existence, if enabled and sets the fileName in the object
''											This function is generic for WebEdit, WebArea
''Function Parameters			    :		NA
''Author							:	    Gallop
''Date								:		6th April 2015
''**************************************************************************************************************************************************

Public Function DownloadFile(ByRef test_object,ByVal strfilename)
	Call changeCreationTime (test_object)
	On error resume Next
	Filename = sBaseDir  & "Resources\"& strfilename
	Set fso=createobject("Scripting.FileSystemObject")
	If fso.FileExists(filepath) then
 		fso.DeleteFile Filename
	End If
	Set fso=nothing
	If test_object.exist(MaxWait) Then
		If test_object.GetROProperty("disabled") = False Then
			Setting.WebPackage("ReplayType") = 2
			test_object.Type Filename
			Setting.WebPackage("ReplayType") = 1				
			rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & Filename & chr(34) & " : value should be entered in " & chr(34) & test_object.ToString & chr(34) & " field."
			strDownlodedFilePath = Filename
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"),test_object.ToString & " object should be enabled."
			strDownlodedFilePath = Filename
			gErrorFlag = True
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		EnterText = false
		gErrorFlag = True
	End If
	On error Goto 0
End Function
RegisterUserFunc "WebEdit", "DownloadFile", "DownloadFile"
RegisterUserFunc "WinEdit", "DownloadFile", "DownloadFile"

''**************************************************************************************************************************************************
''Function Name						:		 changeCreationTime
''Function Description			    :		  check the browsers and set creation time for browser in obejct repository
''											
''Function Parameters			    :		NA
''Author							:	    Gallop
''Date								:		6th April 2015
''**************************************************************************************************************************************************
Function changeCreationTime(ByRef testobject)
Set objBrowser = description.Create
objBrowser("micclass").value = "Browser"
Set BrowserList = desktop.ChildObjects(objBrowser)
If BrowserList.count > 1 Then
	For blist = 0 to BrowserList.count-1 
		if not len(BrowserList(blist).GetROProperty("name"))>0 then
			wait(3)
		End if	
	Next 
	Dim objBrowser
	Set test_object = testobject
	Set objParent = test_object.GetTOProperty("parent")
	
	Do while true
		strclass = objParent.GetTOProperty("micclass")
		If strclass = "Browser" Then
			Set objBrowser = objParent
			Exit Do
		End If
		If Not len(strclass) >0 Then
			Exit Do
		End If	
		set objParent = objParent.GetTOProperty("parent")
	Loop	

	strBrowserName = objBrowser.GetTOProperty("name")
	emptybrowser = 0
	For ic = 0 To BrowserList.count-1 Step 1
		strname = Browser("creationtime:="&ic).GetROProperty("name")		
		bvisible = Browser("creationtime:="&ic).GetROProperty("visible")
		if not len(strname)>0 then
		emptybrowser = emptybrowser+1
		end if 
		If strname = strBrowserName and bvisible = True Then
			objBrowser.setTOProperty "Creationtime",ic-emptybrowser
			Exit For
		End If
	Next
	Set objParent = Nothing
End If
Set objBrowser = Nothing
End Function



''**************************************************************************************************************************************************
''Function Name				:		VerifyDateInRow
''Function Description		:		Checks for the object existence, verifies the Date in row
''									
''Function Parameters		:		NA
''Author					:	    Gallop
''Date						:		5th Mar 2015
''**************************************************************************************************************************************************
Public Function VerifyDateInRow(ByRef test_object,ByVal strValue)
	Call changeCreationTime (test_object)
	On error resume next
	Dim iFoundRow,bResultFlag
	bResultFlag = False
	iFoundRow = 0
	arrvalues = split(strvalue,"|")
	If test_object.exist(MaxWait) Then
		ir = test_object.RowCount
		For iRow = 1 To ir Step 1
			ic = test_object.ColumnCount(iRow)
			For jCol = 1 To ic Step 1
				strcelldata = test_object.GetCellData(iRow,jCol)
				If Trim(Ucase(arrvalues(0))) = Trim(Ucase(strcelldata)) Then
					iFoundRow = iRow
					bResultFlag = True
					Exit for	
				ElseIf InStr(1,Trim(Ucase(strcelldata)),Trim(Ucase(arrvalues(0))),1) > 0 Then
					iFoundRow = iRow
					bResultFlag = True
					Exit for		
				End If
			Next
			If iFoundRow > 0 Then
				Exit For
			End If
		Next
		If ubound(arrvalues)>0 and iFoundRow > 0 Then
			For icols = 1 To ubound(arrvalues) Step 1
				bfound = False
				ic = test_object.ColumnCount(iFoundRow)
				For itablecols = 1 To ic Step 1
					strcelldata = test_object.GetCellData(iFoundRow,itablecols)
					If Ucase(Trim(strcelldata)) = UCase(Trim(arrvalues(icols))) Then
						bfound = True
						Exit For
					ElseIf InStr(1,Ucase(Trim(strcelldata)),UCase(Trim(arrvalues(icols))),1) > 0 Then
						bfound = True
						Exit For
					End If
				Next
				If bfound = True Then
					bResultFlag = True
				Else
					bResultFlag = False					
				End If
			Next
		End If
		If bResultFlag = True Then
			rptWriteReport "PASS", Environment.Value("StepName"), arrvalues(1) &" Value should be verified under " &arrvalues(0)& test_object.ToString		
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), arrvalues(1) &" Value is not verified under " & arrvalues(0)&test_object.ToString		
		End If
	Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Object doesnt Exist: " & test_object.ToString	
	End If
End Function
RegisterUserFunc "webtable", "VerifyDateInRow", "VerifyDateInRow"
''**************************************************************************************************************************************************
''Function Name				:		VerifyvalueInSpecificCell
''Function Description		:		Checks for the object existance, verifiy the value in specific cell
''									This function is generic for WebTable
''Function Parameters		:		iRows,icol,svalue
''Author					:	    Gallop
''Date						:		14th May 2015
''**************************************************************************************************************************************************
Function VerifyvalueInSpecificCell(ByRef test_object,Byval iRow,Byval iCol,Byval svalue)
	
	If test_object.exist(MaxWait) Then
		If Trim(ucase(test_object.GetCellData(iRow,iCol))) = Trim(ucase(svalue)) Then
			rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & Trim(ucase(svalue)) & chr(34) & " Should be in "&iRow&" Row and "&iCol&" Column under "& test_object.ToString & chr(34) 
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), chr(34) & Trim(ucase(svalue)) & chr(34) & " Should be in "&iRow&" Row and "&iCol&" Column under "& test_object.ToString & chr(34) & " but displaayed value is "& Trim(ucase(test_object.GetCellData(iRow,iCol)))
		End If
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & "should exist"
	End If
End Function

RegisterUserFunc "WebTable", "VerifyvalueInSpecificCell","VerifyvalueInSpecificCell"
''**************************************************************************************************************************************************
''Function Name						:		 VerifyValueisNull
''Function Description			    :		 Checks for the object existence, verifies ValueisNull
''															 This function is generic for WebEdit
''Function Parameters			    :		
''Author							:	    Gallop
''Date								:		24th March 2015
''**************************************************************************************************************************************************
Public Function VerifyValueisNull(ByRef test_object)
		
	If test_object.exist(MaxWait) Then		
		If StrComp(test_object.GetROProperty("value"),Empty,1) = 0 Then
			rptWriteReport "PASS", Environment.Value("StepName"), " Null value should be displayed in " & chr(34) & test_object.ToString & chr(34) & " field."
			VerifyValueisNull = true
		ElseIf StrComp(test_object.GetROProperty("innertext"),"select",1) = 0 Then
			rptWriteReport "PASS", Environment.Value("StepName"), " Null value should be displayed in " & chr(34) & test_object.ToString & chr(34) & " field."
			VerifyValueisNull = true
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"),  " Null value should be displayed in " & chr(34) & test_object.ToString & chr(34) & " field."
			VerifyValueisNull = false
		End If
				
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		VerifyValueisNull = false
		gErrorFlag = True
	End If
	
End Function
RegisterUserFunc "WebElement", "VerifyValueisNull", "VerifyValueisNull"
RegisterUserFunc "WebEdit", "VerifyValueisNull", "VerifyValueisNull"

Public Function clickBack(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	wait(10)
	bPrceedflag = false
		Set objBrowser = description.Create
		objBrowser("micclass").value = "Browser"
		Set objProceed = description.Create
		objProceed("micclass").value = "WebButton"
		objProceed("value").value = "Back"
		objProceed("type").value = "submit"
		Set objPage = description.Create
		objPage("micclass").value = "Page"
		Set BrowserList = desktop.ChildObjects(objBrowser)
		For ib = 0 To BrowserList.count-1 Step 1
			Set PageList = BrowserList(ib).ChildObjects(objPage)	
			For ip = 0 To PageList.count-1 Step 1
				Set Proceedobj = PageList(ip).childobjects(objProceed)
				If Proceedobj.count = 1 Then
					Proceedobj(0).click
					bPrceedflag = True
				End If
			Next
		Next
		If bPrceedflag = True Then
			rptWriteReport "PASS", Environment.Value("StepName"),"Should click on Back button"
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"),"Should click on Back button"		
		End If
End Function
RegisterUserFunc "WebButton", "clickBack", "clickBack"


Public Function clickNext(ByRef test_object)
	Call changeCreationTime (test_object)
	On error resume next
	wait(10)
	bPrceedflag = false
		Set objBrowser = description.Create
		objBrowser("micclass").value = "Browser"
		Set objProceed = description.Create
		objProceed("micclass").value = "WebButton"
		objProceed("value").value = "Next"
		objProceed("type").value = "submit"
		Set objPage = description.Create
		objPage("micclass").value = "Page"
		Set BrowserList = desktop.ChildObjects(objBrowser)
		For ib = 0 To BrowserList.count-1 Step 1
			Set PageList = BrowserList(ib).ChildObjects(objPage)	
			For ip = 0 To PageList.count-1 Step 1
				Set Proceedobj = PageList(ip).childobjects(objProceed)
				If Proceedobj.count = 1 Then
					Proceedobj(0).click
					bPrceedflag = True
				End If
			Next
		Next
		If bPrceedflag = True Then
			rptWriteReport "PASS", Environment.Value("StepName"),"Should click on Next button"
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"),"Should click on Next button"		
		End If
End Function
RegisterUserFunc "WebButton", "clickNext", "clickNext"

''**************************************************************************************************************************************************
''Function Name						:		 ValidateState
''Function Description			    :		 Checks for the  status 
''															 This function is generic for WebElement
''Function Parameters			    :		sValue(Type of color) is the value to be set in the object
''Author							:	    Gallop
''Date								:		25th Mar 2015
''**************************************************************************************************************************************************
Public Function ValidateState(ByRef test_object, ByVal sValue)
	
	If test_object.exist(MaxWait) Then
		 		
				If test_object.exist(MaxWait) Then		
					If InStr(1,test_object.GetROProperty("outerhtml"),sValue,1) > 0 Then				
						rptWriteReport "PASS", Environment.Value("StepName"), " State should be displayed in " & chr(34) & test_object.ToString & chr(34) & " field. And the state is "&sValue
					Else
						rptWriteReport "FAIL", Environment.Value("StepName"), " Error should be displayed in " & chr(34) & test_object.ToString & chr(34) & " field. And the state is not "&sValue
					End If
					VerifyColor = true		
				Else
					rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
					VerifyColor = false
					gErrorFlag = True
				End If
			
	Else
		gErrorFlag = True
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.GetROProperty("outertext") & " object should exist."
	End If
	
End Function
RegisterUserFunc "WebElement", "ValidateState", "ValidateState"

''**************************************************************************************************************************************************
''Function Name						:		 VerifyEnabledandVisible
''Function Description			    :		 Checks for the object existence, is enabled and visible
''															 
''Function Parameters			    :		
''Author							:	    Gallop
''Date								:		3rd Aug 2014
''**************************************************************************************************************************************************
Public Function VerifyEnabledandVisible(ByRef test_object)
	
	If test_object.exist(MaxWait) Then
		objVisible = test_object.GetROProperty("visible")	
		objDisabled = test_object.GetROProperty("disabled")
		If objVisible And Not objDisabled Then
			VerifyEnabledandVisible = true		
			rptWriteReport "Pass", Environment.Value("StepName"), test_object.ToString & " should be visible and Enabled"
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " should be visible and Enabled"
		End If
		
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		VerifyEnabledandVisible = false
		gErrorFlag = True
	End If
	
End Function
RegisterUserFunc "WebElement", "VerifyEnabledandVisible", "VerifyEnabledandVisible"

''**************************************************************************************************************************************************
''Function Name						:		 VerifyDisabledandVisible
''Function Description			    :		 Checks for the object existence, is Disabled and visible
''															 
''Function Parameters			    :		
''Author							:	    Gallop
''Date								:		3rd Aug 2014
''**************************************************************************************************************************************************
Public Function VerifyDisabledandVisible(ByRef test_object)
	
	If test_object.exist(MaxWait) Then
		objVisible = test_object.GetROProperty("visible")	
		objDisabled = test_object.GetROProperty("disabled")
		If objVisible And  objDisabled Then
			VerifyDisabledandVisible = true		
			rptWriteReport "Pass", Environment.Value("StepName"), test_object.ToString & " should be visible and Disabled"
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " should be visible and Disabled"
		End If
		
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		VerifyDisabledandVisible = false
		gErrorFlag = True
	End If
	
End Function
RegisterUserFunc "WebEdit", "VerifyDisabledandVisible", "VerifyDisabledandVisible"
RegisterUserFunc "WebElement", "VerifyDisabledandVisible", "VerifyDisabledandVisible"


''**************************************************************************************************************************************************
''Function Name				:		ClickONAttachForSpecificDataInTable
''Function Description		:		Checks for the object existence, check for specific cell data and click on attach button
''									
''Function Parameters		:		NA
''Author					:	    sreedhar metukuru
''Date						:		24 Feb 2016
''**************************************************************************************************************************************************
Public Function ClickONAttachForSpecificDataInTable(ByRef test_object,ByVal strValue)

	On Error Resume Next
	If test_object.exist(MaxWait) Then
		bPassFlag = false
		bFlag=false

		ir = test_object.RowCount
		ic = test_object.ColumnCount(ir)
		For i = 1 To ir Step 1			
			For j = 1 To ic Step 1
				
				strCellValue = test_object.GetCellData(i,j)
							
				If Trim(strCellValue) = Trim(strValue) Then
				   strTitle=test_object.ChildItem(i,8,"WebElement",0).GetROProperty("title")
				   strTag=test_object.ChildItem(i,8,"WebElement",0).GetROProperty("html tag")
				   If strTitle= "Attach Document" And strTag = "A" Then
				   		test_object.ChildItem(i,8,"WebElement",0).click 					   						   						   		
					   	bPassFlag = True	
						bFlag=True						
						Exit for 
				   End If
				End If
			Next
			If bFlag Then
				Exit for
			End If
		Next
		
		If bPassFlag Then
			rptWriteReport "Pass", Environment.Value("StepName"),  strValue & " Cell data displayed and clicked on same row Attach Icon"
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), strValue & " Cell data displayed and clicked on same row Attach Icon"
			gErrorFlag = True
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		gErrorFlag = True
	End If

End Function
RegisterUserFunc "webtable", "ClickONAttachForSpecificDataInTable", "ClickONAttachForSpecificDataInTable"

''**************************************************************************************************************************************************
''Function Name				:		VerifyElementInDataTableAndClickOnIT
''Function Description		:		Checks for the object existence, verifies links intable and click on it
''									This function is generic for WebTable
''Function Parameters		:		column name
''Author					:	    Sreedhar Metukuru
''Date						:		10 March 2016
''**************************************************************************************************************************************************
Public Function ExpandDataInDataTable(ByRef test_object)
	
	On Error Resume Next
	bPassFlag = false
	If test_object.exist(MaxWait) Then
		

		ir = test_object.RowCount
		j=1
		For i = 1 To ir Step 1	
		    test_object.ChildItem(i,j,"WebElement",0).Click
		Next
		bPassFlag = true
		If bPassFlag Then
			rptWriteReport "Pass", Environment.Value("StepName"), "Able to expand data in table"
		Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Able to expand data in table"
			gErrorFlag = True
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		gErrorFlag = True
	End If		
End Function
RegisterUserFunc "WebTable", "ExpandDataInDataTable", "ExpandDataInDataTable"

'***********************************************************************************************************************
''**************************************************************************************************************************************************
''Function Name				:		VerifyDataInRowForSpecificColumnName
''Function Description		:		VerifyDataInRowForSpecificColumnName
''									
''Function Parameters		:		NA
''Author					:	    Sreedhar Metukuru
''Date						:		10 Mar 2016
''**************************************************************************************************************************************************
	
	Public Function VerifyDataInRowForSpecificColumnName(ByRef test_object,ByVal strValue)
	Call changeCreationTime (test_object)
	On error resume next
	Dim iFoundRow,bResultFlag
	bResultFlag = False
	arrvalues = split(strvalue,"|")	
		
	If test_object.exist(MaxWait) Then
		ir = test_object.RowCount
		jCol=1
		For iRow = 1 To ir Step 1
			strData = Trim(test_object.GetCellData(iRow,jCol))
			If Trim(arrvalues(0)) = Trim(strData) Then
				strText = Trim(test_object.GetCellData(iRow,2))
				strValue = Trim(test_object.GetCellData(iRow,3))
				If Trim(arrvalues(1))=Trim(strText) And Trim(arrvalues(2))=Trim(strValue) Then
					bResultFlag = True
					Exit For
				End If		
			End If				
		Next
			If bResultFlag Then
				rptWriteReport "PASS", Environment.Value("StepName"), arrvalues(0) &" Column text and values details follow as " &arrvalues(1)&" , "&arrvalues(2)
				
			Else
				rptWriteReport "FAIL", Environment.Value("StepName"), arrvalues(0) &" Column text and values details follow as " &arrvalues(1)&" , "&arrvalues(2)
				
			End If
	Else
			rptWriteReport "FAIL", Environment.Value("StepName"), "Object doesnt Exist: " & test_object.ToString	
	End If
End Function
RegisterUserFunc "webtable", "VerifyDataInRowForSpecificColumnName", "VerifyDataInRowForSpecificColumnName"

''**************************************************************************************************************************************************
''Function Name						 :	   ClickDebugObject
''Function Description			     :	   Checks for the Dubug object existence, if enabled and Clicks the object
''										   This function is generic for Debug WebElement
''Function Parameters			     :	   NA
''Author					    	 :	   Gallop
''Date								 :	   22 July 2016
''**************************************************************************************************************************************************
Public Function ClickDebugObject(ByRef test_object)
	Call changeCreationTime (test_object)
	wait 10
	On error resume next
	If test_object.exist(MaxWait) Then
		If Instr(1,test_object.GetROProperty("Class Name"),"Win",1) > 0 Then
		
			If test_object.GetROProperty("enabled") = True Then
				Wait 1
				test_object.object.focus
				test_object.Click
				rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " object should be clicked."
				ClickObject = true
			Else
				rptWriteReport "FAIL", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " object should be enabled."
				ClickObject = false
				gErrorFlag = True
			End If
		ElseIf test_object.GetROProperty("disabled") = false Then
		Wait(3)
		        Setting.WebPackage("ReplayType") = 2
				test_object.Click
				Setting.WebPackage("ReplayType") = 1
				rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " object Should be clicked."
				ClickObject = true
'			Else
'				rptWriteReport "FAIL", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " object should be enabled."
'				ClickObject = false
'				gErrorFlag = True
'			End If
		ElseIf Instr(1,test_object.GetROProperty("nativeclass"), "drop", 1) > 0 Then
		    Setting.WebPackage("ReplayType") = 2
			test_object.Click
			Setting.WebPackage("ReplayType") = 1
			rptWriteReport "PASS", Environment.Value("StepName"), chr(34) & test_object.ToString & chr(34) & " object should be clicked."
			ClickObject = true			
		End If
	Else
		rptWriteReport "FAIL", Environment.Value("StepName"), test_object.ToString & " object should exist."
		ClickObject = false
		gErrorFlag = True
	End If
	On error Goto 0
End Function
RegisterUserFunc "WebElement", "ClickDebugObject", "ClickDebugObject"

