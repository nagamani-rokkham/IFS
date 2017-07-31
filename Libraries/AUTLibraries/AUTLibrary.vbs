''###################################################################################################
''Function Name : Login
''Description : Login to application
''Author : Gallop
''###################################################################################################
Public Function Login
Environment.Value("StepName") = "Login"
	Call StartApplication()
	Browser("Login").Page("Login").WebEdit("TXT_user_name").EnterText dtScenario.value("username")
	Browser("Login").Page("Login").WebEdit("TXT_user_password").EnterText dtScenario.value("password")
	Browser("Login").Page("Login").WebButton("BTN_Log In").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB _Home").PageSync 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB _Home").VerifyExists 

End Function
''###################################################################################################
''Function Name : CurrencyValidation
''Description : Currency Truncate
''Author : Prasuna
''###################################################################################################
Public Function CurrencyValidation
Environment.Value("StepName") = "CurrencyValidation"
	Browser("Work Queue").Page("Work Queue").WebElement("Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("Test Question Types").ClickObject 
	Browser("Work Queue").Page("Navigation Page").WebElement("Navigate through Service").SelectListBox "Currency"
	Browser("Work Queue").Page("Navigation Page").WebButton("Next").ClickObject 
	Browser("Currency").Page("Currency").WebEdit("CurrencyOneColumn").EnterText dtScenario.value("CurrencyOneColumn")
	Browser("Currency").Page("Currency").WebEdit("CurrencyOneColumn2").EnterText dtScenario.value("CurrencyOneColumn2")
	Browser("Currency").Page("Currency").WebEdit("CurrencyWholeOneColumn").EnterText dtScenario.value("CurrencyWholeOneColumn")
	Browser("Currency").Page("Currency").WebEdit("CurrencyWholeTwoColumn").EnterText dtScenario.value("CurrencyWholeTwoColumn")
	Browser("Currency").Page("Currency").WebEdit("CurrencyFormattedOneColumn").EnterText dtScenario.value("CurrencyFormattedOneColumn")
	Browser("Currency").Page("Currency").WebEdit("CurrencyFormattedTwoColumn").EnterText dtScenario.value("CurrencyFormattedTwoColumn")
	Browser("Currency").Page("Currency").WebEdit("CurrencyTruncateOneColumn").EnterText dtScenario.value("CurrencyTruncateOneColumn")
	Browser("Currency").Page("Currency").WebEdit("CurrencyTruncateTwoColumn").EnterText dtScenario.value("CurrencyTruncateTwoColumn")
	Browser("Currency").Page("Currency").WebEdit("CurrencyTruncateMinMax").EnterText dtScenario.value("CurrencyTruncateMinMax")
	Browser("Currency").Page("Currency").WebButton("Next").ClickObject 
	Browser("Currency").Page("Currency").WebElement("Currency Truncate Test").VerifyMandatoryMessage "Entry cannot exceed 5"
End Function
''###################################################################################################
''Function Name : NavigateToYesNO
''Description : Navigate to Yes No page
''Author : Gallop
''###################################################################################################
Public Function NavigateToYesNO
Environment.Value("StepName") = "NavigateToYesNO"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_Test Question Types").ClickObject 
	Browser("Navigation Page").Page("Navigation Page").WebElement("DWL_Navigate through Service").SelectListBox "Yes/NO"
	Browser("Navigation Page").Page("Navigation Page").WebButton("BTN_Next").ClickObject 
	Browser("Yes/No").Page("Yes/No").WebElement("LBL_Yes/No").VerifyPageName 
End Function
''###################################################################################################
''Function Name : VerifyDeafultValuesInYesNo
''Description : Verify default radio buttons selected in Yes No Page
''Author : Gallop
''###################################################################################################
Public Function VerifyDeafultValuesInYesNo
Environment.Value("StepName") = "VerifyDeafultValuesInYesNo"
	Browser("Yes/No").Page("Yes/No").WebButton("BTN_Next").WaitForObjectToLoad dtScenario.value("TimeInSecs")
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_Yes/NoDefaulted").VerifyRadioButtonIsSelected "YES"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_Yes/NoDefaultedTwoColumn").VerifyRadioButtonIsSelected "NO"
End Function
''###################################################################################################
''Function Name : LogoutApplication
''Description : Logout from application and close browsers
''Author : Gallop
''###################################################################################################
Public Function LogoutApplication
Environment.Value("StepName") = "LogoutApplication"
	Call Logout()
	Call CloseAllBrowsers()

End Function
''###################################################################################################
''Function Name : NavigateToCurrency
''Description : Navigate to Currency page
''Author : Gallop
''###################################################################################################
Public Function NavigateToCurrency
Environment.Value("StepName") = "NavigateToCurrency"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_Test Question Types").ClickObject 
	Browser("Navigation Page").Page("Navigation Page").WebElement("DWL_Navigate through Service").SelectListBox "Currency"
	Browser("Navigation Page").Page("Navigation Page").WebButton("BTN_Next").ClickObject 
End Function
''###################################################################################################
''Function Name : CurrencyTruncateValidation
''Description : Verify Currency Truncate Test Min/Max field
''Author : Gallop
''###################################################################################################
Public Function CurrencyTruncateValidation
Environment.Value("StepName") = "CurrencyTruncateValidation"
	Browser("Currency").Page("Currency").WebEdit("TXT_CurrencyOneColumn").EnterText dtScenario.value("CurrencyOneColumn")
	Browser("Currency").Page("Currency").WebEdit("TXT_CurrencyOneColumn2").EnterText dtScenario.value("CurrencyOneColumn2")
	Browser("Currency").Page("Currency").WebEdit("TXT_CurrencyWholeOneColumn").EnterText dtScenario.value("CurrencyWholeOneColumn")
	Browser("Currency").Page("Currency").WebEdit("TXT_CurrencyWholeTwoColumn").EnterText dtScenario.value("CurrencyWholeTwoColumn")
	Browser("Currency").Page("Currency").WebEdit("TXT_CurrencyFormattedOneColumn").EnterText dtScenario.value("CurrencyFormattedOneColumn")
	Browser("Currency").Page("Currency").WebEdit("TXT_CurrencyFormattedTwoColumn").EnterText dtScenario.value("CurrencyFormattedTwoColumn")
	Browser("Currency").Page("Currency").WebEdit("TXT_CurrencyTruncateOneColumn").EnterText dtScenario.value("CurrencyTruncateOneColumn")
	Browser("Currency").Page("Currency").WebEdit("TXT_CurrencyTruncateTwoColumn").EnterText dtScenario.value("CurrencyTruncateTwoColumn")
	Browser("Currency").Page("Currency").WebEdit("TXT_CurrencyTruncateMinMax").EnterText dtScenario.value("CurrencyTruncateMinMax")
	Browser("Currency").Page("Currency").WebButton("BTN_Next").ClickObject 
	Browser("Currency").Page("Currency").WebElement("TWL_Currency Truncate Test").VerifyMandatoryMessage "Entry cannot exceed 5"
	Browser("Currency").Page("Currency").WebEdit("TXT_CurrencyTruncateMinMax").EnterText dtScenario.value("CurrencyTruncateMinMaxCorrectValue")
	Browser("Currency").Page("Currency").WebEdit("TXT_CurrencyTruncateTwoColumn").ClickObject 
	Browser("Currency").Page("Currency").WebElement("TWL_Currency Truncate Test").VerifyMessageNotExists "Entry cannot exceed 5"
End Function
''###################################################################################################
''Function Name : VerifyMessagesInStrings
''Description : Verify all mandatory messages in strings
''Author : Gallop
''###################################################################################################
Public Function VerifyMessagesInStrings
Environment.Value("StepName") = "VerifyMessagesInStrings"
	Browser("Strings").Page("Strings").WebButton("BTN_Next").WaitForObjectToLoad 5
	Browser("Strings").Page("Strings").WebButton("BTN_Next").ClickObject 
	Browser("Strings").Page("Strings").WebElement("TWL_One column string Required").VerifyMandatoryMessage "This is a required field"
	Browser("Strings").Page("Strings").WebElement("TWL_One Column String Tooltip").VerifyMandatoryMessage "This is a required field"
	Browser("Strings").Page("Strings").WebElement("TWL_Two Column String Required").VerifyMandatoryMessage "This is a required field"
	Browser("Strings").Page("Strings").WebElement("TWL_Alpha String- One Column").VerifyMandatoryMessage "This is a required field"
	Browser("Strings").Page("Strings").WebElement("TWL_Two Column Required Alpha").VerifyMandatoryMessage "This is a required field"
	Browser("Strings").Page("Strings").WebElement("TWL_Numeric String- One Column").VerifyMandatoryMessage "This is a required field"
	Browser("Strings").Page("Strings").WebElement("TWL_Numeric String- Two Column").VerifyMandatoryMessage "This is a required field"
End Function
''###################################################################################################
''Function Name : NavigateToStrings
''Description : avigate to Strings page
''Author : Gallop
''###################################################################################################
Public Function NavigateToStrings
Environment.Value("StepName") = "NavigateToStrings"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB _Home").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_Test Question Types").ClickObject 
	Browser("Navigation Page").Page("Navigation Page").WebElement("DWL_Navigate through Service").SelectListBox "Strings"
	Browser("Navigation Page").Page("Navigation Page").WebButton("BTN_Next").ClickObject 
	Browser("Strings").Page("Strings").WebElement("LBL_Strings").VerifyPageName 
End Function
''###################################################################################################
''Function Name : FillDataInStrings
''Description : Fill data in Strings
''Author : Gallop
''###################################################################################################
Public Function FillDataInStrings
Environment.Value("StepName") = "FillDataInStrings"
	Browser("Strings").Page("Strings").WebEdit("TXT_OneColumnRequired").EnterText dtScenario.value("OneColumnRequired")
	Browser("Strings").Page("Strings").WebEdit("TXT_TwoColumnRequired").EnterText dtScenario.value("TwoColumnRequired")
	Browser("Strings").Page("Strings").WebEdit("TXT_OneColumnToolTip").EnterText dtScenario.value("OneColumnToolTip")
	Browser("Strings").Page("Strings").WebEdit("TXT_AlphaOneColumnRequired").EnterText dtScenario.value("AlphaOneColumnRequired")
	Browser("Strings").Page("Strings").WebEdit("TXT_AlphaTwoColumnRequired").EnterText dtScenario.value("AlphaTwoColumnRequired")
	Browser("Strings").Page("Strings").WebEdit("TXT_NumericOneColumnRequired").EnterText dtScenario.value("NumericOneColumnRequired")
	Browser("Strings").Page("Strings").WebEdit("TXT_NumericTwoColumnRequired").EnterText dtScenario.value("NumericTwoColumnRequired")
End Function
''###################################################################################################
''Function Name : VerifyReadOnlyValuesInStrings
''Description : Verify read only values in strings
''Author : Gallop
''###################################################################################################
Public Function VerifyReadOnlyValuesInStrings
Environment.Value("StepName") = "VerifyReadOnlyValuesInStrings"
	Browser("Strings").Page("Strings").WebEdit("TXT_OneColumnRequired").ClickObject 
	Browser("Strings").Page("Strings").WebElement("TWL_One column string Required").VerifyMessageNotExists "This is a required field"
	Browser("Strings").Page("Strings").WebElement("TWL_One Column String Tooltip").VerifyMessageNotExists "This is a required field"
	Browser("Strings").Page("Strings").WebElement("TWL_Two Column String Required").VerifyMessageNotExists "This is a required field"
	Browser("Strings").Page("Strings").WebElement("TWL_Alpha String- One Column").VerifyMessageNotExists "This is a required field"
	Browser("Strings").Page("Strings").WebElement("TWL_Two Column Required Alpha").VerifyMessageNotExists "This is a required field"
	Browser("Strings").Page("Strings").WebElement("TWL_Numeric String- One Column").VerifyMessageNotExists "This is a required field"
	Browser("Strings").Page("Strings").WebElement("TWL_Numeric String- Two Column").VerifyMessageNotExists "This is a required field"
	Browser("Strings").Page("Strings").WebButton("BTN_Next").ClickObject 
	Browser("Strings Read Only").Page("Strings Read Only").WebElement("RWL_One column string Required").VerifyExists 
	Browser("Strings Read Only").Page("Strings Read Only").WebElement("RWL_One column string Required").VerifyDisplayedText dtScenario.value("OneColumnRequired")
	Browser("Strings Read Only").Page("Strings Read Only").WebElement("RWL_Two Column String Required").VerifyDisplayedText dtScenario.value("TwoColumnRequired")
	Browser("Strings Read Only").Page("Strings Read Only").WebElement("RWL_One Column String Tooltip").VerifyDisplayedText dtScenario.value("OneColumnToolTip")
	Browser("Strings Read Only").Page("Strings Read Only").WebElement("RWL_Alpha String- One Column Required").VerifyDisplayedText dtScenario.value("AlphaOneColumnRequired")
	Browser("Strings Read Only").Page("Strings Read Only").WebElement("RWL_Two Column Required Alpha").VerifyDisplayedText dtScenario.value("AlphaTwoColumnRequired")
	Browser("Strings Read Only").Page("Strings Read Only").WebElement("RWL_Numeric String- One Column").VerifyDisplayedText dtScenario.value("NumericOneColumnRequired")
	Browser("Strings Read Only").Page("Strings Read Only").WebElement("RWL_Numeric String- Two Column").VerifyDisplayedText dtScenario.value("NumericTwoColumnRequired")
End Function
''###################################################################################################
''Function Name : NavigateToDate
''Description : Navigate to Date page
''Author : Gallop
''###################################################################################################
Public Function NavigateToDate
Environment.Value("StepName") = "NavigateToDate"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_Test Question Types").ClickObject 
	Browser("Navigation Page").Page("Navigation Page").WebElement("DWL_Navigate through Service").SelectListBox "Date"
	Browser("Navigation Page").Page("Navigation Page").WebButton("BTN_Next").ClickObject 
End Function
''###################################################################################################
''Function Name : ValidateDatesinReadonlyPage
''Description : It Verifies whether the dates are displayed in read only page
''Author : Gallop
''###################################################################################################
Public Function ValidateDatesinReadonlyPage
Environment.Value("StepName") = "ValidateDatesinReadonlyPage"
	Browser("Dates Read Only").Page("Dates Read Only").WebElement("Date One Column Required").VerifyDisplayedText gfGetCurrentDate
	Browser("Dates Read Only").Page("Dates Read Only").WebElement("Date One Column Required2").VerifyDisplayedText gfGetCurrentDate
	Browser("Dates Read Only").Page("Dates Read Only").WebElement("Date Two Column Required").VerifyDisplayedText gfGetCurrentDate
	Browser("Dates Read Only").Page("Dates Read Only").WebElement("Date Min/Max Test - Max Size1").VerifyDisplayedText gfGetCurrentDate
	Browser("Dates Read Only").Page("Dates Read Only").WebElement("Date Min/Max Test - Min Size 0").VerifyDisplayedText gfGetCurrentDate
	Browser("Dates Read Only").Page("Dates Read Only").WebElement("Biz Date 1 Col Min 2 Max 5").VerifyDisplayedText gfGetBiz5Date
	Browser("Dates Read Only").Page("Dates Read Only").WebElement("Business Date One Column Required 2").VerifyDisplayedText gfGetCurrentDate
	Browser("Dates Read Only").Page("Dates Read Only").WebElement("Two Column Business Date").VerifyDisplayedText gfGetCurrentDate
	Browser("Dates Read Only").Page("Dates Read Only").WebElement("Business Date Min Test").VerifyDisplayedText gfGetCurrentDate
	Browser("Dates Read Only").Page("Dates Read Only").WebElement("Business Date Max Test").VerifyDisplayedText gfGetCurrentDate
	Browser("Dates Read Only").Page("Dates Read Only").WebElement("Date Month/Day One Column").VerifyDisplayedText gfGetDateMonth
	Browser("Dates Read Only").Page("Dates Read Only").WebElement("Date Month/Day One Column 2").VerifyDisplayedText gfGetDateMonth
	Browser("Dates Read Only").Page("Dates Read Only").WebElement("Date Month/Day Two Column").VerifyDisplayedText gfGetDateMonth
End Function
''###################################################################################################
''Function Name : ValidateMinMaxTestMinSize0
''Description : Verifying calender,previous and next month arrows,previous date validation in Min Max Test Min Size 0 field
''Author : Gallop
''###################################################################################################
Public Function ValidateMinMaxTestMinSize0
Environment.Value("StepName") = "ValidateMinMaxTestMinSize0"
	Browser("Dates").Page("Dates").WebElement("DTL_Date Min/Max Test - Max Size0").ClickOnCalenderIcon 
	Browser("Dates").Page("Dates").Link("CalendarLeftArrowIcon").VerifyCalenderNavigation "Not Movable"
	Browser("Dates").Page("Dates").Link("CalendarRightArrowIcon").VerifyCalenderNavigation "Movable"
	Browser("Dates").Page("Dates").WebTable("CalendarDates").VerifyPreviousDateNotShownInCalendar 
	Browser("Dates").Page("Dates").WebEdit("DateMinMaxTestMaxSize0").EnterText gfGetPastDate
	Browser("Dates").Page("Dates").WebEdit("DateMinMaxTestMaxSize1").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("BizDate1ColMin2Max5").EnterText gfGetBiz5Date
	Browser("Dates").Page("Dates").WebEdit("DateOneColoumnRequired").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("DateOneColoumnRequired2").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("DateTwoColumnRequired").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("BusinessDateOneColumnRequired2").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("TwoColumnBusinessDate").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("BusinessDateMinTest").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("BusinessDateMaxTest1").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("DateMonthDayOneColumn").EnterText gfGetDateMonth
	Browser("Dates").Page("Dates").WebEdit("DateMonthDayOneColumn2").EnterText gfGetDateMonth
	Browser("Dates").Page("Dates").WebEdit("DateMonthDayTwoColumn").EnterText gfGetDateMonth
	Browser("Dates").Page("Dates").WebButton("Next").ClickObject 
	Browser("Dates").Page("Dates").WebElement("DTL_Date Min/Max Test - Max Size0").VerifyMandatoryMessage "Date entry must be on or after "
	Browser("Dates").Page("Dates").WebEdit("DateMinMaxTestMaxSize0").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebButton("Next").ClickObject 
End Function
''###################################################################################################
''Function Name : VerifyDateOneColumnReqd
''Description : Verifying invalid date in  "Date One Column Required" 
''Author : Gallop
''###################################################################################################
Public Function VerifyDateOneColumnReqd
Environment.Value("StepName") = "VerifyDateOneColumnReqd"
	Browser("Dates").Page("Dates").WebEdit("DateOneColoumnRequired").EnterText "20/25/14"
	Browser("Dates").Page("Dates").WebEdit("DateOneColoumnRequired2").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("DateTwoColumnRequired").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("DateMinMaxTestMaxSize1").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("DateMinMaxTestMaxSize0").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("BizDate1ColMin2Max5").EnterText gfGetBiz5Date
	Browser("Dates").Page("Dates").WebEdit("BusinessDateOneColumnRequired2").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("TwoColumnBusinessDate").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("BusinessDateMinTest").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("BusinessDateMaxTest1").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebEdit("DateMonthDayOneColumn").EnterText gfGetDateMonth
	Browser("Dates").Page("Dates").WebEdit("DateMonthDayOneColumn2").EnterText gfGetDateMonth
	Browser("Dates").Page("Dates").WebEdit("DateMonthDayTwoColumn").EnterText gfGetDateMonth
	Browser("Dates").Page("Dates").WebButton("Next").ClickObject 
	Browser("Dates").Page("Dates").WebElement("DTL_Date One Column Required").VerifyMandatoryMessage "Invalid date"
	Browser("Dates").Page("Dates").WebEdit("DateOneColoumnRequired").EnterText gfGetCurrentDate
	Browser("Dates").Page("Dates").WebButton("Next").ClickObject 
End Function
''###################################################################################################
''Function Name : OneColumnStringRequiredValidation
''Description : Verify message when less than 2 characters is entered in one colum string required field in strings
''Author : Gallop
''###################################################################################################
Public Function OneColumnStringRequiredValidation
Environment.Value("StepName") = "OneColumnStringRequiredValidation"
	Browser("Strings").Page("Strings").WebButton("BTN_Next").VerifyExists 
	Browser("Strings").Page("Strings").WebEdit("TXT_OneColumnRequired").EnterText dtScenario.value("OneColumnRequired")
	Browser("Strings").Page("Strings").WebEdit("TXT_TwoColumnRequired").ClickObject 
	Browser("Strings").Page("Strings").WebElement("TWL_One column string Required").VerifyMandatoryMessage "Entry must be at least 2 character(s) long"
End Function
''###################################################################################################
''Function Name : VerifyTootlTipOnOneColumnStringTooltip
''Description : Verify tool tip on One Column String Tooltip Spellcheck
''Author : Gallop
''###################################################################################################
Public Function VerifyTootlTipOnOneColumnStringTooltip
Environment.Value("StepName") = "VerifyTootlTipOnOneColumnStringTooltip"
	Browser("Strings").Page("Strings").WebEdit("TXT_OneColumnToolTip").VerifyExists 
	Browser("Strings").Page("Strings").WebEdit("TXT_OneColumnToolTip").VerifyToolTip "This is a tooltip for a one column string"
End Function
''###################################################################################################
''Function Name : VerifyMaxLengthInOneColumMinMaxString
''Description : Verify max length allowed is 10  in One Column Min/Max String
''Author : Gallop
''###################################################################################################
Public Function VerifyMaxLengthInOneColumMinMaxString
Environment.Value("StepName") = "VerifyMaxLengthInOneColumMinMaxString"
	Browser("Strings").Page("Strings").WebElement("TWL_One column string Required").VerifyExists 
	Browser("Strings").Page("Strings").WebEdit("TXT_OneColumnMinMax").VerifyMaxLength 10,"ALPHANUMERIC"
End Function
''###################################################################################################
''Function Name : VerifySpellCheckInOneColumnStringSpellCheck
''Description : Verify spell check in One Column String Spell Check field
''Author : Gallop
''###################################################################################################
Public Function VerifySpellCheckInOneColumnStringSpellCheck
Environment.Value("StepName") = "VerifySpellCheckInOneColumnStringSpellCheck"
	Browser("Strings").Page("Strings").WebEdit("TXT_OneColumnSpellCheck").VerifyExists 
	Browser("Strings").Page("Strings").WebEdit("TXT_OneColumnSpellCheck").EnterText dtScenario.value("WrongWord")
	Browser("Strings").Page("Strings").WebElement("TWL_OneColumnStringSpell").VerifyExists 
	Browser("Strings").Page("Strings").WebElement("TWL_OneColumnStringSpell").ClickSpellCheckIcon 
	Browser("Strings").Frame("Window").WebList("DRP_dialogControl$Suggestions").VerifyExists 
	Browser("Strings").Frame("Window").WebList("DRP_dialogControl$Suggestions").VerifySpellCheck dtScenario.value("CorrectSpelling")
	Browser("Strings").Frame("Window").WebButton("BTN_Change").ClickObject 
	Browser("Strings").Dialog("SpellcheckPopUp").Static("MSG_TheSpellCheckiscomplete!").VerifyExists 
	Browser("Strings").Dialog("SpellcheckPopUp").WinButton("BTN_OK").ClickObject 
End Function
''###################################################################################################
''Function Name : VerifySplCharactersNotAllowedInStringsAlpha
''Description : Verify special characters are not allowed in all fields in String alpha only 
''Author : Gallop
''###################################################################################################
Public Function VerifySplCharactersNotAllowedInStringsAlpha
Environment.Value("StepName") = "VerifySplCharactersNotAllowedInStringsAlpha"
	Browser("Strings").Page("Strings").WebEdit("TXT_AlphaTwoColumTooltip").VerifyExists 
	Browser("Strings").Page("Strings").WebEdit("TXT_AlphaOneColumnRequired").VerifySpecialCharacterNotAllowed dtScenario.value("SpecialCharcters")
	Browser("Strings").Page("Strings").WebEdit("TXT_AlphaOneColumnNotRequired").VerifySpecialCharacterNotAllowed dtScenario.value("SpecialCharcters")
	Browser("Strings").Page("Strings").WebEdit("TXT_AlphaOneColumnSpellCheck").VerifySpecialCharacterNotAllowed dtScenario.value("SpecialCharcters")
	Browser("Strings").Page("Strings").WebEdit("TXT_AlphaOneColumnTooltip").VerifySpecialCharacterNotAllowed dtScenario.value("SpecialCharcters")
	Browser("Strings").Page("Strings").WebEdit("TXT_AlphaTwoColumTooltip").VerifySpecialCharacterNotAllowed dtScenario.value("SpecialCharcters")
	Browser("Strings").Page("Strings").WebEdit("TXT_AlphaTwoColumnRequired").VerifySpecialCharacterNotAllowed dtScenario.value("SpecialCharcters")
	Browser("Strings").Page("Strings").WebEdit("TXT_AlphaTwoColumnSpellCheck").VerifySpecialCharacterNotAllowed dtScenario.value("SpecialCharcters")
	Browser("Strings").Page("Strings").WebEdit("TXT_AlphaTwoColumnMinMax").VerifySpecialCharacterNotAllowed dtScenario.value("SpecialCharcters")
End Function
''###################################################################################################
''Function Name : VerifyWidthInStrings
''Description : Verify that One column questions are one column in width, two column questions are one column in width.
''Author : Gallop
''###################################################################################################
Public Function VerifyWidthInStrings
Environment.Value("StepName") = "VerifyWidthInStrings"
	Call VerifyColumnWidthInStrings()
End Function
''###################################################################################################
''Function Name : VerifyReadOnlyFields
''Description : Verify read only fields in String : TwoColumnReadOnlyString.one column String RO
''Author : Gallop
''###################################################################################################
Public Function VerifyReadOnlyFields
Environment.Value("StepName") = "VerifyReadOnlyFields"
	Browser("Strings").Page("Strings").WebElement("TWL_OneColStringRO").VerifyFieldIsReadOnly 
	Browser("Strings").Page("Strings").WebElement("TWL_TwoColumnReadOnlyString").VerifyFieldIsReadOnly 
End Function
''###################################################################################################
''Function Name : VerifyAlphaCharactersNotAllowedInStringsNumeric
''Description : Verify alpha characters are not allowed in all fields in String alpha only 
''Author : Gallop
''###################################################################################################
Public Function VerifyAlphaCharactersNotAllowedInStringsNumeric
Environment.Value("StepName") = "VerifyAlphaCharactersNotAllowedInStringsNumeric"
	Browser("Strings").Page("Strings").WebEdit("TXT_NumericTwoColumnRequired").VerifyExists 
	Browser("Strings").Page("Strings").WebEdit("TXT_NumericOneColumnRequired").VerifyAlphaCharacterNotAllowed dtScenario.value("AlphaCharchters")
	Browser("Strings").Page("Strings").VerifyAlphaCharacterNotAllowed dtScenario.value("AlphaCharchters")
	Browser("Strings").Page("Strings").WebEdit("TXT_NumericTwoColumnRequired").VerifyAlphaCharacterNotAllowed dtScenario.value("AlphaCharchters")
	Browser("Strings").Page("Strings").VerifyAlphaCharacterNotAllowed dtScenario.value("AlphaCharchters")
End Function
''###################################################################################################
''Function Name : VerifyMandatoryMessagesInDropDown
''Description : Verify mandatory messages displayed when clicking Next without selecting any value
''Author : Gallop
''###################################################################################################
Public Function VerifyMandatoryMessagesInDropDown
Environment.Value("StepName") = "VerifyMandatoryMessagesInDropDown"
	Browser("Drop Downs").Page("Drop Downs").WebElement("LBL_Drop Downs").VerifyPageName 
	Browser("Drop Downs").Page("Drop Downs").WebButton("BTN_Next").ClickObject 
	Browser("Drop Downs").Page("Drop Downs").WebElement("DWL_DropDownLongString").VerifyMandatoryMessage "Selection required"
	Browser("Drop Downs").Page("Drop Downs").WebElement("DWL_DropDown2Columnselect").VerifyMandatoryMessage "Selection required"
End Function
''###################################################################################################
''Function Name : NavigateToDropDown
''Description : Navigate to dropdown page
''Author : Gallop
''###################################################################################################
Public Function NavigateToDropDown
Environment.Value("StepName") = "NavigateToDropDown"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_Test Question Types").ClickObject 
	Browser("Navigation Page").Page("Navigation Page").WebElement("DWL_Navigate through Service").SelectListBox "Drop down lookup"
	Browser("Navigation Page").Page("Navigation Page").WebButton("BTN_Next").ClickObject 
End Function
''###################################################################################################
''Function Name : VerifyReadOnlyDropDown
''Description : Verify readOnly drop down and its tect
''Author : Gallop
''###################################################################################################
Public Function VerifyReadOnlyDropDown
Environment.Value("StepName") = "VerifyReadOnlyDropDown"
	Browser("Drop Downs").Page("Drop Downs").WebElement("LBL_Drop Downs").VerifyPageName 
	Browser("Drop Downs").Page("Drop Downs").WebElement("RWL_ReadOnlyDropDown").VerifyFieldIsReadOnly 
	Browser("Drop Downs").Page("Drop Downs").WebElement("RWL_ReadOnlyDropDown").VerifyDisplayedText "This is a long default value to determine if the value looks good once it is wrapped. This is the maximum default value length of two hundred and fifty characters. This is the maximum default value length of two hundred and fifty characters. Thanks."
End Function
''###################################################################################################
''Function Name : FillRequiredDetailsInYesNo
''Description : Fill all mandatory fields and selct No for all fields in Yes No Alert and click Next
''Author : Gallop
''###################################################################################################
Public Function FillRequiredDetailsInYesNo
Environment.Value("StepName") = "FillRequiredDetailsInYesNo"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_Yes/NoOneColumn").SelectRadioButton "YES"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_Yes/NoOneColumn2").SelectRadioButton "YES"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_Yes/NoTwoColumnUsually").SelectRadioButton "YES"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_ALERTTwoColumn").SelectRadioButton "NO"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_ALERTOneColumn").SelectRadioButton "NO"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_ALERTOneColumn2").SelectRadioButton "NO"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_ALERTDefaulted").SelectRadioButton "NO"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_ALERTDefaultedTwoColumn").SelectRadioButton "NO"
	Browser("Yes/No").Page("Yes/No").WebButton("BTN_Next").ClickObject 
	Browser("Yes/No Read Only").Page("Yes/No Read Only").WebElement("LBL_Yes/No").VerifyPageName 
End Function
''###################################################################################################
''Function Name : VerifyFirstValueInDropDownLongString
''Description : Select first value inDrop Down Long String and verify its value
''Author : Gallop
''###################################################################################################
Public Function VerifyFirstValueInDropDownLongString
Environment.Value("StepName") = "VerifyFirstValueInDropDownLongString"
	Browser("Drop Downs").Page("Drop Downs").WebElement("DWL_DropDownLongString").SelectListBoxByIndex 1
	Browser("Drop Downs").Page("Drop Downs").WebElement("DWL_DropDownLongString").VerifyValueSelectedInDropDown "This is a test for a two column lookup"
End Function
''###################################################################################################
''Function Name : VerifyMandatoryMessagesInYesNo
''Description : Verimanadatory messages displayed in Yes No page
''Author : Gallop
''###################################################################################################
Public Function VerifyMandatoryMessagesInYesNo
Environment.Value("StepName") = "VerifyMandatoryMessagesInYesNo"
	Browser("Yes/No").Page("Yes/No").WebButton("BTN_Next").ClickObject 
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_Yes/NoOneColumn").VerifyMandatoryMessage "Selection required"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_Yes/NoOneColumn2").VerifyMandatoryMessage "Selection required"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_Yes/NoTwoColumnUsually").VerifyMandatoryMessage "Selection required"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_ALERTOneColumn").VerifyMandatoryMessage "Selection required"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_ALERTOneColumn2").VerifyMandatoryMessage "Selection required"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_ALERTTwoColumn").VerifyMandatoryMessage "Selection required"
End Function
''###################################################################################################
''Function Name : VerifyYellowBackground
''Description : For every Yes answer in the "Yes/No ALERT" field, Verify that there should be a yellow alert icon next to the value in the read-only page
''Author : Gallop
''###################################################################################################
Public Function VerifyYellowBackground
Environment.Value("StepName") = "VerifyYellowBackground"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_Yes/NoOneColumn").SelectRadioButton "NO"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_Yes/NoOneColumn2").SelectRadioButton "YES"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_Yes/NoTwoColumnUsually").SelectRadioButton "NO"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_ALERTOneColumn").SelectRadioButton "YES"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_ALERTTwoColumn").SelectRadioButton "YES"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_ALERTOneColumn2").SelectRadioButton "YES"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_ALERTDefaulted").SelectRadioButton "NO"
	Browser("Yes/No").Page("Yes/No").WebElement("RDL_ALERTDefaultedTwoColumn").SelectRadioButton "YES"
	Browser("Yes/No").Page("Yes/No").WebButton("BTN_Next").ClickObject 
	Browser("Yes/No Read Only").Page("Yes/No Read Only").WebElement("LBL_Yes/No").VerifyPageName 
	Browser("Yes/No Read Only").Page("Yes/No Read Only").WebElement("RWL_ALERTOneColumn").VerifyYellowAlert 
	Browser("Yes/No Read Only").Page("Yes/No Read Only").WebElement("RWL_ALERTOneColumn2").VerifyYellowAlert 
	Browser("Yes/No Read Only").Page("Yes/No Read Only").WebElement("RWL_ALERTTwoColumn").VerifyYellowAlert 
	Browser("Yes/No Read Only").Page("Yes/No Read Only").WebElement("RWL_ALERTDefaulted").VerifyYellowAlertNotExist 
	Browser("Yes/No Read Only").Page("Yes/No Read Only").WebElement("RWL_ALERTDefaultedTwoColumn").VerifyYellowAlert 
End Function
''###################################################################################################
''Function Name : NavigatetoDocument
''Description : Navigates to Document Workflow Testing
''Author : Gallop
''###################################################################################################
Public Function NavigatetoDocument
Environment.Value("StepName") = "NavigatetoDocument"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_Test Document Tab").ClickObject 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").PageSync 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").WaitForObjectToLoad 

End Function
''###################################################################################################
''Function Name : ClickonNextandSubmit
''Description : Click on Next and Submit Button
''Author : Gallop
''###################################################################################################
Public Function ClickonNextandSubmit
Environment.Value("StepName") = "ClickonNextandSubmit"
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
End Function
''###################################################################################################
''Function Name : ValidateEntryPage
''Description : Document Workflow Testing Entry Page� displays properly.
''Author : Gallop
''###################################################################################################
Public Function ValidateEntryPage
Environment.Value("StepName") = "ValidateEntryPage"
	Browser("Entry").Page("Entry").WebElement("LBL_DocumentWorkflowTesting").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_RequestDetails").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_RequestDetails").VerifyExists 
	Browser("Summary").Page("Summary").WebTable("TBL_RequestDetails").verifyNumberexistIncellByText "ID"
	Browser("Summary").Page("Summary").WebTable("TBL_RequestDetails").VerifyValuesInRow "Status|Approval,Document Pending Receipt"
	Browser("Summary").Page("Summary").WebTable("TBL_RequestDetails").VerifyValuesInRow "Created By|Chris Jarrell"
	Browser("Summary").Page("Summary").WebTable("TBL_RequestDetails").VerifyValuesInRow "Created On|"&gfGetCurrentDate
	Browser("Summary").Page("Summary").WebTable("TBL_RequestDetails").VerifyDueDate "Due Date"
	Browser("Summary").Page("Summary").WebTable("TBL_WorkflowDetails").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("IMG_DocumentTestEntry").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("IMG_Approval").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Details").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_DocumentTest").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Documents").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LBL_DocumentsDue").WaitForObjectToLoad "10"
	Browser("Summary").Page("Summary").VerifyDisplayedText "DocumentDue ByAddedNotesAttachmentsFormAttachment"

End Function
''###################################################################################################
''Function Name : ValidateDocuments
''Description : Verifies document types and their initial attachements
''Author : Gallop
''###################################################################################################
Public Function ValidateDocuments
Environment.Value("StepName") = "ValidateDocuments"
	Browser("Entry").Page("Entry").WebElement("LBL_DocumentWorkflowTesting").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").EnterText "Testing"
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_Documents").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_Documents").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_PrintCoverSheets").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_ViewPrintAllForms").VerifyExists 

End Function
''###################################################################################################
''Function Name : ValidateAttachDocument
''Description : All documents are added correctly
''Author : Gallop
''###################################################################################################
Public Function ValidateAttachDocument
Environment.Value("StepName") = "ValidateAttachDocument"
	Browser("Entry").Page("Entry").WebElement("LBL_DocumentWorkflowTesting").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").EnterText "Testing"
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyValuesInRow "Attachment Only|0"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyandClickLinkInRow "Attachment Only|Attach Document"
	Browser("Summary").Page("Summary").WebFile("WFL_Uploader").clickobject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").EnterFileName 
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyStatusofAttachment "Attachment Only|status: Complete"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyValuesInRow "Attachment Only|1"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
End Function
''###################################################################################################
''Function Name : ValidateAddDocument
''Description : Add Document button - miscellaneous document
''Author : Gallop
''###################################################################################################
Public Function ValidateAddDocument
Environment.Value("StepName") = "ValidateAddDocument"
	Browser("Entry").Page("Entry").WebElement("LBL_DocumentWorkflowTesting").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").EnterText "Testing"
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_Documents").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_Documents").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").ClickObject 
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Misc. Document (name input)"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Miscellaneous"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Optional Document for Doc Test"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").SelectWebList "Miscellaneous"
	Browser("Summary").Page("Summary").WebEdit("TXT_MiscellanousInput").VerifyExists 
	Browser("Summary").Page("Summary").WebEdit("TXT_MiscellanousInput").EnterText "Testing"
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").EnterFileName "pass.png"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").WaitForObjectToLoad 5
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").WaitForObjectToLoad 5
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyStatusofAttachment "Testing|status: Complete"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyandClickLinkInRow "Testing|view attachment"
	Browser("Summary").WinObject("Notification").WinButton("BTN_SaveAsDropDown").WaitinSec 5
	Browser("Summary").WinObject("Notification").WinButton("BTN_SaveAsDropDown").VerifyExists 
	Browser("Summary").WinObject("Notification").WinButton("BTN_SaveAsDropDown").ClickObject 
	Call DownloadSaveAs()
	Call StaticWait(5)
	Browser("Summary").Dialog("Save As").WinEdit("TXT_FileName").VerifyExists 
	Browser("Summary").Dialog("Save As").WinEdit("TXT_FileName").DownloadFile "Download.png"
	Browser("Summary").Dialog("Save As").WinButton("BTN_Save").WaitinSec 5
	Browser("Summary").Dialog("Save As").WinButton("BTN_Save").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").WaitForObjectToLoad 5
	Call VerifyDownlodedFileExist()

End Function
''###################################################################################################
''Function Name : ValidateAddDocumentOptional
''Description : Add Document button - Add optional document
''Author : Gallop
''###################################################################################################
Public Function ValidateAddDocumentOptional
Environment.Value("StepName") = "ValidateAddDocumentOptional"
	Browser("Entry").Page("Entry").WebElement("LBL_DocumentWorkflowTesting").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").EnterText "Testing"
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyandClickLinkInRow "Word Template w Review|Attach Document"
	Browser("Summary").Page("Summary").WebFile("WFL_Uploader").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").EnterFileName 
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyStatusofAttachment "Word Template w Review|status: Pending Review (Special)"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Word Template w Review|Accept/Approve document"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Word Template w Review|Reject document"
End Function
''###################################################################################################
''Function Name : ValidateRejectDocument
''Description :  Reject Document button brings you to the Reject Document Page
''Author : Gallop
''###################################################################################################
Public Function ValidateRejectDocument
Environment.Value("StepName") = "ValidateRejectDocument"
	Browser("Entry").Page("Entry").WebElement("LBL_DocumentWorkflowTesting").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").EnterText "Testing"
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyValuesInRow "PDF Template w Review|0"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyandClickLinkInRow "PDF Template w Review|Attach Document"
	Browser("Summary").Page("Summary").WebFile("WFL_Uploader").clickobject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").EnterFileName 
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyValuesInRow "PDF Template w Review|1"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyandClickLinkInRow "PDF Template w Review|Reject document"
	Browser("Summary").Page("Summary").WebEdit("TXT_RejectReason").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_RejectDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_DoNOTRejectDocument").VerifyExists 
End Function
''###################################################################################################
''Function Name : ValidateAddDocumentMiscellanous
''Description : Add Document button - miscellaneous document
''Author : Gallop
''###################################################################################################
Public Function ValidateAddDocumentMiscellanous
Environment.Value("StepName") = "ValidateAddDocumentMiscellanous"
	Browser("Entry").Page("Entry").WebElement("LBL_DocumentWorkflowTesting").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").EnterText "Testing"
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").ClickObject 
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Misc. Document (name input)"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Miscellaneous"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Optional Document for Doc Test"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").SelectWebList "Miscellaneous"
	Browser("Summary").Page("Summary").WebEdit("TXT_MiscellanousInput").VerifyExists 
	Browser("Summary").Page("Summary").WebEdit("TXT_MiscellanousInput").EnterText "Testing"
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").EnterFileName 
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyStatusofAttachment "Testing|status: Complete"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyValuesInRow "Testing|1"
End Function
''###################################################################################################
''Function Name : ValidateAddDocumentMisc
''Description : Add Document button - miscellaneous document
''Author : Gallop
''###################################################################################################
Public Function ValidateAddDocumentMisc
Environment.Value("StepName") = "ValidateAddDocumentMisc"
	Browser("Entry").Page("Entry").WebElement("LBL_DocumentWorkflowTesting").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").EnterText "Testing"
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").ClickObject 
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Misc. Document (name input)"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Miscellaneous"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Optional Document for Doc Test"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").SelectWebList "Miscellaneous"
	Browser("Summary").Page("Summary").WebEdit("TXT_MiscellanousInput").VerifyExists 
	Browser("Summary").Page("Summary").WebEdit("TXT_MiscellanousInput").EnterText "Testing"
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").EnterFileName 
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyStatusofAttachment "Testing|status: Complete"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyValuesInRow "Testing|1"
End Function
''###################################################################################################
''Function Name : NavigatetoWorkflow
''Description : Navigates to Workflow Testing
''Author : Gallop
''###################################################################################################
Public Function NavigatetoWorkflow
Environment.Value("StepName") = "NavigatetoWorkflow"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_TestWorkflowTab").ClickObject 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").PageSync 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").WaitForObjectToLoad 

End Function
''###################################################################################################
''Function Name : ValidateWorkflow
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateWorkflow
Environment.Value("StepName") = "ValidateWorkflow"
	Browser("Entry").Page("Entry").WebButton("BTN_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").VerifyExists 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyExists 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyTextInTable "Step 1"
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyValuesInRow "Step 1 |In Progress"
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyValuesInRow "Step 1 |"&gfGetCurrentDateDDMMYY
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyValuesInRow "Step 1 |Chris Jarrell"
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyValuesInRow "Step 1 |"&gfGetCurrentDateDDMMYY
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyButtonInRow "Step 1 |BUTTON"
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyandClickButtonInRow "Step 1 |BUTTON"
	Browser("Summary").Page("Summary").WebElement("LNK_Work").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LNK_Work").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LNK_Assign").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LNK_DueDate").VerifyExists 

End Function
''###################################################################################################
''Function Name : ValidateAssign
''Description : This component validates Assign button functionality
''Author : Gallop
''###################################################################################################
Public Function ValidateAssign
Environment.Value("StepName") = "ValidateAssign"
	Browser("Entry").Page("Entry").WebButton("BTN_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").PageSync 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").VerifyExists 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyandClickButtonInRow "Step 1 |BUTTON"
	Browser("Summary").Page("Summary").WebElement("LNK_Assign").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LNK_Assign").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_SaveAssignment").PageSync 
	Browser("Summary").Page("Summary").WebButton("BTN_SaveAssignment").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebButton("BTN_SaveAssignment").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_ClearAssignment").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_Cancel").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("DWL_NewAssignment").SelectListBox "Admin, Sally"
	Browser("Summary").Page("Summary").WebButton("BTN_SaveAssignment").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_SaveAssignment").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyExists 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyValuesInRow "Step 1|Sally Admin"
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyDataInCell "Step 1|Assigned To|Sally Admin"

End Function
''###################################################################################################
''Function Name : ValidateDueDate
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateDueDate
Environment.Value("StepName") = "ValidateDueDate"
	Browser("Entry").Page("Entry").WebButton("BTN_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").PageSync 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").VerifyExists 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyandClickButtonInRow "Step 1 |BUTTON"
	Browser("Summary").Page("Summary").WebElement("LNK_DueDate").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LNK_DueDate").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LNK_DueDate").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LBL_TaskName").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LBL_CreatedOn").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LBL_OriginalDueDate").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LBL_NewDueDate").VerifyExists 
	Browser("Summary").Page("Summary").WebEdit("TXT_NewDueDate").VerifyExists 
	Browser("Summary").Page("Summary").WebEdit("TXT_DueDateReason").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_UpdateDueDate").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_Cancel").VerifyExists 
	Browser("Summary").Page("Summary").WebEdit("TXT_NewDueDate").EnterText gfGetNextDate
	Browser("Summary").Page("Summary").WebEdit("TXT_DueDateReason").EnterText "Testing"
	Browser("Summary").Page("Summary").WebButton("BTN_UpdateDueDate").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyValuesInRow "Step 1 |"&gfGetNextDateMMDDYY

End Function
''###################################################################################################
''Function Name : ValidateUserPermissions
''Description :  user cannot work a task if they do not have permission.
''Author : Gallop
''###################################################################################################
Public Function ValidateUserPermissions
Environment.Value("StepName") = "ValidateUserPermissions"
	Browser("Entry").Page("Entry").WebButton("BTN_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LBL_Step1_Work").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_Proceed").WaitForObjectToLoadWithTime 60
	Browser("Summary").Page("Summary").WebButton("BTN_Proceed").PageSync 
	Browser("Summary").Page("Summary").WebButton("BTN_Proceed").clickProceed 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").WaitForObjectToLoadWithTime 10
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").PageSync 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyTextInTable "Step 1 "
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyValuesInRow "Step 1 |Proceed"
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyActionsareNotDisplayedInRow "Step 1|Proceed"

End Function
''###################################################################################################
''Function Name : ValidateWorkingTask
''Description : Working a task completes the task and creates the next task in the workflow
''Author : Gallop
''###################################################################################################
Public Function ValidateWorkingTask
Environment.Value("StepName") = "ValidateWorkingTask"
	Browser("Entry").Page("Entry").WebButton("BTN_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").PageSync 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").PageSync 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LBL_Step1_Work").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_Proceed").WaitForObjectToLoadWithTime 60
	Browser("Summary").Page("Summary").WebButton("BTN_Proceed").PageSync 
	Browser("Summary").Page("Summary").WebButton("BTN_Proceed").clickProceed 
	Browser("Summary").Page("Summary").WebElement("TAB_Audit").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("TAB_Audit").WaitForObjectToLoadWithTime 20
	Browser("Summary").Page("Summary").WebElement("TAB_Audit").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_ViewAll").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebButton("BTN_ViewAll").PageSync 
	Browser("Summary").Page("Summary").WebButton("BTN_ViewAll").ClickObject 
	Browser("Summary").Page("Summary").WebElement("CWL_ViewAll").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("CWL_ViewAll").PageSync 
	Browser("Summary").Page("Summary").WebElement("CWL_ViewAll").SelectCheckBox 
	Browser("Summary").Page("Summary").WebTable("TBL_AuditHistory").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_AuditHistory").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_AuditHistory").VerifyValuesInRow "Completed Task|Performed action 'Proceed' on task 'Step 1'"
	Browser("Summary").Page("Summary").WebTable("TBL_AuditHistory").VerifyValuesInRow "Created Regular Task|Added task Step 2 - Path 1"
	Browser("Summary").Page("Summary").WebTable("TBL_AuditHistory").VerifyValuesInRow "Added task Step 2 - Path 2|Created Regular Task"
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").PageSync 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyTextInTable "Step 1 "
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyValuesInRow "Step 1 |Proceed"
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyandClickButtonInRow "Step 1 |BUTTON"
	Browser("Summary").Page("Summary").WebElement("LNK_Reopen").VerifyExists 

End Function
''###################################################################################################
''Function Name : ValidateAddOptionalDocument
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateAddOptionalDocument
Environment.Value("StepName") = "ValidateAddOptionalDocument"
	Browser("Entry").Page("Entry").WebElement("LBL_DocumentWorkflowTesting").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").EnterText "QA Testing"
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_Documents").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_Documents").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").ClickObject 
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Misc. Document (name input)"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Miscellaneous" 
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Optional Document for Doc Test"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").SelectWebList "Optional Document for Doc Test"
	Browser("Summary").Page("Summary").WebEdit("TXT_MiscellanousInput").VerifyNotExists 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").EnterFileName "Optional Document for Doc Test.pdf"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").WaitForObjectToLoad 5
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").WaitForObjectToLoad 5
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyStatusofAttachment "Optional Document for Doc Test|status: Pending Review"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|view attachment"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyandClickLinkInRow "Optional Document for Doc Test|view attachment"
	Browser("Summary").WinObject("Notification").WinButton("BTN_SaveAsDropDown").WaitinSec 5
	Call StaticWait(5)
	Browser("Summary").WinObject("Notification").WinButton("BTN_SaveAsDropDown").VerifyExists 
	Browser("Summary").WinObject("Notification").WinButton("BTN_SaveAsDropDown").ClickObject 
	Call StaticWait()
	Call DownloadSaveAs()
	Browser("Summary").Dialog("Save As").WinEdit("TXT_FileName").DownloadFile "IFS_Office_Version.png"
	Browser("Summary").Dialog("Save As").WinButton("BTN_Save").WaitinSec 5
	Browser("Summary").Dialog("Save As").WinButton("BTN_Save").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|view attachment"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|View document details and version history"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|Accept/Approve document"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|Reject document"

End Function
''###################################################################################################
''Function Name : ValidateDocumentApproval
''Description : Accept Document functionality is working as intended
''Author : Gallop
''###################################################################################################
Public Function ValidateDocumentApproval
Environment.Value("StepName") = "ValidateDocumentApproval"
	Browser("Entry").Page("Entry").WebElement("LBL_DocumentWorkflowTesting").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").EnterText " QA Rules!"
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").ClickObject 
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Optional Document for Doc Test"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").SelectWebList "Optional Document for Doc Test"
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").EnterFileName "Optional Document for Doc Test.pdf"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyStatusofAttachment "Optional Document for Doc Test|status: Pending Review"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|view attachment"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|Reject document"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|View document details and version history"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|Accept/Approve document"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyandClickLinkInRow "Optional Document for Doc Test|Accept/Approve document"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyStatusofAttachment "Optional Document for Doc Test|status: Pending Imaging"

End Function
''###################################################################################################
''Function Name : ValidateDocumentReject
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateDocumentReject
Environment.Value("StepName") = "ValidateDocumentReject"
	Browser("Entry").Page("Entry").WebElement("LBL_DocumentWorkflowTesting").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").EnterText "QA Rules!"
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_Documents").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").ClickObject 
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Optional Document for Doc Test"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").SelectWebList "Optional Document for Doc Test"
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").EnterFileName "Optional Document for Doc Test.pdf"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyStatusofAttachment "Optional Document for Doc Test|status: Pending Review"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|view attachment"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|Reject document"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|View document details and version history"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|Accept/Approve document"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyStatusofAttachment "Optional Document for Doc Test|status: Pending Review"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyandClickLinkInRow "Optional Document for Doc Test|Reject document"
	Browser("Summary").Page("Summary").WebEdit("TXT_RejectReason").PageSync 
	Browser("Summary").Page("Summary").WebEdit("TXT_RejectReason").EnterText "QA Testing document rejection reason"
	Browser("Summary").Page("Summary").WebButton("BTN_RejectDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_DoNOTRejectDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_RejectDocument").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyStatusofAttachment "Optional Document for Doc Test|status: Rejected"

End Function
''###################################################################################################
''Function Name : ValidateRejectCanBeCancelled
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateRejectCanBeCancelled
Environment.Value("StepName") = "ValidateRejectCanBeCancelled"
	Browser("Entry").Page("Entry").WebElement("LBL_DocumentWorkflowTesting").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").EnterText "IFS QA Automated Testing" 
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_Documents").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").ClickObject 
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").VerifyValueinWebList "Optional Document for Doc Test"
	Browser("Summary").Page("Summary").WebList("DRP_DocumentOptions").SelectWebList "Optional Document for Doc Test"
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").EnterFileName "Optional Document for Doc Test.pdf"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").WaitForObjectToLoad 5
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyStatusofAttachment "Optional Document for Doc Test|status: Pending Review"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|view attachment"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|Reject document"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|View document details and version history"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|Accept/Approve document"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyStatusofAttachment "Optional Document for Doc Test|status: Pending Review"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyandClickLinkInRow "Optional Document for Doc Test|Reject document"
	Browser("Summary").Page("Summary").WebButton("BTN_DoNOTRejectDocument").PageSync 
	Browser("Summary").Page("Summary").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_RejectDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebEdit("TXT_RejectReason").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_DoNOTRejectDocument").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").VerifyLinkInRow "Optional Document for Doc Test|Reject document"

End Function
''###################################################################################################
''Function Name : ValidateCancelAssignment
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateCancelAssignment
Environment.Value("StepName") = "ValidateCancelAssignment"
	Browser("Entry").Page("Entry").WebButton("BTN_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").VerifyExists 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyandClickButtonInRow "Step 1 |BUTTON"
	Browser("Summary").Page("Summary").WebElement("LNK_Assign").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LNK_Assign").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_Cancel").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_Cancel").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyValuesInRow "Step 1 |Chris Jarrell"

End Function
''###################################################################################################
''Function Name : ValidateCancelDueDate
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateCancelDueDate
Environment.Value("StepName") = "ValidateCancelDueDate"
	Browser("Entry").Page("Entry").WebButton("BTN_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").VerifyExists 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyandClickButtonInRow "Step 1 |BUTTON"
	Browser("Summary").Page("Summary").WebElement("LNK_DueDate").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LNK_DueDate").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LNK_DueDate").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LBL_TaskName").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LBL_CreatedOn").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LBL_OriginalDueDate").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LBL_NewDueDate").VerifyExists 
	Browser("Summary").Page("Summary").WebEdit("TXT_NewDueDate").VerifyExists 
	Browser("Summary").Page("Summary").WebEdit("TXT_DueDateReason").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_UpdateDueDate").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_Cancel").VerifyExists 
	Browser("Summary").Page("Summary").WebEdit("TXT_NewDueDate").EnterText gfGetNextDate
	Browser("Summary").Page("Summary").WebEdit("TXT_DueDateReason").EnterText "QA Automated Testing"
	Browser("Summary").Page("Summary").WebButton("BTN_Cancel").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyDueDate "Due Date"

End Function
''###################################################################################################
''Function Name : ValidateClearAssignment
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateClearAssignment
Environment.Value("StepName") = "ValidateClearAssignment"
	Browser("Entry").Page("Entry").WebButton("BTN_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_Workflow").VerifyExists 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyValuesInRow "Step 1 |Chris Jarrell"
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyandClickButtonInRow "Step 1 |BUTTON"
	Browser("Summary").Page("Summary").WebElement("LNK_Assign").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LNK_Assign").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LNK_Assign").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_SaveAssignment").PageSync 
	Browser("Summary").Page("Summary").WebButton("BTN_SaveAssignment").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebButton("BTN_SaveAssignment").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_ClearAssignment").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_Cancel").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_ClearAssignment").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyDataInCell "Step 1 |Assigned To|"
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyandClickButtonInRow "Step 1 |BUTTON"
	Browser("Summary").Page("Summary").WebElement("LNK_Assign").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LNK_Assign").ClickObject 
	Browser("Summary").Page("Summary").WebElement("DWL_NewAssignment").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("DWL_NewAssignment").SelectListBox "Admin, Sally"
	Browser("Summary").Page("Summary").WebButton("BTN_SaveAssignment").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_SaveAssignment").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyExists 
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyValuesInRow "Step 1|Sally Admin"
	Browser("Summary").Page("Summary").WebTable("TBL_Task").VerifyDataInCell "Step 1|Assigned To|Sally Admin"

End Function
''###################################################################################################
''Function Name : AddNoteinAudit
''Description : Audit tab - Can Add note
''Author : Gallop
''###################################################################################################
Public Function AddNoteinAudit
Environment.Value("StepName") = "AddNoteinAudit"
	Browser("Entry").Page("Entry").WebButton("BTN_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Audit").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("TAB_Audit").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Audit").ClickObject 
	Browser("Summary").Page("Summary").WebElement("TAB_Audit").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("TAB_Audit").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_ViewAll").PageSync 
	Browser("Summary").Page("Summary").WebButton("BTN_ViewAll").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_AddNote").PageSync 
	Browser("Summary").Page("Summary").WebElement("LBL_AddNote").ClickObject 
	Browser("Summary").Page("Summary").WebEdit("TXT_AddNote").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_CancelinAddNoteAudit").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_AddNote").VerifyExists 
	Browser("Summary").Page("Summary").WebEdit("TXT_AddNote").EnterText "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc congue, est nec congue consectetur, ex erat mattis nisi, quis sodales"
	Browser("Summary").Page("Summary").WebButton("BTN_AddNote").ClickObject 
	Browser("Summary").Page("Summary").WebTable("TBL_AuditHistory").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_History").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_History").VerifyvalueInSpecificCell 2,4,"Note"
	Browser("Summary").Page("Summary").WebTable("TBL_History").VerifyvalueInSpecificCell 2,5,"Entered Note: Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc congue, est nec congue consectetur, ex erat mattis nisi, quis sodales"

End Function
''###################################################################################################
''Function Name : VerifyTaskHistoryinAudit
''Description :  Audit tab - displays task history
''Author : Gallop
''###################################################################################################
Public Function VerifyTaskHistoryinAudit
Environment.Value("StepName") = "VerifyTaskHistoryinAudit"
	Browser("Entry").Page("Entry").WebButton("BTN_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_WorkFlow").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Audit").ClickObject 
	Browser("Summary").Page("Summary").WebElement("TAB_Audit").PageSync 
	Browser("Summary").Page("Summary").WebElement("TAB_Audit").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("TAB_Audit").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("TAB_Audit").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_ViewAll").PageSync 
	Browser("Summary").Page("Summary").WebButton("BTN_ViewAll").ClickObject 
	Browser("Summary").Page("Summary").WebElement("CWL_ViewAll").PageSync 
	Browser("Summary").Page("Summary").WebElement("CWL_ViewAll").SelectCheckBox 
	Browser("Summary").Page("Summary").WebTable("TBL_History").PageSync 
	Browser("Summary").Page("Summary").WebTable("TBL_History").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_History").VerifyDatainAuditHistory "Date|"&gfGetCurrentDateDDMMYY
	Browser("Summary").Page("Summary").WebTable("TBL_History").VerifyvalueInSpecificCell 2,3, "Jarrell, Chris (chrisj)"
	Browser("Summary").Page("Summary").WebTable("TBL_History").VerifyvalueInSpecificCell 2,4, "Assigned Task"
	Browser("Summary").Page("Summary").WebTable("TBL_History").VerifyvalueInSpecificCell 2,5,"Chris Jarrell assigned to task Step 1"
	Browser("Summary").Page("Summary").WebTable("TBL_History").VerifyDateInRow "Created Regular Task|"&gfGetCurrentDateDDMMYY
	Browser("Summary").Page("Summary").WebTable("TBL_History").VerifyvalueInSpecificCell 3,3,"Jarrell, Chris (chrisj)"
	Browser("Summary").Page("Summary").WebTable("TBL_History").VerifyvalueInSpecificCell 3,4, "Created Regular Task"
	Browser("Summary").Page("Summary").WebTable("TBL_History").VerifyvalueInSpecificCell 3,5, "Added task Step 1"

End Function
''###################################################################################################
''Function Name : VerifyCheckRequestPage
''Description : Verify check request page
''Author : Gallop
''###################################################################################################
Public Function VerifyCheckRequestPage
Environment.Value("StepName") = "VerifyCheckRequestPage"
	Browser("Account Home").Page("Account Home").WebButton("BTN_Check").ClickObject 
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("LBL_AccountSummary").VerifyPageName 
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("LBL_AccountSummary").VerifyExists 
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("LBL_AccountDetails").VerifyExists 
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("LBL_FutureDatedTransaction").VerifyExists 
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("LBL_CheckInstructions").VerifyExists 
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("LBL_DistributionAmount").VerifyExists 
	Browser("Check Request Entry").Page("Check Request Entry").WebEdit("TXT_DistributionAmount").EnterText dtScenario.value("DistributionAmount")
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("DWL_DeliveryMethod").ClickObject 
	Browser("Check Request Entry").Page("Check Request Entry").WebEdit("TXT_DistributionAmount").VerifyValue dtScenario.value("DistributionAmountFormatted")
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("RWL_Total DebitToAccount").VerifyDisplayedText dtScenario.value("DistributionAmountFormatted")
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("RWL_NetCashAmount").VerifyDisplayedText dtScenario.value("DistributionAmountFormatted")
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("DWL_DeliveryMethod").SelectListBox dtScenario.value("DeliveryMethod")
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("DWL_IsASignatureRequired").VerifyExists 
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("LBL_Fees").VerifyExists 
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("RWL_Delivery Fee").VerifyDisplayedText dtScenario.value("DeliveryFee")
	Browser("Check Request Entry").Page("Check Request Entry").WebElement("DWL_Charge To").VerifyValueSelectedInDropDown dtScenario.value("ChargeTo")
	Browser("Check Request Entry").Page("Check Request Entry").WebButton("BTN_Next").ClickObject 
End Function
''###################################################################################################
''Function Name : NavigateToAccountSearch
''Description : Navigate to Account Search
''Author : Gallop
''###################################################################################################
Public Function NavigateToAccountSearch
Environment.Value("StepName") = "NavigateToAccountSearch"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_AssetMovementSRs").ClickObject 
	Browser("IFS").Page("IFS").WebElement("LBL_Account Search").PageSync 
	Browser("IFS").Page("IFS").WebElement("LBL_Account Search").VerifyPageName 

End Function
''###################################################################################################
''Function Name : SearchAccountAndVerifyAccountSummary
''Description : Search with account number and verify account summary
''Author : Gallop
''###################################################################################################
Public Function SearchAccountAndVerifyAccountSummary
Environment.Value("StepName") = "SearchAccountAndVerifyAccountSummary"
	Browser("Account Search").Page("Account Search").WebEdit("TXT_AccountSearch").VerifyExists 
	Browser("Account Search").Page("Account Search").WebEdit("TXT_AccountSearch").EnterText dtScenario.value("Account")
	Browser("Account Search").Page("Account Search").WebButton("BTN_Search").ClickObject 
	Browser("Account Home").Page("Account Home").WebElement("LBL_AccountSummary").VerifyPageName 
	Browser("Account Home").Page("Account Home").WebElement("LBL_AccountSummary").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_Rep Number").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_Account #").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_SSN").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_AccountClass").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_Registration").VerifyExists 

End Function
''###################################################################################################
''Function Name : VerifyAccountSummaryPageAfterSubmit
''Description : Verify Account Summary Page After Submit
''Author : Gallop
''###################################################################################################
Public Function VerifyAccountSummaryPageAfterSubmit
Environment.Value("StepName") = "VerifyAccountSummaryPageAfterSubmit"
	Browser("Confirmation").Page("Confirmation").WebElement("LBL_Documents").VerifyPageName 
	Browser("Confirmation").Page("Confirmation").WebElement("RWL_ConfirmationYellowAlert1").VerifyYellowBanner 
	Browser("Confirmation").Page("Confirmation").WebElement("RWL_ConfirmationYellowAlert2").VerifyYellowBanner 
	Browser("Confirmation").Page("Confirmation").WebElement("LBL_Documents").VerifyExists 
	Browser("Confirmation").Page("Confirmation").WebElement("IMG_CompletedIcon").VerifyExists 
	Browser("Confirmation").Page("Confirmation").WebButton("BTN_Take").VerifyExists 
	Browser("Confirmation").Page("Confirmation").WebButton("BTN_Done").ClickObject 
	Browser("Account Home").Page("Account Home").WebElement("LBL_AccountSummary").VerifyPageName 
	Browser("Account Home").Page("Account Home").WebElement("LBL_AccountSummary").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_Rep Number").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_Account #").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_SSN").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_AccountClass").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_Registration").VerifyExists 
End Function
''###################################################################################################
''Function Name : VerifyCheckRequestReviewPage
''Description : Verify check request review page
''Author : Gallop
''###################################################################################################
Public Function VerifyCheckRequestReviewPage
Environment.Value("StepName") = "VerifyCheckRequestReviewPage"
	Browser("Submit").Page("Submit").WebElement("LBL_AccountSummary").VerifyPageName 
	Browser("Submit").Page("Submit").WebElement("LBL_AccountSummary").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_AccountDetails").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_FutureDatedTransaction").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_CheckInstructions").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_CheckRecipient").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_CheckMemo").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_DistributionSummary").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_DistributionAmount").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_DistributionSource").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_Fees").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("RWL_TotalDebitToAccount").VerifyDisplayedText dtScenario.value("TotalDebitToAccount")
	Browser("Submit").Page("Submit").WebElement("RWL_TotalFees/Adjustments").VerifyDisplayedText dtScenario.value("TotalFeesAdjustments")
	Browser("Submit").Page("Submit").WebElement("RWL_NetCashAmount").VerifyDisplayedText dtScenario.value("NetCashAmount")
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
End Function
''###################################################################################################
''Function Name : VerifyAccountSummaryPageInJournal
''Description : Verify Account Summary Page In Journal
''Author : Gallop
''###################################################################################################
Public Function VerifyAccountSummaryPageInJournal
Environment.Value("StepName") = "VerifyAccountSummaryPageInJournal"
	Browser("Account Home").Page("Account Home").WebElement("LBL_AccountSummary").PageSync 
	Browser("Account Home").Page("Account Home").WebElement("LBL_AccountSummary").VerifyPageName 
	Browser("Account Home").Page("Account Home").WebElement("RWL_Rep Number").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_Account #").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_AccountClass").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_Registration").VerifyExists 
	Browser("Account Home").Page("Account Home").WebElement("RWL_SSN").VerifyExists
End Function
''###################################################################################################
''Function Name : VerifyConfirmationPageInJournal
''Description : Verify Confirmation Page In Journal
''Author : Gallop
''###################################################################################################
Public Function VerifyConfirmationPageInJournal
Environment.Value("StepName") = "VerifyConfirmationPageInJournal"
	Browser("Confirmation").Page("Confirmation").WebElement("LBL_Documents").PageSync 
	Browser("Confirmation").Page("Confirmation").WebElement("LBL_Documents").VerifyPageName 
	Browser("Confirmation").Page("Confirmation").WebElement("RWL_ConfirmationYellowAlert1").VerifyYellowBanner 
	Browser("Confirmation").Page("Confirmation").WebElement("RWL_ConfirmationYellowAlert2").VerifyYellowBanner 
	Browser("Confirmation").Page("Confirmation").WebElement("LBL_Documents").VerifyExists 
	Browser("Confirmation").Page("Confirmation").WebTable("TBL_Document").VerifyTableIsEmpty 
	Browser("Confirmation").Page("Confirmation").WebButton("BTN_AddDocument").VerifyExists 
	Browser("Confirmation").Page("Confirmation").WebElement("LBL_CashControlLowRisk").VerifyExists 
	Browser("Confirmation").Page("Confirmation").WebButton("BTN_Take").VerifyExists 
	Browser("Confirmation").Page("Confirmation").WebButton("BTN_Done").VerifyExists 
	Browser("Confirmation").Page("Confirmation").WebButton("BTN_Done").ClickObject 
End Function
''###################################################################################################
''Function Name : VerifyJournalRequestPage
''Description : Verify journal request  page
''Author : Gallop
''###################################################################################################
Public Function VerifyJournalRequestPage
Environment.Value("StepName") = "VerifyJournalRequestPage"
	Browser("Account Home").Page("Account Home").WebButton("BTN_Journal").ClickObject 
	Browser("Journal Entry").Page("Journal Entry").WebElement("LBL_AccountSummary").PageSync 
	Browser("Journal Entry").Page("Journal Entry").WebElement("LBL_AccountSummary").VerifyPageName 
	Browser("Journal Entry").Page("Journal Entry").WebElement("LBL_AccountSummary").VerifyExists 
	Browser("Journal Entry").Page("Journal Entry").WebElement("LBL_AccountDetails").VerifyExists 
	Browser("Journal Entry").Page("Journal Entry").WebElement("LBL_FutureDatedTransaction").VerifyExists 
	Browser("Journal Entry").Page("Journal Entry").WebElement("LBL_JournalInstructions").VerifyExists 
	Browser("Journal Entry").Page("Journal Entry").WebElement("LBL_ReceivingAccount").VerifyExists 
	Browser("Journal Entry").Page("Journal Entry").WebElement("LBL_DistributionAmount").VerifyExists 
	Browser("Journal Entry").Page("Journal Entry").WebButton("BTN_WorkFolderIcon").VerifyExists 
	Browser("Journal Entry").Page("Journal Entry").WebTable("TBL_Assigned").VerifyValuesInRow "Assigned|Tim Parsons"
	Browser("Journal Entry").Page("Journal Entry").WebTable("TBL_Assigned").VerifyDueDate "Due Date"
	Browser("Journal Entry").Page("Journal Entry").WebElement("TGL_Journal").ClickOnToggles "Securities"
	Browser("Journal Entry").Page("Journal Entry").WebElement("TGL_JournalSecurities").VerifyExists 
	Browser("Journal Entry").Page("Journal Entry").WebElement("TGL_JournalSecurities").VerifyNoValuesSelectedForToggle 
	Browser("Journal Entry").Page("Journal Entry").WebElement("LBL_DistributionAmount").VerifyNotExists 
	Browser("Journal Entry").Page("Journal Entry").WebElement("LBL_DistributionSummary").VerifyNotExists 
	Browser("Journal Entry").Page("Journal Entry").WebElement("TGL_Journal").ClickOnToggles "Cash"
	Browser("Journal Entry").Page("Journal Entry").WebElement("TGL_JournalSecurities").VerifyNotExists 
	Browser("Journal Entry").Page("Journal Entry").WebElement("LBL_DistributionAmount").VerifyExists 
	Browser("Journal Entry").Page("Journal Entry").WebElement("LBL_DistributionSummary").VerifyExists 
	Browser("Journal Entry").Page("Journal Entry").WebEdit("TXT_DistributionAmount").EnterText dtScenario.value("DistributionAmount")
	Browser("Journal Entry").Page("Journal Entry").WebElement("LBL_DistributionAmount").ClickObject 
	Browser("Journal Entry").Page("Journal Entry").WebEdit("TXT_DistributionAmount").VerifyValue dtScenario.value("DistributionAmountFormatted")
	Browser("Journal Entry").Page("Journal Entry").WebElement("RWL_TotalDebitToAccount").VerifyDisplayedText dtScenario.value("DistributionAmountFormatted")
	Browser("Journal Entry").Page("Journal Entry").WebElement("RWL_GrossAfterFees").VerifyDisplayedText dtScenario.value("DistributionAmountFormatted")
	Browser("Journal Entry").Page("Journal Entry").WebElement("RWL_NetCashAmount").VerifyDisplayedText dtScenario.value("DistributionAmountFormatted")
	Browser("Journal Entry").Page("Journal Entry").WebButton("BTN_Next").ClickObject 
End Function
''###################################################################################################
''Function Name : VerifyJournalRequestReceivingAccountPage
''Description : Verify Journla Request Receiving Account Page
''Author : Gallop
''###################################################################################################
Public Function VerifyJournalRequestReceivingAccountPage
Environment.Value("StepName") = "VerifyJournalRequestReceivingAccountPage"
	Browser("Add New LOA").Page("Add New LOA").WebElement("LBL_AccountSummary").VerifyPageName 
	Browser("Add New LOA").Page("Add New LOA").WebElement("LBL_AccountSummary").VerifyExists 
	Browser("Add New LOA").Page("Add New LOA").WebElement("LBL_AccountDetails").VerifyExists 
	Browser("Add New LOA").Page("Add New LOA").WebElement("LBL_ReceivingAccount").VerifyExists 
	Browser("Add New LOA").Page("Add New LOA").WebButton("BTN_Next").ClickObject 
End Function
''###################################################################################################
''Function Name : VerifyJournalReviewAndSubmitPage
''Description : Verify Journla Review And Submit Page
''Author : Gallop
''###################################################################################################
Public Function VerifyJournalReviewAndSubmitPage
Environment.Value("StepName") = "VerifyJournalReviewAndSubmitPage"
	Browser("Submit").Page("Submit").WebElement("LBL_AccountSummary").VerifyPageName 
	Browser("Submit").Page("Submit").WebElement("LBL_AccountSummary").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_AccountDetails").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_FutureDatedTransaction").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_JournalDetails").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_ReceivingAccount").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_DistributionSummary").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("RWL_JournalType").VerifyDisplayedText "Account to Account"
	Browser("Submit").Page("Submit").WebElement("RWL_NetSecuritiesValue").VerifyDisplayedText "$0.00"
	Browser("Submit").Page("Submit").WebElement("RWL_FromAccountType").VerifyDisplayedText "Cash"
	Browser("Submit").Page("Submit").WebElement("RWL_NetCashAmount").VerifyDisplayedText "$5.00"
	Browser("Submit").Page("Submit").WebButton("BTN_Back").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Save").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
End Function
''###################################################################################################
''Function Name : LoginToUser
''Description : Login to application
''Author : Gallop
''###################################################################################################
Public Function LoginToUser
Environment.Value("StepName") = "LoginToUser"
	Call StartApplication()
	Browser("Login").Page("Login").WebEdit("TXT_user_name").EnterText dtScenario.value("username")
	Browser("Login").Page("Login").WebEdit("TXT_user_password").EnterText dtScenario.value("password")
	Browser("Login").Page("Login").WebButton("BTN_Log In").ClickObject 
	Browser("Work In Progress").Page("Work In Progress").WebElement("Home").PageSync 
	Browser("Work In Progress").Page("Work In Progress").WebElement("Home").VerifyExists 

End Function
''###################################################################################################
''Function Name : NavigatetoWorkflowUser
''Description : Navigates to Workflow Testing
''Author : Gallop
''###################################################################################################
Public Function NavigatetoWorkflowUser
Environment.Value("StepName") = "NavigatetoWorkflowUser"
	Browser("Work In Progress").Page("Work In Progress").WebElement("Tests").ClickObject 
	Browser("Work In Progress").Page("Work In Progress").WebElement("Test Workflow Tab").ClickObject 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").PageSync 
	Browser("Entry").Page("Entry").WebButton("BTN_Next").WaitForObjectToLoad 

End Function
''###################################################################################################
''Function Name : VerifyWireRequestPage
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function VerifyWireRequestPage
Environment.Value("StepName") = "VerifyWireRequestPage"
	Browser("Account Home").Page("Account Home").WebButton("BTN_Wire").ClickObject 
	Browser("FedWire Entry").Page("FedWire Entry").WebElement("LBL_AccountSummary").PageSync 
	Browser("FedWire Entry").Page("FedWire Entry").WebElement("LBL_AccountSummary").VerifyPageName 
	Browser("FedWire Entry").Page("FedWire Entry").WebElement("LBL_AccountSummary").VerifyExists 
	Browser("FedWire Entry").Page("FedWire Entry").WebElement("LBL_AccountDetails").VerifyExists 
	Browser("FedWire Entry").Page("FedWire Entry").WebElement("LBL_FutureDatedTransaction").VerifyExists 
	Browser("FedWire Entry").Page("FedWire Entry").WebElement("LBL_WireInstructions").VerifyExists 
	Browser("FedWire Entry").Page("FedWire Entry").WebButton("BTN_WorkFolderIcon").VerifyExists 

End Function
''###################################################################################################
''Function Name : NavigateToButtonTesting
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function NavigateToButtonTesting
Environment.Value("StepName") = "NavigateToButtonTesting"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_ButtonTesting").ClickObject 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Next").PageSync 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Next").WaitForObjectToLoad 

End Function
''###################################################################################################
''Function Name : ValidateHideButton
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateHideButton
Environment.Value("StepName") = "ValidateHideButton"
	Browser("Entry").Page("Entry").WebElement("LBL_QATestingPageHeader").PageSync 
	Browser("Entry").Page("Entry").WebElement("LBL_QATestingPageHeader").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebElement("LBL_QATestingPageHeader").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Cancel").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Back").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Edit").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Delete").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Next").VerifyExists 

End Function
''###################################################################################################
''Function Name : ValidateButtonsTopAndBottomOfPage
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateButtonsTopAndBottomOfPage
Environment.Value("StepName") = "ValidateButtonsTopAndBottomOfPage"
	Browser("Entry").Page("Entry").WebElement("LBL_QATestingPageHeader").PageSync 
	Browser("Entry").Page("Entry").WebElement("LBL_QATestingPageHeader").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebElement("LBL_QATestingPageHeader").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Cancel").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Back").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Edit").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Delete").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BOT_BTN_Back").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BOT_BTN_Edit").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BOT_BTN_Cancel").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BOT_BTN_Delete").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BOT_BTN_Next").VerifyExists 

End Function
''###################################################################################################
''Function Name : ValidateHideButtonRule
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateHideButtonRule
Environment.Value("StepName") = "ValidateHideButtonRule"
	Browser("Entry").Page("Entry").WebElement("LBL_QATestingPageHeader").PageSync 
	Browser("Entry").Page("Entry").WebElement("LBL_QATestingPageHeader").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BOT_BTN_Edit").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Edit").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("DWL_TaxIndicator").SelectListBox dtScenario.value("DropDownValue")
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Next").ClickObject 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Edit").VerifyNotExist 
	Browser("Entry").Page("Entry").WebButton("BOT_BTN_Edit").VerifyNotExist 

End Function
''###################################################################################################
''Function Name : ValidateDeleteButton
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateDeleteButton
Environment.Value("StepName") = "ValidateDeleteButton"
	Browser("Entry").Page("Entry").WebElement("LBL_QATestingPageHeader").PageSync 
	Browser("Entry").Page("Entry").WebElement("LBL_QATestingPageHeader").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Delete").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BOT_BTN_Delete").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("DWL_TaxIndicator").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("DWL_TaxIndicator").SelectListBox dtScenario.value("DropDownValue")
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Delete").ClickObject 
	Browser("Entry").Dialog("Message from webpage").Static("MSG_This action will permanently").VerifyExists 
	Browser("Entry").Dialog("Message from webpage").WinButton("BTN_OK").VerifyExists 
	Browser("Entry").Dialog("Message from webpage").WinButton("BTN_Cancel").VerifyExists 
	Browser("Entry").Dialog("Message from webpage").WinButton("BTN_OK").ClickObject 
	Browser("Entry").Page("Entry").WebElement("DWL_TaxIndicator").VerifyValueSelectedInDropDown ""

End Function
''###################################################################################################
''Function Name : ValidateDeleteCanBeCancelled
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateDeleteCanBeCancelled
Environment.Value("StepName") = "ValidateDeleteCanBeCancelled"
	Browser("Entry").Page("Entry").WebElement("LBL_QATestingPageHeader").PageSync 
	Browser("Entry").Page("Entry").WebElement("LBL_QATestingPageHeader").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("TOP_BTN_Delete").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BOT_BTN_Delete").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("DWL_TaxIndicator").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("DWL_TaxIndicator").SelectListBox dtScenario.value("DropDownValue")
	Browser("Entry").Page("Entry").WebButton("BOT_BTN_Delete").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BOT_BTN_Delete").ClickObject 
	Browser("Entry").Dialog("Message from webpage").Static("MSG_This action will permanently").VerifyExists 
	Browser("Entry").Dialog("Message from webpage").WinButton("BTN_Cancel").VerifyExists 
	Browser("Entry").Dialog("Message from webpage").WinButton("BTN_Cancel").ClickObject 
	Browser("Entry").Page("Entry").WebElement("DWL_TaxIndicator").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("DWL_TaxIndicator").VerifyValueSelectedInDropDown dtScenario.value("ValueAfterCancel")

End Function
''###################################################################################################
''Function Name : NavigateToDraftMode
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function NavigateToDraftMode
Environment.Value("StepName") = "NavigateToDraftMode"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB _Home").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_Draft Mode").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_Draft Mode").ClickObject 

End Function
''###################################################################################################
''Function Name : VerifyWarningMessageRemoved
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function VerifyWarningMessageRemoved
Environment.Value("StepName") = "VerifyWarningMessageRemoved"
	Browser("Entry").Page("Entry").WebElement("LBL_Header_DraftModeEntry").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebElement("LBL_Header_DraftModeEntry").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("LBL_PageMessage_Thispagecontainrequired").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("LBL_Warning_Answer theDropDownquestion").VerifyYellowBanner 
	Browser("Entry").Page("Entry").WebElement("DWL_SelectValue").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("DWL_SelectValue").SelectListBoxByIndex 2
	Browser("Entry").Page("Entry").WebElement("LBL_Warning_Answer theDropDownquestion").VerifyNotExists 
	Browser("Entry").Page("Entry").WebElement("NAV_RequiredQuestions").ClickObject 
	Browser("Required Questions").Page("Required Questions").WebElement("LBL_Answer * Required Checkbox").VerifyExists 
	Browser("Required Questions").Page("Required Questions").WebCheckBox("ChkBox_Brian").VerifyExists 
	Browser("Required Questions").Page("Required Questions").WebCheckBox("ChkBox_Brian").ClickObject 
	Browser("Required Questions").Page("Required Questions").WebElement("RDL_RequiredInterview").VerifyExists 
	Browser("Required Questions").Page("Required Questions").WebElement("RDL_RequiredInterview").SelectRadioButton "small value"
	Browser("Required Questions").Page("Required Questions").WebElement("LBL_Answer * Required Checkbox").VerifyNotExists 

End Function
''###################################################################################################
''Function Name : VerifiyNavBarStatus
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function VerifiyNavBarStatus
Environment.Value("StepName") = "VerifiyNavBarStatus"
	Browser("Entry").Page("Entry").WebElement("NAV_Entry").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("NAV_RequiredQuestions").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("NAV_Submit").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("NAV_AdditionalInformation").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("NAV_Summary").VerifyExists 
	Browser("Entry").Page("Entry").WebTable("TBL_RepeatingModuleDataTable").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_NewItem").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_NewItem").ClickObject 
	Browser("DraftRM").Page("DraftRM").WebElement("NAV_(NEW)").VerifyExists 
	Browser("DraftRM").Page("DraftRM").WebEdit("TXT_FirstName").EnterText dtScenario.value("FirstName")
	Browser("DraftRM").Page("DraftRM").WebEdit("TXT_LastName").EnterText dtScenario.value("LastName")
	Browser("DraftRM").Page("DraftRM").WebEdit("TXT_ClientID").EnterText "1111-1111"
	Browser("DraftRM").Page("DraftRM").WebButton("BTN_Next_DraftMode").ClickObject 
	Browser("Entry").Page("Entry").WebElement("NAV_TitleDanger").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("NAV_TitleDanger").ClickObject 
	Browser("DraftRM").Page("DraftRM").WebEdit("TXT_ClientID").EnterText dtScenario.value("ClientID")
	Browser("DraftRM").Page("DraftRM").WebButton("BTN_Next_DraftMode").ClickObject 
	Browser("Entry").Page("Entry").WebElement("NAV_TitleSuccess").VerifyExists 
	Browser("Entry").Page("Entry").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("TRKcircle_Button_Entry").ClickObject 
	Browser("Entry").Page("Entry").WebElement("Entry_StepPopUp").VerifyExists 

End Function
''###################################################################################################
''Function Name : VerifyNavMatchesStepTracker
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function VerifyNavMatchesStepTracker
Environment.Value("StepName") = "VerifyNavMatchesStepTracker"
	Browser("Entry").Page("Entry").WebElement("NAV_Entry").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("NAV_RequiredQuestions").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("NAV_Submit").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("NAV_AdditionalInformation").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("NAV_Summary").VerifyExists 
	Browser("Entry").Page("Entry").WebTable("TBL_RepeatingModuleDataTable").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_NewItem").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_NewItem").ClickObject 
	Browser("DraftRM").Page("DraftRM").WebElement("NAV_(NEW)").VerifyExists 
	Browser("DraftRM").Page("DraftRM").WebEdit("TXT_FirstName").EnterText dtScenario.value("FirstName")
	Browser("DraftRM").Page("DraftRM").WebEdit("TXT_LastName").EnterText dtScenario.value("LastName")
	Browser("DraftRM").Page("DraftRM").WebEdit("TXT_ClientID").EnterText dtScenario.value("ClientID")
	Browser("DraftRM").Page("DraftRM").WebButton("BTN_Next_DraftMode").ClickObject 
	Browser("Entry").Page("Entry").WebElement("DWL_SelectValue").SelectListBoxByIndex 2
	Browser("Entry").Page("Entry").WebElement("RDL_SelectionRequired").SelectRadioButton "small value"
	Browser("Entry").Page("Entry").WebElement("TGL_Yes").ClickObject 
	Browser("Entry").Page("Entry").WebEdit("TXT_SelectADate").EnterText gfGetCurrentDate
	Browser("Entry").Page("Entry").WebElement("TRKcircleEntry").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("TRKcircleEntry").ValidateState "valid"

End Function
''###################################################################################################
''Function Name : SaveAndNewCanAdd
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function SaveAndNewCanAdd
Environment.Value("StepName") = "SaveAndNewCanAdd"
	Browser("IFS").Page("IFS").WebElement("LBL_RepeatingModTable").VerifyExists 
	Browser("IFS").Page("IFS").WebElement("DWL_DroppingDown").VerifyExists 
	Browser("IFS").Page("IFS").WebButton("BTN_NewItem").VerifyExists 
	Browser("IFS").Page("IFS").WebButton("BTN_NewItem").ClickObject 
	Browser("IFS").Page("IFS_2").WebEdit("TXT_RM_Test1").EnterText dtScenario.value("Test1")
	Browser("IFS").Page("IFS_2").WebEdit("TXT_RM_Test2").EnterText dtScenario.value("Test2")
	Browser("IFS").Page("IFS_2").WebEdit("TXT_RM_Test3").EnterText dtScenario.value("Test3")
	Browser("IFS").Page("IFS_2").WebElement("DWL_RM _TestQuestion5").SelectListBoxByIndex 5
	Browser("IFS").Page("IFS_2").WebElement("RDL_RM_Test4").SelectRadioButton "No"
	Browser("IFS").Page("IFS_2").WebButton("BTN_Next").ClickObject 
	Browser("IFS").Page("IFS").WebTable("TBL_SaveAndNew").WaitForObjectToLoad 
	Browser("IFS").Page("IFS").WebTable("TBL_SaveAndNew").PageSync 
	Browser("IFS").Page("IFS").WebTable("TBL_SaveAndNew").VerifyvalueInSpecificCell 2,4,"False"
	Browser("IFS").Page("IFS").WebTable("TBL_SaveAndNew").VerifyvalueInSpecificCell 2,5,"Bi-Monthly"
	Browser("IFS").Page("IFS").WebTable("TBL_SaveAndNew").VerifyvalueInSpecificCell 2,3,"123456789"
	Browser("IFS").Page("IFS").WebTable("TBL_SaveAndNew").VerifyDateInRow 2,2,"(704) 997-3603"

End Function
''###################################################################################################
''Function Name : NavigatetoRepeatingSaveAndNew
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function NavigatetoRepeatingSaveAndNew
Environment.Value("StepName") = "NavigatetoRepeatingSaveAndNew"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB _Home").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_SaveAndNew").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_SaveAndNew").ClickObject 

End Function
''###################################################################################################
''Function Name : SaveAndNewCanAddMultiple
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function SaveAndNewCanAddMultiple
Environment.Value("StepName") = "SaveAndNewCanAddMultiple"
	.VerifyExists 
	.VerifyExists 
	.VerifyExists 
	.ClickObject 
	.EnterText dtScenario.value("Test1")
	.EnterText dtScenario.value("Test2")
	.EnterText dtScenario.value("Test3")
	.SelectListBoxByIndex 5
	.SelectRadioButton "No"
	.VerifyExists 
	.ClickObject 
	.VerifyValueisNull 
	.VerifyValueisNull 
	.VerifyValueisNull 
	.VerifyValueisNull 
	.EnterText dtScenario.value("Test1B")
	.EnterText dtScenario.value("Test2B")
	.EnterText dtScenario.value("Test3B")
	.SelectRadioButton "No"
	.SelectListBoxByIndex 4
	.VerifyExists 
	.ClickObject 
	.EnterText dtScenario.value("Test1C")
	.EnterText dtScenario.value("Test2C")
	.EnterText dtScenario.value("Test3C")
	.SelectRadioButton "No"
	.SelectListBoxByIndex 8
	.ClickObject 
	.EnterText dtScenario.value("Test1D")
	.EnterText dtScenario.value("Test2D")
	.EnterText dtScenario.value("Test3D")
	.SelectRadioButton "No"
	.SelectListBoxByIndex 7
	.ClickObject 
	.WaitForObjectToLoad 
	.PageSync 
	.VerifyvalueInSpecificCell 2,4,"False"
	.VerifyvalueInSpecificCell 2,5,"Bi-Monthly"
	.VerifyvalueInSpecificCell 2,3,"123456789"
	.VerifyvalueInSpecificCell 2,2,"(704) 997-3603"
	.VerifyvalueInSpecificCell 3,2,"(876) 879-6876 "
	.VerifyvalueInSpecificCell 3,3,"987654321"
	.VerifyvalueInSpecificCell 3,4,"False"
	.VerifyvalueInSpecificCell 3,5,"Monthly"
	.VerifyvalueInSpecificCell 4,2,"(879) 879-8987"
	.VerifyvalueInSpecificCell 4,3,"123456789"
	.VerifyvalueInSpecificCell 5,3,"456712345"
	.VerifyvalueInSpecificCell 5,2,"(704) 997-3603"

End Function
''###################################################################################################
''Function Name : ValidateNotRequiredDocument
''Description : Verifies document types and their initial attachements
''Author : Gallop
''###################################################################################################
Public Function ValidateNotRequiredDocument
Environment.Value("StepName") = "ValidateNotRequiredDocument"
	Browser("Entry").Page("Entry").WebElement("LBL_DocumentWorkflowTesting").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_TestingtheDescription").EnterText "Testing"
	Browser("Entry").Page("Entry").WebButton("BTN_NextQADoumentTestNext").ClickObject 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit").ClickObject 
	Browser("Summary").Page("Summary").Link("TAB_Documents").WaitForObjectToLoadWithTime 5
	Browser("Summary").Page("Summary").Link("TAB_Documents").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").VerifyExists 
	Browser("Summary").Page("Summary").Link("TAB_Documents").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_AddDocument").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_PrintCoverSheets").VerifyExists 
	Browser("Summary").Page("Summary").WebButton("BTN_ViewPrintAllForms").VerifyExists 
	Browser("Summary").Page("Summary").WebCheckBox("CWL_Not_Required").VerifyExists 
	Browser("Summary").Page("Summary").WebCheckBox("CWL_Not_Required").ClickObject 

End Function
''###################################################################################################
''Function Name : VerifySkipValidation
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function VerifySkipValidation
Environment.Value("StepName") = "VerifySkipValidation"
	Browser("Entry").Page("Entry").WebElement("LBL_Header_DraftModeEntry").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebElement("LBL_Header_DraftModeEntry").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("LBL_PageMessage_Thispagecontainrequired").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("LBL_Warning_Answer theDropDownquestion").VerifyYellowBanner 
	Browser("Entry").Page("Entry").WebElement("DWL_SelectValue").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("DWL_SelectValue").SelectListBoxByIndex 2
	Browser("Entry").Page("Entry").WebElement("LBL_Warning_Answer theDropDownquestion").VerifyNotExists 
	Browser("Entry").Page("Entry").WebElement("TGL_Yes").ClickObject 
	Browser("Entry").Page("Entry").WebEdit("TXT_SelectADate").EnterText gfGetCurrentDate
	Browser("Entry").Page("Entry").WebButton("BTN_SkipValidation").ClickObject 
	Browser("Required Questions").Page("Required Questions").WebElement("TWL_RequiredCheckboxGroup").SelectCheckBox 
	Browser("Required Questions").Page("Required Questions").WebElement("LBL_Answer * Required Checkbox").VerifyExists 
	Browser("Required Questions").Page("Required Questions").WebButton("BTN_SkipValidation").ClickObject 
	Browser("Additional Information").Page("Additional Information").WebEdit("TXT_PetsName").EnterText "Fred"
	Browser("Additional Information").Page("Additional Information").WebButton("BTN_Next_AdditionalInfo").ClickObject 
	Browser("Additional Information2").Page("Additional Information2_2").WebEdit("TXT_DateBowlCleaned").EnterText gfGetCurrentDate
	Browser("Additional Information2").Page("Additional Information2_2").WebButton("BTN_SkipValidation").ClickObject 
	Browser("Submit").Page("Submit").WebElement("PNL_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("TRK_BTN_Submit").ClickObject 
	Browser("Submit").Page("Submit").WebElement("STP_PopUp_Submit").VerifyExists 
	Browser("Submit").Page("Submit").WebButton("BTN_Submit_DftMd").ClickObject 
	Browser("Submit").Page("Submit").WebElement("LBL_MSG_CannotSubmit").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("TRKcircle_Entry_Error").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("TRKcircle_AdditionalInfo_Error").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("TRKcircle_Submit_Error").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("NAV_Entry_Error").ClickObject 
	Browser("Entry").Page("Entry").WebButton("BTN_NewItem").ClickObject 

End Function
''###################################################################################################
''Function Name : NavigatetoShowHide
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function NavigatetoShowHide
Environment.Value("StepName") = "NavigatetoShowHide"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB _Home").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_ShowHide").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_ShowHide").ClickObject 

End Function
''###################################################################################################
''Function Name : ShowHide_QuestionsWithToggle
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ShowHide_QuestionsWithToggle
Environment.Value("StepName") = "ShowHide_QuestionsWithToggle"
	Browser("Entry").Page("Entry").WebElement("LBL_EntryPageforShow/Hide").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebElement("LBL_EntryPageforShow/Hide").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("LBL_Warning").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("TWL_HiddenStringQuestion").VerifyNotExists 
	Browser("Entry").Page("Entry").WebElement("QG_EditableNonEditable").VerifyNotExists 
	Browser("Entry").Page("Entry").WebElement("TWL_Numeric").VerifyNotExists 
	Browser("Entry").Page("Entry").WebElement("TGL_Toggle1").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("TGL_Toggle1").ClickOnToggles "Yes"
	Browser("Entry").Page("Entry").WebElement("TWL_HiddenStringQuestion").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("TGL_Toggle2").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("TGL_Toggle2").ClickOnToggles "Yes"
	Browser("Entry").Page("Entry").WebElement("TWL_Numeric").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("TWL_String").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("TGL_Toggle4").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("TGL_Toggle4").ClickOnToggles "Yes"
	Browser("Entry").Page("Entry").WebElement("QG_EditableNonEditable").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("TGL_Toggle3").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("TGL_Toggle3").ClickOnToggles "Yes"
	Browser("Entry").Page("Entry").WebElement("TWL_String").VerifyNotExists 
	Browser("Entry").Page("Entry").WebElement("TGL_Toggle3").ClickOnToggles "No"
	Browser("Entry").Page("Entry").WebElement("TWL_String").VerifyExists 

End Function
''###################################################################################################
''Function Name : EditNonEdit_ValidateToggles
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function EditNonEdit_ValidateToggles
Environment.Value("StepName") = "EditNonEdit_ValidateToggles"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("LBL_MakeQuestionsNonEditable").WaitForObjectToLoad 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("LBL_MakeQuestionsNonEditable").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TWL_Numeric").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_QANumeric").EnterText "345"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#1").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#1").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_QANumeric").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NumericInput2Decimal").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NumericInput2Decimal").EnterText "345"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#2").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#2").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NumericInput2Decimal").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NumericInputNonDecimal").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NumericInputNonDecimal").EnterText "345"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#3").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#3").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NumericInputNonDecimal").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NumericInputCommaDelimted").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NumericInputCommaDelimted").EnterText "345"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#4").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#4").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NumericInputCommaDelimted").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_YesNoFireTheRule").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("DWL_DropDown1").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_YesNoFireTheRule").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("DWL_DropDown1").VerifyFieldIsReadOnly 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NameAlphaSpecialChar").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NameAlphaSpecialChar").EnterText "Testing"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#37").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#37").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NameAlphaSpecialChar").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_Currency").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_Currency").EnterText "$345.00"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#5").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#5").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_Currency").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_CurrencyWholeNumber").EnterText "$345.00"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#6").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_CurrencyWholeNumber").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_Currecy_Formated_NEG").EnterText "$345.00"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#7").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_Currecy_Formated_NEG").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_Currecy_Truncate").EnterText "$345.00"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#8").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_Currecy_Truncate").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_String").EnterText "Testing"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#9").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_String").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_String_alpha_Only").EnterText "Testing"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#10").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_String_alpha_Only").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_String_NumericOnly").EnterText "345"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#11").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_String_NumericOnly").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_Date").EnterText gfGetCurrentDate
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#12").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_Date").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_DateMonthDayOnly").EnterText gfGetDateMonth
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#14").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_DateMonthDayOnly").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_BusinessDate").EnterText gfGetCurrentDate
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#15").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_BusinessDate").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_BirthDate").EnterText gfGetCurrentDate
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#35").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_BirthDate").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_MonthYearDate").EnterText gfGetMonthYear
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#36").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_MonthYearDate").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_PhoneNumber").EnterText "8885555555"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#17").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_PhoneNumber").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_BLAccountNumber").EnterText "1010-1010"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#22").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_BLAccountNumber").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_Email").EnterText dtScenario.value("EmailID")
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#23").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_Email").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_Social").EnterText "111-11-1111"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#26").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_Social").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_SocialwDashes").EnterText "111-11-1111"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#27").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_SocialwDashes").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_TaxID").EnterText "11-1111111"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#28").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_TaxID").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_TaxIDwDashes").EnterText "11-1111111"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#29").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_TaxIDwDashes").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_Zip").EnterText "32232-2223"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#30").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_Zip").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_XML").EnterText "Testing"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#33").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_XML").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_AuditNote").EnterText "Testing"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TGL_#34").ClickOnToggles "Yes"
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_AuditNote").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebButton("BTN_Next").ClickObject 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TWL_Editable/Non-Editable").WaitForObjectToLoad 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TWL_Editable/Non-Editable").VerifyExists 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebButton("BTN_Back").ClickObject 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("LBL_MakeQuestionsNonEditable").WaitForObjectToLoad 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_QANumeric").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NumericInput2Decimal").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NumericInputNonDecimal").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NumericInputCommaDelimted").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebElement("TWL_DropDown#1").VerifyFieldIsReadOnly 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_NameAlphaSpecialChar").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_Currency").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_CurrencyWholeNumber").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_Currecy_Formated_NEG").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_Currecy_Truncate").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_String").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_String_alpha_Only").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_String_NumericOnly").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_Date").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_DateMonthDayOnly").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_BusinessDate").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_BirthDate").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_MonthYearDate").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_PhoneNumber").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_BLAccountNumber").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_Email").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_Social").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_SocialwDashes").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_TaxID").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_TaxIDwDashes").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_Zip").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TWL_XML").VerifyDisabledandVisible 
	Browser("Questions Editable/Non-Editabl").Page("Questions Editable/Non-Editabl").WebEdit("TXT_AuditNote").VerifyDisabledandVisible 

End Function
''###################################################################################################
''Function Name : NavigatetoEditNonEdit
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function NavigatetoEditNonEdit
Environment.Value("StepName") = "NavigatetoEditNonEdit"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB _Home").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_QuestionsReadOnly").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_QuestionsReadOnly").ClickObject 

End Function
''###################################################################################################
''Function Name : NavigatetoDateStringEval
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function NavigatetoDateStringEval
Environment.Value("StepName") = "NavigatetoDateStringEval"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB _Home").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_Date/String Evaluation").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_Date/String Evaluation").ClickObject 

End Function
''###################################################################################################
''Function Name : Component_Template
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function Component_Template
Environment.Value("StepName") = "Component_Template"
	Browser("Entry").Page("Entry").WebElement("LBL_AuditBubbleTest").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_AuditBubble1").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_AuditBubble2").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_AuditBubble1").EnterText "123"
	Browser("Entry").Page("Entry").WebEdit("TXT_AuditBubble2").EnterText "123"

End Function
''###################################################################################################
''Function Name : DateStringEval_Verify_RuleSetsDate
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function DateStringEval_Verify_RuleSetsDate
Environment.Value("StepName") = "DateStringEval_Verify_RuleSetsDate"
	Browser("Entry").Page("Entry").WebElement("LBL_DateNumeric").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("LBL_Description").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("DWL_MakeSelection").SelectListBoxByIndex 3
	Browser("Entry").Page("Entry").WebElement("DTL_Date1").VerifyExists 
	Browser("Entry").Page("Entry").VerifyExists 
	Browser("Entry").Page("Entry").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_NumericCompare").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_NumericCompare").VerifyValue dtScenario.value("NumericCompare")
	Browser("Entry").Page("Entry").WebEdit("TXT_DateMonthDay").VerifyValue dtScenario.value("DateMonthDay")
	Browser("Entry").Page("Entry").WebEdit("TXT_DateBizDate").VerifyValue dtScenario.value("DateBizDate")
	Browser("Entry").Page("Entry").WebEdit("TXT_Date1").VerifyValue dtScenario.value("Date1")
	Browser("Entry").Page("Entry").WebButton("Eval_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("Eval_Next").ClickObject 
	Browser("Entry").Page("Entry").WebElement("MSG_NumericCompareFailed").VerifyNotExists 
	Browser("Entry").Page("Entry").WebElement("MSG_DateCompareTestFailed").VerifyNotExists 
	Browser("Entry").Page("Dummy").WebElement("DTL_QADateTest").VerifyDisplayedText dtScenario.value("DummyDate")
	Browser("Entry").Page("Dummy").WebElement("DTL_MonthDay").VerifyDisplayedText dtScenario.value("DummyMonthDay")
	Browser("Entry").Page("Dummy").WebElement("DTL_BizDate").VerifyDisplayedText dtScenario.value("DummyBizDate")
	Browser("Entry").Page("Dummy").WebElement("TWL_DummyNumeric").VerifyDisplayedText dtScenario.value("DummyNumeric")
	Browser("Entry").Page("Dummy").WebButton("BTN_Gear").ClickObject 
	Browser("Entry").Page("Dummy").WebElement("LNK_Debug").VerifyExists 
	Browser("Entry").Page("Dummy").WebElement("LNK_Debug").ClickDebugObject 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").WaitForObjectToLoadWithTime 5
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyExists 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 3,2,"TEST.NUMERIC.COMPARE2"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 3,3,".01"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 6,2,"QA.DATETEST3"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 6,3,"09/23/2015"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 9,2,"QA.DATE.TEST.DUMMY"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 9,3,"09/01/2015"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 11,2,"QA.DATE.TEST.DUMMY2"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 11,3,"9/24"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 4,2,"QA.DATETEST1"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 4,3,"09/01/2015"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 7,2,"QADATE.TESTDROPDOWN"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 7,3,"3"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 5,2,"QA.DATETEST2"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 5,3,"09/24"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 8,2,"TEST.NUMERIC.COMPARE"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 8,3,".1"

End Function
''###################################################################################################
''Function Name : DateStringEval_Verify_ManualEnterValue
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function DateStringEval_Verify_ManualEnterValue
Environment.Value("StepName") = "DateStringEval_Verify_ManualEnterValue"
	Browser("Entry").Page("Entry").WebElement("LBL_DateNumeric").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("LBL_Description").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("DTL_Date1").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("DTL_Date2").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("DTL_Date3").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_NumericCompare").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_NumericCompare").EnterText ".1"
	Browser("Entry").Page("Entry").WebEdit("TXT_NumericCompare").VerifyValue dtScenario.value("NumericCompare")
	Browser("Entry").Page("Entry").WebEdit("TXT_Date1").VerifyExists dtScenario.value("Date1")
	Browser("Entry").Page("Entry").WebEdit("TXT_Date1").EnterText "09/01/02015"
	Browser("Entry").Page("Entry").WebButton("Eval_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("Eval_Next").ClickObject 
	Browser("Entry").Page("Entry").WebElement("MSG_NumericCompareFailed").VerifyNotExists 
	Browser("Entry").Page("Entry").WebElement("MSG_DateCompareTestFailed").VerifyNotExists 
	Browser("Entry").Page("Dummy").WebEdit("TXT_QADateTEst").VerifyValue dtScenario.value("DummyDate")
	Browser("Entry").Page("Dummy").WebEdit("TXT_DummyNumeric").VerifyValue dtScenario.value("DummyNumeric")

End Function
''###################################################################################################
''Function Name : DateStringEval_Verify_ManualEnterValues
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function DateStringEval_Verify_ManualEnterValues
Environment.Value("StepName") = "DateStringEval_Verify_ManualEnterValues"
	Browser("Entry").Page("Entry").WebElement("LBL_DateNumeric").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("LBL_Description").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("DTL_Date1").VerifyExists 
	Browser("Entry").Page("Entry").VerifyExists 
	Browser("Entry").Page("Entry").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_NumericCompare").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_NumericCompare").EnterText ".1"
	Browser("Entry").Page("Entry").WebEdit("TXT_Date1").EnterText "09/01/2015"
	Browser("Entry").Page("Entry").WebEdit("TXT_DateMonthDay").EnterText "09/24"
	Browser("Entry").Page("Entry").WebEdit("TXT_DateBizDate").EnterText "09/23/2015"
	Browser("Entry").Page("Entry").WebEdit("TXT_NumericCompare").VerifyValue dtScenario.value("NumericCompare")
	Browser("Entry").Page("Entry").WebEdit("TXT_DateMonthDay").VerifyValue dtScenario.value("DateMonthDay")
	Browser("Entry").Page("Entry").WebEdit("TXT_DateBizDate").VerifyValue dtScenario.value("DateBizDate")
	Browser("Entry").Page("Entry").WebEdit("TXT_Date1").VerifyValue dtScenario.value("Date1")
	Browser("Entry").Page("Entry").WebButton("Eval_Next").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("Eval_Next").ClickObject 
	Browser("Entry").Page("Entry").WebElement("MSG_NumericCompareFailed").VerifyNotExists 
	Browser("Entry").Page("Entry").WebElement("MSG_DateCompareTestFailed").VerifyNotExists 
	Browser("Entry").Page("Dummy").WebElement("DTL_QADateTest").VerifyDisplayedText dtScenario.value("DummyDate")
	Browser("Entry").Page("Dummy").WebElement("DTL_MonthDay").VerifyDisplayedText dtScenario.value("DummyMonthDay")
	Browser("Entry").Page("Dummy").WebElement("DTL_BizDate").VerifyDisplayedText dtScenario.value("DummyBizDate")
	Browser("Entry").Page("Dummy").WebElement("TWL_DummyNumeric").VerifyDisplayedText dtScenario.value("DummyNumeric")
	Browser("Entry").Page("Dummy").WebButton("BTN_Gear").ClickObject 
	Browser("Entry").Page("Dummy").WebElement("LNK_Debug").VerifyExists 
	Browser("Entry").Page("Dummy").WebElement("LNK_Debug").ClickDebugObject 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").WaitForObjectToLoadWithTime 5
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyExists 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 3,2,"QA.DATETEST1"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 3,3,"09/01/2015"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 6,2,"TEST.NUMERIC.COMPARE"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 6,3,".1"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 9,2,"QA.DATE.TEST.DUMMY2"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 9,3,"9/24"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 10,2,"QA.DATE.TEST.DUMMY3"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 10,3,"09/23/2015"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 4,2,"QA.DATETEST2"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 4,3,"09/24"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 7,2,"QA.DATE.TEST.DUMMY"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 7,3,"09/01/2015"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 5,2,"QA.DATETEST3"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 5,3,"09/23/2015"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 8,2,"QA.NUMERIC.TEST.DUMMY"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_DebugQuestions").VerifyvalueInSpecificCell 8,3,"0.1"

End Function
''###################################################################################################
''Function Name : NavigateToAuditBubble
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function NavigateToAuditBubble
Environment.Value("StepName") = "NavigateToAuditBubble"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB _Home").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_AuditTestWorkflow").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_AuditTestWorkflow").ClickObject 

End Function
''###################################################################################################
''Function Name : AuditBubble_VerifyAuditValues
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function AuditBubble_VerifyAuditValues
Environment.Value("StepName") = "AuditBubble_VerifyAuditValues"
	Browser("Entry").Page("Entry").WebElement("LBL_AuditBubbleTest").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_AuditBubble1").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_AuditBubble2").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_AuditBubble1").EnterText "123"
	Browser("Entry").Page("Entry").WebEdit("TXT_AuditBubble2").EnterText "123"
	Browser("Entry").Page("Entry").WebButton("Audit1_Next").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebButton("Audit1_Next").ClickObject 
	Browser("Entry").Page("Submit").WebButton("Audit1_Submit").WaitForObjectToLoad 
	Browser("Entry").Page("Submit").WebButton("Audit1_Submit").ClickObject 
	Browser("Entry").Page("Summary").WebElement("LBL_QAAuditBubbleTest").WaitForObjectToLoad 
	Browser("Entry").Page("Summary").WebElement("LBL_QAAuditBubbleTest").VerifyExists 
	Browser("Entry").Page("Summary").WebElement("TWL_QA_Audit1").VerifyDisplayedText "123"
	Browser("Entry").Page("Summary").WebElement("TWL_QA_Audit2").VerifyDisplayedText "123"

End Function
''###################################################################################################
''Function Name : VerifyDraftValidationPanelErrorList
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function VerifyDraftValidationPanelErrorList
Environment.Value("StepName") = "VerifyDraftValidationPanelErrorList"
	Browser("Entry").Page("Entry").WebElement("LBL_Header_DraftModeEntry").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebElement("LBL_Header_DraftModeEntry").VerifyExists 
	Browser("Additional Information2").Page("Additional Information2_2").WebButton("BTN_SkipValidation").WaitForObjectToLoad 
	Browser("Additional Information2").Page("Additional Information2_2").WebButton("BTN_SkipValidation").ClickObject 
	Browser("Required Questions").Page("Required Questions").WebElement("LBL_Answer * Required Checkbox").WaitForObjectToLoad 
	Browser("Required Questions").Page("Required Questions").WebElement("LBL_Answer * Required Checkbox").VerifyExists 
	Browser("Required Questions").Page("Required Questions").WebElement("TRKcircle_State_error").WaitForObjectToLoad 
	Browser("Required Questions").Page("Required Questions").WebElement("TRKcircle_State_error").VerifyExists 
	Browser("Required Questions").Page("Required Questions").WebButton("BTN_SkipValidation").WaitForObjectToLoad 
	Browser("Required Questions").Page("Required Questions").WebButton("BTN_SkipValidation").ClickObject 
	Browser("Additional Information").Page("Additional Information").WebElement("LBL_This is an PreEvent infoMessage").WaitForObjectToLoad 
	Browser("Additional Information").Page("Additional Information").WebElement("LBL_This is an PreEvent infoMessage").VerifyExists 
	Browser("Additional Information").Page("Additional Information").WebElement("TRKcircle_AdditionalInfo").WaitForObjectToLoad 
	Browser("Additional Information").Page("Additional Information").WebElement("TRKcircle_AdditionalInfo").VerifyExists 
	Browser("Additional Information").Page("Additional Information").WebButton("BTN_SkipValidation").ClickObject 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("LBL_DraftMode.AdditionalInfo2").WaitForObjectToLoad 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("LBL_DraftMode.AdditionalInfo2").VerifyExists 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("TRKCircle_AdditionalInfoError").WaitForObjectToLoad 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("TRKCircle_AdditionalInfoError").VerifyExists 
	Browser("Additional Information2").Page("Additional Information2_2").WebButton("BTN_SkipValidation").ClickObject 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("LBL_IRARegistration").WaitForObjectToLoad 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("LBL_IRARegistration").VerifyExists 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("TRKcircle_IRARegistration").WaitForObjectToLoad 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("TRKcircle_IRARegistration").VerifyExists 
	Browser("Additional Information2").Page("Additional Information2_2").WebButton("BTN_SkipValidation").ClickObject 
	Browser("Submit").Page("Submit").WebElement("LBL_Error Summary").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebElement("LBL_Error Summary").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("TRKcircle_IRARegistrationError").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("LBL_Error Summary").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebElement("LBL_Error Summary").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("NAV_Entry_Error").WaitForObjectToLoad 
	Browser("Submit").Page("Submit").WebElement("NAV_Entry_Error").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("NAV_Entry_Error").ClickObject 
	Browser("Entry").Page("Entry").WebElement("DWL_SelectValue").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebElement("DWL_SelectValue").SelectListBoxByIndex 2
	Browser("Entry").Page("Entry").WebEdit("TXT_SelectADate").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebEdit("TXT_SelectADate").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_SelectADate").EnterText gfGetCurrentDate
	Browser("Entry").Page("Entry").WebElement("RDL_SelectionRequired").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebElement("RDL_SelectionRequired").SelectRadioButton "small value"
	Browser("Entry").Page("Entry").WebElement("NAV_RequiredQuestions").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebElement("NAV_RequiredQuestions").ClickObject 
	Browser("Required Questions").Page("Required Questions").WebElement("LBL_Answer * Required Checkbox").WaitForObjectToLoad 
	Browser("Required Questions").Page("Required Questions").WebElement("LBL_Answer * Required Checkbox").VerifyExists 
	Browser("Required Questions").Page("Required Questions").WebElement("TWL_RequiredCheckboxGroup").WaitForObjectToLoad 
	Browser("Required Questions").Page("Required Questions").WebElement("TWL_RequiredCheckboxGroup").VerifyExists 
	Browser("Required Questions").Page("Required Questions").WebCheckBox("ChkBox_Brian").ClickObject 
	Browser("Required Questions").Page("Required Questions").WebEdit("TXT_RequiredString").VerifyExists 
	Browser("Required Questions").Page("Required Questions").WebEdit("TXT_RequiredString").EnterText "Testing"
	Browser("Required Questions").Page("Required Questions").WebEdit("TXT_EmailAddress").VerifyExists 
	Browser("Required Questions").Page("Required Questions").WebEdit("TXT_EmailAddress").EnterText dtScenario.value("EmailID")
	Browser("Required Questions").Page("Required Questions").WebElement("RDL_RequiredInterview").SelectRadioButton "small value"
	Browser("Required Questions").Page("Required Questions").WebElement("NAV_AdditionalInformation").WaitForObjectToLoad 
	Browser("Required Questions").Page("Required Questions").WebElement("NAV_AdditionalInformation").ClickObject 
	Browser("Additional Information").Page("Additional Information").WebElement("LBL_This is an PreEvent infoMessage").WaitForObjectToLoad 
	Browser("Additional Information").Page("Additional Information").WebElement("LBL_This is an PreEvent infoMessage").VerifyExists 
	Browser("Additional Information").Page("Additional Information").WebElement("TRKCircle_EntrySectionValid").WaitForObjectToLoad 
	Browser("Additional Information").Page("Additional Information").WebElement("TRKCircle_EntrySectionValid").VerifyExists 
	Browser("Additional Information").Page("Additional Information").WebElement("TRKcircle_SubmitSectionValid").WaitForObjectToLoad 
	Browser("Additional Information").Page("Additional Information").WebElement("TRKcircle_SubmitSectionValid").VerifyExists 
	Browser("Additional Information").Page("Additional Information").WebElement("TRKCircle_AdditionalInfoError").WaitForObjectToLoad 
	Browser("Additional Information").Page("Additional Information").WebElement("TRKCircle_AdditionalInfoError").VerifyExists 
	Browser("Additional Information").Page("Additional Information").WebElement("TRKcircle_IRARegistrationError").WaitForObjectToLoad 
	Browser("Additional Information").Page("Additional Information").WebElement("TRKcircle_IRARegistrationError").VerifyExists 
	Browser("Additional Information").Page("Additional Information").WebEdit("TXT_PetsName").EnterText "Testing"
	Browser("Additional Information").Page("Additional Information").WebElement("NAV_AdditionalInformation2").ClickObject 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("LBL_DraftMode.AdditionalInfo2").WaitForObjectToLoad 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("LBL_DraftMode.AdditionalInfo2").VerifyExists 
	Browser("Additional Information2").Page("Additional Information2_2").WebEdit("TXT_GoldFishName").EnterText "Testing"
	Browser("Additional Information2").Page("Additional Information2_2").WebEdit("TXT_DateBowlCleaned").EnterText gfGetCurrentDate
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("NAV_IRA").WaitForObjectToLoad 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("NAV_IRA").ClickObject 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("LBL_IRARegistration").WaitForObjectToLoad 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("LBL_IRARegistration").VerifyExists 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("TRKCircle_AdditionalInfoValid").VerifyExists 
	Browser("Additional Information2").Page("Additional Information2_2").WebEdit("TXT_AccountNumber").VerifyExists 
	Browser("Additional Information2").Page("Additional Information2_2").WebEdit("TXT_AccountNumber").EnterText "1111-1111"
	Browser("Additional Information2").Page("Additional Information2_2").WebButton("BTN_Next").ClickObject 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("TRKcircle_IRARegistrationValid").WaitForObjectToLoad 
	Browser("Additional Information2").Page("Additional Information2_2").WebElement("TRKcircle_IRARegistrationValid").VerifyExists 
	Browser("Submit").Page("Submit").WebElement("TRKcircle_Entry_Valid").ValidateState "valid"
	Browser("Submit").Page("Submit").WebElement("TRKcircle_AdditionalInfo_Valid").ValidateState "valid"
	Browser("Submit").Page("Submit").WebElement("TRKcircle_SubmitSection_Valid").ValidateState "valid"

End Function
''###################################################################################################
''Function Name : ValidateInvalidDocumentsUpload
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateInvalidDocumentsUpload
Environment.Value("StepName") = "ValidateInvalidDocumentsUpload"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.bat"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".bat"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xlsx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xlsx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xls"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xls"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.png"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".png"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.jpg"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".jpg"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.zip"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".zip"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.pdf"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".pdf"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.csv"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".csv"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.bat"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".bat"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xlsx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xlsx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xls"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xls"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.png"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".png"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.jpg"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".jpg"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.zip"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".zip"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.pdf"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".pdf"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.csv"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".csv"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.bat"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".bat"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xlsx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xlsx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xls"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xls"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.png"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".png"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.jpg"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".jpg"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.zip"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".zip"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.doc"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".doc"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.docx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".docx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.csv"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".csv"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.bat"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".bat"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xlsx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xlsx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xls"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xls"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.png"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".png"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.jpg"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".jpg"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.zip"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").WaitForObjectToLoadWithTime 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".zip"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.doc"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".doc"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.docx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".docx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.csv"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".csv"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.bat"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".bat"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xlsx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xlsx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xls"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xls"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.jpg"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".jpg"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.zip"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".zip"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.doc"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".doc"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.docx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".docx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.csv"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".csv"" is not allowed."

End Function
''###################################################################################################
''Function Name : UploadInvalidDocumentsForWordTempate
''Description : Upload InvalidDocuments ForWordTempate and validate
''Author : Gallop
''###################################################################################################
Public Function UploadInvalidDocumentsForWordTempate
Environment.Value("StepName") = "UploadInvalidDocumentsForWordTempate"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.bat"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyExists 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".bat"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xlsx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xlsx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xls"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xls"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.png"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".png"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.jpg"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".jpg"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.zip"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".zip"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.pdf"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".pdf"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Tempate NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.csv"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".csv"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.bat"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".bat"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xlsx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xlsx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xls"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xls"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.png"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".png"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.jpg"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".jpg"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.zip"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".zip"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.pdf"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".pdf"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Word Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.csv"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".csv"" is not allowed."

End Function
''###################################################################################################
''Function Name : UploadInvalidDocumentsForPDFTemplate
''Description : Upload InvalidDocuments ForPDF Template
''Author : Gallop
''###################################################################################################
Public Function UploadInvalidDocumentsForPDFTemplate
Environment.Value("StepName") = "UploadInvalidDocumentsForPDFTemplate"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.bat"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".bat"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xlsx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xlsx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xls"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xls"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.png"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".png"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.jpg"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".jpg"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.zip"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".zip"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.doc"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".doc"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.docx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".docx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template NO Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.csv"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".csv"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.bat"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".bat"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xlsx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xlsx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xls"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xls"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.png"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".png"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.jpg"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".jpg"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.zip"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".zip"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.doc"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".doc"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.docx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".docx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "PDF Template w Review"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.csv"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".csv"" is not allowed."

End Function
''###################################################################################################
''Function Name : UploadInvalidDocumetsForAttachements
''Description : Upload Invalid documents for Attachements
''Author : Gallop
''###################################################################################################
Public Function UploadInvalidDocumetsForAttachements
Environment.Value("StepName") = "UploadInvalidDocumetsForAttachements"
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.bat"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".bat"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xlsx"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xlsx"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.xls"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".xls"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.jpg"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".jpg"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.zip"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".zip"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.doc"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".doc"" is not allowed."
	Browser("Summary").Page("Summary").WebTable("TBL_Documents").ClickONAttachForSpecificDataInTable "Attachment Only"
	Browser("Summary").Page("Summary").WebElement("LBL_Documents").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebFile("WFL_UploadFileAddDoc").ClickObject 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").VerifyExists 
	Browser("Summary").Dialog("Choose File to Upload").WinEdit("TXT_Filename").TypeText dtScenario.value("ResourcePath")&"\Testing.csv"
	Browser("Summary").Dialog("Choose File to Upload").WinButton("BTN_Open").ClickObject 
	Browser("Summary").Page("Summary").WebButton("BTN_UploadDocument").ClickObject 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").WaitForObjectToLoad 
	Browser("Summary").Page("Summary").WebElement("LBL_FileExtenstionNoAllowed").VerifyDisplayedText "Error adding document attachment. File extension "".csv"" is not allowed."

End Function
''###################################################################################################
''Function Name : ValidateAuditBubble
''Description : Audit Bubble Testing
''Author : Gallop
''###################################################################################################
Public Function ValidateAuditBubble
Environment.Value("StepName") = "ValidateAuditBubble"
	Browser("Entry").Page("Summary").WebButton("BTN_Back").WaitForObjectToLoad 
	Browser("Entry").Page("Summary").WebButton("BTN_Back").ClickObject 
	Browser("Entry").Page("Submit").WebElement("DTL_SetQuestionValues1").WaitForObjectToLoad 
	Browser("Entry").Page("Submit").WebElement("DTL_SetQuestionValues1").SelectListBox "12345"
	Browser("Entry").Page("Submit").WebElement("DTL_SetQuestionValues2").SelectListBox "12345"
	Browser("Entry").Page("Submit").WebEdit("TXT_AuditBubble1").VerifyValue "12345"
	Browser("Entry").Page("Submit").WebEdit("TXT_AuditBubble2").VerifyValue "12345"
	Browser("Entry").Page("Submit").WebButton("Audit1_Submit").ClickObject 
	Browser("Entry").Page("Summary").WebElement("LBL_QAAuditBubbleTest").WaitForObjectToLoad 
	Browser("Entry").Page("Summary").WebElement("TWL_QA_Audit1").VerifyDisplayedText "12345"
	Browser("Entry").Page("Summary").WebElement("TWL_QA_Audit2").VerifyDisplayedText "12345"
	Browser("Entry").Page("Summary").WebElement("DTL_SetQuestionValues1").VerifyDisplayedText "12345"
	Browser("Entry").Page("Summary").WebElement("DTL_SetQuestionValues2").VerifyDisplayedText "12345"
	Browser("Entry").Page("Summary").WebElement("BBL_Audit1").ClickObject 
	Browser("Entry").Page("Summary").WebTable("TBL_AuditHistory1").VerifyTextInTable "123"
	Browser("Entry").Page("Summary").WebElement("BBL_Audit2").ClickObject 
	Browser("Entry").Page("Summary").WebTable("TBL_AuditHistory2").VerifyTextInTable "123"
	Browser("Entry").Page("Summary").WebButton("BTN_Back").ClickObject 
	Browser("Entry").Page("Submit").WebElement("DTL_SetQuestionValues1").WaitForObjectToLoad 
	Browser("Entry").Page("Submit").WebElement("DTL_SetQuestionValues1").SelectListBox "1234"
	Browser("Entry").Page("Submit").WebElement("DTL_SetQuestionValues2").SelectListBox "1234"
	Browser("Entry").Page("Submit").WebEdit("TXT_AuditBubble1").VerifyValue "1234"
	Browser("Entry").Page("Submit").WebEdit("TXT_AuditBubble2").VerifyValue "1234"
	Browser("Entry").Page("Submit").WebButton("Audit1_Submit").ClickObject 
	Browser("Entry").Page("Summary").WebElement("LBL_QAAuditBubbleTest").WaitForObjectToLoad 
	Browser("Entry").Page("Summary").WebElement("TWL_QA_Audit1").VerifyDisplayedText "1234"
	Browser("Entry").Page("Summary").WebElement("TWL_QA_Audit2").VerifyDisplayedText "1234"
	Browser("Entry").Page("Summary").WebElement("DTL_SetQuestionValues1").VerifyDisplayedText "1234"
	Browser("Entry").Page("Summary").WebElement("DTL_SetQuestionValues2").VerifyDisplayedText "1234"
	Browser("Entry").Page("Summary").WebElement("BBL_Audit1").ClickObject 
	Browser("Entry").Page("Summary").WebTable("TBL_AuditHistory1").VerifyTextInTable "12345"
	Browser("Entry").Page("Summary").WebTable("TBL_AuditHistory1").VerifyTextInTable "123"
	Browser("Entry").Page("Summary").WebElement("BBL_Audit2").ClickObject 
	Browser("Entry").Page("Summary").WebTable("TBL_AuditHistory2").VerifyTextInTable "12345"
	Browser("Entry").Page("Summary").WebTable("TBL_AuditHistory2").VerifyTextInTable "123"
	Browser("Entry").Page("Summary").Link("LNK_Audit").ClickObject 
	Browser("Entry").Page("Summary").WebElement("LBL_AuditHistory").WaitForObjectToLoad 
	Browser("Entry").Page("Summary").WebElement("LBL_AuditHistory").VerifyExists 
	Browser("Entry").Page("Summary").WebButton("BTN_ContextMenu").ClickObject 
	Browser("Entry").Page("Summary").WebElement("CB_View All").ClickObject 
	Browser("Entry").Page("Summary").WebTable("TBL_Audit").WaitForObjectToLoad 
	Browser("Entry").Page("Summary").WebTable("TBL_Audit").ExpandDataInDataTable 
	Browser("Entry").Page("Summary").WebTable("TBL_Name").VerifyDataInRowForSpecificColumnName "QA.Audit.TestOne|12345|1234"
	Browser("Entry").Page("Summary").WebTable("TBL_Name").VerifyDataInRowForSpecificColumnName "QA.Audit.TestTwo|12345|1234"
	Browser("Entry").Page("Summary").WebTable("TBL_Name").VerifyDataInRowForSpecificColumnName "QA.DropDown.SetValue1|12345|1234"
	Browser("Entry").Page("Summary").WebTable("TBL_Name").VerifyDataInRowForSpecificColumnName "QA.DropDown.SetValue2|12345|1234"
	Browser("Entry").Page("Summary").WebTable("TBL_Name_2").VerifyDataInRowForSpecificColumnName "QA.Audit.TestOne|123|12345"
	Browser("Entry").Page("Summary").WebTable("TBL_Name_2").VerifyDataInRowForSpecificColumnName "QA.Audit.TestTwo|123|12345"
	Browser("Entry").Page("Summary").WebTable("TBL_Name_2").VerifyDataInRowForSpecificColumnName "QA.DropDown.SetValue1|Select a value|12345"
	Browser("Entry").Page("Summary").WebTable("TBL_Name_2").VerifyDataInRowForSpecificColumnName "QA.DropDown.SetValue2|Select a value|12345"

End Function
''###################################################################################################
''Function Name : EnterInvalidData
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function EnterInvalidData
Environment.Value("StepName") = "EnterInvalidData"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_LegalAddress").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_LA_Address").TypeText "722 Jetton"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_LA_City").TypeText "Davidson"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_LA_State").TypeText "NC"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_LA_Zip").TypeText "28036"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebButton("BTN_Validate").ClickObject 

End Function
''###################################################################################################
''Function Name : LoginInvalidUser
''Description : Login to application with invalid password's
''Author : Gallop
''###################################################################################################
Public Function LoginInvalidUser
Environment.Value("StepName") = "LoginInvalidUser"
	Call StartApplication()
	Browser("Login").Page("Login").WebEdit("TXT_user_name").EnterText dtScenario.value("username")
	Browser("Login").Page("Login").WebEdit("TXT_user_password").EnterText dtScenario.value("password1")
	Browser("Login").Page("Login").WebButton("BTN_Log In").ClickObject 
	Browser("Login").Page("Login").WebEdit("TXT_user_password").EnterText dtScenario.value("password2")
	Browser("Login").Page("Login").WebButton("BTN_Log In").ClickObject 
	Browser("Login").Page("Login").WebEdit("TXT_user_password").EnterText dtScenario.value("password3")
	Browser("Login").Page("Login").WebButton("BTN_Log In").ClickObject 
	Browser("Login").Page("Login").WebEdit("TXT_user_password").EnterText dtScenario.value("password4")
	Browser("Login").Page("Login").WebButton("BTN_Log In").ClickObject 
	Browser("Login").Page("Login").WebEdit("TXT_user_password").EnterText dtScenario.value("password5")
	Browser("Login").Page("Login").WebButton("BTN_Log In").ClickObject 
	Browser("Login").Page("Login").WebEdit("TXT_user_password").EnterText dtScenario.value("password6")
	Browser("Login").Page("Login").WebButton("BTN_Log In").ClickObject 
	Browser("Login").Page("Login").WebEdit("TXT_user_password").EnterText dtScenario.value("password7")
	Browser("Login").Page("Login").WebButton("BTN_Log In").ClickObject 
	Browser("Login").Page("Login").WebEdit("TXT_user_password").EnterText dtScenario.value("password8")
	Browser("Login").Page("Login").WebButton("BTN_Log In").ClickObject 
	Browser("Login").Page("Login").WebEdit("TXT_user_password").EnterText dtScenario.value("password9")
	Browser("Login").Page("Login").WebButton("BTN_Log In").ClickObject 
	Browser("Login").Page("Login").WebEdit("TXT_user_password").WaitForObjectToLoad 

End Function
''###################################################################################################
''Function Name : ValidateLockOutMessage
''Description : Logout from application and close browsers
''Author : Gallop
''###################################################################################################
Public Function ValidateLockOutMessage
Environment.Value("StepName") = "ValidateLockOutMessage"
	Browser("Login").Page("Login").WebElement("DWL_LockOutErrMessage").VerifyDisplayedText "Sorry, maximum number of retries were exceeded for this account. Please contact your administrator to unlock your account and reset your password."
	Call CloseAllBrowsers()

End Function
''###################################################################################################
''Function Name : ValidateUserAccountUnlock
''Description : Unlock user account 
''Author : Gallop
''###################################################################################################
Public Function ValidateUserAccountUnlock
Environment.Value("StepName") = "ValidateUserAccountUnlock"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_SecurityAdministration").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_ModifyUser").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_ModifyUser").ClickObject 
	Browser("FindEmployee").Page("FindEmployee").WebElement("LBL_ModifyUser").WaitForObjectToLoad 5
	Browser("FindEmployee").Page("FindEmployee").WebEdit("TXT_UserId").EnterText "user"
	Browser("FindEmployee").Page("FindEmployee").WebButton("BTN_Search").ClickObject 
	Browser("FindEmployee").Page("FindEmployee").WebTable("TBL_Filter").WaitForObjectToLoad 5
	Browser("FindEmployee").Page("FindEmployee").WebTable("TBL_Filter").VerifyExists 
	Browser("FindEmployee").Page("FindEmployee").WebTable("TBL_Filter").VerifyvalueInSpecificCell 2,6,"LOCKED"
	Browser("FindEmployee").Page("FindEmployee").WebTable("TBL_Filter").VerifyandClickLinkInRow "Joe lockout User|  "
	Browser("FindEmployee").Page("ModifyUserInformation").WebElement("LBL_UserMaintenance").WaitForObjectToLoad 5
	Browser("FindEmployee").Page("ModifyUserInformation").WebButton("BTN_UnlockAccount").VerifyExists 
	Browser("FindEmployee").Page("ModifyUserInformation").WebButton("BTN_UnlockAccount").ClickObject 
	Browser("FindEmployee").Page("FindEmployee").WebElement("LBL_ModifyUser").WaitForObjectToLoad 5
	Browser("FindEmployee").Page("FindEmployee").WebElement("LBL_ModifyUser").VerifyExists 
	Browser("FindEmployee").Page("FindEmployee").WebEdit("TXT_UserId").EnterText "user"
	Browser("FindEmployee").Page("FindEmployee").WebButton("BTN_Search").ClickObject 
	Browser("FindEmployee").Page("FindEmployee").WebTable("TBL_Filter").WaitForObjectToLoad 5
	Browser("FindEmployee").Page("FindEmployee").WebTable("TBL_Filter").VerifyvalueInSpecificCell 2,6,"ACTIVE"
	Browser("FindEmployee").Page("FindEmployee").WebButton("BTN_WebGearButton").ClickObject 
	Browser("FindEmployee").Page("FindEmployee").WebElement("LNK_Logout").VerifyExists 
	Browser("FindEmployee").Page("FindEmployee").WebElement("LNK_Logout").ClickObject 
	Browser("Login").Page("Login").WebButton("BTN_Log In").WaitForObjectToLoad 
	Browser("Login").Page("Login").WebEdit("TXT_user_name").EnterText "user"
	Browser("Login").Page("Login").WebEdit("TXT_user_password").EnterText "password"
	Browser("Login").Page("Login").WebButton("BTN_Log In").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("TAB _Home").VerifyExists 

End Function
''###################################################################################################
''Function Name : NavigateToSmartyStreetTesting
''Description : NavigateToSmartyStreetTesting
''Author : Gallop
''###################################################################################################
Public Function NavigateToSmartyStreetTesting
Environment.Value("StepName") = "NavigateToSmartyStreetTesting"
	Browser("Work Queue").Page("Work Queue").WebElement("TAB_Tests").ClickObject 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_SmartyStreetsTesting").VerifyExists 
	Browser("Work Queue").Page("Work Queue").WebElement("LNK_SmartyStreetsTesting").ClickObject 

End Function
''###################################################################################################
''Function Name : ForeignAddressRetryValidation
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ForeignAddressRetryValidation
Environment.Value("StepName") = "ForeignAddressRetryValidation"
	Browser("Entry").Page("Entry").WebButton("BTN_Retry").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("LBL_FA_SorryButThereWas").VerifyDisplayedText "Sorry, but there was "
	Browser("Entry").Page("Entry").WebButton("BTN_Skip").VerifyExists 
	Browser("Entry").Page("Entry").WebEdit("TXT_FA_Address1").EnterText "Flat  No 100"
	Browser("Entry").Page("Entry").WebEdit("TXT_FA_Address2").EnterText "TriVeni Apartments"
	Browser("Entry").Page("Entry").WebEdit("TXT_FA_City").EnterText "Pitam Pura"
	Browser("Entry").Page("Entry").WebEdit("TXT_FA_Zip").EnterText "32232"
	Browser("Entry").Page("Entry").WebButton("BTN_Retry").ClickObject 
	Browser("Entry").Page("Entry").WebButton("BTN_OverRide").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_OverRide").ClickObject 
	Browser("Entry").Page("Entry").WebButton("BTN_Accept").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebButton("BTN_Accept").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Cancel").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("LBL_AddressVerificationlBypass").VerifyDisplayedText "Address verification will be bypassed"
	Browser("Entry").Page("Entry").WebButton("BTN_Accept").ClickObject 
	Browser("Entry").Page("Entry").Link("LNK_Close").VerifyExists 
	Browser("Entry").Page("Entry").Link("LNK_Close").ClickObject 

End Function
''###################################################################################################
''Function Name : ForeignAddressSkipValidation
''Description : Validating the foreign address details by clicking skip button
''Author : Gallop
''###################################################################################################
Public Function ForeignAddressSkipValidation
Environment.Value("StepName") = "ForeignAddressSkipValidation"
	Browser("Entry").Page("Entry").WebButton("BTN_Retry").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebButton("BTN_Retry").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("LBL_FA_SorryButThereWas").VerifyDisplayedText "Sorry, but there was an error"
	Browser("Entry").Page("Entry").WebButton("BTN_Skip").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Skip").ClickObject 
	Browser("Entry").Page("Entry").WebElement("LBL_AddressVerificationlBypass").VerifyDisplayedText "Address verification will be temporarily bypassed"
	Browser("Entry").Page("Entry").WebButton("BTN_Accept").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Cancel").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Accept").ClickObject 
	Browser("Entry").Page("Entry").Link("LNK_Close").VerifyExists 
	Browser("Entry").Page("Entry").Link("LNK_Close").ClickObject 

End Function
''###################################################################################################
''Function Name : LegalAddressOverRideValidation
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function LegalAddressOverRideValidation
Environment.Value("StepName") = "LegalAddressOverRideValidation"
	Browser("Entry").Page("Entry").WebButton("BTN_Retry").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebElement("LBL_LA_LegalAddressIsNotValid").VerifyDisplayedText "Legal Address is not valid"
	Browser("Entry").Page("Entry").WebButton("BTN_OverRide").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_OverRide").ClickObject 
	Browser("Entry").Page("Entry").WebElement("LBL_AddressVerificationlBypass").VerifyDisplayedText "Address verification will be bypassed"
	Browser("Entry").Page("Entry").WebButton("BTN_Accept").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Cancel").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Accept").ClickObject 

End Function
''###################################################################################################
''Function Name : LegalAddressRetryValidation
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function LegalAddressRetryValidation
Environment.Value("StepName") = "LegalAddressRetryValidation"
	Browser("Entry").Page("Entry").WebButton("BTN_Retry").WaitForObjectToLoad 
	Browser("Entry").Page("Entry").WebButton("BTN_Retry").VerifyExists 
	Browser("Entry").Page("Entry").WebElement("LBL_LA_LegalAddressIsNotValid").VerifyDisplayedText "Legal Address is not valid"
	Browser("Entry").Page("Entry").WebEdit("TXT_LA_Address").EnterText "721 Jetton"
	Browser("Entry").Page("Entry").WebButton("BTN_Retry").ClickObject 
	Browser("Entry").Page("Entry").WebElement("LBL_AddressValidationSuccess").VerifyDisplayedText "Address validated successfully"
	Browser("Entry").Page("Entry").WebButton("BTN_Accept").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Cancel").VerifyExists 
	Browser("Entry").Page("Entry").WebButton("BTN_Accept").ClickObject 
	Browser("Entry").Page("Entry").Link("LNK_Close").VerifyExists 
	Browser("Entry").Page("Entry").Link("LNK_Close").ClickObject 

End Function
''###################################################################################################
''Function Name : ValidateCorrectLater
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function ValidateCorrectLater
Environment.Value("StepName") = "ValidateCorrectLater"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_LegalAddress").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_LA_Address").EnterText "722 Jetton"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebButton("BTN_CorrectLater").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebButton("BTN_CorrectLater").ClickObject 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Address").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Address").VerifyDisplayedText dtScenario.value("LA_Address")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_ValidationStatus").VerifyDisplayedText dtScenario.value("LA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Overridden").VerifyDisplayedText dtScenario.value("LA_Overridden")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_EA_OverriddenStatus").VerifyDisplayedText dtScenario.value("EA_OverriddenStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_EA_ValidationStatus").VerifyDisplayedText dtScenario.value("EA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FFA_Overridden").VerifyDisplayedText dtScenario.value("FFA_Overridden")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FFA_ValidationStatus").VerifyDisplayedText dtScenario.value("FFA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_State").VerifyDisplayedText dtScenario.value("FA_State")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_ValidationStatus").VerifyDisplayedText dtScenario.value("FA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_AddressOverridden").VerifyDisplayedText dtScenario.value("FA_AddressOverridden")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebButton("BTN_Back").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebButton("BTN_Back").ClickObject 

End Function
''###################################################################################################
''Function Name : VerifyDefaultValues
''Description : VerifydefaultFieldsAndValues
''Author : Gallop
''###################################################################################################
Public Function VerifyDefaultValues
Environment.Value("StepName") = "VerifyDefaultValues"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_LegalAddress").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_LegalAddress").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_EmployerAddress").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_Free-formAddress").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_ForeignAddress").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_LA_ValidationStatus").VerifyValue "Not Run"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_FA_StateAlabamaselect").VerifyDisplayedText "Alabama"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebButton("BTN_Validate").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebButton("BTN_Next").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebButton("BTN_CorrectLater").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebButton("BTN_Back").VerifyExists 

End Function
''###################################################################################################
''Function Name : VerifyForeignAddressRetryDataOnReadOnlyPage
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function VerifyForeignAddressRetryDataOnReadOnlyPage
Environment.Value("StepName") = "VerifyForeignAddressRetryDataOnReadOnlyPage"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_LegalAddress").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebButton("BTN_Next").ClickObject 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Address").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Address").VerifyDisplayedText dtScenario.value("LA_Address")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_City").VerifyDisplayedText dtScenario.value("LA_City")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_State").VerifyDisplayedText dtScenario.value("LA_State")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Zip").VerifyDisplayedText dtScenario.value("LA_Zip")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_ValidationStatus").VerifyDisplayedText dtScenario.value("LA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Overridden").VerifyDisplayedText dtScenario.value("LA_Overridden")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_EA_OverriddenStatus").VerifyDisplayedText dtScenario.value("EA_OverriddenStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_EA_ValidationStatus").VerifyDisplayedText dtScenario.value("EA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FFA_Overridden").VerifyDisplayedText dtScenario.value("FFA_Overridden")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FFA_ValidationStatus").VerifyDisplayedText dtScenario.value("FFA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_AddressLine1").VerifyDisplayedText dtScenario.value("FA_AddressLine1")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_AddressLine2").VerifyDisplayedText dtScenario.value("FA_AddressLine2")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_City").VerifyDisplayedText dtScenario.value("FA_City")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_State").VerifyDisplayedText dtScenario.value("FA_State")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_Zip").VerifyDisplayedText dtScenario.value("FA_Zip")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_ValidationStatus").VerifyDisplayedText dtScenario.value("FA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_AddressOverridden").VerifyDisplayedText dtScenario.value("FA_AddressOverridden")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebButton("WebButton_Gear").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebButton("WebButton_Gear").ClickObject 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("LNK_Debug").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("LNK_Debug").ClickDebugObject 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").WaitForObjectToLoad 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyExists 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 3,3,"722 Jetton"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 4,3,"Davidson"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 5,3,"NC"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 6,3,"28036"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 7,3,"Invalid"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 8,3,"Overridden"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 9,3,"Not Run"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 10,3,"False"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 11,3,"Not Run"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 12,3,"False"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 20,3,"Flat No 100"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 21,3,"TriVeni Apartments"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 22,3,"Pitam Pura"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 24,3,"32232"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 26,3,"SetOverridden"

End Function
''###################################################################################################
''Function Name : VerifyForeignAddressSkipDataOnReadOnlyPage
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function VerifyForeignAddressSkipDataOnReadOnlyPage
Environment.Value("StepName") = "VerifyForeignAddressSkipDataOnReadOnlyPage"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_LegalAddress").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebButton("BTN_Next").ClickObject 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Address").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Address").VerifyDisplayedText dtScenario.value("LA_Address")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_City").VerifyDisplayedText dtScenario.value("LA_City")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_State").VerifyDisplayedText dtScenario.value("LA_State")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Zip").VerifyDisplayedText dtScenario.value("LA_Zip")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_ValidationStatus").VerifyDisplayedText dtScenario.value("LA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Overridden").VerifyDisplayedText dtScenario.value("LA_Overridden")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_EA_OverriddenStatus").VerifyDisplayedText dtScenario.value("EA_OverriddenStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_EA_ValidationStatus").VerifyDisplayedText dtScenario.value("EA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FFA_Overridden").VerifyDisplayedText dtScenario.value("FFA_Overridden")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FFA_ValidationStatus").VerifyDisplayedText dtScenario.value("FFA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_State").VerifyDisplayedText dtScenario.value("FA_State")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_ValidationStatus").VerifyDisplayedText dtScenario.value("FA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_AddressOverridden").VerifyDisplayedText dtScenario.value("FA_AddressOverridden")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebButton("WebButton_Gear").ClickObject 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("LNK_Debug").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("LNK_Debug").ClickDebugObject 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").WaitForObjectToLoad 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyExists 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 3,3,"722 Jetton"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 4,3,"Davidson"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 5,3,"NC"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 6,3,"28036"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 7,3,"Invalid"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 8,3,"OVERRIDDEN"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 9,3,"Not Run"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 10,3,"False"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 11,3,"Not Run"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 20,3,"AL"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 21,3,"Error"

End Function
''###################################################################################################
''Function Name : VerifyLegalAddressOverRideDataOnReadOnlyPage
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function VerifyLegalAddressOverRideDataOnReadOnlyPage
Environment.Value("StepName") = "VerifyLegalAddressOverRideDataOnReadOnlyPage"
	Browser("Entry").Page("Entry").Link("LNK_Close").VerifyExists 
	Browser("Entry").Page("Entry").Link("LNK_Close").ClickObject 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_LegalAddress").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebButton("BTN_Next").ClickObject 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Address").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Address").VerifyDisplayedText dtScenario.value("LA_Address")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_City").VerifyDisplayedText dtScenario.value("LA_City")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_State").VerifyDisplayedText dtScenario.value("LA_State")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Zip").VerifyDisplayedText dtScenario.value("LA_Zip")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_ValidationStatus").VerifyDisplayedText dtScenario.value("LA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Overridden").VerifyDisplayedText dtScenario.value("LA_Overridden")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_EA_OverriddenStatus").VerifyDisplayedText dtScenario.value("EA_OverriddenStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_EA_ValidationStatus").VerifyDisplayedText dtScenario.value("EA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FFA_Overridden").VerifyDisplayedText dtScenario.value("FFA_Overridden")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FFA_ValidationStatus").VerifyDisplayedText dtScenario.value("FFA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_State").VerifyDisplayedText dtScenario.value("FA_State")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_ValidationStatus").VerifyDisplayedText dtScenario.value("FA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebButton("WebButton_Gear").ClickObject 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("LNK_Debug").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("LNK_Debug").ClickDebugObject 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").WaitForObjectToLoad 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyExists 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 3,3,"722 Jetton"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 4,3,"Davidson"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 5,3,"NC"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 6,3,"28036"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 7,3,"INVALID"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 8,3,"Overridden"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 9,3,"NOT RUN"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 10,3,"False"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 11,3,"NOT RUN"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 20,3,"AL"

End Function
''###################################################################################################
''Function Name : VerifyLegalAddressRetryDataOnReadOnlyPage
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function VerifyLegalAddressRetryDataOnReadOnlyPage
Environment.Value("StepName") = "VerifyLegalAddressRetryDataOnReadOnlyPage"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_LegalAddress").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebButton("BTN_Next").ClickObject 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Address").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Address").VerifyDisplayedText dtScenario.value("LA_Address")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_City").VerifyDisplayedText dtScenario.value("LA_City")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_State").VerifyDisplayedText dtScenario.value("LA_State")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Zip").VerifyDisplayedText "28036-7107"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_ValidationStatus").VerifyDisplayedText dtScenario.value("LA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_EA_OverriddenStatus").VerifyDisplayedText dtScenario.value("EA_OverriddenStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_EA_ValidationStatus").VerifyDisplayedText dtScenario.value("EA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FFA_Overridden").VerifyDisplayedText dtScenario.value("FFA_Overridden")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FFA_ValidationStatus").VerifyDisplayedText dtScenario.value("FFA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_State").VerifyDisplayedText dtScenario.value("FA_State")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_ValidationStatus").VerifyDisplayedText dtScenario.value("FA_ValidationStatus")
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebButton("WebButton_Gear").ClickObject 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("LNK_Debug").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("LNK_Debug").ClickDebugObject 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").WaitForObjectToLoad 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyExists 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 3,3,"721 Jetton St"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 4,3,"Davidson"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 5,3,"NC"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 6,3,"28036-7107"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 7,3,"Valid"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 8,3,"Not Run"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 9,3,"False"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 10,3,"Not Run"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 11,3,"False"

End Function
''###################################################################################################
''Function Name : VerifyReadOnlyValuesInSmartyStreets
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function VerifyReadOnlyValuesInSmartyStreets
Environment.Value("StepName") = "VerifyReadOnlyValuesInSmartyStreets"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_LegalAddress").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_LA_Address").EnterText "721 Jetton St"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_LA_City").TypeText "Davidson"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_LA_State").TypeText "NC"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_LA_Zip").TypeText "28036"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_EA_Street").TypeText "1971 University Blvd"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_EA_CityZipState").EnterText "Lynchburg, VA, 24515"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_FFA_FullName").EnterText "1971 University Blvd, Lynchburg, VA, 24515"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("DWL_FA_CountrySelect").ClickObject 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("DWL_FA_CountrySelect").SelectListBox "India"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_FA_Address1").TypeText "Flat No.100"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_FA_Address2").TypeText "Triveni Apartments"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_FA_City").TypeText "Pitam Pura"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebEdit("TXT_FA_ForgienZip").TypeText "NEW DELHI 110034"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebButton("BTN_Next").ClickObject 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Address").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Address").VerifyDisplayedText "721 Jetton St"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_City").VerifyDisplayedText "Davidson"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_State").VerifyDisplayedText "NC"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_LA_Zip").VerifyDisplayedText "28036-7107"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_EA_Street").VerifyDisplayedText "1971 University Blvd"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_EA_CityStateZip").VerifyDisplayedText "Lynchburg VA 24515-0002"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FFA_FullAddress").VerifyDisplayedText "1971 University Blvd Lynchburg VA 24515-0002"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_Country").VerifyDisplayedText "India"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_AddressLine1").VerifyDisplayedText "Flat No.100"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_AddressLine2").VerifyDisplayedText "Triveni Apartments"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_City").VerifyDisplayedText "Pitam Pura"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("RWL_FA_ForeignZip").VerifyDisplayedText "NEW DELHI 110034"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebButton("WebButton_Gear").ClickObject 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("LNK_Debug").VerifyExists 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTestingReadOnly").WebElement("LNK_Debug").ClickDebugObject 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").PageSync 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyExists 
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 3,3,"721 Jetton St"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 4,3,"Davidson"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 5,3,"NC"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 6,3,"28036-7107"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 7,3,"Valid"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 8,3,"1971 University Blvd"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 9,3,"Lynchburg VA 24515-0002"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 10,3,"Valid"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 11,3,"1971 University Blvd Lynchburg VA 24515-0002"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 12,3,"Valid"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 20,3,"Flat No.100"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 21,3,"IN"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 22,3,"Triveni Apartments"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 23,3,"Pitam Pura"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 24,3,"AL"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 25,3,"NEW DELHI 110034"
	Browser("IFS Page Debugging").Page("IFS Page Debugging").WebTable("TBL_Debug").VerifyvalueInSpecificCell 26,3,"Not Run"

End Function
''###################################################################################################
''Function Name : VerifyRequiredErrorMeesage
''Description : Description of Component
''Author : Gallop
''###################################################################################################
Public Function VerifyRequiredErrorMeesage
Environment.Value("StepName") = "VerifyRequiredErrorMeesage"
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_LegalAddress").WaitForObjectToLoad 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebButton("BTN_Next").clickNext 
	Browser("SmartyStreetsTesting").Page("SmartyStreetsTesting").WebElement("LBL_LA_Address").VerifyDisplayedText "This is a required field"

End Function