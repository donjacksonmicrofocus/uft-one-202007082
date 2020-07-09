'===========================================================
'Function for creating a number at run time based on current time down to the second, to allow for a unique number each time the script is run
'===========================================================
Function fnRandomNumberWithDateTimeStamp()

'Find out the current date and time
Dim sDate : sDate = Day(Now)
Dim sMonth : sMonth = Month(Now)
Dim sYear : sYear = Year(Now)
Dim sHour : sHour = Hour(Now)
Dim sMinute : sMinute = Minute(Now)
Dim sSecond : sSecond = Second(Now)

'Create Random Number
fnRandomNumberWithDateTimeStamp = Int(sDate & sMonth & sYear & sHour & sMinute & sSecond)

'======================== End Function =====================
End Function

'===========================================================
'Function for debugging properties at run time to output to the log, once script is working, can be deleted
'===========================================================
Function PropertiesDebug

'Dim CPActual, CPExpected
'Debug code to determine why the checkpoint was failing, turns out that there is a trailing space in the application code that the result HTML was trimming when displaying expected vs. actual
'CPExpected = "'" & DataTable.Value ("FullName") & "'"										'Set the variable for what is in the data table, enclose with single quotes so we can find leading/trailing spaces
'CPActual = Browser("Browser").Page("SuccessFactors: Candidates").Link("CandidateName").GetROProperty("text")	'Get the actual text from the object at run time
'CPActual = "'" & CPActual & "'"															'Set the variable for what is the object property at run time enclosed with a single quotes so we can find leading/trailing spaces
'Print "Expected is " & CPExpected															'Output the expected value to the output log
'Print "Actual is " & CPActual																'Output the actual value to the output log
	
End Function

Dim FirstName, LastName, Email

FirstName = "FN" & fnRandomNumberWithDateTimeStamp											'Create a unique first name
LastName = "LN" & fnRandomNumberWithDateTimeStamp											'Create a unique last name

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
SystemUtil.Run "CHROME.exe" ,"","","",3														'launch Chrome, could be data drive to launch other browser (e.g. Firefox)
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate "http://www.sage.com/"													'Navigate to the application URL @@ hightlight id_;_1311600_;_script infofile_;_ZIP::ssf1.xml_;_
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning

AIUtil.SetContext AppContext																'Tell the AI engine to point against the application
AIUtil.FindTextBlock("Products V").Click													'Click the Products menu item, the OCR recognizes the down arrow as a V
AIUtil.FindTextBlock("Desktop accounting software").Click									'Click the text
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil("button", "Try it now").Click														'Click the button
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.FindTextBlock("First Name").Click													'Click on the First Name text to have the application move the text above the text box, to ensure the AI engine will select the right field
AIUtil("text_box", "First Name").Type FirstName												'Enter the unique first name into the field
AIUtil("text_box", "Last Name").Highlight													'Highlight the Last Name field, this gets the AI engine to rescan the page as well
AIUtil("text_box", "Last Name").Type LastName												'Enter the unique last name into the field
AIUtil.FindTextBlock("Zip or Postal Code").Click											'Click on the Zip or Postal Code text to have the application move the text above the text box, to ensure the AI engine will select the right field
AIUtil("text_box", "Zip or Postal Code").Type "90210"										'Enter the zip code

AppContext.Close																			'Close the application
