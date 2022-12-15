With Browser("CreationTime:=0")
	while .Exist(0)	'Loop to close all open browsers
		.Close 
	Wend
End With

If Not Browser("Logon").Page("Logon").Exist(1) Then
	SystemUtil.Run  "chrome.exe","","","",3		'launch the chrome
	Set AppContext=Browser("CreationTime:=0")	'Set the variable for what application (in this case the browser) we are acting upon
	wait 5 
	'URL = "https://sap-hana.mfdemoportal.com:44300/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home"
	URL = "https://sap-hana.mfdemoportal.com:44300/sap/bc/ui5_ui5/ui2/ushell/shells/abap/FioriLaunchpad.html"
	AppContext.ClearCache		'Clear the browser cache to ensure you're getting the latest forms from the application
	AppContext.Navigate URL		'Navigate to the application URL
	AppContext.Maximize			'Maximize the application to give the best chance that the fields will be visible on the screen
	AppContext.Sync				'Wait for the browser to stop spinning
End If

With Browser("Logon")
	.Page("Logon").WebEdit("sap-user").Set Parameter("userName")
	.Page("Logon").WebEdit("sap-password").Set Parameter("password")
	.Page("Logon").WebButton("Log On").Click
End With
