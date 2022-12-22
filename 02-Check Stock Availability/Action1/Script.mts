wait(5)
With Browser("Home")
	While Not .Page("Home").SAPUIButton("More groups").GetROProperty("disabled") = "False"
		wait(1)
	Wend
	wait(2)
	.Page("Home").SAPUIButton("More groups").Highlight @@ script infofile_;_ZIP::ssf13.xml_;_
	.Page("Home").SAPUIButton("More groups").HoverTap
	.Page("Home").SAPUIButton("More groups").Click @@ script infofile_;_ZIP::ssf14.xml_;_
	
	If Not .Page("Home").WebElement("Stock Monitoring").Exist(1) Then
		Reporter.ReportEvent micFail, "LogIn User", "User is incorrect"
		ExitAction
	End If
	.Page("Home").WebElement("Stock Monitoring").Click
	.Page("Home").Link("StockMultiple Materials").Click
	.Page("Home").SAPUIButton("Expand Header").Click
	wait(5)
	.Page("Home").WebEdit("Material Number").Set "EWMS4-01"
	pressButton("TAB")
	wait 5
	.Page("Home").WebEdit("Material Number").Set "EWMS4-02"
	pressButton("TAB")
	'Browser("Home").Page("Home").SAPUITextEdit("Storage Location").Set "171S"
	.Page("Home").WebEdit("Storage Location").Set "171S"
	pressButton("TAB")
	.Page("Home").SAPUIButton("Go").Click @@ script infofile_;_ZIP::ssf11.xml_;_
	wait(5)

	'	Check Stock Availability
	For i = 2 To 3
		materialNumber = .Page("Home").WebTable("tblMaterials").GetCellData(i,1)
		blockedStock = .Page("Home").WebTable("tblMaterials").GetCellData(i,11)
		arrblockedStock = split(blockedStock, ".")
		blockedStock = arrblockedStock(0)
		If blockedStock > 1 Then
			Reporter.ReportEvent micPass, "Blocked Stock is greater than 1", "Blocked Stock of material "&materialNumber& " is: "&blockedStock
			If materialNumber = "EWMS4-02" Then
				blockedStock_EWMS4_02 = blockedStock
				Parameter("blockedStock_EWMS4_02") = blockedStock_EWMS4_02
				ElseIf materialNumber = "EWMS4-01" Then
					blockedStock_EWMS4_01 = blockedStock
					Parameter("blockedStock_EWMS4_01") = blockedStock_EWMS4_01
			End If
			Else
				Reporter.ReportEvent micFail, "Stock is not available", "Blocked Stock of material "&materialNumber& " is: "&blockedStock
				ExitAction
		End If
	Next
End With





'Function - Press Button {ENTER} / {TAB}
Public Sub pressButton(buttonName)
	Set WinShell = CreateObject("WScript.Shell")
	MyVar = Ucase (buttonName)
   	Select Case MyVar
      		Case "ENTER"
      			WinShell.SendKeys "{ENTER}"   	
		Case "TAB"
      			WinShell.SendKeys "{TAB}"        			
       	Case Else
   	End Select
   	Set WinShell = Nothing
End Sub
