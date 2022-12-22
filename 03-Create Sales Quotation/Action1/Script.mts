blockedStock_EWMS4_01 = Parameter("blockedStock_EWMS4_01")
blockedStock_EWMS4_02 = Parameter("blockedStock_EWMS4_02")
If blockedStock_EWMS4_01 > 1 and blockedStock_EWMS4_02 > 1 Then
	Reporter.ReportEvent micPass, "Stock is Available", "Stock is available to create Quote"
Else
	Reporter.ReportEvent micFail, "Stock is not available", "Stock is not available to create Quote"
	'ExitAction
End If
	
With Browser("Home")
	.Page("Home").Link("My Sales Overview").Click
	While Not .Page("Home").WebElement("quickActionsOriginal_CreateSalesQuotation").Exist(1)
		wait(1)
	Wend

	.Page("Home").WebElement("quickActionsOriginal_CreateSalesQuotation").Click
	While Not .Page("Home").SAPFrame("Create Quotations").SAPEdit("Quotation Type").Exist(5)
		wait(1)
	Wend
	.Page("Home").SAPFrame("Create Quotations").SAPEdit("Quotation Type").Set DataTable.Value("quoteType",dtLocalSheet)
	'Browser("Home").Page("Home").SAPFrame("Create Quotations").SAPEdit("Sales Organization").Set "1710"
	.Page("Home").SAPFrame("Create Quotations").SAPEdit("Distribution Channel").Set DataTable.Value("distriChannel",dtLocalSheet)
	.Page("Home").SAPFrame("Create Quotations").SAPEdit("Division").Set "00"
	'	Browser("Home").Page("Home").SAPFrame("Create Quotations").SAPButton("Continue").Click

	AIUtil.SetContext Browser("creationtime:=0")
	AIUtil("button", "Continue").Click

	For Iterator = 1 To 1
		DataTable.SetCurrentRow Iterator
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPEdit("Sold-To Party").Set DataTable.Value("soldToParty",dtLocalSheet)
		pressButton("TAB")		
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPEdit("Ship-To Party").Set DataTable.Value("shipToParty",dtLocalSheet) 
		pressButton("TAB")
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPEdit("Cust. Reference").Set DataTable.Value("custReference",dtLocalSheet)
		pressButton("TAB")		
		currentDate =  Month(Now) &"/"&Day(Now)&"/"&Year(Now)
		custReferenceDate = DateAdd("d",-1, currentDate)
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPEdit("Cust. Ref. Date").Set custReferenceDate 
		reqDeliveryDate = DateAdd("m",1, custReferenceDate)
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPEdit("Requested Delivery Date").Set reqDeliveryDate
		pressButton("TAB")
		validToDate = DateAdd("ww",6,custReferenceDate)
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPEdit("Valid To").Set validToDate
		pressButton("TAB")

		'	Enter Order Details
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPTable("All Items").SelectCell 2,2 @@ script infofile_;_ZIP::ssf1.xml_;_
		If .Page("Home").SAPFrame("Create Quotation: Overview").WebTable("Information").Exist(1) Then
			.Page("Home").SAPFrame("Create Quotation: Overview").SAPButton("Continue").Click	
		End If

		wait(2)
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPTable("All Items").SetCellData 2,2, DataTable.Value("prodItem1",dtLocalSheet) @@ script infofile_;_ZIP::ssf11.xml_;_
		wait(2)
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPTable("All Items").SetCellData 2,3, DataTable.Value("prodMaterial1",dtLocalSheet) @@ script infofile_;_ZIP::ssf15.xml_;_
		pressButton("TAB")
		wait(5)
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPTable("All Items").SetCellData 2,5, DataTable.Value("prodQty1",dtLocalSheet) @@ script infofile_;_ZIP::ssf17.xml_;_
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPTable("All Items").SetCellData 2,6, DataTable.Value("prodSU1",dtLocalSheet) @@ script infofile_;_ZIP::ssf19.xml_;_
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPTable("All Items").SetCellData 2,8, DataTable.Value("prodDesc1",dtLocalSheet) @@ script infofile_;_ZIP::ssf22.xml_;_
		wait(2)
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPTable("All Items").SetCellData 3,2, DataTable.Value("prodItem2",dtLocalSheet) @@ script infofile_;_ZIP::ssf24.xml_;_
		wait(2)
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPTable("All Items").SetCellData 3,3, DataTable.Value("prodMaterial2",dtLocalSheet) @@ script infofile_;_ZIP::ssf26.xml_;_
		pressButton("TAB")
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPTable("All Items").SetCellData 3,5, DataTable.Value("prodQty2",dtLocalSheet) @@ script infofile_;_ZIP::ssf28.xml_;_
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPTable("All Items").SetCellData 3,6, DataTable.Value("prodSU2",dtLocalSheet) @@ script infofile_;_ZIP::ssf30.xml_;_
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPTable("All Items").SetCellData 3,8, DataTable.Value("prodDesc2",dtLocalSheet) @@ script infofile_;_ZIP::ssf32.xml_;_
		.Page("Home").SAPFrame("Create Quotation: Overview").SAPButton("Save").Click @@ script infofile_;_ZIP::ssf33.xml_;_
		If .Page("Home").SAPFrame("Create Quotation: Overview").WebTable("Information").Exist(5) Then
			.Page("Home").SAPFrame("Create Quotation: Overview").SAPButton("Continue").Click	
		End If
		If .Page("Home").SAPFrame("Create Quotation: Overview").WebElement("msgQuotation").Exist(2) Then
			msgQuotation = .Page("Home").SAPFrame("Create Quotation: Overview").WebElement("msgQuotation").GetROProperty("innertext")
			quoteNumber = Trim(Split(Split(msgQuotation,"Quotation")(1),"has")(0))
			DataTable.Value("quoteNum",dtLocalSheet) = quoteNumber
			Parameter("quoteNumber") =  quoteNumber
			Reporter.ReportEvent micPass, "Save Quotation", "New Quotation is created, quote number is: "&quoteNumber
			Else
				Reporter.ReportEvent micPass, "Save Quotation", "New Quotation is not created"
		End  If
	Next

	Wait(5)

	.Page("Home").SAPFrame("Create Quotation: Overview").SAPButton("Cancel").Click
	If .Page("Home").SAPFrame("Create Quotations").SAPButton("Exit").Exist(10) Then
		.Page("Home").SAPFrame("Create Quotations").SAPButton("Exit").Click
	End If
	.Page("Home").Image("Company Logo").Click
	.Refresh
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

