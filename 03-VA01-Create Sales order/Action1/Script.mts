﻿	quoteNumber = Parameter("quoteNumber")
	Browser("Home").Page("Home").Link("My Sales Overview").Click
	Browser("Home").Page("Home").WebElement("openQuotationsOriginal").Click
	quoteRowNum = Browser("Home").Page("Home").SAPUITable("Sales Quotations").FindRowByCellContent(1,quoteNumber, 1)
	Browser("Home").Page("Home").SAPUITable("Sales Quotations").SelectRow quoteRowNum
	Browser("Home").Page("Home").SAPUIButton("Create Subsequent Order").Click
	Browser("Home").Page("Home").SAPUIToolbar("SAPUIToolbar").OpenOverflow ' Create Subsequent Order
	Browser("Home").Page("Home").SAPUIMenu("SAPUIMenu").Select "Standard Order (OR)" @@ script infofile_;_ZIP::ssf4.xml_;_
	Browser("Home").Page("Home").SAPUIButton("OK").Click @@ script infofile_;_ZIP::ssf7.xml_;_
	While Not Browser("Home").Page("Review Availability Check").Frame("frameReviewAvailabilityCheck").SAPUIButton("Apply").Exist(1)
		wait(1)
	Wend
	
	wait(5)
	Browser("Home").Page("Review Availability Check").Frame("frameReviewAvailabilityCheck").SAPUIButton("Apply").Click
	Browser("Home").Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPEdit("Cust. Reference").Set "4599009999"
	currentDate =  Month(Now) &"/"&Day(Now)&"/"&Year(Now)
	custReferenceDate = currentDate
	custReferenceDate = DateAdd("d",-1, currentDate)
	Browser("Home").Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPEdit("Cust. Ref. Date").Set custReferenceDate
	reqDeliveryDate = DateAdd("m",1, custReferenceDate)
	Browser("Home").Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPEdit("Requested Delivery Date").Set reqDeliveryDate
	Browser("Home").Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPButton("More").Click @@ script infofile_;_ZIP::ssf14.xml_;_
	Browser("Home").Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPDropDownMenu("Sales A").Select "Goto;Item;Shipping" @@ script infofile_;_ZIP::ssf15.xml_;_
	wait(5)
	Browser("Home").Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPEdit("Delivery Prior.").Set "1"
	Browser("Home").Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPEdit("Stor. Loc.").Set "171S"
	Browser("Home").Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPButton("Save").Click @@ script infofile_;_ZIP::ssf21.xml_;_
	
	While Not Browser("Home").Page("Review Availability Check").Frame("frameReviewAvailabilityCheck").SAPUIButton("Apply").Exist(1)
		wait(1)
	Wend
	
	wait(2)
	Browser("Home").Page("Review Availability Check").Frame("frameReviewAvailabilityCheck").SAPUIButton("Apply").Click
	If Browser("Home").Page("Review Availability Check").SAPFrame("Create Standard Order:").WebElement("msgStandardOrder").Exist(2) Then
		Reporter.ReportEvent micPass, "Save Order", "New Order is created"
		msgOrderCreation = Browser("Home").Page("Review Availability Check").SAPFrame("Create Standard Order:").WebElement("msgStandardOrder").GetROProperty("innertext")
		orderNumber = Split(Split(msgOrderCreation,"Order")(1),"has")(0)
		DataTable.Value("orderNumber",dtLocalSheet) = Trim(orderNumber)
		Parameter("orderNumber") = orderNumber
	Else
		Reporter.ReportEvent micPass, "Save Order", "New Order is not created"
	End  If	
	
	Browser("Home").Page("Home").SAPFrame("Create Standard Order:").SAPButton("Exit").Click
	Browser("Home").Page("Home").Image("Company Logo").Click
