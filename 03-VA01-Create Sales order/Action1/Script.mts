quoteNumber = Parameter("quoteNumber")
With	Browser("Home")
	.Page("Home").Link("My Sales Overview").Click
	While Not .Page("Home").WebElement("openQuotationsOriginal").Exist(1)
		wait(1)
	Wend
	.Page("Home").WebElement("openQuotationsOriginal").Click
	RowsSalesQuotations = .Page("Home").SAPUITable("Sales Quotations").RowCount
	While Not .Page("Home").SAPUITable("Sales Quotations").RowCount > 1
		wait(1)
	Wend
	quoteRowNum = .Page("Home").SAPUITable("Sales Quotations").FindRowByCellContent(1,quoteNumber, 1)
	If quoteRowNum >=1 Then
		Reporter.ReportEvent micDone, "Search Quote Number in Sales Quotations", "Quote number: "&quoteNumber& "is available" 	
	Else
		Reporter.ReportEvent micDone, "Search Quote Number in Sales Quotations", "Quote number: "&quoteNumber& "is not available" 
		ExitAction
	End If

	.Page("Home").SAPUITable("Sales Quotations").SelectRow quoteRowNum
	wait(10)
	.Page("Home").SAPUIButton("Create Subsequent Order").Click
	wait(5)
	'Browser("Home").Page("Home").SAPUIToolbar("SAPUIToolbar").OpenOverflow ' Create Subsequent Order
	.Page("Home").SAPUIMenu("SAPUIMenu").Select "Standard Order (OR)" @@ script infofile_;_ZIP::ssf4.xml_;_
	.Page("Home").SAPUIButton("OK").Click @@ script infofile_;_ZIP::ssf7.xml_;_
	wait(10)
	While Not .Page("Review Availability Check").Frame("frameReviewAvailabilityCheck").SAPUIButton("Apply").Exist(1)
		wait(1)
	Wend

	wait(5)
	.Page("Review Availability Check").Frame("frameReviewAvailabilityCheck").SAPUIButton("Apply").Click
	wait(2)
	.Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPEdit("Cust. Reference").Set "4599009999"
	currentDate =  Month(Now) &"/"&Day(Now)&"/"&Year(Now)
	custReferenceDate = currentDate
	custReferenceDate = DateAdd("d",-1, currentDate)
	.Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPEdit("Cust. Ref. Date").Set custReferenceDate
	reqDeliveryDate = DateAdd("m",1, custReferenceDate)
	.Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPEdit("Requested Delivery Date").Set reqDeliveryDate
	wait(5)
	.Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPButton("More").Click @@ script infofile_;_ZIP::ssf14.xml_;_
	.Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPDropDownMenu("Sales A").Select "Goto;Item;Shipping" @@ script infofile_;_ZIP::ssf15.xml_;_
	wait(2)
	.Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPEdit("Delivery Prior.").Set "01"
	.Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPEdit("Stor. Loc.").Set "171S"
	'	AIUtil.SetContext Browser("creationtime:=0")
	'	AIUtil("button", "Save").Click

	.Page("Review Availability Check").SAPFrame("Create Standard Order:").SAPButton("Save").Click @@ script infofile_;_ZIP::ssf21.xml_;_
	While Not .Page("Review Availability Check").Frame("frameReviewAvailabilityCheck").SAPUIButton("Apply").Exist(1)
		wait(1)
	Wend

	wait(2)
	.Page("Review Availability Check").Frame("frameReviewAvailabilityCheck").SAPUIButton("Apply").Click
	If .Page("Review Availability Check").SAPFrame("Create Standard Order:").WebElement("msgStandardOrder").Exist(2) Then
		msgOrderCreation = .Page("Review Availability Check").SAPFrame("Create Standard Order:").WebElement("msgStandardOrder").GetROProperty("innertext")
		orderNumber = Trim(Split(Split(msgOrderCreation,"Order")(1),"has")(0))
		DataTable.Value("orderNumber",dtLocalSheet) = orderNumber
		Parameter("orderNumber") = orderNumber
		Reporter.ReportEvent micPass, "Save Order", "New Order is created with order number : "&orderNumber
	Else
		Reporter.ReportEvent micPass, "Save Order", "New Order is not created"
	End  If

	wait(5)
	.Page("Home").SAPFrame("Create Standard Order:").SAPButton("Exit").Click
	.Page("Home").Image("Company Logo").Click
	.Refresh
End With

