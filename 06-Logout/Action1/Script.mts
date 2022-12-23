With Browser("Home").Page("Home")
	.Image("Navigate to Home Page").Click
	.WebElement("Profile").Click
	If .WebElement("Sign Out").Exist(10) Then
		.WebElement("Sign Out").Click
	Else
		Wait(20)
		.WebElement("Sign Out").Click
	End  If
	.WebElement("msg_Ok_Btn").Click
End With

