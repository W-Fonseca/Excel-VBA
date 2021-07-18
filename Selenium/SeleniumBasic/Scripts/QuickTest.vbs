
'
' Basic test to check that the installation succeed.
'

Class Script
	Dim driver
    
	Sub Class_Initialize
        Set driver = CreateObject("Selenium.FirefoxDriver")
        driver.Get "https://www.google.co.uk"
        driver.FindElementByName("q").SendKeys "Eiffel tower" & vbLf
        WScript.Echo "Title=" & driver.Title & vbLF & "Click OK to terminate"
	End Sub
	
	Sub Class_Terminate
		driver.Quit
	End Sub
End Class

Set s = New Script
