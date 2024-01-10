'DataTable.AddSheet "Test Data"
'DataTable.AddSheet "Test Case"

'Import data from an external file(Test Data)......
'DataTable.ImportSheet "C:\Users\mirza\Documents\Booking_Framework\Test Data\TestDataFile.xlsx","AppList","Test Data"
'DataTable.ImportSheet "C:\Users\mirza\Documents\Booking_Framework\Test Case\Test Case.xlsx","Test Execution","Test Case"

'Dim nRow, iApplication, iExecute, iBrowser
'nRow = DataTable.GetSheet("Test Data").GetRowCount

'For i  = 1 To nRow Step 1 
' DataTable.SetCurrentRow(i)
' iApplication = DataTable.Value("Application","Test Data")
' iExecute = DataTable.Value("Execute","Test Data")
'iBrowser = DataTable.Value("Browser","Test Data")  
'
'If iExecute = "YES" Then
'  Select Case iApplication
'  	Case "WebTour"
'  	Print "Now Running Application: "&iApplication
'  	RunAction "WebTour", oneIteration
'  	 
'  Select Case "Facebook"
'  	print "Now Running Application: "&iApplication
'  	RunAction "Facebook", oneIteration
'  	 
' Select Case "Chase App"
'  	print "Now Running Application: "&iApplication
'  	 RunAction "Chase App" , oneIteration
' 
' Select Case "OrangeHRM"
'     print "Now Running Appliation: "&iApplication
'    RunAction "OrangeHRM" , oneIteration
'       
'   Case Else
'Print "No Application Selected"
'    Reporter.ReportEvent micPass &"Landing Page","Found"
'   DataTable.Value("Application","Test Data") = "Pass"
'    Else 
'   Reporter.ReportEvent micFail &"Not Landing Page","Not Found"
' DataTable.Value("Application","Test Data") = "Fail"
' ExitTest 
' 
' End Select 

'
'
'Next
'End If