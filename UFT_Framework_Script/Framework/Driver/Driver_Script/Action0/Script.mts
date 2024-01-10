'*************************************************************************************************************************
'Script: Web_Tour_Booking_Reservatipn
'Author: Faris Howlader
'Date: 05/23/2020
'Parameters: jojo and bean from a user of Joseph Marshall
'******************************************************************************
strUserNamePath = "C:\Users\mirza\Documents\Regression_Framework\Environment Variable\Web_Tour_InFo.xml"

Environment.LoadFromFile(strUserNamePath)
SystemUtil.CloseProcessByName "iexplore.exe"
SystemUtil.Run Environment("Web_Tour")

nameoftravel = Environment("Travel")
username = Environment("username")
password = Environment("password")
Passengers = Environment("No_of_Passengers")
DepartFlight = Environment("Flight_Departing")

print "---------------------------------------------------------------------------------------------------------------------------"
print " username " & username
print " password " & password
print " No_of_Passengers " & Passengers
print " Flight_Departing " & DepartFlight

'Call fEnterEdit(wBrowser,wPage,"username",username)

'Call fEnterEdit(wBrowser,wPage,"password",password)
'Wait 2
'Call fClickImage(wBrowser,wPage,"Login")
'Read: User and System process
'User Environment
OS = Environment.Value("OS")
OSVersion = Environment("OSVersion")
Username = Environment("UserName")
TestName = Environment("TestName")
ActionName = Environment("ActionName")
TestName = Environment("TestName")

print OS
print OSVersion
print Username
Print TestName
Print ActionName
Print TestName
