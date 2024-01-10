

Dim fso, file_Notepad, file_location
Const ForWriting=2
Const ForReading=1
file_location = "C:\Users\mirza\Desktop\UFT Class-Noyon Sat&Sun\Data File.txt"
Set fso = CreateObject("Scripting.FileSystemObject")
'Set file = fso.CreateTextFile(file_location, True)

Set file_Notepad= fso.OpenTextFile(file_location, ForReading, True)

allLine=file_Notepad.ReadAll
allLine=Split(allLine,VBnewline)

ReDim myData(Ubound(allLine)-1)

For i = 0 To Ubound(allLine)-1

	myData(i)=allLine(i+1)
	
	
Next






Dim fso, file_Notepad, file_location
Const ForWriting=2
Const ForReading=1
file_location = "C:\Users\mirza\Desktop\UFT Class-Noyon Sat&Sun\MyQTPResult.txt"
Set fso = CreateObject("Scripting.FileSystemObject")
'Set file = fso.CreateTextFile(file_location, True)

'Set file_Notepad= fso.OpenTextFile(file_location, ForWriting, True)

'file_Notepad.Write("My name is Jaman")

'fso.CreateFolder("C:\Users\mirza\Desktop\UFT Class-Noyon Sat&Sun\Israfil")


If fso.FileExists(file_location) Then
	fso.DeleteFile(file_location)
End If


'Dim mydata()
'
'While Not file_Notepad.AtEndOfLine
'	allData=file_Notepad.ReadLine
'Wend
'file_Notepad.

