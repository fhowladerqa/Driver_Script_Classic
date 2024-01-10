
''Excel Validation

'Excel  

'1.Excel Object
'2.Workbook
'3.Work Sheet
'4. Rows
'5.Columns
'6.Cells

Dim xlDatabaseFile,XlReportFile
xlDatabaseFile="C:\Users\mirza\Desktop\UFT Class 2019\Excel Validation\Database Data.xls"
XlReportFile="C:\Users\mirza\Desktop\UFT Class 2019\Excel Validation\Report.xls"

Set obj = createobject("Excel.Application")   'Creating an Excel Object
obj.visible=True                                    'Making an Excel Object visible

Set WBookDB = obj.Workbooks.open(xlDatabaseFile)    'Opening DB file
Set wSheetDB=WBookDB.Worksheets("Sheet1")    'Referring Sheet1 of DB excel file

Set WBookReport = obj.Workbooks.open(XlReportFile)    'Opening DB file
Set wSheetReport=WBookReport.Worksheets("Sheet1")    'Referring Sheet1 of DB excel file


nRowReport=wSheetReport.UsedRange.Rows.Count  ' Row count
nColReport=wSheetReport.UsedRange.Columns.Count  'column Count

For i = 1 To nRowReport

	For j = 1 To nColReport
		CellValue=wSheetReport.Cells(i,j)
		If CellValue<>"" Then
			If Trim(CellValue)="Patient ID" Then
				Redim PatientId(nRowReport-(i+1))
				
				For k = i To nRowReport-1
					PatientId(k-i)=wSheetReport.Cells(k+1,j)
					
					Call fgetDBDataforPatientId(wSheetDB,PatientId(k-i),mMonth,dbCellValue)
					  For m = j+1 To nColReport
					  		mMonth=wSheetReport.Cells(k,m)
					  		mCellValue=wSheetReport.Cells(k+1,m)
					  		
					  		If mCellValue=dbCellValue(m-j-1) Then
					  			
					  			print "Data Match "& "Report "&mCellValue&" With DB "&" "&dbCellValue(m-j-1)
'								XLSht_report.Cells(k+1,j).interior.colorindex=32
								wSheetReport.Cells(k+1,m).Font.Color=vbBlue
								Else
								print "Data Not Match "& "Report "&mCellValue&" With DB "&" "&dbCellValue(m-j-1)
'								XLSht_report.Cells(k+1,j).interior.colorindex=32
								wSheetReport.Cells(k+1,m).Font.Color=vbRed					  			
					  			
					  		End If
					  		
					  	
					  Next


				Next
				Exit For
				
			End If
		End If
	Next
	
Next


Function fgetDBDataforPatientId(wSheetDB,PatientId,mMonth,dbCellValue)
  
  nRow=wSheetDB.UsedRange.Rows.Count  ' Row count
  nColm=wSheetDB.UsedRange.Columns.Count  'column Count
  
  For n = 1 To nRow
  	For p = 1 To nColm
  		CellValue=wSheetDB.Cells(n,p)
  		If CellValue="Patiend_ID" Then
  			p_Id=wSheetDB.Cells(n+1,p)
  			
  			ReDim dbCellValue(nColm-(p+1))
  			If p_Id=PatientId Then
  				For a = p+1 To nColm
  					dbCellValue(a-2)=wSheetDB.Cells(n+1,a)
  					
  				Next
  				
  				
  				Exit Function
  			End If
  		End If
  		
  	Next
  	
  Next


End Function













'Msgbox obj2.Cells(2,2).Value  ‘Value from the specified cell will be read and shown
'obj1.Close                                             ‘Closing a Workbook
'obj.Quit                                                  ‘Exit from Excel Application
'Set obj1=Nothing                                 ‘Releasing Workbook object
'Set obj2 = Nothing                               ‘Releasing Worksheet object
'Set obj=Nothing   



Set WshShell = CreateObject("WScript.Shell")
WshShell.AppActivate '"Put the label of the browser" ' Activate the browser window
wait(3)
WshShell.SendKeys"" ' The caret (^) represents the CTRL key.

Call fSendKey("TAB")

Call fSendKey("Enter")
Function fSendKey(Key)
Dim mySendKeys
set mySendKeys = CreateObject("WScript.shell")
mySendKeys.SendKeys("{&Key&}")
End Function

