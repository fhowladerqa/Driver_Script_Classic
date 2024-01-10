
'' Object mapping Concept


Dim Result
Systemutil.Run "iexplore.exe","http://127.0.0.1:1080/WebTours/index.htm"

Wait 2
'Enter user id
Result=fMasterObjectMapping("WebEdit_username","WebTourFindFlight",ObjUid)
Result=fMasterObjectMapping("WebEdit_password","WebTourFindFlight",Objpw)
Result=fMasterObjectMapping("Image_Login","WebTourFindFlight",ObjLogin)


Call fPerformAction(ObjUid,"jojo")
Call fPerformAction(Objpw,"bean")
Call fPerformAction(ObjLogin,"")
