Set ws = CreateObject("Wscript.Shell") 
Msgbox "��������Խ��LOL�ĸ����ٶ����ơ��������Ϊ���������ٶ�û�дﵽ���ޣ������������԰ﵽ�㡣",64,"��ӭʹ��"
If vbOK = Msgbox ("ʹ�ñ�������ܵ������ٲ��ȶ���Ҫ������",68,"����") Then
ws.run "taskkill -f -t -im TenioDL.exe",vbhide
Msgbox "���������ٶȿ��ܻῨס�������������뵣�ģ��Ǹ����ֻ��������ڿ�ʼ�����䶯��",64,"�������"
End If
End If
End If If