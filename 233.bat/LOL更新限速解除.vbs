Set ws = CreateObject("Wscript.Shell") 
Msgbox "本程序可以解除LOL的更新速度限制。如果你认为更新下载速度没有达到极限，本程序或许可以帮到你。",64,"欢迎使用"
If vbOK = Msgbox ("使用本程序可能导致网速不稳定。要继续吗？",68,"提问") Then
ws.run "taskkill -f -t -im TenioDL.exe",vbhide
Msgbox "现在下载速度可能会卡住不动。不过无须担心，那个数字会在数秒内开始继续变动。",64,"加速完成"
End If
End If
End If If