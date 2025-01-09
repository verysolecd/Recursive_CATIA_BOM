Sub aMsgBox(msgx)
    Dim Msgwd, Title, lg, wd
    Set Msgwd = CreateObject("wscript.shell")
    wd = "0.5s后自动关闭对话框"
    lg = 0.514
    Msgwd.popup msgx, lg, wd
    Set Msgwd = Nothing
End Sub