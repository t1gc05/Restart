Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colOS = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOS in colOS
  msg = "Are you sure you want to restart your computer?"
  title = "Restart Computer"
  style = vbYesNo + vbQuestion
  response = MsgBox(msg, style, title)
  If response = vbYes Then
    psCmd = "powershell.exe -WindowStyle Hidden -Command ""Restart-Computer"""
    Set obiShell = CreateObject("WScript.Shell")
    obiShell.Run psCmd, 0, True
  End If
Next