Set WshShell = CreateObject("WScript.Shell")
plinkPath = "D:\\Program Files\\putty\\plink.exe"  ' 修改为实际路径
cmd = plinkPath & " -N root@111.222.123.123 -pw 123456 -D 0.0.0.0:1080"

' 确保窗口标题唯一
windowTitle = "plink.exe - " & cmd

Do While True
    Set oExec = WshShell.Exec(cmd)

    ' 等待首次连接响应（约2秒）
    WScript.Sleep 2000

    ' 尝试激活窗口并发送回车
    On Error Resume Next
    WshShell.AppActivate windowTitle
    If Err.Number = 0 Then
        WshShell.SendKeys "{ENTER}"
        WScript.Sleep 300 ' 等待窗口激活
        WshShell.SendKeys "{ENTER}"
        WScript.Sleep 300 ' 等待窗口激活
    End If
    On Error GoTo 0

    ' 监控进程状态
    Do While oExec.Status = 0
        WScript.Sleep 1000
        If InStr(oExec.StdOut.ReadAll, "disconnected") > 0 Then Exit Do
    Loop

    ' 异常等待后重连
    WScript.Sleep 5000
Loop
