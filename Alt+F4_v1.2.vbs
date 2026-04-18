' =============================================
' 防拖课：定时强制关闭所有非系统程序
' =============================================
Option Explicit

' 1. 请求管理员权限
If Not WScript.Arguments.Named.Exists("ELEVATED") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /ELEVATED", "", "runas", 1
    WScript.Quit
End If

' 2. 输入分钟数
Dim sInput, iMins
Do
    sInput = InputBox("请输入分钟后关闭所有程序（正整数）：", "定时关闭", "5")
    If IsEmpty(sInput) Then WScript.Quit
    If IsNumeric(sInput) Then
        iMins = CLng(sInput)
        If iMins > 0 Then Exit Do
    End If
    MsgBox "分钟数必须是大于0的整数！", vbExclamation + vbOKOnly, "输入错误"
Loop

' 3. 询问是否输错时间
If MsgBox("你输错了吗？要结束本程序吗？" & vbCrLf & vbCrLf & _
          "选「是」则退出定时关闭，选「否」则继续。", _
          vbYesNo + vbQuestion, "确认") = vbYes Then
    WScript.Quit
End If

' 4. 提示如何取消定时
MsgBox "定时关闭已启动，倒计时中..." & vbCrLf & vbCrLf & _
       "如需取消，请按 Ctrl+Shift+Esc 打开任务管理器，" & vbCrLf & _
       "在「进程」中找到 wscript.exe，然后右键结束任务。", _
       vbInformation + vbOKOnly, "提示"

' 5. 倒计时（完全无窗口，任务栏无图标）
WScript.Sleep iMins * 60 * 1000

' 6. 强制结束非系统进程，并重启 explorer.exe 以刷新桌面和任务栏
Dim oShell, sCmd
Set oShell = CreateObject("WScript.Shell")

' 白名单：只保留系统核心进程（不包括 explorer，因为后面要重启它）
sCmd = "taskkill /F /FI ""USERNAME eq %USERNAME%"" /FI ""IMAGENAME ne taskmgr.exe"" /FI ""IMAGENAME ne wscript.exe"" /FI ""IMAGENAME ne csrss.exe"" /FI ""IMAGENAME ne winlogon.exe"" /FI ""IMAGENAME ne services.exe"" /FI ""IMAGENAME ne svchost.exe"" /FI ""IMAGENAME ne lsass.exe"""
oShell.Run sCmd, 0, True

' 重启 explorer.exe 以恢复桌面和任务栏图标
oShell.Run "taskkill /F /IM explorer.exe", 0, True
oShell.Run "explorer.exe", 0, False

Set oShell = Nothing
WScript.Quit