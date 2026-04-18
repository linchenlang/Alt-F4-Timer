' =============================================
' 防拖课：定时强制关闭所有非系统程序（统一输入框版）
' 绝对可靠，纯VBS，兼容Win7/Win11/Server
' 输入格式：时:分:秒 或 分:秒 或 秒，例如 "1:30" 或 "5" 或 "0:5:0"
' =============================================
Option Explicit

' 1. 请求管理员权限
If Not WScript.Arguments.Named.Exists("ELEVATED") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /ELEVATED", "", "runas", 1
    WScript.Quit
End If

' 2. 获取时间字符串并解析
Dim sInput, totalSeconds
Do
    sInput = InputBox("请输入定时时间（格式：时:分:秒 或 分:秒 或 秒）" & vbCrLf & _
                      "例如：1:30:0 表示1小时30分；5:30 表示5分30秒；45 表示45秒", _
                      "定时关闭", "5")
    If IsEmpty(sInput) Then WScript.Quit
    sInput = Trim(sInput)
    If sInput = "" Then
        MsgBox "请输入有效时间！", vbExclamation, "输入错误"
        sInput = ""
    Else
        totalSeconds = ParseTimeString(sInput)
        If totalSeconds > 0 Then Exit Do
        MsgBox "时间格式错误！请使用数字和冒号组合，例如：1:30 或 45", vbExclamation, "输入错误"
    End If
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

' 5. 倒计时
WScript.Sleep totalSeconds * 1000

' 6. 强制结束非系统进程并重启explorer
Dim oShell, sCmd
Set oShell = CreateObject("WScript.Shell")

' 结束所有用户进程（保留关键系统进程）
sCmd = "taskkill /F /FI ""USERNAME eq %USERNAME%"" /FI ""IMAGENAME ne taskmgr.exe"" /FI ""IMAGENAME ne wscript.exe"" /FI ""IMAGENAME ne csrss.exe"" /FI ""IMAGENAME ne winlogon.exe"" /FI ""IMAGENAME ne services.exe"" /FI ""IMAGENAME ne svchost.exe"" /FI ""IMAGENAME ne lsass.exe"" /FI ""IMAGENAME ne DownKyi.exe"""
oShell.Run sCmd, 0, True

' 重启explorer
oShell.Run "taskkill /F /IM explorer.exe", 0, True
oShell.Run "explorer.exe", 0, False

Set oShell = Nothing
WScript.Quit

' 解析时间字符串函数
Function ParseTimeString(str)
    Dim parts, h, m, s, i
    ParseTimeString = 0
    parts = Split(str, ":")
    Dim count
    count = UBound(parts) + 1
    If count = 1 Then
        ' 只有秒
        If IsNumeric(parts(0)) Then
            s = CLng(parts(0))
            If s >= 0 Then ParseTimeString = s
        End If
    ElseIf count = 2 Then
        ' 分:秒
        If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
            m = CLng(parts(0))
            s = CLng(parts(1))
            If m >= 0 And s >= 0 And s < 60 Then ParseTimeString = m * 60 + s
        End If
    ElseIf count = 3 Then
        ' 时:分:秒
        If IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
            h = CLng(parts(0))
            m = CLng(parts(1))
            s = CLng(parts(2))
            If h >= 0 And m >= 0 And m < 60 And s >= 0 And s < 60 Then ParseTimeString = h * 3600 + m * 60 + s
        End If
    End If
End Function