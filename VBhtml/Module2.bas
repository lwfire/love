Attribute VB_Name = "Module1"
Public Type SECURITY_ATTRIBUTES  '(createprocess)
    长度 As Long
    权限 As Long
    句柄 As Long
End Type
Public Declare Function 创建匿名管道 Lib "kernel32" Alias "CreatePipe" (输出管道 As Long, 输入管道 As Long, _
管道属性 As SECURITY_ATTRIBUTES, ByVal 尺寸 As Long) As Long
Public Type t启动信息
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Public Type PROCESS_INFORMATION  '(creteprocess)
    进程号 As Long
    线程号 As Long
    进程ID As Long
    线程ID As Long
End Type


Public Declare Sub 获取启动信息_ Lib "kernel32" Alias "GetStartupInfoA" (启动信息 As t启动信息)
Declare Function 创建进程 Lib "kernel32" Alias "CreateProcessA" (ByVal 程序名称 As Long, _
ByVal 命令行 As String, ByVal 进程属性 As Long, ByVal 线程属性 As Long, ByVal 参数一 As Long, _
ByVal 参数二 As Long, ByVal 参数三 As Long, ByVal 运行目录 As String, 启动信息 As t启动信息, 进程信息 As PROCESS_INFORMATION) As Long
Declare Function 读文件 Lib "kernel32" Alias "ReadFile" (ByVal 文件号 As Long, ByVal 缓存 As Long, _
ByVal 读取尺寸 As Long, _
实际尺寸 As Long, ByVal 参数 As Long) As Long
Declare Function 写文件 Lib "kernel32" Alias "WriteFile" (ByVal 文件号 As Long, ByVal 写入内容 As Long, _
    ByVal 写入尺寸 As Long, 实际尺寸 As Long, ByVal 参数 As Long) As Long
Declare Function 关闭句柄 Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long

Declare Function 结束进程 Lib "kernel32" Alias "TerminateProcess" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Function 打开进程句柄 Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Public Declare Function PeekNamedPipe Lib "kernel32" _
                          (ByVal hNamedPipe As Long, _
                           ByVal lpBuffer As Long, _
                           ByVal nBufferSize As Long, _
                           ByRef lpBytesRead As Long, _
                           ByRef lpTotalBytesAvail As Long, _
                           ByRef lpBytesLeftThisMessage As Long _
                           ) As Long

Function 寻找文本_取文本中间(ByVal str1 As String, ByVal str2 As String, ByVal str3 As String)
    Dim p开始 As Long
    Dim p结束 As Long
    p开始 = InStr(str1, str2)
    If p开始 = 0 Then Exit Function
    p结束 = InStr(p开始 + Len(str2), str1, str3)
    If p开始 > 0 And p结束 > 0 Then
        If p结束 > p开始 Then
           寻找文本_取文本中间 = Mid(str1, p开始 + Len(str2), p结束 - p开始 - Len(str2))
        End If
    End If
End Function
