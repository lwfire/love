Attribute VB_Name = "Module1"
Public Type SECURITY_ATTRIBUTES  '(createprocess)
    ���� As Long
    Ȩ�� As Long
    ��� As Long
End Type
Public Declare Function ���������ܵ� Lib "kernel32" Alias "CreatePipe" (����ܵ� As Long, ����ܵ� As Long, _
�ܵ����� As SECURITY_ATTRIBUTES, ByVal �ߴ� As Long) As Long
Public Type t������Ϣ
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
    ���̺� As Long
    �̺߳� As Long
    ����ID As Long
    �߳�ID As Long
End Type


Public Declare Sub ��ȡ������Ϣ_ Lib "kernel32" Alias "GetStartupInfoA" (������Ϣ As t������Ϣ)
Declare Function �������� Lib "kernel32" Alias "CreateProcessA" (ByVal �������� As Long, _
ByVal ������ As String, ByVal �������� As Long, ByVal �߳����� As Long, ByVal ����һ As Long, _
ByVal ������ As Long, ByVal ������ As Long, ByVal ����Ŀ¼ As String, ������Ϣ As t������Ϣ, ������Ϣ As PROCESS_INFORMATION) As Long
Declare Function ���ļ� Lib "kernel32" Alias "ReadFile" (ByVal �ļ��� As Long, ByVal ���� As Long, _
ByVal ��ȡ�ߴ� As Long, _
ʵ�ʳߴ� As Long, ByVal ���� As Long) As Long
Declare Function д�ļ� Lib "kernel32" Alias "WriteFile" (ByVal �ļ��� As Long, ByVal д������ As Long, _
    ByVal д��ߴ� As Long, ʵ�ʳߴ� As Long, ByVal ���� As Long) As Long
Declare Function �رվ�� Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long

Declare Function �������� Lib "kernel32" Alias "TerminateProcess" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Function �򿪽��̾�� Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Public Declare Function PeekNamedPipe Lib "kernel32" _
                          (ByVal hNamedPipe As Long, _
                           ByVal lpBuffer As Long, _
                           ByVal nBufferSize As Long, _
                           ByRef lpBytesRead As Long, _
                           ByRef lpTotalBytesAvail As Long, _
                           ByRef lpBytesLeftThisMessage As Long _
                           ) As Long

Function Ѱ���ı�_ȡ�ı��м�(ByVal str1 As String, ByVal str2 As String, ByVal str3 As String)
    Dim p��ʼ As Long
    Dim p���� As Long
    p��ʼ = InStr(str1, str2)
    If p��ʼ = 0 Then Exit Function
    p���� = InStr(p��ʼ + Len(str2), str1, str3)
    If p��ʼ > 0 And p���� > 0 Then
        If p���� > p��ʼ Then
           Ѱ���ı�_ȡ�ı��м� = Mid(str1, p��ʼ + Len(str2), p���� - p��ʼ - Len(str2))
        End If
    End If
End Function
