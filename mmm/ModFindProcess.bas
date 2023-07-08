Attribute VB_Name = "ModFindProcess"

Option Explicit

Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const PROCESS_TERMINATE = 1

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
   th32ProcessID As Long
   th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
   th32ParentProcessID As Long
   pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

Private Type MyProcess
    ExeName As String
    PID As Long
End Type

Public Function 终止进程(Optional ByVal ProName As String, Optional ByVal PID As Long = -1) As Integer
   '传入进程名或PID,结束相应进程
    Dim tPID As Long
    Dim tPHwnd As Long
    Dim ProArr() As String, PIDArr() As Long
    Dim I As Long
    
    Call ListProcess(ProArr, PIDArr)
    For I = 1 To UBound(ProArr)
        Debug.Print ProArr(I)
       If PIDArr(I) = PID Or LCase(ProArr(I)) = LCase(ProName) Then     '配对进程ID或进程名
           Exit For
       End If
    Next I
   
    If I > UBound(PIDArr) Then Exit Function
    tPID = PIDArr(I)
   
    tPHwnd = OpenProcess(PROCESS_TERMINATE, False, tPID)
    Debug.Print tPHwnd
    If tPHwnd Then
       终止进程 = TerminateProcess(tPHwnd, 0)
    End If
End Function

Public Function FindProcess(ByVal ProName As String, Optional ByRef PID As Long) As Boolean
   '传入进程名,如果进程存在,在PID里返回进程ID,函数返回True,否则返回Flase
    'ProName: 指定进程名
    'PID: 如果进程名存在,返回其PID
    '返回值: 进程名存在返回TRUE,否则返回FALSE
    Dim ProArr() As String, PIDArr() As Long
    Dim I As Long
   
    Call ListProcess(ProArr, PIDArr)
    For I = 1 To UBound(ProArr)
       If ProArr(I) = ProName Then
           PID = PIDArr(I)
           FindProcess = True
           Exit For
       End If
    Next I
End Function

Public Function ListProcess(ByRef ProExeName() As String, ByRef ProPid() As Long)
   '列出进程以及相应PID
   'ProExeName(): 进程名
    'ProPid(): 相应的PID
    Dim MyProcess As PROCESSENTRY32
    Dim mySnapshot As Long
    Dim ProData() As MyProcess
    Dim I As Long
   
    ReDim ProData(0)
   
   MyProcess.dwSize = Len(MyProcess)
    mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    ProcessFirst mySnapshot, MyProcess
   
    ReDim Preserve ProData(UBound(ProData) + 1)
   
   ProData(UBound(ProData)).ExeName = Left(MyProcess.szexeFile, InStr(MyProcess.szexeFile, Chr(0)) - 1)
   ProData(UBound(ProData)).PID = MyProcess.th32ProcessID
   
    'Debug.Print ProData(UBound(ProData)).ExeName
   
   MyProcess.szexeFile = ""
   
    While ProcessNext(mySnapshot, MyProcess)
       ReDim Preserve ProData(UBound(ProData) + 1)
       
       ProData(UBound(ProData)).ExeName = Left(MyProcess.szexeFile, InStr(MyProcess.szexeFile, Chr(0)) - 1)
       ProData(UBound(ProData)).PID = MyProcess.th32ProcessID
       
   '    Debug.Print ProData(UBound(ProData)).ExeName
       
       MyProcess.szexeFile = ""
    Wend
   
    ReDim ProExeName(UBound(ProData))
    ReDim ProPid(UBound(ProData))
   
    For I = 1 To UBound(ProData)
       With ProData(I)
           ProExeName(I) = .ExeName
           ProPid(I) = .PID
       End With
    Next I
End Function
