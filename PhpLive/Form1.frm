VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vjms"
   ClientHeight    =   5040
   ClientLeft      =   -30
   ClientTop       =   315
   ClientWidth     =   6165
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6165
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   612
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2775
      Left            =   600
      TabIndex        =   4
      Top             =   480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4895
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Form1.frx":08CA
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   612
      Left            =   600
      TabIndex        =   2
      Top             =   3360
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim 安全属性a As SECURITY_ATTRIBUTES
Dim a输出管道 As Long
Dim a输入管道 As Long
Dim 安全属性b As SECURITY_ATTRIBUTES
Dim b输出管道 As Long
Dim b输入管道 As Long
Dim 启动信息 As t启动信息
Dim 进程信息 As PROCESS_INFORMATION
Dim 线程句柄tvb As Long
Dim 序号 As Long
Dim 管道a As Long
Dim 管道b As Long
Dim 创建进程 As Long
Dim 运行 As Boolean
Dim ip As String
Dim id As String
Dim bbb As String
Dim 连接地址 As String
Public Sub Command1_Click()
    Text1 = ""
    
    
    
    If InStr(Text2.Text, "vjms:") > 0 Then
        连接地址 = Text2.Text
        tvbus = App.Path & "\codecs\mmmj.exe " + "id=vjms//url== "
        执行 (tvbus)
    End If
End Sub

Private Sub Command2_Click()
    运行 = False
    关闭句柄 (a输出管道)
    关闭句柄 (b输出管道)
    关闭句柄 (a输入管道)
    关闭句柄 (b输入管道)
    Do
        DoEvents
    Loop Until 终止进程("cmd.exe") = 0
    Do
        DoEvents
    Loop Until 终止进程("mmmj.exe") = 0
    Do
        DoEvents
    Loop Until 终止进程("conhost.exe") = 0
    Do
        DoEvents
    Loop Until 终止进程("vjocx3.dll") = 0
    Do
        DoEvents
    Loop Until 终止进程("VJStream.exe") = 0
    Call shell(App.Path & "\codecs\mmmj.exe " & "0", 1)
    Call 结束进程(打开进程句柄(1, False, 进程信息.进程ID), 3389)
   
   
End Sub

Private Sub 子程序1()
    序号 = 0
    安全属性a.长度 = 12
    安全属性a.权限 = 0
    安全属性a.句柄 = -1
    安全属性b.长度 = 12
    安全属性b.权限 = 0
    安全属性b.句柄 = -1
    管道a = 创建匿名管道(a输出管道, a输入管道, 安全属性a, 0)
    管道b = 创建匿名管道(b输出管道, b输入管道, 安全属性b, 0)
    Call 获取启动信息_(启动信息)
    启动信息.dwFlags = 257
    启动信息.hStdInput = a输出管道
    启动信息.hStdOutput = b输入管道
    启动信息.hStdError = b输入管道
    启动信息.wShowWindow = 0
    创建进程 = Module1.创建进程(0, "cmd.exe", 0, 0, -1, 0, 0, "C:\WINDOWS\system32\", 启动信息, 进程信息)
   
   

End Sub

Sub 执行(ByVal 内容 As String)
    Dim 信息 As String
    运行 = True
    Call 子程序1
    信息 = 内容 + 连接地址
    写管道 信息
  
End Sub

Sub 写管道(ByVal 命令名 As String)
    Dim shell() As Byte
    shell = StrConv(命令名 & vbCrLf, vbFromUnicode)
    Call 写文件(a输入管道, VarPtr(shell(0)), UBound(shell) + 1, 实际尺寸, 0)
End Sub

Sub 读管道()
  Dim ret As Long, TmpBuf As String * 128, BtRead As Long, BtTotal As Long, BtLeft As Long
  Dim rtn As Long, lngbytesread As Long
  
  rtn = PeekNamedPipe(b输出管道, StrPtr(TmpBuf), 128, BtRead, BtTotal, BtLeft)
  If rtn = 0 Then '查询信息量
    DosOutput = ERROR_QUERY_INFO_SIZE
    Exit Sub
  End If
  
  If BtTotal = 0 Then
    Exit Sub
  End If
  
    Dim 实际尺寸 As Long
    Dim 缓存() As Byte
    Dim 加入文本 As String
    
    ReDim 缓存(260) As Byte
    If (读文件(b输出管道, VarPtr(缓存(0)), 260, 实际尺寸, 0&) <> 0) Then
        If 实际尺寸 <> 0 Then
        ReDim Preserve 缓存(实际尺寸 - 1)
        加入文本 = StrConv(缓存, vbUnicode)
       
        With Text1
        .SelStart = Len(.Text)
        .SelText = 加入文本
        .SelLength = 0
        End With

        End If
    End If

End Sub

Private Sub Form_Load()
Text3.Visible = False


End Sub

Private Sub Form_Unload(Cancel As Integer)
    运行 = False
    关闭句柄 (a输出管道)
    关闭句柄 (b输出管道)
    关闭句柄 (a输入管道)
    关闭句柄 (b输入管道)
    Do
        DoEvents
    Loop Until 终止进程("cmd.exe") = 0
    Do
        DoEvents
    Loop Until 终止进程("mmmj.exe") = 0
    Do
        DoEvents
    Loop Until 终止进程("conhost.exe") = 0
    Do
        DoEvents
    Loop Until 终止进程("vjocx3.dll") = 0
    Do
        DoEvents
    Loop Until 终止进程("VJStream.exe") = 0
    Call shell(App.Path & "\codecs\mmmj.exe " & "0", 1)
    Call 结束进程(打开进程句柄(1, False, 进程信息.进程ID), 3389)
   
    
   
End Sub

Private Sub Timer1_Timer()
    
    Call 读管道
    
     Text3.Text = 寻找文本_取文本中间(Text1.Text, "debug: `", "' gives")
    
    HTMLx.Caption = 寻找文本_取文本中间(Text3.Text, "http://127.0.0.1:", "/1.ts")
    
    
    
End Sub



    
    
    
    


