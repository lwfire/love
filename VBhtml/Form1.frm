VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vjms"
   ClientHeight    =   4590
   ClientLeft      =   -30
   ClientTop       =   315
   ClientWidth     =   6165
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6165
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   612
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   612
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   1560
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
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Form1.frx":08CA
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ��ȫ����a As SECURITY_ATTRIBUTES
Dim a����ܵ� As Long
Dim a����ܵ� As Long
Dim ��ȫ����b As SECURITY_ATTRIBUTES
Dim b����ܵ� As Long
Dim b����ܵ� As Long
Dim ������Ϣ As t������Ϣ
Dim ������Ϣ As PROCESS_INFORMATION
Dim �߳̾��tvb As Long
Dim ��� As Long
Dim �ܵ�a As Long
Dim �ܵ�b As Long
Dim �������� As Long
Dim ���� As Boolean
Dim ip As String
Dim id As String
Dim bbb As String
Dim ���ӵ�ַ As String
Public Sub Command1_Click()
    Text1 = ""
    
    
    
    If InStr(Text2.Text, "vjms:") > 0 Then
        ���ӵ�ַ = Text2.Text
        tvbus = App.Path & "\codecs\mmmj.exe " + "id=vjms//url== "
        ִ�� (tvbus)
    End If
End Sub

Private Sub Command2_Click()
    ���� = False
    �رվ�� (a����ܵ�)
    �رվ�� (b����ܵ�)
    �رվ�� (a����ܵ�)
    �رվ�� (b����ܵ�)
    Do
        DoEvents
    Loop Until ��ֹ����("cmd.exe") = 0
    Do
        DoEvents
    Loop Until ��ֹ����("mmmj.exe") = 0
    Do
        DoEvents
    Loop Until ��ֹ����("conhost.exe") = 0
    Do
        DoEvents
    Loop Until ��ֹ����("vjocx3.dll") = 0
    Do
        DoEvents
    Loop Until ��ֹ����("VJStream.exe") = 0
    Call shell(App.Path & "\codecs\mmmj.exe " & "0", 1)
    Call ��������(�򿪽��̾��(1, False, ������Ϣ.����ID), 3389)
   
   
End Sub

Private Sub �ӳ���1()
    ��� = 0
    ��ȫ����a.���� = 12
    ��ȫ����a.Ȩ�� = 0
    ��ȫ����a.��� = -1
    ��ȫ����b.���� = 12
    ��ȫ����b.Ȩ�� = 0
    ��ȫ����b.��� = -1
    �ܵ�a = ���������ܵ�(a����ܵ�, a����ܵ�, ��ȫ����a, 0)
    �ܵ�b = ���������ܵ�(b����ܵ�, b����ܵ�, ��ȫ����b, 0)
    Call ��ȡ������Ϣ_(������Ϣ)
    ������Ϣ.dwFlags = 257
    ������Ϣ.hStdInput = a����ܵ�
    ������Ϣ.hStdOutput = b����ܵ�
    ������Ϣ.hStdError = b����ܵ�
    ������Ϣ.wShowWindow = 0
    �������� = Module1.��������(0, "cmd.exe", 0, 0, -1, 0, 0, "C:\WINDOWS\system32\", ������Ϣ, ������Ϣ)
   
   

End Sub

Sub ִ��(ByVal ���� As String)
    Dim ��Ϣ As String
    ���� = True
    Call �ӳ���1
    ��Ϣ = ���� + ���ӵ�ַ
    д�ܵ� ��Ϣ
  
End Sub

Sub д�ܵ�(ByVal ������ As String)
    Dim shell() As Byte
    shell = StrConv(������ & vbCrLf, vbFromUnicode)
    Call д�ļ�(a����ܵ�, VarPtr(shell(0)), UBound(shell) + 1, ʵ�ʳߴ�, 0)
End Sub

Sub ���ܵ�()
  Dim ret As Long, TmpBuf As String * 128, BtRead As Long, BtTotal As Long, BtLeft As Long
  Dim rtn As Long, lngbytesread As Long
  
  rtn = PeekNamedPipe(b����ܵ�, StrPtr(TmpBuf), 128, BtRead, BtTotal, BtLeft)
  If rtn = 0 Then '��ѯ��Ϣ��
    DosOutput = ERROR_QUERY_INFO_SIZE
    Exit Sub
  End If
  
  If BtTotal = 0 Then
    Exit Sub
  End If
  
    Dim ʵ�ʳߴ� As Long
    Dim ����() As Byte
    Dim �����ı� As String
    
    ReDim ����(260) As Byte
    If (���ļ�(b����ܵ�, VarPtr(����(0)), 260, ʵ�ʳߴ�, 0&) <> 0) Then
        If ʵ�ʳߴ� <> 0 Then
        ReDim Preserve ����(ʵ�ʳߴ� - 1)
        �����ı� = StrConv(����, vbUnicode)
       
        With Text1
        .SelStart = Len(.Text)
        .SelText = �����ı�
        .SelLength = 0
        End With

        End If
    End If

End Sub

Private Sub Form_Load()
Text3.Visible = False
Text2.Visible = False



End Sub

Private Sub Form_Unload(Cancel As Integer)
    ���� = False
    �رվ�� (a����ܵ�)
    �رվ�� (b����ܵ�)
    �رվ�� (a����ܵ�)
    �رվ�� (b����ܵ�)
    Do
        DoEvents
    Loop Until ��ֹ����("cmd.exe") = 0
    Do
        DoEvents
    Loop Until ��ֹ����("mmmj.exe") = 0
    Do
        DoEvents
    Loop Until ��ֹ����("conhost.exe") = 0
    Do
        DoEvents
    Loop Until ��ֹ����("vjocx3.dll") = 0
    Do
        DoEvents
    Loop Until ��ֹ����("VJStream.exe") = 0
    Call shell(App.Path & "\codecs\mmmj.exe " & "0", 1)
    Call ��������(�򿪽��̾��(1, False, ������Ϣ.����ID), 3389)
   
    
   
End Sub

Private Sub Timer1_Timer()
    
    Call ���ܵ�
    
     Text3.Text = Ѱ���ı�_ȡ�ı��м�(Text1.Text, "debug: `", "' gives")
    
    HTMLx.Caption = Ѱ���ı�_ȡ�ı��м�(Text3.Text, "http://127.0.0.1:", "/1.ts")
    
    
    
End Sub



    
    
    
    


