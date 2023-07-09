VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "iptv"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14325
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   14325
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.ListBox List2 
      Height          =   3120
      ItemData        =   "Form1.frx":08CA
      Left            =   11520
      List            =   "Form1.frx":08CC
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7200
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   10800
      ExtentX         =   19050
      ExtentY         =   12700
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ListBox List1 
      Height          =   4020
      ItemData        =   "Form1.frx":08CE
      Left            =   11520
      List            =   "Form1.frx":08D0
      TabIndex        =   1
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Menu s1 
      Caption         =   "File(F)"
      Begin VB.Menu s2 
         Caption         =   "Login"
      End
      Begin VB.Menu s7 
         Caption         =   "Enter"
      End
   End
   Begin VB.Menu s3 
      Caption         =   "Play(P)"
      Begin VB.Menu s4 
         Caption         =   "H.265 Player"
      End
   End
   Begin VB.Menu s5 
      Caption         =   "Help(H)"
      Begin VB.Menu s6 
         Caption         =   "VJStream Close"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub List2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Form1.PopupMenu s3

End Sub


Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Form1.PopupMenu s3

End Sub




Private Sub Form_Load()
Form1.WebBrowser1.Visible = False

Form1.List1.Visible = False
Form1.List2.Visible = False


End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell "cmd.exe /c taskkill /im VJStream.exe /f", vbHide
End Sub

Private Sub s2_Click()

Form1.List1.Visible = True

Form1.List2.Visible = True




Dim strl1_string

Open App.Path & "\list.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, strl1
Form1.List1.AddItem strl1
Loop
Close #1

Dim strl2_string

Open App.Path & "\ip.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, strl2
Form1.List2.AddItem strl2
Loop
Close #1
End Sub

Private Sub s4_Click()
Form1.WebBrowser1.Visible = True

Form1.WebBrowser1.Navigate App.Path & "\pp.htm?id=vjms://" & List2.Text & ":8500:3502/live/cid=" & List1.Text

Form1.WebBrowser1.Silent = True
End Sub

Private Sub s6_Click()
Shell "cmd.exe /c taskkill /im VJStream.exe /f", vbHide

End Sub

Private Sub s7_Click()
List1.Visible = False
List1.Clear

List2.Visible = False
List2.Clear
End Sub
