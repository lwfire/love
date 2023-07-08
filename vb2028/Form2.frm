VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form2 
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11520
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7290
   ScaleWidth      =   11520
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3255
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   5535
      ExtentX         =   9763
      ExtentY         =   5741
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
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form2.WebBrowser1.Visible = False



End Sub
Private Sub Form_Activate()
With WebBrowser1


.Top = 0
.Left = 0
.Width = Me.ScaleWidth
.Height = Me.ScaleHeight
End With





End Sub







Private Sub Form_Resize()
With WebBrowser1

.Top = 0
.Left = 0
.Width = Me.ScaleWidth
.Height = Me.ScaleHeight
End With





End Sub


Private Sub Form_Unload(Cancel As Integer)

Shell "cmd.exe /c taskkill /im VJStream.exe /f", vbHide



End Sub
