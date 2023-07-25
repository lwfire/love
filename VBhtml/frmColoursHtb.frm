VERSION 5.00
Begin VB.Form frmColoursHtb 
   Caption         =   "Background Colours"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3480
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   15
      Left            =   2640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   15
      Top             =   2280
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   14
      Left            =   1800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   14
      Top             =   2280
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   13
      Left            =   960
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   13
      Top             =   2280
      Width           =   735
   End
   'Download by http://www.codefans.net
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   12
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   12
      Top             =   2280
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   11
      Left            =   2640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   11
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   10
      Left            =   1800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   9
      Left            =   960
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   8
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   7
      Left            =   2640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   6
      Left            =   1800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   5
      Left            =   960
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   4
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   3
      Left            =   2640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   2
      Left            =   1800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   1
      Left            =   960
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   0
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmColoursHtb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
 Call sCenterForm(Me)
 Dim intI As Integer ' counter
    For intI = 0 To 15 '16 colors
        ' set color
        Picture1(intI).BackColor = QBColor(intI)
    Next intI
End Sub

Private Sub Picture1_Click(Index As Integer)
  'MainForm.Text1
    Select Case Index
       Case 0
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "000000"">" + HTMLx.Text1.SelText
       Case 1
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "000080"">" + HTMLx.Text1.SelText
       Case 2
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "008000"">" + HTMLx.Text1.SelText
       Case 3
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "408080"">" + HTMLx.Text1.SelText
       Case 4
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "800000"">" + HTMLx.Text1.SelText
       Case 5
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "800080"">" + HTMLx.Text1.SelText
       Case 6
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "808000"">" + HTMLx.Text1.SelText
       Case 7
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "C0C0C0"">" + HTMLx.Text1.SelText
       Case 8
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "808080"">" + HTMLx.Text1.SelText
       Case 9
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "0000FF"">" + HTMLx.Text1.SelText
       Case 10
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "00FF00"">" + HTMLx.Text1.SelText
       Case 11
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "000FFF"">" + HTMLx.Text1.SelText
       Case 12
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "FF0000"">" + HTMLx.Text1.SelText
       Case 13
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "FF00FF"">" + HTMLx.Text1.SelText
       Case 14
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "FFFF00"">" + HTMLx.Text1.SelText
       Case 15
       HTMLx.Text1.SelText = "<BODY BGCOLOR=""#" + "FFFFFF"">" + HTMLx.Text1.SelText
      
    End Select
    
    Unload Me ' unload form
End Sub
