VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add a picture"
   ClientHeight    =   4350
   ClientLeft      =   1755
   ClientTop       =   2400
   ClientWidth     =   6165
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4350
   ScaleWidth      =   6165
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   600
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3135
      Left            =   1560
      ScaleHeight     =   3075
      ScaleWidth      =   4395
      TabIndex        =   7
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton cmdPicture 
      Caption         =   "Browse"
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cancel 
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton ok 
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox alt 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox pic 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label imgalt 
      BackStyle       =   0  'Transparent
      Caption         =   "Image Alt:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label picture 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.codefans.net
Option Explicit

Sub cancel_Click()
Unload Me
End Sub

Private Sub cmdPicture_Click()

Dialog.DialogTitle = "Browse for Picture..." 'set the dialog title
Dialog.Filter = "BMP|*.bmp*|JPG|*.jpg*|GIF|*.gif"
Dialog.ShowOpen 'show the dialog box
pic.Text = Dialog.Filename 'set the target text box to the file chosen

Picture1.Picture = LoadPicture(Dialog.Filename)
End Sub

Private Sub Form_Load()
Call sCenterForm(Me)
End Sub

Sub ok_Click()
On Error GoTo colorerror

HTMLx.Text1.SelText = HTMLx.Text1.SelText + "<img alt=" & Chr(34) & alt.Text & Chr(34) & ", src=" & Chr(34) & pic.Text & Chr(34) & " Border=0>"
Form4.pic.Text = ""
Form4.alt.Text = ""

Unload Me
Exit Sub
colorerror:
End Sub

