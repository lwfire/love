VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add a Link"
   ClientHeight    =   3315
   ClientLeft      =   1740
   ClientTop       =   2670
   ClientWidth     =   5940
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3315
   ScaleWidth      =   5940
   Begin VB.Frame Frame1 
      Caption         =   "Examples"
      Height          =   1935
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   4935
      Begin VB.Label Label9 
         Caption         =   "Mail Link                            mailto:userName@host"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Label8 
         Caption         =   "Ftp Link                             ftp://server/dir/file.txt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label7 
         Caption         =   "Anchor in same file            #anchoreName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Label6 
         Caption         =   "Link to Anchor                   http://.../file.html#anchoreName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label5 
         Caption         =   "External Link                      http://server/dir/file.html"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Local link                           file.html"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.CommandButton cancel 
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton ok 
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox link 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox address 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   720
      TabIndex        =   6
      Top             =   720
      Width           =   15
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Link text:"
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
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "URL:"
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
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
'Download by http://www.codefans.net
Sub cancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call sCenterForm(Me)
End Sub

Sub ok_Click()
On Error GoTo colorerror

HTMLx.Text1.SelText = HTMLx.Text1.SelText + "<a href=" & Chr(34) & address.Text & Chr(34) & ">" + link.Text + "</a>"
Form2.address.Text = ""
Form2.link.Text = ""
Unload Me
Exit Sub
colorerror:
End Sub

