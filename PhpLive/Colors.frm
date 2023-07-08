VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document Colors"
   ClientHeight    =   4755
   ClientLeft      =   1005
   ClientTop       =   1350
   ClientWidth     =   8025
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
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4755
   ScaleWidth      =   8025
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   360
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2415
      Left            =   3720
      ScaleHeight     =   2355
      ScaleWidth      =   4155
      TabIndex        =   18
      Top             =   2280
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   17
      Top             =   240
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   3720
      TabIndex        =   13
      Top             =   1200
      Width           =   4095
      Begin VB.CommandButton cmdPicture 
         Caption         =   "Browse"
         Height          =   375
         Left            =   3120
         TabIndex        =   19
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Background Picture:"
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
         TabIndex        =   14
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   1200
      TabIndex        =   10
      Text            =   "Colors:"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   5640
      TabIndex        =   8
      Text            =   "Colors:"
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   5640
      TabIndex        =   7
      Text            =   "Colors:"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cancel 
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton ok 
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Text            =   "Colors:"
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Text            =   "Colors:"
      Top             =   1200
      Width           =   2295
   End
  'Download by http://www.codefans.net
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   960
      TabIndex        =   12
      Top             =   1680
      Width           =   2655
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Background Color:"
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
         TabIndex        =   11
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
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
      TabIndex        =   16
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Visited Link Color:"
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
      Left            =   3840
      TabIndex        =   9
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Active Link Color:"
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
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Link Color:"
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
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Text Color: "
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
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub cancel_Click()
Unload Me
End Sub
Sub Command1_Click()
End Sub

Private Sub cmdPicture_Click()
Dialog.DialogTitle = "Browse for Picture..." 'set the dialog title
Dialog.Filter = "BMP|*.bmp*|JPG|*.jpg*|GIF|*.gif"
Dialog.ShowOpen 'show the dialog box
Text1.Text = Dialog.Filename 'set the target text box to the file chosen

Picture1.Picture = LoadPicture(Dialog.Filename)
End Sub

Sub Form_Load()
Call sCenterForm(Me)
Combo1.AddItem "Red"
Combo1.AddItem "Green"
Combo1.AddItem "Blue"
Combo1.AddItem "Yellow"
Combo1.AddItem "Black"
Combo1.AddItem "Magenta"
Combo1.AddItem "Brown"
Combo1.AddItem "Cyan"
Combo1.AddItem "White"
Combo1.AddItem "Blue Violet"
Combo1.AddItem "Dark Brown"
Combo1.AddItem "Dark Green"
Combo1.AddItem "Dark Purple"
Combo1.AddItem "Orange"
Combo1.AddItem "Tan"
Combo1.AddItem "Gold"
Combo2.AddItem "Red"
Combo2.AddItem "Green"
Combo2.AddItem "Blue"
Combo2.AddItem "Yellow"
Combo2.AddItem "Black"
Combo2.AddItem "Magenta"
Combo2.AddItem "Brown"
Combo2.AddItem "Cyan"
Combo2.AddItem "White"
Combo2.AddItem "Blue Violet"
Combo2.AddItem "Dark Brown"
Combo2.AddItem "Dark Green"
Combo2.AddItem "Dark Purple"
Combo2.AddItem "Orange"
Combo2.AddItem "Tan"
Combo2.AddItem "Gold"
Combo3.AddItem "Red"
Combo3.AddItem "Green"
Combo3.AddItem "Blue"
Combo3.AddItem "Yellow"
Combo3.AddItem "Black"
Combo3.AddItem "Magenta"
Combo3.AddItem "Brown"
Combo3.AddItem "Cyan"
Combo3.AddItem "White"
Combo3.AddItem "Blue Violet"
Combo3.AddItem "Dark Brown"
Combo3.AddItem "Dark Green"
Combo3.AddItem "Dark Purple"
Combo3.AddItem "Orange"
Combo3.AddItem "Tan"
Combo3.AddItem "Gold"
Combo4.AddItem "Red"
Combo4.AddItem "Green"
Combo4.AddItem "Blue"
Combo4.AddItem "Yellow"
Combo4.AddItem "Black"
Combo4.AddItem "Magenta"
Combo4.AddItem "Brown"
Combo4.AddItem "Cyan"
Combo4.AddItem "White"
Combo4.AddItem "Blue Violet"
Combo4.AddItem "Dark Brown"
Combo4.AddItem "Dark Green"
Combo4.AddItem "Dark Purple"
Combo4.AddItem "Orange"
Combo4.AddItem "Tan"
Combo4.AddItem "Gold"
Combo5.AddItem "Red"
Combo5.AddItem "Green"
Combo5.AddItem "Blue"
Combo5.AddItem "Yellow"
Combo5.AddItem "Black"
Combo5.AddItem "Magenta"
Combo5.AddItem "Brown"
Combo5.AddItem "Cyan"
Combo5.AddItem "White"
Combo5.AddItem "Blue Violet"
Combo5.AddItem "Dark Brown"
Combo5.AddItem "Dark Green"
Combo5.AddItem "Dark Purple"
Combo5.AddItem "Orange"
Combo5.AddItem "Tan"
Combo5.AddItem "Gold"
End Sub

Sub ok_Click()
On Error GoTo colorerror
Dim F As Form
Set F = HTMLx
Dim l0038 As Variant
Dim l003C As Variant
Dim l0040 As Variant
Dim l0044 As Variant
Dim l0048 As Variant
Dim l0052 As Variant
If Combo1.Text = "Red" Then l0038 = "#FF0000"
If Combo1.Text = "Green" Then l0038 = "#00FF00"
If Combo1.Text = "Blue" Then l0038 = "#0000FF"
If Combo1.Text = "Magenta" Then l0038 = "#FF00FF"
If Combo1.Text = "Cyan" Then l0038 = "#00FFFF"
If Combo1.Text = "Yellow" Then l0038 = "#FFFF00"
If Combo1.Text = "Black" Then l0038 = "#000000"
If Combo1.Text = "Blue Violet" Then l0038 = "#9F5F9F"
If Combo1.Text = "Brown" Then l0038 = "#A62A2A"
If Combo1.Text = "Dark Brown" Then l0038 = "#5C4033"
If Combo1.Text = "Dark Green" Then l0038 = "#2F4F2F"
If Combo1.Text = "Dark Purple" Then l0038 = "#871F78"
If Combo1.Text = "Gold" Then l0038 = "#CD7F32"
If Combo1.Text = "Tan" Then l0038 = "#DB9370"
If Combo1.Text = "Orange" Then l0038 = "#FF7F00"
If Combo1.Text = "Red" Then l0038 = "#FF0000"
If Combo1.Text = "White" Then l0038 = "#FFFFFF"
If Combo2.Text = "Red" Then l003C = "#FF0000"
If Combo2.Text = "Green" Then l003C = "#00FF00"
If Combo2.Text = "Blue" Then l003C = "#0000FF"
If Combo2.Text = "Magenta" Then l003C = "#FF00FF"
If Combo2.Text = "Cyan" Then l003C = "#00FFFF"
If Combo2.Text = "Yellow" Then l003C = "#FFFF00"
If Combo2.Text = "Black" Then l003C = "#000000"
If Combo2.Text = "Blue Violet" Then l003C = "#9F5F9F"
If Combo2.Text = "Brown" Then l003C = "#A62A2A"
If Combo2.Text = "Dark Brown" Then l003C = "#5C4033"
If Combo2.Text = "Dark Green" Then l003C = "#2F4F2F"
If Combo2.Text = "Dark Purple" Then l003C = "#872F78"
If Combo2.Text = "Gold" Then l003C = "#CD7F32"
If Combo2.Text = "Tan" Then l003C = "#DB9370"
If Combo2.Text = "Orange" Then l003C = "#FF7F00"
If Combo2.Text = "Red" Then l003C = "#FF0000"
If Combo2.Text = "White" Then l003C = "#FFFFFF"
If Combo3.Text = "Red" Then l0040 = "#FF0000"
If Combo3.Text = "Green" Then l0040 = "#00FF00"
If Combo3.Text = "Blue" Then l0040 = "#0000FF"
If Combo3.Text = "Magenta" Then l0040 = "#FF00FF"
If Combo3.Text = "Cyan" Then l0040 = "#00FFFF"
If Combo3.Text = "Yellow" Then l0040 = "#FFFF00"
If Combo3.Text = "Black" Then l0040 = "#000000"
If Combo3.Text = "Blue Violet" Then l0040 = "#9F5F9F"
If Combo3.Text = "Brown" Then l0040 = "#A62A2A"
If Combo3.Text = "Dark Brown" Then l0040 = "#5C4033"
If Combo3.Text = "Dark Green" Then l0040 = "#2F4F2F"
If Combo3.Text = "Dark Purple" Then l0040 = "#873F78"
If Combo3.Text = "Gold" Then l0040 = "#CD7F32"
If Combo3.Text = "Tan" Then l0040 = "#DB9370"
If Combo3.Text = "Orange" Then l0040 = "#FF7F00"
If Combo3.Text = "Red" Then l0040 = "#FF0000"
If Combo3.Text = "White" Then l0040 = "#FFFFFF"
If Combo4.Text = "Red" Then l0044 = "#FF0000"
If Combo4.Text = "Green" Then l0044 = "#00FF00"
If Combo4.Text = "Blue" Then l0044 = "#0000FF"
If Combo4.Text = "Magenta" Then l0044 = "#FF00FF"
If Combo4.Text = "Cyan" Then l0044 = "#00FFFF"
If Combo4.Text = "Yellow" Then l0044 = "#FFFF00"
If Combo4.Text = "Black" Then l0044 = "#000000"
If Combo4.Text = "Blue Violet" Then l0044 = "#9F5F9F"
If Combo4.Text = "Brown" Then l0044 = "#A62A2A"
If Combo4.Text = "Dark Brown" Then l0044 = "#5C4033"
If Combo4.Text = "Dark Green" Then l0044 = "#2F4F2F"
If Combo4.Text = "Dark Purple" Then l0044 = "#874F78"
If Combo4.Text = "Gold" Then l0044 = "#CD7F32"
If Combo4.Text = "Tan" Then l0044 = "#DB9370"
If Combo4.Text = "Orange" Then l0044 = "#FF7F00"
If Combo4.Text = "Red" Then l0044 = "#FF0000"
If Combo4.Text = "White" Then l0044 = "#FFFFFF"
If Combo5.Text = "Red" Then l0048 = "#FF0000"
If Combo5.Text = "Green" Then l0048 = "#00FF00"
If Combo5.Text = "Blue" Then l0048 = "#0000FF"
If Combo5.Text = "Magenta" Then l0048 = "#FF00FF"
If Combo5.Text = "Cyan" Then l0048 = "#00FFFF"
If Combo5.Text = "Yellow" Then l0048 = "#FFFF00"
If Combo5.Text = "Black" Then l0048 = "#000000"
If Combo5.Text = "Blue Violet" Then l0048 = "#9F5F9F"
If Combo5.Text = "Brown" Then l0048 = "#A62A2A"
If Combo5.Text = "Dark Brown" Then l0048 = "#5C4033"
If Combo5.Text = "Dark Green" Then l0048 = "#2F4F2F"
If Combo5.Text = "Dark Purple" Then l0048 = "#875F78"
If Combo5.Text = "Gold" Then l0048 = "#CD7F32"
If Combo5.Text = "Tan" Then l0048 = "#DB9370"
If Combo5.Text = "Orange" Then l0048 = "#FF7F00"
If Combo5.Text = "Red" Then l0048 = "#FF0000"
If Combo5.Text = "White" Then l0048 = "#FFFFFF"
l0052 = Text1.Text
HTMLx.Text1.SelText = "<HTML>" + Chr(13) + Chr(10) + "<HEAD>" + Chr(13) + Chr(10) + "<TITLE>" + Text2.Text + "</TITLE>" + Chr(13) + Chr(10) + "</HEAD>" + Chr(13) + Chr(10) + "<BODY BGCOLOR=" & Chr(34) & l0048 + Chr(34) & " " + "BACKGROUND=" + Chr(34) & l0052 + Chr(34) & " " + "TEXT=" & Chr(34) & l0038 + Chr(34) & " " + "LINK=" & Chr(34) & l003C + Chr(34) & " " + "ALINK=" & Chr(34) + l0040 + Chr(34) & " " + "VLINK=" + Chr(34) & l0044 & Chr(34) & ">" + " " + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "<P>  Type your Text here  </P>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</BODY>" + Chr(13) + Chr(10) + "</HTML>" + HTMLx.Text1.SelText
Unload Me
colorerror:
End Sub

Sub sub0876()
Frame2.Visible = False
Frame1.Visible = True
Combo5.Visible = True
End Sub

Sub sub0887()
Frame2.Visible = True
Frame1.Visible = False
Combo5.Visible = False
End Sub

Sub sub0898()
Frame1.Visible = True
Frame2.Visible = False
Combo5.Visible = True
End Sub

