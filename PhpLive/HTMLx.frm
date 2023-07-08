VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form HTMLx 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   1695
   ClientTop       =   1515
   ClientWidth     =   8880
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FF0000&
   HelpContextID   =   7890
   Icon            =   "HTMLx.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   9360
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7500
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   8880
      ExtentX         =   15663
      ExtentY         =   13229
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
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "HTMLx.frx":08CA
      Top             =   0
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   8820
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7995
      Width           =   8880
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   6000
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Play "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   2880
      End
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   9990
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "HTMLx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Main Form"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const VK_CONTROL = &H11
Const KEYEVENTF_KEYUP = &H2
Const VK_ESCAPE = &H1B

' Windows API call used to control textbox
'
#If Win16 Then
   Private Declare Function SendMessage Lib "User" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
#ElseIf Win32 Then
   Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If
'
' Edit Control Messages
'
Const WM_CUT = &H300
Const WM_COPY = &H301
Const WM_PASTE = &H302
Const WM_CLEAR = &H303
Const WM_UNDO = &H304
#If Win16 Then
   Const EM_CANUNDO = &H416     'WM_USER + 22
   Const EM_GETMODIFY = &H408   'WM_USER + 8
#ElseIf Win32 Then
   Const EM_CANUNDO = &HC6
   Const EM_GETMODIFY = &HB8
#End If
'
' Edit menu array constants
'
Const mUndo = 0
Const mCut = 2
Const mCopy = 3
Const mPaste = 4
Const mDelete = 5
'
' Flag to track status of Control key
Const ATTR_NORMAL = 0
Const ATTR_READONLY = 1
Const ATTR_HIDDEN = 2
Const ATTR_SYSTEM = 4
Const ATTR_VOLUME = 8
Const ATTR_DIRECTORY = 16
Const ATTR_ARCHIVE = 32

Private m_ControlKey As Boolean

Public UserMsgChoice As String
Public MsgMode As String    'Resolves who called msg dialog
Public ReplaceFlag As Boolean    'Tells btnFind_Click() whether or not

Private Filename As String  ' The full file name.
Private FileTitle As String ' The file name without path.

Private DataModified As Boolean














Private Sub Command5_Click()

WebBrowser1.Visible = True

 


Text3.Text = 寻找文本_取文本中间(Text1.Text, "value='", "' />")



Text1.Text = Replace(Text1, Text3.Text, Text2)







Dim Filenum
Filenum = FreeFile

Open "Preview.htm" For Output As Filenum
Print #Filenum, HTMLx.Text1.Text
Close Filenum
WebBrowser1.Navigate App.Path & "\Preview.htm"
''Open App.Path & "\preview.html" For Output As #1
''Print #1, Text1.Text
''Close #1
''Load frmBrowser
''frmBrowser.Show
''frmBrowser.brwWebBrowser.Navigate App.Path & "\preview.html"
End Sub







Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   ' Watch for Control key, set flag
   '
   If KeyCode = vbKeyControl Then
      m_ControlKey = True
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   '
   ' Watch for Control key, clear flag
   '
   If KeyCode = vbKeyControl Then
      m_ControlKey = False
   End If
End Sub

Private Sub Form_Load()

 Dim DirPath As String
  DirPath = App.Path
If Right$(DirPath, 1) <> "\" Then
        DirPath = DirPath + "\"
 End If 'Text1.RightMargin = 8670
'mControl_Click 1
Dim wid As Single
Dim hgt As Single
    
    ' Get the last directory the program accessed.
    ' If there is no entry, use the App.Path.
    '
    ' @@@ Change the application and section names.
    FileDialog.InitDir = GetSetting( _
        "SimpleEditor", "Directories", _
        "SaveDir", App.Path)
        HTMLx.Caption = "PHPLive"
        
  WebBrowser1.Navigate "about:blank"
  
 
   
  
        
        
End Sub


Private Sub Form_Resize()
On Error GoTo errhandler

Text1.Height = ScaleHeight - Picture2.Height - 500
Exit Sub
errhandler:
End Sub




Private Sub mnuColor_Click(Index As Integer)

    Select Case Index
           Case 0
               Text1.BackColor = &HFF&
           Case 1
               Text1.BackColor = &HFF00&
           Case 2
               Text1.BackColor = &HFF0000
           Case 3
               Text1.BackColor = &HFFFF00
           Case 4
               Text1.BackColor = &HFF00FF
           Case 5
               Text1.BackColor = &HFFFFFF
           Case 6
               Text1.BackColor = &HC0C0C0

    End Select
    
End Sub





Private Sub mEdit1_Click()
SendKeys "%{BS}"
End Sub

Private Sub mEdit2_Click()
SendKeys "+{DEL}"
End Sub

Private Sub mEdit3_Click()
SendKeys "^{INSERT}"
End Sub

Private Sub mEdit4_Click()
SendKeys "+{INSERT}"
End Sub
















Private Sub mnuFileNew_Click()
 Text1.Text = ""
 HTMLx.Caption = "Untitled"
End Sub



Private Sub mnuFileExit_Click()
   Unload Me
End Sub




Private Sub mnuFileOpen_Click()

'Declare the variables

Dim strFilename As String

'Set up the common dialog control
'so that if the cancel button is
'pressed, it generated a runtime
'error that can be caught

FileDialog.CancelError = True
On Error GoTo errhandler

'Load the open dialog box and return
'the selected file path into the
'variable strFilename

FileDialog.Filter = _
"HTML Files|*.htm*|Text Files|*.txt*"
FileDialog.ShowOpen
strFilename = FileDialog.Filename

'Read in the text file strFilename
'into the Text1 text box

Open strFilename For Input As #1
HTMLx.Text1 = Input(LOF(1), 1)
Close #1
HTMLx.Caption = "" & UCase(strFilename)
Exit Sub

'User pressed the Cancel button
errhandler:



End Sub


Private Sub mnuFileSave_Click()
If MsgBox("Do not forget to add the EXTENTION .txt or .htm to the file. Example: Whatever.txt or Myfile.htm", vbYesNo + vbCritical, "Warning") = vbYes Then
 
 Dim strFilename As String
 'Declare variableDim strFilename As String'Do the error handler
FileDialog.CancelError = True
On Error GoTo errhandler
'Set the properties of the text control
FileDialog.Filter = _
"Text Files|*.txt*|HTML Files|*.htm*"
'Show the save-as dialog box
FileDialog.ShowSave 'retrieve the filename
strFilename = FileDialog.Filename
'save the file
Open strFilename For Output As #1
Print #1, HTMLx.Text1
Close #1
End If
Exit Sub
errhandler:
End Sub















Private Sub mnuPopCopy_Click()
   'Call mnuEditCopy_Click
End Sub

Private Sub mnuPopCut_Click()
    'Call mnuEditCut_Click
End Sub

Private Sub mnuPopDel_Click()
    Screen.ActiveControl.SelText = ""
End Sub

Private Sub mnuPopFind_Click()
    'Call mnufind_Click
End Sub

Private Sub mnuPopFindN_Click()
        'Call mnuFindNext_Click
End Sub


Private Sub mnuPopPaste_Click()
    'Call mnuEditPaste_Click
End Sub

Private Sub mnuPopUndo_Click()
    'Call mnuEditUndo_Click
End Sub



Private Sub mnuSellTR_Click()
If Text1.Visible = True Then
         
         Text1.Visible = False
        Text1.Visible = True
HTMLx.Text1.SelStart = 0
HTMLx.Text1.SelLength = Len(HTMLx.Text1.Text)
      
    Else
         Text1.Visible = True
         Text1.Visible = False
       
HTMLx.Text1.SelStart = 0
HTMLx.Text1.SelLength = Len(HTMLx.Text1.Text)
   End If
End Sub





Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If Left$(HTMLx.Caption, 12) = "Ezzahir" Then
    HTMLx.Caption = "Untitled"
Else
    HTMLx.Caption = HTMLx.Caption
End If


End Sub





Private Sub mExit_Click()
   Unload Me
End Sub

Private Sub mnuH1_Click()
HTMLx.Text1.SelText = "<h1>" + HTMLx.Text1.SelText + "</h1>"
End Sub

Private Sub mnuH2_Click()
HTMLx.Text1.SelText = "<h2>" + HTMLx.Text1.SelText + "</h2>"
End Sub

Private Sub mnuH3_Click()
HTMLx.Text1.SelText = "<h3>" + HTMLx.Text1.SelText + "</h3>"
End Sub

Private Sub mnuH4_Click()
HTMLx.Text1.SelText = "<h4>" + HTMLx.Text1.SelText + "</h4>"
End Sub

Private Sub mnuH5_Click()
HTMLx.Text1.SelText = "<h5>" + HTMLx.Text1.SelText + "</h5>"
End Sub

Private Sub mnuH6_Click()
HTMLx.Text1.SelText = "<h6>" + HTMLx.Text1.SelText + "</h6>"
End Sub
Private Sub mnuBlack_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#000000>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuBlue_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#0000FF>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuBlueViolet_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#9F5F9F>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuBrown_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#A62A2A>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuCyan_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#00FFFF>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuDarkBrown_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#5C4033>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuDarkGreen_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#2F4F2F>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuDarkPurple_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#871F78>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuGold_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#CD7F32>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuGreen_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#00FF00>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuMagenta_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#FF00FF>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuOrange_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#FF7F00>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuRed_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#FF0000>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuTan_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#DB9370>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuWhite_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#FFFFFF>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuYellow_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#FFFF00>" + HTMLx.Text1.SelText + "</font>"
End Sub
Private Sub mnuLeft_Click()
HTMLx.Text1.SelText = "<p align=left>" + HTMLx.Text1.SelText + "</p>"
End Sub
Private Sub mnuCenter_Click()
HTMLx.Text1.SelText = "<center>" + HTMLx.Text1.SelText + "</center>"
End Sub
Private Sub mnuRight_Click()
HTMLx.Text1.SelText = "<p align=right>" + HTMLx.Text1.SelText + "</p>"
End Sub
Private Sub mnuBlink_Click()
HTMLx.Text1.SelText = "<blink>" + HTMLx.Text1.SelText + "</blink>"
End Sub
Private Sub mnuBold_Click()
HTMLx.Text1.SelText = "<b>" + HTMLx.Text1.SelText + "</b>"
End Sub
Private Sub mnuCite_Click()
HTMLx.Text1.SelText = "<cite>" + HTMLx.Text1.SelText + "</cite>"
End Sub
Private Sub mnuItalic_Click()
HTMLx.Text1.SelText = "<i>" + HTMLx.Text1.SelText + "</i>"
End Sub
Private Sub mnuStrong_Click()
HTMLx.Text1.SelText = "<strong>" + HTMLx.Text1.SelText + "</strong>"
End Sub
Private Sub mnuStrikeThrough_Click()
HTMLx.Text1.SelText = "<strike>" + HTMLx.Text1.SelText + "</strike>"
End Sub
Private Sub mnuTypeWriter_Click()
HTMLx.Text1.SelText = "<pre>" + HTMLx.Text1.SelText + "</pre>"
End Sub
Private Sub mnuunderline_Click()
HTMLx.Text1.SelText = "<u>" + HTMLx.Text1.SelText + "</u>"
End Sub
Private Sub mnuCells_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=1>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub
Private Sub mnuAddCH_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "</TR>" + Chr(13) + Chr(10)
End Sub
Private Sub mnuAddCol_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub











Private Sub mnuTime1_Click()
HTMLx.Text1.SelText = Format(Time$, "Short Time")
End Sub

Private Sub mnuTime2_Click()
HTMLx.Text1.SelText = Format(Time$, "Medium Time")
End Sub

Private Sub mnuTime3_Click()
HTMLx.Text1.SelText = Format(Time$, "Long Time")
End Sub
Private Sub mnuDate1_Click()
HTMLx.Text1.SelText = Format(Date$, "Short Date")
End Sub

Private Sub mnuDate2_Click()
HTMLx.Text1.SelText = Format(Date$, "Medium Date")
End Sub

Private Sub mnuDate3_Click()
HTMLx.Text1.SelText = Format(Date$, "Long Date")
End Sub









