VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "iptv"
   ClientHeight    =   7740
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   4245
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   4245
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text2 
      Height          =   1575
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":08CA
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   6900
      ItemData        =   "Form1.frx":08D0
      Left            =   360
      List            =   "Form1.frx":08D2
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin VB.Menu t1 
      Caption         =   "File(F)"
      Begin VB.Menu t12 
         Caption         =   "Login"
      End
   End
   Begin VB.Menu t2 
      Caption         =   "Play(P)"
      Begin VB.Menu t24 
         Caption         =   "P2P    H.265 A"
      End
      Begin VB.Menu t25 
         Caption         =   "P2P    H.265 B"
      End
      Begin VB.Menu t26 
         Caption         =   "P2P    H.265 C"
      End
      Begin VB.Menu t27 
         Caption         =   "P2P    H.265 D"
      End
      Begin VB.Menu t28 
         Caption         =   "P2P    H.265 E"
      End
   End
   Begin VB.Menu t3 
      Caption         =   "Help(H)"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Function strGetDate() As String
    Dim XmlHttp As Object
    Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
    XmlHttp.Open "Get", "http://www.lwfire.cn/date.asp", False
    XmlHttp.Send
    strGetDate = StrConv(XmlHttp.ResponseBody, vbUnicode)
    Set XmlHttp = Nothing
End Function
Private Sub Form_Load()

If strGetDate >= #12/31/2026# Then
    
    
    t1.Visible = False
    
    t2.Visible = False
    
    t3.Visible = False
   
        
      Else
      
  t1.Visible = True
      
  t2.Visible = True
  
    
  t3.Visible = True
  
  
    End If
  




Text1.Visible = False


Text2.Visible = False





List1.Visible = False



End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Form2.PopupMenu t2

End Sub

Private Sub t12_Click()

    
List1.Visible = True


Set xmlobject = CreateObject("Microsoft.XMLHTTP")
strURL = "http://www.lwfire.cn/b/list.txt"
xmlobject.Open "GET", strURL, False
xmlobject.Send
If xmlobject.ReadyState = 4 Then
a = xmlobject.Responsetext
Form1.Text2.Text = a
End If

Dim s() As String
s = Split(Text2.Text, vbCrLf)

For i = 0 To UBound(s)

List1.AddItem (s(i))

Next i

    
  
    
   




End Sub










Private Sub t24_Click()
Set xmlobject = CreateObject("Microsoft.XMLHTTP")
strURL = "http://www.lwfire.cn/b/line.txt"
xmlobject.Open "GET", strURL, False
xmlobject.Send
If xmlobject.ReadyState = 4 Then
a = xmlobject.Responsetext
Form1.Text1.Text = a
End If
Form2.Show



Form2.WebBrowser1.Visible = True

Form2.WebBrowser1.Navigate App.Path & "\pp.htm?id=vjms://" & Text1.Text & ":8500:3502/live/cid=" & List1.Text


Form2.WebBrowser1.Silent = True
End Sub


Private Sub t25_Click()
Set xmlobject = CreateObject("Microsoft.XMLHTTP")
strURL = "http://www.lwfire.cn/b/line1.txt"
xmlobject.Open "GET", strURL, False
xmlobject.Send
If xmlobject.ReadyState = 4 Then
a = xmlobject.Responsetext
Form1.Text1.Text = a
End If
Form2.Show



Form2.WebBrowser1.Visible = True

Form2.WebBrowser1.Navigate App.Path & "\pp.htm?id=vjms://" & Text1.Text & ":8500:3502/live/cid=" & List1.Text


Form2.WebBrowser1.Silent = True
End Sub

Private Sub t26_Click()
Set xmlobject = CreateObject("Microsoft.XMLHTTP")
strURL = "http://www.lwfire.cn/b/line2.txt"
xmlobject.Open "GET", strURL, False
xmlobject.Send
If xmlobject.ReadyState = 4 Then
a = xmlobject.Responsetext
Form1.Text1.Text = a
End If
Form2.Show



Form2.WebBrowser1.Visible = True

Form2.WebBrowser1.Navigate App.Path & "\pp.htm?id=vjms://" & Text1.Text & ":8500:3502/live/cid=" & List1.Text


Form2.WebBrowser1.Silent = True
End Sub

Private Sub t27_Click()
Set xmlobject = CreateObject("Microsoft.XMLHTTP")
strURL = "http://www.lwfire.cn/b/line3.txt"
xmlobject.Open "GET", strURL, False
xmlobject.Send
If xmlobject.ReadyState = 4 Then
a = xmlobject.Responsetext
Form1.Text1.Text = a
End If
Form2.Show



Form2.WebBrowser1.Visible = True

Form2.WebBrowser1.Navigate App.Path & "\pp.htm?id=vjms://" & Text1.Text & ":8500:3502/live/cid=" & List1.Text


Form2.WebBrowser1.Silent = True
End Sub

Private Sub t28_Click()
Set xmlobject = CreateObject("Microsoft.XMLHTTP")
strURL = "http://www.lwfire.cn/b/line4.txt"
xmlobject.Open "GET", strURL, False
xmlobject.Send
If xmlobject.ReadyState = 4 Then
a = xmlobject.Responsetext
Form1.Text1.Text = a
End If
Form2.Show



Form2.WebBrowser1.Visible = True

Form2.WebBrowser1.Navigate App.Path & "\pp.htm?id=vjms://" & Text1.Text & ":8500:3502/live/cid=" & List1.Text


Form2.WebBrowser1.Silent = True
End Sub






