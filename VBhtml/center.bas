Attribute VB_Name = "basCenter"

'Download by http://www.codefans.net
Public Sub sCenterForm(tmpF As Form)

' Declare Screen Cortdinates
Dim x As Integer, y As Integer

' The "/" is to divide 2 Integers and returns a
' floating Point Result.  "\" is a quicker division.
y = (Screen.Height - tmpF.Height) \ 2
x = (Screen.Width - tmpF.Width) \ 2

' Use Move because is is faster than setting
' single Properties.
tmpF.Move x, y

End Sub

