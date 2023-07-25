VERSION 5.00
Begin VB.Form HTML 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   1020
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuSize 
      Caption         =   "Size"
      Begin VB.Menu mnuH1 
         Caption         =   "H1"
      End
      Begin VB.Menu mnuH2 
         Caption         =   "H2"
      End
      Begin VB.Menu mnuH3 
         Caption         =   "H3"
      End
      Begin VB.Menu mnuH4 
         Caption         =   "H4"
      End
      Begin VB.Menu mnuH5 
         Caption         =   "H5"
      End
      Begin VB.Menu mnuH6 
         Caption         =   "H6"
      End
      Begin VB.Menu mnuOthers 
         Caption         =   "Font Size"
      End
   End
   Begin VB.Menu mnuTables 
      Caption         =   "Tables"
      Begin VB.Menu mnuCells 
         Caption         =   "Add the first column"
         Begin VB.Menu mnuColCells 
            Caption         =   "Choose Background"
            Begin VB.Menu mnu1 
               Caption         =   "Black"
            End
            Begin VB.Menu mnu2 
               Caption         =   "Blue"
            End
            Begin VB.Menu mnu3 
               Caption         =   "Blue violet"
            End
            Begin VB.Menu mnu4 
               Caption         =   "Brown"
            End
            Begin VB.Menu mnu5 
               Caption         =   "Cyan"
            End
            Begin VB.Menu mnu6 
               Caption         =   "Dark browm"
            End
            Begin VB.Menu mnu7 
               Caption         =   "Dark green"
            End
            Begin VB.Menu mnu8 
               Caption         =   "Dark blue"
            End
            Begin VB.Menu mnu9 
               Caption         =   "Gold"
            End
            Begin VB.Menu mnu10 
               Caption         =   "Green"
            End
            Begin VB.Menu mnu11 
               Caption         =   "Magenta"
            End
            Begin VB.Menu mnu12 
               Caption         =   "Orange"
            End
            Begin VB.Menu mnu13 
               Caption         =   "Red"
            End
            Begin VB.Menu mnu14 
               Caption         =   "Tan"
            End
            Begin VB.Menu mnu15 
               Caption         =   "White"
            End
            Begin VB.Menu mnu16 
               Caption         =   "Yellow"
            End
         End
      End
      Begin VB.Menu mnuAddCol 
         Caption         =   "Add new column"
         Begin VB.Menu mnuColBac 
            Caption         =   "Choose backgound"
            Begin VB.Menu mnu1a 
               Caption         =   "Black"
            End
            Begin VB.Menu mnu2a 
               Caption         =   "Blue"
            End
            Begin VB.Menu mnu3a 
               Caption         =   "Blue violet"
            End
            Begin VB.Menu mnu4a 
               Caption         =   "Brown"
            End
            Begin VB.Menu mnu5a 
               Caption         =   "Cyan"
            End
            Begin VB.Menu mnu6a 
               Caption         =   "Dark browm"
            End
            Begin VB.Menu mnu7a 
               Caption         =   "Dark Green"
            End
            Begin VB.Menu mnu8a 
               Caption         =   "Dark blue"
            End
            Begin VB.Menu mnu9a 
               Caption         =   "Gold"
            End
            Begin VB.Menu mnu10a 
               Caption         =   "Green"
            End
            Begin VB.Menu mnu11a 
               Caption         =   "Magenta"
            End
            Begin VB.Menu mnu12a 
               Caption         =   "Orange"
            End
            Begin VB.Menu mnu13a 
               Caption         =   "Red"
            End
            Begin VB.Menu mnu14a 
               Caption         =   "Tan"
            End
            Begin VB.Menu mnu15a 
               Caption         =   "White"
            End
            Begin VB.Menu mnu16a 
               Caption         =   "Yellow"
            End
         End
      End
      Begin VB.Menu mnuAddCH 
         Caption         =   "Add cells"
      End
   End
   Begin VB.Menu mnuPositions 
      Caption         =   "Position of the text"
      Begin VB.Menu mnuRight 
         Caption         =   "Right"
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "Left"
      End
      Begin VB.Menu mnuCenter 
         Caption         =   "Center"
      End
   End
   Begin VB.Menu mnuFunctions 
      Caption         =   "Functions"
      Begin VB.Menu mnuLink 
         Caption         =   "Link"
      End
      Begin VB.Menu mnuPictureH 
         Caption         =   "Picture"
      End
   End
   Begin VB.Menu mnuFonts 
      Caption         =   "Fonts"
      Begin VB.Menu mnuBlack 
         Caption         =   "Black"
      End
      Begin VB.Menu mnuBlue 
         Caption         =   "Blue"
      End
      Begin VB.Menu mnuBlueViolet 
         Caption         =   "Blue violet"
      End
      Begin VB.Menu mnuBrown 
         Caption         =   "Brown"
      End
      Begin VB.Menu mnuCyan 
         Caption         =   "Cyan"
      End
      Begin VB.Menu mnuDarkBrown 
         Caption         =   "Dark brown"
      End
      Begin VB.Menu mnuDarkGreen 
         Caption         =   "Dark green"
      End
      Begin VB.Menu mnuDarkPurple 
         Caption         =   "Dark purple"
      End
      Begin VB.Menu mnuGold 
         Caption         =   "Gold"
      End
      Begin VB.Menu mnuGreen 
         Caption         =   "Green"
      End
      Begin VB.Menu mnuMagenta 
         Caption         =   "Magenta"
      End
      Begin VB.Menu mnuOrange 
         Caption         =   "Orange"
      End
      Begin VB.Menu mnuRed 
         Caption         =   "Red"
      End
      Begin VB.Menu mnuTan 
         Caption         =   "Tan"
      End
      Begin VB.Menu mnuWhite 
         Caption         =   "White"
      End
      Begin VB.Menu mnuYellow 
         Caption         =   "Yellow"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItalics 
         Caption         =   "Italic"
      End
      Begin VB.Menu mnuBolds 
         Caption         =   "Bold"
      End
      Begin VB.Menu mnuunderlines 
         Caption         =   "Underline"
      End
   End
   Begin VB.Menu mnuStyles 
      Caption         =   "Styles"
      Begin VB.Menu mnuBlink 
         Caption         =   "Blinking"
      End
      Begin VB.Menu mnuBold 
         Caption         =   "Bold"
      End
      Begin VB.Menu mnuCite 
         Caption         =   "Citation"
      End
      Begin VB.Menu mnuItalic 
         Caption         =   "Italic"
      End
      Begin VB.Menu mnuStrikeThrough 
         Caption         =   "Strikethrough"
      End
      Begin VB.Menu mnuStrong 
         Caption         =   "Strong"
      End
      Begin VB.Menu mnuTypeWriter 
         Caption         =   "Typewriter"
      End
      Begin VB.Menu mnuunderline 
         Caption         =   "Underline"
      End
   End
   Begin VB.Menu mnuOtherss 
      Caption         =   "Others"
      Begin VB.Menu mnuUnnumberesLists 
         Caption         =   "Unnumbered Lists"
      End
      Begin VB.Menu mnuNumberedLists 
         Caption         =   "Numbered Lists"
      End
      Begin VB.Menu mnuDefinitionLists 
         Caption         =   "Definition Lists"
      End
      Begin VB.Menu mnuNestedLists 
         Caption         =   "Nested Lists"
      End
      Begin VB.Menu mnuExtendedQuotations 
         Caption         =   "Extended Quotations"
      End
      Begin VB.Menu mnuBreaks 
         Caption         =   "Forced Line Breaks"
      End
      Begin VB.Menu mnuHorRules 
         Caption         =   "Horizontal Rules"
      End
      Begin VB.Menu mnuAnchor1 
         Caption         =   "Anchor"
         Begin VB.Menu mnuAnchor 
            Caption         =   "Anchor of the phrase to be linked"
         End
         Begin VB.Menu mnuAnchor2 
            Caption         =   "Anchor where the phrase will be linked to"
         End
      End
      Begin VB.Menu mnuWhitespace 
         Caption         =   "White space"
      End
   End
   Begin VB.Menu mnuLinkEmail 
      Caption         =   "Link and E-Mail "
      Begin VB.Menu mnuLink1 
         Caption         =   "Link without image"
      End
      Begin VB.Menu mnuLink2 
         Caption         =   "Link with image"
      End
   End
   Begin VB.Menu mnuFonth 
      Caption         =   "Fonts"
      Begin VB.Menu mnuFonts1 
         Caption         =   "Abadi MT Condensed"
      End
      Begin VB.Menu mnuFonts2 
         Caption         =   "Arial"
      End
      Begin VB.Menu mnuFonts3 
         Caption         =   "Arial Black"
      End
      Begin VB.Menu mnuFonts4 
         Caption         =   "Arial Narrow"
      End
      Begin VB.Menu mnuFonts5 
         Caption         =   "Bookman Old Style"
      End
      Begin VB.Menu mnuFonts6 
         Caption         =   "Comic Sans MS"
      End
      Begin VB.Menu mnuFonts7 
         Caption         =   "Courier"
      End
      Begin VB.Menu mnuFonts8 
         Caption         =   "Courier New"
      End
      Begin VB.Menu mnuFonts9 
         Caption         =   "Fixedsys"
      End
      Begin VB.Menu mnuFonts10 
         Caption         =   "Garamond"
      End
      Begin VB.Menu mnuFonts11 
         Caption         =   "Impact"
      End
      Begin VB.Menu mnuFonts12 
         Caption         =   "MS Sans Serif"
      End
      Begin VB.Menu mnuFonts13 
         Caption         =   "MS Serif"
      End
      Begin VB.Menu mnuFonts14 
         Caption         =   "Marlett"
      End
      Begin VB.Menu mnuFonts15 
         Caption         =   "Small Fonts"
      End
      Begin VB.Menu mnuFonts16 
         Caption         =   "Symbol"
      End
      Begin VB.Menu mnuFonts17 
         Caption         =   "System"
      End
      Begin VB.Menu mnuFonts18 
         Caption         =   "Tahoma"
      End
      Begin VB.Menu mnuFonts19 
         Caption         =   "Terminal"
      End
      Begin VB.Menu mnuFonts20 
         Caption         =   "Times New Roman"
      End
      Begin VB.Menu mnuFonts21 
         Caption         =   "Verdana"
      End
      Begin VB.Menu mnuFonts22 
         Caption         =   "Webdings"
      End
      Begin VB.Menu mnuFonts23 
         Caption         =   "Wingdings"
      End
      Begin VB.Menu mnuFonts24 
         Caption         =   "Wingdings 2"
      End
      Begin VB.Menu mnuFonts25 
         Caption         =   "Wingdings 3"
      End
   End
   Begin VB.Menu mnuTables1 
      Caption         =   "Tables1"
      Begin VB.Menu mnuCol1 
         Caption         =   "Add One Column"
      End
      Begin VB.Menu mnuCol2 
         Caption         =   "Add Two Columns"
      End
      Begin VB.Menu mnuCol3 
         Caption         =   "Add Three Columns"
      End
      Begin VB.Menu mnuCol4 
         Caption         =   "Add Four Columns"
      End
      Begin VB.Menu mnuCol5 
         Caption         =   "Add more Columns"
      End
      Begin VB.Menu mnuSepCol6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCol7 
         Caption         =   "Add Rows"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy background codes"
      End
   End
End
Attribute VB_Name = "HTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnu1_Click()
'HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#000000>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
HTMLx.Text1.SelText = HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#CD7F32>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + "Here add new cells for the first column" + Chr(13) + Chr(10) + "Here add the second column" + Chr(13) + Chr(10) + "Here add new cells for the second column, and so on " + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu10_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#00FF00>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu10a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#00FF00>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu11_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FF00FF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu11a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FF00FF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu12_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FF7F00>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu12a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FF7F00>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu13_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FF0000>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu13a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FF0000>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu14_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#DB9370>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu14a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#DB9370>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu15_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FFFFFF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu15a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FFFFFF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu16_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FFFF00>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu16a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FFFF00>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu1a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#000000>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu2_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#0000FF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu2a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#0000FF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu3_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#9F5F9F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu3a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#9F5F9F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu4_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#A62A2A>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu4a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#A62A2A>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu5_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#00FFFF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu5a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#00FFFF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu6_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#5C4033>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu6a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#5C4033>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu7_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#2F4F2F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu7a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#2F4F2F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu8_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#871F78>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu8a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#871F78>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu9_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#CD7F32>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu9a_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#CD7F32>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnuAnchor_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<P><A HREF=#anchoreName (eg. #1)>Enter the text to be linked, eg. Editor</A></P>"
End Sub

Private Sub mnuAnchor2_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<P><A NAME=anchoreName></A>Enter the text to be linked to, eg. Editor</P>"
End Sub

Private Sub mnuBold_Click()
HTMLx.Text1.SelText = "<b>" + HTMLx.Text1.SelText + "</b>"
End Sub

Private Sub mnuBreaks_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<BR>"
End Sub

Private Sub mnuCol1_Click()
HTMLx.Text1.SelText = HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<P><TABLE BORDER=1>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the code of the cell's background>" + Chr(13) + Chr(10) + "<P>Your your text in the first cell" + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the code of the cell's background>" + Chr(13) + Chr(10) + "<P>Write your text in the second cell" + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + "xxxxxxxxxxxxxxxxxxxxxxxxxxx" + Chr(13) + Chr(10) + "Copy and Paste the following code to add more cells" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the code of the cell's background>" + Chr(13) + Chr(10) + "<P>Your your text in the cell" + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + "xxxxxxxxxxxxxxxxxxxxxxxxx" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub mnuCol2_Click()
HTMLx.Text1.SelText = HTMLx.Text1.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub mnuCol3_Click()
HTMLx.Text1.SelText = HTMLx.Text1.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub mnuCol4_Click()
HTMLx.Text1.SelText = HTMLx.Text1.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub mnuCol5_Click()
HTMLx.Text1.SelText = HTMLx.Text1.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "ADD HERE COLUMNS. Select and Paste one of the two lines <P></TD><TD>" + Chr(13) + Chr(10) + "<P>Write your Text here (The cell of the LAST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub mnuCol7_Click()
HTMLx.Text1.SelText = HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor= >" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the added ROW)" + Chr(13) + Chr(10) + "</TD><TD bgcolor= >" + Chr(13) + Chr(10) + "ADD HERE CELLS. Select and Paste the two lines above: <P>...</TD><TD>" + Chr(13) + Chr(10) + "<P>Write your Text here (Last cell of the added ROW)" + Chr(13) + Chr(10) + "</TD></TR>ADD HERE ROWS"
End Sub

Private Sub mnuCol8_Click()
HTMLx.Text1.SelText = HTMLx.Text1.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (4th cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub mnuCopy_Click()
Codes.Visible = True
End Sub

Private Sub mnuDefinitionLists_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<DL>" + Chr(13) + Chr(10) + "<DT> Paragraph Title" + Chr(13) + Chr(10) + "<DD> Your Text Here" + Chr(13) + Chr(10) + "<DT> Second Paragraph Title" + Chr(13) + Chr(10) + "<DD> Your Text Here" + Chr(13) + Chr(10) + "</DL>"
End Sub

Private Sub mnuExtendedQuotations_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<P>Your text" + Chr(13) + Chr(10) + "<BLOCKQUOTE>" + Chr(13) + Chr(10) + "<P> Write your text here to include lengthy quotations in a separate block on the screen" + Chr(13) + Chr(10) + "</P>" + Chr(13) + Chr(10) + "<P> Add more text here if you want</P>" + Chr(13) + Chr(10) + "</BLOCKQUOTE>"
End Sub

Private Sub mnuFonh_Click(Index As Integer)

End Sub

Private Sub mnuFonts1_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Abadi MT Condensed"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts10_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Garamond"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts11_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Impact"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts12_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""MS Sans Serif"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts13_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""MS Serif"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts14_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Marlett"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts15_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Small Fonts"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts16_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Symbol"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts17_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""System"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts18_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Tahoma"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts19_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Terminal"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts2_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Arial"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts20_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Times New Roman"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts21_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Verdana"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts22_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Webdings"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts23_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Wingdings"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts24_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Wingdings 2"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts25_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Wingdings 3"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts3_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Arial Black"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts4_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Arial Narrow"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts5_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Bookman Old Style"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts6_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Comic Sans MS"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts7_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Courier"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts8_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Courier New"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
End Sub

Private Sub mnuFonts9_Click()
HTMLx.Text1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Fixedsys"">" + "Type your Text Here" + "</FONT>" + HTMLx.Text1.SelText
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

Private Sub mnuAddCH_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10)
End Sub


Private Sub mnuHorRules_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<HR SIZE= Enter the desired size    WIDTH=" + "Enter a number %>"
End Sub

Private Sub mnuItalic_Click()
HTMLx.Text1.SelText = "<i>" + HTMLx.Text1.SelText + "</i>"
End Sub

Private Sub mnuLink1_Click()
Form2.Visible = True
End Sub

Private Sub mnuLink2_Click()
Form2A.Visible = True
End Sub

Private Sub mnuNestedLists_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<UL>" + Chr(13) + Chr(10) + "<LI> Sub-heading" + Chr(13) + Chr(10) + "<UL>" + Chr(13) + Chr(10) + "<LI> Your Text Here" + Chr(13) + Chr(10) + "<LI> Your Text Here" + Chr(13) + Chr(10) + "<LI> Your Text Here. Add more <LI> if necessary" + Chr(13) + Chr(10) + "</UL>" + Chr(13) + Chr(10) + "<LI> Second Sub-heading" + Chr(13) + Chr(10) + "<UL>" + Chr(13) + Chr(10) + "<LI> Your Text Here" + Chr(13) + Chr(10) + "<LI> Your Text Here" + Chr(13) + Chr(10) + "<LI> Your Text Here. Add more <LI> if necessary" + Chr(13) + Chr(10) + "</UL>" + Chr(13) + Chr(10) + "</UL>"
End Sub

Private Sub mnuNumberedLists_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<OL>" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here and add more <LI> if necessary" + Chr(13) + Chr(10) + "</OL>"
End Sub

Private Sub mnuOthers_Click()
Form11.Visible = True
End Sub

Private Sub mnuRight_Click()
HTMLx.Text1.SelText = "<p align=right>" + HTMLx.Text1.SelText + "</p>"
End Sub
Private Sub mnuCenter_Click()
HTMLx.Text1.SelText = "<center>" + HTMLx.Text1.SelText + "</center>"
End Sub
Private Sub mnuLeft_Click()
HTMLx.Text1.SelText = "<p align=left>" + HTMLx.Text1.SelText + "</p>"
End Sub
Private Sub mnuLink_Click()
Form2.Visible = True
End Sub

Private Sub mnuPictureH_Click()
Form4.Visible = True
End Sub
Private Sub mnuBlack_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#000000>" + HTMLx.Text1.SelText + "</FONT>"
End Sub
Private Sub mnuBlue_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#0000FF>" + HTMLx.Text1.SelText + "</FONT>"
End Sub
Private Sub mnuBlueViolet_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#9F5F9F>" + HTMLx.Text1.SelText + "</FONT>"
End Sub
Private Sub mnuBrown_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#A62A2A>" + HTMLx.Text1.SelText + "</FONT>"
End Sub
Private Sub mnuCyan_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#00FFFF>" + HTMLx.Text1.SelText + "</FONT>"
End Sub
Private Sub mnuDarkBrown_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#5C4033>" + HTMLx.Text1.SelText + "</FONT>"
End Sub
Private Sub mnuDarkGreen_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#2F4F2F>" + HTMLx.Text1.SelText + "</FONT>"
End Sub
Private Sub mnuDarkPurple_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#871F78>" + HTMLx.Text1.SelText + "</FONT>"
End Sub
Private Sub mnuGold_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#CD7F32>" + HTMLx.Text1.SelText + "</FONT>"
End Sub
Private Sub mnuGreen_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#00FF00>" + HTMLx.Text1.SelText + "</FONT>"
End Sub
Private Sub mnuMagenta_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#FF00FF>" + HTMLx.Text1.SelText + "</FONT>"
End Sub
Private Sub mnuOrange_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#FF7F00>" + HTMLx.Text1.SelText + "</FONT>"
End Sub
Private Sub mnuRed_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#FF0000>" + HTMLx.Text1.SelText + "</FONT>"
End Sub

Private Sub mnuTan_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#DB9370>" + HTMLx.Text1.SelText + "</FONT>"
End Sub

Private Sub mnuunderline_Click()
HTMLx.Text1.SelText = "<u>" + HTMLx.Text1.SelText + "</u>"
End Sub

Private Sub mnuUnnumberesLists_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<UL>" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here and add more <LI> if necessary" + Chr(13) + Chr(10) + "</UL>"
End Sub

Private Sub mnuWhite_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#FFFFFF>" + HTMLx.Text1.SelText + "</FONT>"
End Sub

Private Sub mnuWhitespace_Click()
HTMLx.Text1.SelText = Chr(13) + Chr(10) + HTMLx.Text1.SelText + Chr(13) + Chr(10) + "<P>&nbsp;</P>"
End Sub

Private Sub mnuYellow_Click()
HTMLx.Text1.SelText = "<FONT COLOR=#FFFF00>" + HTMLx.Text1.SelText + "</FONT>"
End Sub
Private Sub mnuBlink_Click()
HTMLx.Text1.SelText = "<blink>" + HTMLx.Text1.SelText + "</blink>"
End Sub
Private Sub mnuBolds_Click()
HTMLx.Text1.SelText = "<b>" + HTMLx.Text1.SelText + "</b>"
End Sub
Private Sub mnuCite_Click()
HTMLx.Text1.SelText = "<cite>" + HTMLx.Text1.SelText + "</cite>"
End Sub
Private Sub mnuItalics_Click()
HTMLx.Text1.SelText = "<i>" + HTMLx.Text1.SelText + "</i>"
End Sub
Private Sub mnuStrikeThrough_Click()
HTMLx.Text1.SelText = "<strike>" + HTMLx.Text1.SelText + "</strike>"
End Sub
Private Sub mnuStrong_Click()
HTMLx.Text1.SelText = "<strong>" + HTMLx.Text1.SelText + "</strong>"
End Sub
Private Sub mnuTypeWriter_Click()
HTMLx.Text1.SelText = "<pre>" + HTMLx.Text1.SelText + "</pre>"
End Sub
Private Sub mnuunderlines_Click()
HTMLx.Text1.SelText = "<u>" + HTMLx.Text1.SelText + "</u>"
End Sub

