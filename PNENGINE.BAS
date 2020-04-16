Attribute VB_Name = "PrintEngine"
Option Explicit

' To print call the following routines (in this order !!):
'
'  PrintStartDoc
'  ... Your printing using PrintAt, PrintLine, etc
'  ... PrintNewPage
'  ... Your printing using PrintAt, PrintLine, etc
'  PrintEndDoc
'
' No physical printing occurs here. This is just a "formatter".
' All page data will be stored in "Layout()" variable and nLineCount will keep track of
' new elements. "Layout()" variable is defined in PnCommon.bas (PrintCommon module)
'
' An easy accessable picturebox control must be available. See font section why.
' This modules uses picPaper in the frmMain form (this form never unloads until program is finished).
'
' "Text only" reports are 1 based coordinates, eg position (1, 1) is top-left corner.
' Thus making zero nothing. Use a OEM font for previewing on screen.
'

' Constants used in PrintLine()
Public Const LINE_HORINZONTAL As Integer = 0
Public Const LINE_VERTICAL As Integer = 1

Dim lp As Integer                ' Layout pointer (into it's array)
Dim nCurrentX As Long            ' X,Y position holders (X = left|width, Y = top|height)
Dim nCurrentY As Long

Sub PrintStartDoc()
   ' Initialise Layout() array
   ClearLayout

   If Page.Cancelled Then Exit Sub
   If Not PrintStatus(PU_STARTPRINT) Then Exit Sub

   nCurrentX = Page.Margin.Left
   nCurrentY = Page.Margin.Top
End Sub

' Like PrintStartDoc(), but does not intefere with actual printing process
Sub PrintStartPage()
   ' Initialise Layout() array
   ClearLayout

   If Page.Cancelled Then Exit Sub

   nCurrentX = Page.Margin.Left
   nCurrentY = Page.Margin.Top
End Sub

Private Sub ClearLayout()
   If Page.Ruler = RULER_CHAR Then
      ' "Text only" used text output only
      Dim i As Integer

      ' Create "virtual" page with spaces
      ReDim Layout(1 To Page.Height)

      For i = 1 To Page.Height
         Layout(i).Mode = LYO_TEXT
         Layout(i).X = 0
         Layout(i).Y = Line2Twips(i - 1)
         Layout(i).Text.Text = Space$(Page.Width)
      Next

      nLineCount = Page.Height

   Else
      Erase Layout
      nLineCount = 0
   End If
End Sub

Sub PrintNewPage()
   If Page.Cancelled Then Exit Sub
   If Not PrintStatus(PU_NEWPAGE) Then Exit Sub

   ' Clear layout _after_ "new page" command is to user-interfaced form
   ClearLayout

   nCurrentX = Page.Margin.Left
   nCurrentY = Page.Margin.Top
End Sub

' Same as PrintNewPage(), but does not intefere with actual printing process
Sub PrintClearPage()
   If Page.Cancelled Then Exit Sub

   ' Clear layout _after_ "new page" command is to user-interfaced form
   ClearLayout

   nCurrentX = Page.Margin.Left
   nCurrentY = Page.Margin.Top
End Sub

Sub PrintKillDoc()         ' You could just use PrintEndDoc()....
   Page.Cancelled = True
   PrintStatus PU_ENDPRINT
End Sub

Sub PrintEndDoc()
   PrintStatus PU_ENDPRINT
End Sub

' Image print * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Sub PrintPicture(nIndex As Integer, ByVal nWidth As Single, ByVal nHeight As Single)
   If Page.Ruler = RULER_CHAR Then Exit Sub        ' No pictures form "text only" reports
   If Page.Cancelled Then Exit Sub

   lp = AddLayout(LYO_IMAGE, nCurrentX, nCurrentY)
   Layout(lp).Image.Index = nIndex                 ' if nIndex = -1 => sample icon from main form.
   Layout(lp).Image.Width = nWidth
   Layout(lp).Image.Height = nHeight

   nCurrentX = Page.Margin.Left
   nCurrentY = nCurrentY + Layout(lp).Image.Height
End Sub

' Circle print  * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Using absolute coordinates - whatever that means...
Sub PrintCircle(ByVal nLeft As Single, ByVal nTop As Single, ByVal nRadius As Single, Optional nColor)
   If Page.Ruler = RULER_CHAR Then Exit Sub        ' No circles in "text only" reports
   If Page.Cancelled Then Exit Sub
   If IsMissing(nColor) Then nColor = QBColor(0)   ' Default: black

   lp = AddLayout(LYO_CIRCLE, Page.Margin.Left + CTwips(nLeft), Page.Margin.Top + CTwips(nTop))
   Layout(lp).Circles.Radius = CTwips(nRadius)
   Layout(lp).Circles.Color = nColor
End Sub

' Line/Box print  * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Box with no contents
' Using absolute coordinates - whatever that means...
Sub PrintBox(ByVal nLeft As Single, ByVal nTop As Single, ByVal nWidth As Single, ByVal nHeight As Single, Optional nColor, Optional nStyle)
   If Page.Cancelled Then Exit Sub
   If IsMissing(nColor) Then nColor = QBColor(0)   ' Default: black
   If IsMissing(nStyle) Then nStyle = vbSolid      ' Solid line

   If Page.Ruler = RULER_CHAR Then
      Dim sLine As String
      Dim i As Integer

      ' Top line
      Select Case nWidth
      Case Is > 2
         sLine = "+" & String$(nWidth - 2, "-") & "+"
      Case Is = 2
         sLine = "++"
      Case Is < 1
         Exit Sub
      Case Else
         sLine = "+"
      End Select

      PutText sLine, Page.Margin.Left + nLeft, Page.Margin.Top + nTop

      If nHeight > 1 Then
         If nHeight > 2 Then
            For i = 1 To (nHeight - 2)
               PutText "|", Page.Margin.Left + nLeft, Page.Margin.Top + nTop + i
               If nWidth > 1 Then
                  PutText "|", Page.Margin.Left + nLeft + (nWidth - 1), Page.Margin.Top + nTop + i
               End If
            Next
         End If
         PutText sLine, Page.Margin.Left + nLeft, Page.Margin.Top + nTop + (nHeight - 1)
      End If
   
   Else
      lp = AddLayout(LYO_BOX, Page.Margin.Left + CTwips(nLeft), Page.Margin.Top + CTwips(nTop))
      Layout(lp).Line.Width = Page.Margin.Left + CTwips(nLeft + nWidth)
      Layout(lp).Line.Height = Page.Margin.Top + CTwips(nTop + nHeight)
      Layout(lp).Line.Color = nColor
      Layout(lp).Line.Style = nStyle
   End If
End Sub

' Valid styles: (see VB help)
'   0  Solid.
'   1  (Default) Transparent.
'   2  Horizontal Line.
'   3  Vertical Line.
'   4  Upward Diagonal.
'   5  Downward Diagonal.
'   6  Cross.
'   7  Diagonal Cross.
'
Sub PrintFilledBox(ByVal nLeft As Single, ByVal nTop As Single, ByVal nWidth As Single, ByVal nHeight As Single, Optional nColor, Optional nStyle)
   If Page.Cancelled Then Exit Sub
   If IsMissing(nColor) Then nColor = QBColor(0)   ' Default: black
   If IsMissing(nStyle) Then nStyle = vbSolid      ' Solid fill

   If Page.Ruler = RULER_CHAR Then
      Dim sLine As String, sBlock As String
      Dim i As Integer

      Select Case nStyle
      Case 0  ' Solid.
         sBlock = "#"
      Case 1  ' (Default) Transparent.
         PrintBox nLeft, nTop, nWidth, nHeight
         Exit Sub
      Case 2  ' Horizontal Line.
         sBlock = "-"
      Case 3  ' Vertical Line.
         sBlock = "|"
      Case 4  ' Upward Diagonal.
         sBlock = "/"
      Case 5  ' Downward Diagonal.
         sBlock = "\"
      Case 6  ' Cross.
         sBlock = "+"
      Case 7  ' Diagonal Cross.
         sBlock = "X"
      Case Else
         Exit Sub
      End Select

      sLine = String$(nWidth, sBlock)

      For i = 0 To (nHeight - 1)
         PutText sLine, Page.Margin.Left + nLeft, Page.Margin.Top + nTop + i
      Next

   Else
      lp = AddLayout(LYO_FILLBOX, Page.Margin.Left + CTwips(nLeft), Page.Margin.Top + CTwips(nTop))
      Layout(lp).Line.Width = Page.Margin.Left + CTwips(nLeft + nWidth)
      Layout(lp).Line.Height = Page.Margin.Top + CTwips(nTop + nHeight)
      Layout(lp).Line.Color = nColor
      Layout(lp).Line.Style = nStyle
   End If
End Sub

' Good for non-vertical or non-horinzontal lines (use PrintLine() instead for the other ones [easier])
Sub PrintDraw(ByVal nLeft As Single, ByVal nTop As Single, ByVal nRight As Single, ByVal nBottom As Single, Optional nColor, Optional nStyle)
   If Page.Ruler = RULER_CHAR Then Exit Sub        ' No "crooked" lines in "text only" reports
   If Page.Cancelled Then Exit Sub
   If IsMissing(nColor) Then nColor = QBColor(0)   ' Default: black
   If IsMissing(nStyle) Then nStyle = vbSolid      ' Solid fill
   
   lp = AddLayout(LYO_LINE, Page.Margin.Left + CTwips(nLeft), Page.Margin.Top + CTwips(nTop))
   Layout(lp).Line.Width = Page.Margin.Left + CTwips(nRight)
   Layout(lp).Line.Height = Page.Margin.Top + CTwips(nBottom)
   Layout(lp).Line.Color = nColor
   Layout(lp).Line.Style = nStyle
End Sub

' Line is always vertical or horinzontal, else use PrintDraw()
' Assumes horinzontal as default - Use this instead of PrintDraw() [it's easier]
Sub PrintLine(ByVal nLeft As Single, ByVal nTop As Single, Optional nLength, Optional nDirection, Optional nColor, Optional nStyle)
   If Page.Cancelled Then Exit Sub
   If IsMissing(nDirection) Then nDirection = LINE_HORINZONTAL
   If IsMissing(nColor) Then nColor = QBColor(0)   ' Default: black
   If IsMissing(nStyle) Then nStyle = vbSolid      ' Solid fill

   If Page.Ruler = RULER_CHAR Then
      If nDirection = LINE_VERTICAL Then           ' Going down...
         Dim i As Integer, nJump As Integer
         Select Case nStyle
         Case 0, 1, 3, 6   ' Solid, Dash, Dash-Dot, inside solid.
            nJump = 1
         Case 2, 4         ' Dot, Dash-Dot-Dot.
            nJump = 2
         Case 5            ' Transparent.
            Exit Sub
         End Select

         If IsMissing(nLength) Then nLength = GetHeight - nTop    ' Until edge of printable page
         nTop = Page.Margin.Top + nTop

         For i = nTop To (nTop + (nLength - 1)) Step nJump
            PutText "|", Page.Margin.Left + nLeft, i
         Next

      Else
         If IsMissing(nLength) Then nLength = GetWidth - nLeft    ' Until edge of printable page

         Dim sLine As String
         Select Case nStyle
         Case 0, 1, 3, 6   ' Solid, Dash, Dash-Dot, inside solid.
            sLine = String$(nLength, "-")
         Case 2, 4         ' Dot, Dash-Dot-Dot.
            sLine = Replicate(CInt(nLength), "- ")
         Case 5            ' Transparent.
            Exit Sub
         End Select

         PutText sLine, Page.Margin.Left + nLeft, Page.Margin.Top + nTop
      End If

   Else
      lp = AddLayout(LYO_LINE, Page.Margin.Left + CTwips(nLeft), Page.Margin.Top + CTwips(nTop))
      Layout(lp).Line.Color = nColor
      Layout(lp).Line.Style = nStyle

      If nDirection = LINE_VERTICAL Then
         If IsMissing(nLength) Then nLength = GetHeight - nTop    ' Until edge of printable page
         Layout(lp).Line.Width = Page.Margin.Left + CTwips(nLeft)
         Layout(lp).Line.Height = Page.Margin.Top + CTwips(nTop + nLength)
      Else
         If IsMissing(nLength) Then nLength = GetWidth - nLeft    ' Until edge of printable page
         Layout(lp).Line.Width = Page.Margin.Left + CTwips(nLeft + nLength)
         Layout(lp).Line.Height = Page.Margin.Top + CTwips(nTop)
      End If
   End If
End Sub

' Text print  * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Sub PrintPrint(sString As String, Optional bCrLf)
   If Page.Cancelled Then Exit Sub
   If IsMissing(bCrLf) Then bCrLf = True

   If Page.Ruler = RULER_CHAR Then
      If Not EmptyString(sString) Then
         ' Why store a command for no text.
         PutText sString, nCurrentX, nCurrentY
      End If

      If bCrLf Then
         ' Go to next line
         nCurrentX = Page.Margin.Left
         nCurrentY = nCurrentY + 1
      Else
         ' Shift the left position - we don't do wrap
         nCurrentX = nCurrentX + Len(sString)
      End If

   Else
      If Not EmptyString(sString) Then
         ' Why store a command for no text.
         lp = AddLayout(LYO_TEXT, nCurrentX, nCurrentY)
         Layout(lp).Text.Text = sString
      End If

      If bCrLf Then
         ' New line
         nCurrentX = Page.Margin.Left
         nCurrentY = nCurrentY + GetTextHeight(sString, vbTwips)
      Else
         ' Continue on same line, but further up
         nCurrentX = nCurrentX + GetTextWidth(sString, vbTwips)
      End If
   End If
End Sub

Sub PrintAt(ByVal nXVal As Single, ByVal nYVal As Single, sString As String, Optional bCrLf)
   If Page.Cancelled Then Exit Sub
   If IsMissing(bCrLf) Then bCrLf = True

   PrintPSet nXVal, nYVal
   PrintPrint sString, bCrLf
End Sub

Sub PrintPSet(ByVal nXVal As Single, ByVal nYVal As Single)
   If Page.Ruler = RULER_CHAR Then
      nCurrentX = Page.Margin.Left + nXVal
      nCurrentY = Page.Margin.Top + nYVal
   Else
      nCurrentX = Page.Margin.Left + CTwips(nXVal)
      nCurrentY = Page.Margin.Top + CTwips(nYVal)
   End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Used for "text only" reports only (This is used instead of AddLayout())
'
Private Sub PutText(sText As String, ByVal nPosX As Long, ByVal nPosY As Long)

   ' Do some clipping
   If nPosY < 1 Or nPosY > Page.Height Then Exit Sub
   If nPosX > Page.Width Then Exit Sub
   If nPosX < 1 Then
      If Len(sText) < (Abs(nPosX) + 2) Then Exit Sub
      sText = Mid$(sText, (Abs(nPosX) + 2))
      nPosX = 1
   End If

   ' Cut-off excess...
   If Len(sText) > (Page.Width - (nPosX - 1)) Then
      sText = Left(sText, Page.Width - (nPosX - 1))
   End If

   ' Place the text
   Mid(Layout(nPosY).Text.Text, nPosX, Len(sText)) = sText

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Control "frmMain.picPaper" is being used as a font buffer and text size measurements

' Be smart with PrintFont(). Do some on-the-run optimisation here
Sub PrintFont()
   If Page.Ruler = RULER_CHAR Then
      ' "Text only" doesn't use fonts
      Exit Sub

   ElseIf nLineCount < 1 Then                          ' No instructions recorded yet - let's be the first one
      lp = AddLayout(LYO_FONT, -1, -1)

   ElseIf Layout(nLineCount).Mode = LYO_FONT Then  ' Previous command was a font, overwrite it.
      lp = nLineCount

   Else
      ' Check it there was a previous font which was the same
      Dim i As Integer
      Dim bFound As Boolean
      bFound = False

      ' Only go one font back.
      For i = nLineCount To 1 Step -1
         If Layout(i).Mode = LYO_FONT Then
            ' Compare...
            bFound = (Layout(i).Font.Name = frmMain.picPaper.FontName And _
                      Layout(i).Font.Size = frmMain.picPaper.FontSize And _
                      Layout(i).Font.Color = frmMain.picPaper.ForeColor And _
                      Layout(i).Font.Bold = frmMain.picPaper.FontBold And _
                      Layout(i).Font.Italic = frmMain.picPaper.FontItalic And _
                      Layout(i).Font.Strikethru = frmMain.picPaper.FontStrikethru And _
                      Layout(i).Font.Underline = frmMain.picPaper.FontUnderline)
            Exit For
         End If
      Next

      If bFound Then Exit Sub                      ' Found same font - No add.

      ' No, different - so add it.
      lp = AddLayout(LYO_FONT, -1, -1)
   End If

   Layout(lp).Font.Name = frmMain.picPaper.FontName
   Layout(lp).Font.Size = frmMain.picPaper.FontSize
   Layout(lp).Font.Color = frmMain.picPaper.ForeColor
   Layout(lp).Font.Bold = frmMain.picPaper.FontBold
   Layout(lp).Font.Italic = frmMain.picPaper.FontItalic
   Layout(lp).Font.Strikethru = frmMain.picPaper.FontStrikethru
   Layout(lp).Font.Underline = frmMain.picPaper.FontUnderline
End Sub

Sub SetFontName(sFontName As String)
   If Page.Ruler = RULER_CHAR Then Exit Sub
   frmMain.picPaper.FontName = sFontName
End Sub

Sub SetFontSize(nSize As Integer)
   If Page.Ruler = RULER_CHAR Then Exit Sub
   frmMain.picPaper.FontSize = nSize
End Sub

Sub SetFontColor(nForeColor As Long)
   If Page.Ruler = RULER_CHAR Then Exit Sub
   frmMain.picPaper.ForeColor = nForeColor
End Sub

Sub SetFontBold(bFontBold As Boolean)
   If Page.Ruler = RULER_CHAR Then Exit Sub
   frmMain.picPaper.FontBold = bFontBold
End Sub

Sub SetFontItalic(bFontItalic As Boolean)
   If Page.Ruler = RULER_CHAR Then Exit Sub
   frmMain.picPaper.FontItalic = bFontItalic
End Sub

Sub SetFontStrikethru(bFontStrikethru As Boolean)
   If Page.Ruler = RULER_CHAR Then Exit Sub
   frmMain.picPaper.FontStrikethru = bFontStrikethru
End Sub

Sub SetFontUnderline(bFontUnderline As Boolean)
   If Page.Ruler = RULER_CHAR Then Exit Sub
   frmMain.picPaper.FontUnderline = bFontUnderline
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Function GetTextWidth(sText As String, Optional nScale) As Single
   If IsMissing(nScale) Then
      If Page.Ruler = RULER_CHAR Then
         nScale = vbCharacters
      Else
         nScale = vbMillimeters
      End If
   End If

   Select Case nScale
   Case vbCharacters
      GetTextWidth = Len(sText)
   Case Else
      Dim nPrevScale As Integer
      nPrevScale = frmMain.picPaper.ScaleMode
      frmMain.picPaper.ScaleMode = nScale
      GetTextWidth = frmMain.picPaper.TextWidth(sText)
      frmMain.picPaper.ScaleMode = nPrevScale
   End Select
End Function

Function GetTextHeight(Optional sText, Optional nScale) As Single
   If IsMissing(nScale) Then
      If Page.Ruler = RULER_CHAR Then
         nScale = vbCharacters
      Else
         nScale = vbMillimeters
      End If
   End If

   Select Case nScale
   Case vbCharacters
      GetTextHeight = 1
   Case Else
      If IsMissing(sText) Then
         sText = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" ' Just testing my alfabet
      End If

      Dim nPrevScale As Integer
      nPrevScale = frmMain.picPaper.ScaleMode
      frmMain.picPaper.ScaleMode = nScale
      GetTextHeight = frmMain.picPaper.TextHeight(sText)
      frmMain.picPaper.ScaleMode = nPrevScale
   End Select
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Convert into twips

Sub SetCurrentX(ByVal nXVal As Single)
   If Page.Ruler = RULER_CHAR Then
      nCurrentX = Page.Margin.Left + nXVal
   Else
      nCurrentX = Page.Margin.Left + CTwips(nXVal)
   End If
End Sub

Sub SetCurrentY(ByVal nYVal As Single)
   If Page.Ruler = RULER_CHAR Then
      nCurrentY = Page.Margin.Top + nYVal
   Else
      nCurrentY = Page.Margin.Top + CTwips(nYVal)
   End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' These position functions take margin in mind and
' converts back to millimeters or characters/lines

Function GetCurrentX() As Single
   If Page.Ruler = RULER_CHAR Then
      GetCurrentX = nCurrentX - Page.Margin.Left
   Else
      GetCurrentX = (nCurrentX - Page.Margin.Left) / 56.7
   End If
End Function

Function GetCurrentY() As Single
   If Page.Ruler = RULER_CHAR Then
      GetCurrentY = nCurrentY - Page.Margin.Top
   Else
      GetCurrentY = (nCurrentY - Page.Margin.Top) / 56.7
   End If
End Function

' Great for drawing lines across the page - just look at PrintLine()
Function GetWidth() As Single
   If Page.Ruler = RULER_CHAR Then
      GetWidth = Page.Width - (Page.Margin.Left + Page.Margin.Right)
   Else
      GetWidth = (Page.Width - (Page.Margin.Left + Page.Margin.Right)) / 56.7
   End If
End Function

Function GetHeight() As Single
   If Page.Ruler = RULER_CHAR Then
      GetHeight = Page.Height - (Page.Margin.Top + Page.Margin.Bottom)
   Else
      GetHeight = (Page.Height - (Page.Margin.Top + Page.Margin.Bottom)) / 56.7
   End If
End Function

' E.O.M
