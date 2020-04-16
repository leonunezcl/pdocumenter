Attribute VB_Name = "TextEngine"
Option Explicit

' To (text-)print call the following routines (in this order !!):
'
'  bSuccess = TextInit(picName, True/False)
'  TextSetMargins LeftMargin, RightMargin, TopMargin, BottomMargin
'  TextStartDoc
'  ... Your printing using TextAt, TextLine, etc
'  ... TextNewPage
'  ... Your printing using TextAt, TextLine, etc
'  TextEndDoc
'  ... or TextKillDoc to abort last page
'

Public Const TXLINE_HORINZONTAL As Integer = 0     ' Use "-" (dash)
Public Const TXLINE_VERTICAL As Integer = 1        ' Use "|" (pipe)

Dim nMargin As MARGINSTATE

Private Type TXFONTSTATE
   Bold As Boolean
   Italic As Boolean
   Strikethru As Boolean
   Underline As Boolean
End Type
Dim FontMem As TXFONTSTATE

Dim bSendtoPrinter As Boolean    ' Flag indicating Printing or Previewing

Dim nPgWidth As Integer           ' The actual paper size from calculated dimensions
Dim nPgHeight As Integer

Dim bAbort As Boolean            ' True when abort is requested

Dim ObjPrint As Control          ' Object used for Print Preview
Dim nPreviousObject As Long
Dim bScreenView As Boolean       ' True if screenview is possible
Dim nFontWidth As Long
Dim nFontHeight As Long

Dim PageBuffer() As String       ' The virtual page buffer for printer!!
Dim nPgCurrentX As Integer         ' Left position in characters [zero based]
Dim nPgCurrentY As Integer         ' Top position in characters [zero based]
Dim nBufferWidth As Integer      ' Number of characters in each array element (columns) [zero based]
Dim nBufferHeight As Integer     ' Number of elements in array (lines) [zero based]

Sub TextInit(bPrint As Boolean, objToPrintOn As Control)
   Dim nHeightRatio As Double, nWidthRatio As Double, nRatio As Double

   'Set the object used for preview
   Set ObjPrint = objToPrintOn

   bAbort = False

   'Set the flag that determines whether printing or previewing
   bSendtoPrinter = bPrint

   nBufferWidth = frmMain.cboWidth - 1
   nBufferHeight = frmMain.cboHeight - 1

   ' Reset the margins
   nMargin.Left = 0
   nMargin.Right = 0
   nMargin.Top = 0
   nMargin.Bottom = 0

   nPgCurrentX = 0
   nPgCurrentY = 0

   'Get the calculated page size (to be used in screen view) - Printer only uses characters
   nPgWidth = frmMain.cboWidth * 120
   nPgHeight = frmMain.cboHeight * 240

   ObjPrint.ScaleMode = vbTwips
   ObjPrint.FontName = "Courier New"
   ObjPrint.FontSize = 10
   ObjPrint.ForeColor = QBColor(0)
   ObjPrint.FontBold = False
   ObjPrint.FontItalic = False
   ObjPrint.FontStrikethru = False
   ObjPrint.FontUnderline = False

   ' Compare the height and Width ratios to determine the
   ' ratio to use and how to size the picture box font size
   nHeightRatio = ObjPrint.ScaleHeight / nPgHeight
   nWidthRatio = ObjPrint.ScaleWidth / nPgWidth

   ' Obtain smallest ratio
   If nHeightRatio < nWidthRatio Then
      nRatio = nHeightRatio
   Else
      nRatio = nWidthRatio
   End If

   ObjPrint.FontSize = ObjPrint.FontSize * nRatio
   nFontWidth = 120
   nFontHeight = 240

   ' Set default properties of picture box to match printer
   ' There are many that you could add here
   ObjPrint.Scale (0, 0)-(nPgWidth, nPgHeight)     ' Use the same scaling as the printer object !!
   
   'Initialize virtual printer page when required
   'If bSendtoPrinter Then EmptyVirtualBuffer
   EmptyVirtualBuffer

End Sub

Private Function FlushVirtualBuffer() As Boolean
   Dim i As Integer, nLast As Integer
   Dim bOpenPort As Boolean
   Dim nPortHandle As Integer       ' File handle used in Open command

   bOpenPort = False
   On Error GoTo FlushError

   ' Find the last line
   nLast = -1
   For i = nBufferHeight To 0 Step -1
      If Len(Trim(PageBuffer(i))) > 0 Then
         nLast = i
         Exit For
      End If
   Next

   If nLast < 0 Then
      ' Nothing to print... (don't waste any paper)
      EmptyVirtualBuffer            ' Just call it to reset settings
      FlushVirtualBuffer = True
      Exit Function
   End If

   If frmMain.chkFormFeed <> vbChecked Then
      ' If not using formfeed, print ALL lines (so it feeds to next page)
      nLast = nBufferHeight
   End If

   nPortHandle = FreeFile
   Open frmMain.cboPort For Output As #nPortHandle
   bOpenPort = True

   For i = 0 To nLast
      DoEvents
      If bAbort Then Exit For

      Print #nPortHandle, RTrim$(PageBuffer(i))
   Next

   If Not bAbort Then
      If frmMain.chkFormFeed = vbChecked Then
         Print #nPortHandle, vbFormFeed
      End If
   End If

   Close #nPortHandle
   FlushVirtualBuffer = True
   Exit Function

FlushError:
   If bOpenPort Then Close #nPortHandle
   MsgBox "Problems accessing printer. Please check port setting or printer.", vbExclamation, "Printing Error"
   FlushVirtualBuffer = False
End Function

Private Sub EmptyVirtualBuffer()
   Dim i As Integer

   ReDim PageBuffer(0 To nBufferHeight)

   For i = 0 To nBufferHeight
      PageBuffer(i) = Space$(nBufferWidth + 1)
   Next

   nPgCurrentX = 0
   nPgCurrentY = 0
End Sub

Private Sub PrintBuffer(sText As String, Optional nPosX, Optional nPosY, Optional bCrLf)
   If bAbort Then Exit Sub
   If IsMissing(bCrLf) Then bCrLf = True
   If IsMissing(nPosX) Then nPosX = nPgCurrentX
   If IsMissing(nPosY) Then nPosY = nPgCurrentY

   ' Do some clipping
   If nPosY < 0 Or nPosY > nBufferHeight Then Exit Sub
   If nPosX > nBufferWidth Then Exit Sub
   If nPosX < 0 Then
      If Len(sText) < (Abs(nPosX) + 1) Then Exit Sub
      sText = Mid$(sText, Abs(nPosX))
      nPosX = 1
   End If

   If Len(sText) > ((nBufferWidth + 1) - nPosX) Then
      sText = Left(sText, (nBufferWidth + 1) - nPosX)
   End If

   ' Update the screen view
   ScreenPrint nPosX, nPosY, sText

   ' Place the text
   Mid(PageBuffer(nPosY), nPosX + 1, Len(sText)) = sText

   If bCrLf Then
      ' Go to next line
      nPgCurrentX = 0
      nPgCurrentY = nPosY + 1
   Else
      ' Shift the left position - we don't do wrap
      nPgCurrentX = nPosX + Len(sText)
      nPgCurrentY = nPosY
   End If

End Sub

' Only used thru "PrintBuffer()"
Private Sub ScreenPrint(ByVal nPosX As Integer, ByVal nPosY As Integer, sText As String)
   ObjPrint.CurrentX = nPosX * nFontWidth
   ObjPrint.CurrentY = nPosY * nFontHeight

   ObjPrint.Print sText
End Sub

Sub TextSetAbort(Optional bFlag)
   If IsMissing(bFlag) Then
      bAbort = True
   Else
      bAbort = bFlag
   End If

   If bAbort Then TextResetControl
End Sub

Sub TextSetMargins(nLeft As Integer, nRight As Integer, nTop As Integer, nBottom As Integer)
   nMargin.Left = nLeft
   nMargin.Right = nRight
   nMargin.Top = nTop
   nMargin.Bottom = nBottom
End Sub

Sub TextStartDoc()
   If bAbort Then Exit Sub

   EmptyVirtualBuffer
   ObjPrint.Cls
   TextPSet 0, 0
End Sub

Sub TextNewPage()
   If bAbort Then Exit Sub
   If bSendtoPrinter Then
      ' Spool the buffer page...
      If Not FlushVirtualBuffer Then
         bAbort = True
         PrintSystemAbort
      End If
   End If
   EmptyVirtualBuffer
   ObjPrint.Cls
   TextPSet 0, 0
End Sub

Sub TextKillDoc()
   EmptyVirtualBuffer
   TextResetControl
End Sub

Sub TextEndDoc()
   Dim dl As Long
   If bSendtoPrinter Then
      dl = FlushVirtualBuffer
   End If
   EmptyVirtualBuffer
   TextResetControl
End Sub

' Box sample:
'
' 0123456789012      [zero based]
' +-----------+ 0
' |           | 1
' |           | 2
' +-----------+ 3
'
' (nMargin.Left + nLeft, nMargin.Top + nTop)-(nMargin.Left + nLeft + nWidth, nMargin.Top + nTop + nHeight)
'
Sub TxtBox(nLeft As Integer, nTop As Integer, nWidth As Integer, nHeight As Integer)
   If bAbort Then Exit Sub

   Dim sLine As String
   Dim i As Integer

   ' Top line
   Select Case nWidth
   Case Is > 1
      sLine = "+" & String$(nWidth - 1, "-") & "+"
   Case Is = 1
      sLine = "++"
   Case Else
      sLine = "+"
   End Select

   PrintBuffer sLine, nMargin.Left + nLeft, nMargin.Top + nTop

   If nHeight > 0 Then
      If nHeight > 1 Then
         For i = 1 To (nHeight - 1)
            PrintBuffer "|", nMargin.Left + nLeft, nMargin.Top + nTop + i
            If nWidth > 0 Then
               PrintBuffer "|", nMargin.Left + nLeft + nWidth, nMargin.Top + nTop + i
            End If
         Next
      End If

      PrintBuffer sLine, nMargin.Left + nLeft, nMargin.Top + nTop + nHeight
   End If

End Sub

' Valid styles:
'   0  Solid.
'   1  (Default) Transparent.
'   2  Horizontal Line.
'   3  Vertical Line.
'   4  Upward Diagonal.
'   5  Downward Diagonal.
'   6  Cross.
'   7  Diagonal Cross.
'
Sub TxtFilledBox(nLeft As Integer, nTop As Integer, nWidth As Integer, nHeight As Integer, Optional nStyle)
   If bAbort Then Exit Sub
   Dim nPrevStyle As Integer
   If IsMissing(nStyle) Then nStyle = 0   ' Solid fill

   Dim sLine As String, sBlock As String
   Dim i As Integer

   Select Case nStyle
   Case 0  ' Solid.
      sBlock = "#"
   Case 1  ' (Default) Transparent.
      TextBox nLeft, nTop, nWidth, nHeight
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

   sLine = String$(nWidth + 1, sBlock)

   For i = 0 To nHeight
      PrintBuffer sLine, nMargin.Left + nLeft, nMargin.Top + nTop + i
   Next
End Sub

' Line is always vertical or horinzontal
' Assumes horinzontal as default
'
Sub TxtLine(nLeft As Integer, nTop As Integer, nLength As Integer, Optional nDirection, Optional nStyle)
   If bAbort Then Exit Sub
   Dim nPrevStyle As Integer
   On Error Resume Next

   If IsMissing(nDirection) Then nDirection = LINE_HORINZONTAL
   If IsMissing(nStyle) Then nStyle = vbSolid

   If nDirection = LINE_VERTICAL Then        ' Going down...
      Dim i As Integer, nJump As Integer

      Select Case nStyle
      Case 0, 1, 3, 6 ' Solid, Dash, Dash-Dot, inside solid.
         nJump = 1
      Case 2, 4       ' Dot, Dash-Dot-Dot.
         nJump = 2
      Case 5          ' Transparent.
         Exit Sub
      End Select

      nTop = nMargin.Top + nTop

      For i = nTop To (nTop + nLength) Step nJump
         PrintBuffer "|", nMargin.Left + nLeft, i
      Next

   Else
      Dim sLine As String
      Select Case nStyle
      Case 0, 1, 3, 6 ' Solid, Dash, Dash-Dot, inside solid.
         sLine = String$(nLength + 1, "-")
      Case 2, 4       ' Dot, Dash-Dot-Dot.
         sLine = Left(String$(nLength + 1, "- "), nLength + 1)
      Case 5          ' Transparent.
         Exit Sub
       End Select

      PrintBuffer sLine, nMargin.Left + nLeft, nMargin.Top + nTop

   End If
End Sub

Sub TxtCurrentX(nXVal As Integer)
   nPgCurrentX = nXVal + nMargin.Left
End Sub

Sub TxtCurrentY(nYVal As Integer)
   nPgCurrentY = nYVal + nMargin.Top
End Sub

Sub TxtPrint(sString As String, Optional bCrLf)
   If bAbort Then Exit Sub
   If IsMissing(bCrLf) Then bCrLf = True

   If bCrLf Then
      PrintBuffer sString
      TextCurrentX 0
   Else
      PrintBuffer sString, , , False
   End If
End Sub

Sub TxtAt(nXVal As Integer, nYVal As Integer, sString As String, Optional bCrLf)
   If bAbort Then Exit Sub
   If IsMissing(bCrLf) Then bCrLf = True

   nPgCurrentX = nMargin.Left + nXVal
   nPgCurrentY = nMargin.Top + nYVal

   TextPrint sString, bCrLf
End Sub

Sub TxtPSet(nXVal As Integer, nYVal As Integer)
   nPgCurrentX = nMargin.Left + nXVal
 n nPgCurrentY = nMargin.Top + nYVal
End Sub

' -----------------------------------------------------------------------------
' These position functions take margin in mind

Function GetTxtCurrentX() As Integer
   GetTextCurrentX = nPgCurrentX - nMargin.Left
End Function

Function GetTxtCurrentY() As Integer
   GetTextCurrentY = nPgCurrentY - nMargin.Top
End Function

' Great for drawing lines across the page
Function GetTxtWidth() As Integer
   GetTextWidth = (nBufferWidth + 1) - (nMargin.Left + nMargin.Right)
End Function

Function GetTxtHeight() As Integer
   GetTextHeight = (nBufferHeight + 1) - (nMargin.Top + nMargin.Bottom)
End Function

' -----------------------------------------------------------------------------

Function GetTxtTextWidth(sText As String) As Integer
   GetTextTextWidth = Len(sText)
End Function

' OEM is always one character in height
Function GetTxtTextHeight(sText As String) As Integer
   GetTextTextHeight = 1
End Function

' Used for debugging
Sub TextView()
   Dim sText As String
   Dim i As Integer

   sText = "Width: " & nBufferWidth & "  Height: " & nBufferHeight & vbCrLf
   For i = 0 To nBufferHeight
      sText = sText & PageBuffer(i) & vbCrLf
   Next

   Load frmViewFile
   frmViewFile.SetText "Text print", sText
   frmViewFile.Show
End Sub
