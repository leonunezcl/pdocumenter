Attribute VB_Name = "PrintCommon"
Option Explicit

' ---------------------------
' Variables used in frmSelect (simulated Print-Commondialog)
Public Type PSelectState
   Name As String
   Port As String
   RangeAll As Boolean
   FromPage As Integer
   ToPage As Integer
   Min As Integer
   Max As Integer
   EnableCopies As Boolean
   Copies As Integer
   Collate As Boolean
   Cancelled As Boolean
End Type
Public PD As PSelectState

' ---------------------------
' Variables required for index pages
Public Type IxProcState
   Name As String
   Type As String
End Type
Public Type IxIState
   File As String
   Page As Integer
   Procedure As IxProcState      ' Only used in ProcIndex()
End Type
Public Type IndexState
   CICount As Long
   ControlIndex() As IxIState
   DIcount As Long
   DeclareIndex() As IxIState
   PIcount As Long
   ProcIndex() As IxIState
End Type
Public Idx As IndexState

' ---------------------------
' Font buffer - used in PushFont(), PopFont()
Type FONTSTATE
   Name As String
   Size As Single
   Color As Long
   Bold As Boolean
   Italic As Boolean
   Strikethru As Boolean
   Underline As Boolean
End Type
Public FontMem As FONTSTATE

' ---------------------------
' Page view/print control
Type MARGINSTATE
   Left As Long
   Right As Long
   Top As Long
   Bottom As Long
End Type
Public Type PageState
   Title As String            ' Document title
   Output As Integer          ' Output mode - see OUT_* constants
   Ruler As Integer           ' See RULER_* constants
   Width As Long              ' Page dimensions in twips
   Height As Long             ' see Width
   Count As Integer           ' Number of pages
   Cancelled As Boolean       ' True if cancel is invoked
   Show As Boolean            ' True if display form is loaded (either preview or print)
   Form As Form               ' Form object that is in control - check with Page.Show prior useage
   Margin As MARGINSTATE      ' Margin control - in twips
   File As String             ' In case "Output = OUT_RTF" this will contain the filename, with "Output = OUT_PORT" contains port name (eg LPT1 - NO colon!!)
   FormFeed As Boolean        ' Used in "text-only" mode - True if formfeed Ascii 12 is required at end of page
   Sample As Boolean          ' True if page sample is being created
   Copies As Integer          ' Must be > 0
   PageNo As Integer          ' If > 0, use this.
End Type
Public Page As PageState

Public nYAvailable As Double  ' Height available
Public nYHeight As Double     ' Max height

Public Const OUT_DRIVER As Integer = 0
Public Const OUT_RTF As Integer = 1
Public Const OUT_PORT As Integer = 2

Public Const RULER_MM As Integer = 0
Public Const RULER_CHAR As Integer = 1

' ---------------------------
' Layout of a page - element for each command
'
' Printer object methods:  (*) Store these.
'   * Circle Method
'     EndDoc Method
'     KillDoc Method
'   * Line Method
'     NewPage Method
'   * PaintPicture Method
'   * Print Method
'     PSet Method
'     Scale Method
'     ScaleX, ScaleY Methods
'     TextHeight Method
'     TextWidth Method
'
Public Type PnFontState       ' Font
   Name As String             ' Font name
   Size As Single             ' Font size
   Color As Long              ' Font colour
   Bold As Boolean            ' Font bold flag
   Italic As Boolean          ' Font italics flag
   Strikethru As Boolean      ' Font Strikethru flag
   Underline As Boolean       ' Font underline flag
End Type
Public Type PnTextState       ' Text
   Text As String
End Type
Public Type PnLineState       ' Line/Box
   Width As Long
   Height As Long
   Color As Long
   Style As Integer
End Type
Public Type PnImageState      ' Image/Picture
   'Image As String
   Index As Integer           ' Element pointer to Mdl().IconData
   Width As Long
   Height As Long
End Type
Public Type PnCircleState     ' Circle
   Radius As Long
   Color As Long
End Type
Public Type LayoutState       ' All together now...
   Mode As Integer            ' See LYO_* constants
   X As Long                  ' Left offset position
   Y As Long                  ' Top offset position
   Text As PnTextState
   Line As PnLineState
   Image As PnImageState
   Circles As PnCircleState
   Font As PnFontState
   Tabs() As Long
End Type
Public Layout() As LayoutState            ' The engine will fill it, and Preview will copy and Print will read it.
Public nLineCount As Integer              ' Number of lines in this layout

Public Const LYO_TEXT As Integer = 0      ' Type TextState
Public Const LYO_FONT As Integer = 1      ' Type FontState
Public Const LYO_LINE As Integer = 2      ' Type LineState
Public Const LYO_BOX As Integer = 3       ' Type LineState
Public Const LYO_FILLBOX As Integer = 4   ' Type LineState
Public Const LYO_IMAGE As Integer = 5     ' Type ImageState
Public Const LYO_CIRCLE As Integer = 6    ' Type CircleState
Public Const LYO_TABS As Integer = 7      ' RTF only: X will contain number of tab positions, Tabs() will contain designated tab positions in twips
Public Const LYO_TAB As Integer = 8       ' RTF only: Marks a tab character (vbTab or "\tab ")
Public Const LYO_EOL As Integer = 9       ' RTF only : End of line marker (places vbCrlf in .SelText)
Public Const LYO_PAGE As Integer = 10     ' RTF only: inserts "/page " string

' ---------------------------
Public Const PU_IDLE As Integer = 0       ' Just update progress bar and do some events
Public Const PU_NEWPAGE As Integer = 1    ' New page. Save buffer to ImageList
Public Const PU_CANCEL As Integer = 2     ' Cancel invoked from system
Public Const PU_STARTPRINT As Integer = 3 ' Indicating start of print engine
Public Const PU_ENDPRINT As Integer = 4   ' Indicating end of print engine

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Simulate page print
'   PrintStatus PU_STARTPRINT
'   PrintDialog "Formatting page " & i
'   ... Load information into Layout() - update nLineCount
'   PrintStatus PU_NEWPAGE
'   PrintStatus PU_ENDPRINT

' The print controller !!
' This procedure will determine what will be printed and who is going to do it.
' It uses options set in the main form.
'
Sub PrintControl()
'   Dim i As Integer, j As Integer, n As Integer, nIndex As Integer, nFinalIndex As Integer, nFinalProc As Integer
'   Dim nSize As Long
'   Dim sString As String, sUpper As String
'   Dim H As Single, W As Single, nPointY As Single

   Page.Sample = False

   Page.Title = "Imprimiendo " & ExtractFileName(frmMain.txtProject)
   Page.Count = 0                      ' Number of pages
   Page.PageNo = 0
   Page.Copies = 1
   Page.Cancelled = False

   ' Reset the margins
   Page.Margin.Left = 0
   Page.Margin.Right = 0
   Page.Margin.Top = 0
   Page.Margin.Bottom = 0

   If frmMain.chkPreview = vbChecked Then
      frmPreview.Show
   Else
      frmPrint.Show
   End If

   If frmMain.optOutput(0) Then
      ' Windows print driver... (uses millimeters)
      Page.Output = OUT_DRIVER
      Page.File = ""

      Page.Ruler = RULER_MM            ' Set this first prior giving any dimensions/coordinates !!
      Page.Width = Printer.Width       ' Printer.Width and .Height are twips already...
      Page.Height = Printer.Height

      Page.Margin.Left = CTwips(frmMain.lblLeft(0), True)
      Page.Margin.Right = CTwips(frmMain.lblRight(0), True)
      Page.Margin.Top = CTwips(frmMain.lblTop(0))
      Page.Margin.Bottom = CTwips(frmMain.lblBottom(0))

      PrintJob

   ElseIf frmMain.optOutput(1) Then
      ' RTF file
      Page.Output = OUT_RTF
      Page.File = frmMain.txtRTFfile

      Page.Ruler = RULER_MM            ' Set this first prior giving any dimensions/coordinates !!

      ' Leave margins to zero - deduct width and height only.
      Page.Width = Printer.Width - (CTwips(frmMain.lblLeft(0), True) + CTwips(frmMain.lblRight(0), True))
      Page.Height = Printer.Height - (CTwips(frmMain.lblTop(0)) + CTwips(frmMain.lblBottom(0)))

      PrintJob

   ElseIf frmMain.optOutput(2) Then
      ' Text direct to printer
      Page.Output = OUT_PORT
      Page.File = frmMain.cboPort.Text
      Page.FormFeed = (frmMain.chkFormFeed = vbChecked)

      Page.Ruler = RULER_CHAR          ' Set this first prior giving any dimensions/coordinates !!
      Page.Width = frmMain.cboWidth
      Page.Height = frmMain.cboHeight

      Page.Margin.Left = frmMain.lblLeft(1)
      Page.Margin.Right = frmMain.lblRight(1)
      Page.Margin.Top = frmMain.lblTop(1)
      Page.Margin.Bottom = frmMain.lblBottom(1)

      PrintJob

   ElseIf Page.Show Then    ' Just in case if form is not unloaded
      Unload Page.Form
   End If

   If frmMain.chkPreview <> vbChecked Then
      ' Print form normally doesn't wait for the user to close it - it should be done automatically
      If Page.Show Then
         Unload Page.Form     ' And this is the automation
      End If
   End If

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Parameters are millimeters or character/lines....
' (set the scaling mode first "Page.Ruler" !!!)
'
'Sub SetPageMargins(nLeft As Single, nRight As Single, nTop As Single, nBottom As Single)
'   Page.Margin.Left = CTwips(nLeft, True)
'   Page.Margin.Right = CTwips(nRight, True)
'   Page.Margin.Top = CTwips(nTop)
'   Page.Margin.Bottom = CTwips(nBottom)
'End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Sub SetPaperSize(picCtrl As Control, nScaling As Integer)
   picCtrl.Scale  ' Reset scaling mode (to twips)

   Select Case nScaling
   Case RULER_MM
      picCtrl.Height = Printer.Height
      picCtrl.Width = Printer.Width
   Case RULER_CHAR                        ' Use characters
      picCtrl.Height = Line2Twips(frmMain.cboHeight)
      picCtrl.Width = Char2Twips(frmMain.cboWidth)
   End Select
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Function CTwips(ByVal nValue As Single, Optional bCharValue) As Long
   If Page.Ruler = RULER_CHAR Then

      If IsMissing(bCharValue) Then bCharValue = False

      If bCharValue Then
         CTwips = Char2Twips(Int(nValue))
      Else
         CTwips = Line2Twips(Int(nValue))
      End If

   Else
      CTwips = MM2Twips(nValue)
   End If
End Function

Function AddLayout(nMode As Integer, ByVal XPos As Long, ByVal YPos As Long) As Integer
   nLineCount = nLineCount + 1
   ReDim Preserve Layout(1 To nLineCount)

   Layout(nLineCount).Mode = nMode
   Layout(nLineCount).X = XPos
   Layout(nLineCount).Y = YPos

   AddLayout = nLineCount
End Function

Sub LayoutMode(nMode As Integer)
   Dim lp As Integer
   lp = AddLayout(nMode, 0, 0)
End Sub

' Conversion functions * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
'  vbUser         0  User-defined: indicates that the width or height of object is set to a custom value.
'  vbTwips        1  Twip (1440 twips per logical inch; 567 twips per logical centimeter).
'  vbPoints       2  Point (72 points per logical inch; 20 twips per point).
'  vbPixels       3  Pixel (smallest unit of monitor or printer resolution).
'  vbCharacters   4  Character (horizontal = 120 twips per unit; vertical = 240 twips per unit).
'  vbInches       5  Inch. (1440 twips per inch)
'  vbMillimeters  6  Millimeter. (56.7 twips per millimeter)
'  vbCentimeters  7  Centimeter. (567 twips per centimeter)

' The height
Function Line2Twips(ByVal nLines As Integer) As Long
   Line2Twips = nLines * 240
End Function

' The width
Function Char2Twips(ByVal nCharacters As Integer) As Long
   Char2Twips = nCharacters * 120
End Function

' Millimeters being used... (567 twips per centimeter, 1440 twips per inch)
' Height and width are the same
Function MM2Twips(ByVal nMM As Double) As Long
   MM2Twips = nMM * 56.7
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' The "in-between" function for the print engine and form
'
Function PrintStatus(Optional nAction) As Boolean

   If Page.Sample Then
      If Not IsMissing(nAction) Then
        If nAction = PU_NEWPAGE Then ShowSamplePage
      'Else
      '  ShowSamplePage
      End If
      
      PrintStatus = True
      Exit Function
   ElseIf Not Page.Show Then
      ' Form is not active - nothing to do.
      PrintStatus = False
      Exit Function
   End If

   If IsMissing(nAction) Then nAction = PU_IDLE

   Select Case nAction
   Case PU_IDLE                     ' Just update progress bar and do some events
      ' Use "YAvailable" and "nYHeight" to calculate progress
      PrintProgress
   Case PU_NEWPAGE                  ' New page. Save buffer to ImageList
      Page.Form.PrintNewPage
   Case PU_CANCEL                   ' Cancel invoked from system
      Page.Form.PrintCancel
   Case PU_STARTPRINT               ' Indicating start of print engine
      Page.Form.PrintStart
   Case PU_ENDPRINT                 ' Indicating end of print engine
      Page.Form.PrintEnd
   End Select

   DoEvents
   PrintStatus = Page.Show

End Function

Function UserAbort() As Boolean
   If Page.Cancelled Then
      ' Why bother PrintStatus() if job is cancelled anyhow.
      UserAbort = True
      Exit Function
   End If

   If Not PrintStatus Then Page.Cancelled = True
   UserAbort = Page.Cancelled
End Function

Sub PrintDialog(sText As String)
   If Page.Sample Then Exit Sub
   If Page.Show Then Page.Form.ShowDialog sText
End Sub

' Just updates progress-bar - No user interaction check (eg DoEvents)
Sub PrintProgress()
   If Page.Sample Then Exit Sub
   If Page.Show And nYAvailable >= 0 And nYHeight > 0 Then
      Page.Form.ShowProgress (((nYHeight - nYAvailable) * 100) / nYHeight)
   End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
' Remember current font setting
Sub PushFont()
   If Page.Output = OUT_PORT Then Exit Sub
   On Error Resume Next
   FontMem.Name = frmMain.picPaper.FontName
   FontMem.Size = frmMain.picPaper.FONTSIZE
   FontMem.Color = frmMain.picPaper.ForeColor
   FontMem.Bold = frmMain.picPaper.FontBold
   FontMem.Italic = frmMain.picPaper.FontItalic
   FontMem.Strikethru = frmMain.picPaper.FontStrikethru
   FontMem.Underline = frmMain.picPaper.FontUnderline
End Sub

' Restore memorised font setting
Sub PopFont()
   If Page.Output = OUT_PORT Then Exit Sub
   On Error Resume Next
   frmMain.picPaper.FontName = FontMem.Name
   frmMain.picPaper.FONTSIZE = FontMem.Size
   frmMain.picPaper.ForeColor = FontMem.Color
   frmMain.picPaper.FontBold = FontMem.Bold
   frmMain.picPaper.FontItalic = FontMem.Italic
   frmMain.picPaper.FontStrikethru = FontMem.Strikethru
   frmMain.picPaper.FontUnderline = FontMem.Underline
End Sub

Sub CloneFont(ByRef ctrlIn As Control)
   If Page.Output = OUT_PORT Then Exit Sub
   On Error Resume Next
   ctrlIn.FontName = frmMain.picPaper.FontName
   ctrlIn.FONTSIZE = frmMain.picPaper.FONTSIZE
   ctrlIn.ForeColor = frmMain.picPaper.ForeColor
   ctrlIn.FontBold = frmMain.picPaper.FontBold
   ctrlIn.FontItalic = frmMain.picPaper.FontItalic
   ctrlIn.FontStrikethru = frmMain.picPaper.FontStrikethru
   ctrlIn.FontUnderline = frmMain.picPaper.FontUnderline
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
' Test layout (coordinates) of text printing
Sub TestTextPrint()
   If Page.Show Then
      MsgBox "No puedo imprimir pagina de prueba. Hay un trabajo activo.", vbExclamation
      Exit Sub
   End If

   Dim sString As String
   Dim i As Integer, j As Integer, _
       nWidth As Integer, nHeight As Integer, _
       n As Integer, nMax As Integer, _
       nLeft As Integer, nRight As Integer, nTop As Integer, nBottom As Integer

   On Error GoTo TestAbort

   Page.Title = "Pagina solo texto"
   Page.Count = 0                      ' Number of pages
   Page.Cancelled = False

   Page.Show = False                   ' Set by frmPreview or frmPrint
   Set Page.Form = Nothing

   If frmMain.chkPreview = vbChecked Then
      frmPreview.Show
   Else
      frmPrint.Show
   End If

   Page.Output = OUT_PORT
   Page.File = frmMain.cboPort.Text

   Page.Ruler = RULER_CHAR          ' Set this first prior giving any dimensions/coordinates !!
   Page.Width = frmMain.cboWidth
   Page.Height = frmMain.cboHeight

   ' Activate the "margin", so the contents area can be shown.
   Page.Margin.Left = frmMain.lblLeft(1)
   Page.Margin.Right = frmMain.lblRight(1)
   Page.Margin.Top = frmMain.lblTop(1)
   Page.Margin.Bottom = frmMain.lblBottom(1)
   
   nWidth = GetWidth
   nHeight = GetHeight

   nYHeight = Page.Height ' + GetHeight
   nYAvailable = nYHeight
   n = 1

   PrintStartDoc

   If UserAbort Then GoTo TestAbort

   If Page.Show Then PrintDialog "Formateando pagina de prueba"

   ' The "margin"
   
   PrintBox 1, 1, nWidth, nHeight
   
'   sString = "+" & String$(nWidth - 2, "-") & "+"

'   PrintAt 1, 1, sString                        ' Top line
'   For i = 2 To (nHeight - 1)                   ' Body
'      n = n + 1
'      nYAvailable = nYAvailable - n
'      If UserAbort Then Exit For
'      PrintAt 1, i, "|"
'      PrintAt nWidth, i, "|"
'   Next
'   If Page.Cancelled Then GoTo TestAbort
'   PrintAt 1, nHeight, sString                   ' Bottom line

   If GetHeight > 6 And GetWidth > 16 Then
      PrintAt 2, 2, "       Alto: " & frmMain.cboHeight
      PrintAt 2, 3, "      ancho: " & frmMain.cboWidth
      PrintAt 2, 4, "     Arriba: " & frmMain.lblTop(1)
      PrintAt 2, 5, "      abajo: " & frmMain.lblBottom(1)
      PrintAt 2, 6, "        izq: " & frmMain.lblLeft(1)
      PrintAt 2, 7, "    derecha: " & frmMain.lblRight(1)
   End If

   ' Remove margin.
   Page.Margin.Left = 0
   Page.Margin.Right = 0
   Page.Margin.Top = 0
   Page.Margin.Bottom = 0

   nWidth = GetWidth
   nHeight = GetHeight

   ' Page line and columns numbers
   If nWidth > 10 Then
      j = Int(nWidth / 10)
      sString = ""
      For i = 1 To j
         sString = sString & "123456789|"
      Next
      i = nWidth - (j * 10)
      If i > 0 Then sString = sString & Left$("123456789|", i)
   Else
      sString = Left$("123456789|", nWidth)
   End If

   j = 2
   PrintAt 1, 1, sString                        ' Top line
   For i = 2 To (nHeight - 1)
      n = n + 1
      nYAvailable = nYAvailable - n
      If UserAbort Then Exit For

      If j = 10 Then
         PrintAt 1, i, "-"
         j = 1
      Else
         PrintAt 1, i, Trim$(Str$(j))
         j = j + 1
      End If
   Next
   If Page.Cancelled Then GoTo TestAbort
   PrintAt 1, nHeight, sString

   PrintNewPage

   PrintEndDoc

   On Error Resume Next
   GoTo TestFinish

TestAbort:
   On Error Resume Next
   PrintKillDoc                             ' Job is aborted - how about killing the print buffer

TestFinish:
   If frmMain.chkPreview <> vbChecked Then
      ' Print form normally doesn't wait for the user to close it - it should be done automatically
      If Page.Show Then
         Unload Page.Form     ' And this is the automation
      End If
   End If
End Sub
