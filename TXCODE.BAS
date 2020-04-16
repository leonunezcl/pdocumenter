Attribute VB_Name = "TextCode"
Option Explicit

' The Print Module !! (Text only)
'
' Pages (in order):  1) Controls (if applicable)          .chkControlNames
' (per module)       2) Declaration/Procedures            .chkCode
'
'
' Additional pages:  a) Project Information               .chkProject
'                    b) Index                             .chkIndex
'
' Printer order:     a, 1, 2, c
'
' This looks like the exact copy of PnCode.bas, and that is true. But this
' module only handles "pure" text (no graphics) and used characters as coordinates, not millimetres.
' I could use a flag, but this is more simpler (for me at least).
'
' -----------------------------------------------------
' See PnCommon.bas for support procedures and variables
' -----------------------------------------------------

Dim nPage As Integer             ' Page number
Dim nPgPrinted As Integer        ' Number of pages printer (to show on main form)
Dim nMdlIndex As Integer         ' Let the module know which file is being processed (using array element number)

Dim nAvailable As Integer         ' Available printable area on page
Dim bNextPage As Boolean         ' Next line should go to next page
Dim bFinalPage As Boolean        ' True if final page is done (used for preview)
Dim bPreview As Boolean          ' True - Preview mode, False - Print to printer
Dim WrapText() As String         ' Used for the wrapping procedures
Dim nWrapLines As Integer        ' Number of lines in wrapped text

Sub PrintTextJob()
   Dim i As Integer, j As Integer, n As Integer, nIndex As Integer, nFinalIndex As Integer, nFinalProc As Integer
   Dim nSize As Long
   Dim sString As String, sUpper As String
   Dim H As Integer, W As Integer, nPointY As Integer
   
   Page.Cancelled = False     ' Either one of them is set to true - Abort has priority.
   bPrintProceed = False   ' (see above)

   bPreview = (frmMain.chkPreview = vbChecked)

   If bPreview Then
      Load frmPreview
      frmPreview.SetPaperCharSize
      TextInit PRINT_PREVIEW, frmPreview.picPaper
      frmPreview.Show
   Else
      Load frmPrint
      TextInit PRINT_PRINTER, frmPrint.picPaper
      frmPrint.Show
   End If
   TextSetMargins frmMain.lblLeft(1), frmMain.lblRight(1), frmMain.lblTop(1), frmMain.lblBottom(1)

   bFinalPage = False
   nPage = 0
   nPgPrinted = 0
   frmMain.lblPrinted = "none"

'   Pages.Range = (frmMain.chkRange = vbChecked)
'   Pages.From = Val(frmMain.txtFromPage)
'   Pages.To = Val(frmMain.txtToPage)

'   If Pages.From > Pages.To Then
'      n = Pages.From
'      Pages.From = Pages.To
'      Pages.To = n
'   End If

   Erase Idx.ControlIndex
   Erase Idx.DeclareIndex
   Erase Idx.ProcIndex
   Idx.CICount = -1
   Idx.DIcount = -1
   Idx.PIcount = -1

   ' Find the last module array element selected
   nFinalIndex = MdCount
   For i = MdCount To 1 Step -1
      If Mdl(i).Selected <> vbUnchecked Then
         nFinalIndex = i
         Exit For
      End If
   Next

   TextStartDoc

   ' General project information ----------------------------------------------------------------
   
   If frmMain.chkProject = vbChecked Then
      If frmMain.chkResetPage = vbChecked Then nPage = 0
      PrintProjectPage
      If Page.Cancelled Then GoTo PrintAbort
   End If
         
   ' No form icons in text mode -----------------------------------------------------------------
   
   ' Procedures (code) --------------------------------------------------------------------------

   If Not InDevelopmentMode Then
      On Error GoTo PrintError
   End If

   For i = 1 To MdCount

      If UserAbort Then Exit For

      If Mdl(i).Selected <> vbUnchecked Then

         ' Reset pagenumber if user wants it.
         If frmMain.chkResetPage = vbChecked Then nPage = 0

         ' Tell the rest of module which module is being processed
         nMdlIndex = i

         ' Tell the user that something is going on.
         If Not Page.Cancelled Then
            If bPreview Then
               frmPreview.ShowProgress 0
               frmPreview.Refresh
            Else
               frmPrint.ShowJob "Printing " & Mdl(nMdlIndex).File
               frmPrint.ShowPageNumber nPage
               frmPrint.ShowProgress 0
               frmPrint.Refresh
            End If
         End If

         ' Does the user wants to abort?
         If UserAbort Then Exit For

         ' Initialise
         nAvailable = -1                     ' Force header to print on first line to be printed
         bNextPage = True                    ' Force new page

         ' No form icon in text mode -------------------------------------------------------------

         ' Controls ------------------------------------------------------------------------------
         If frmMain.chkControlNames = vbChecked And Mdl(nMdlIndex).CtrlSelect Then
            PrintFormControls
         End If

         If UserAbort Then Exit For

         ' Declaration/Procedures (Code) ---------------------------------------------------------
         If frmMain.chkCode = vbChecked And Mdl(nMdlIndex).ProcCount > 0 Then

            nFinalProc = Mdl(nMdlIndex).ProcCount
            For n = Mdl(nMdlIndex).ProcCount To 1 Step -1
               If Mdl(nMdlIndex).Proc(n).Selected = vbChecked Then
                  nFinalProc = n
                  Exit For
               End If
            Next n

            CheckAreaPrint

            For n = 1 To Mdl(nMdlIndex).ProcCount

               If PrintStatus((n * 100) / Mdl(nMdlIndex).ProcCount) Then Exit For

               If Mdl(nMdlIndex).Proc(n).Selected = vbChecked Then

                  If frmMain.chkProcNames = vbChecked Then                 ' Procedure names only...

                     If Mdl(nMdlIndex).Proc(n).Type <> PT_DECLARE Then     ' Declarations not allowed

                        LinePrint Mdl(nMdlIndex).Proc(n).Syntax

                        If frmMain.chkIndex = vbChecked Then               ' INDEX - Update index-page reference
                           Idx.PIcount = Idx.PIcount + 1
                           ReDim Preserve Idx.ProcIndex(0 To Idx.PIcount)
                           Idx.ProcIndex(Idx.PIcount).File = Mdl(nMdlIndex).File
                           Idx.ProcIndex(Idx.PIcount).Page = nPage
                           Idx.ProcIndex(Idx.PIcount).Procedure.Name = Mdl(nMdlIndex).Proc(n).IndexName
                           Idx.ProcIndex(Idx.PIcount).Procedure.Type = ProcType(nMdlIndex, n)
                        End If

                        If n < nFinalProc Then
                           If frmMain.chkProcPage = vbChecked Then
                              bNextPage = True
                           ElseIf frmMain.chkSeparator = vbChecked Then
                              SeperatorPrint
                           End If
                        End If
                     End If

                  Else                                                     ' All code
                     CheckAreaPrint

                     If Mdl(nMdlIndex).Proc(n).Type = PT_DECLARE Then
                        LinePrint "(Declarations)"
                        LinePrint ""

                        If frmMain.chkIndex = vbChecked Then               ' INDEX - Update index-page reference
                           Idx.DIcount = Idx.DIcount + 1
                           ReDim Preserve Idx.DeclareIndex(0 To Idx.DIcount)
                           Idx.DeclareIndex(Idx.DIcount).File = Mdl(nMdlIndex).File
                           Idx.DeclareIndex(Idx.DIcount).Page = nPage
                        End If
                     End If

                     For j = 1 To Mdl(nMdlIndex).Proc(n).Lines

                        sString = Mdl(nMdlIndex).Proc(n).Code(j)

                        If IsProcedure(UCase$(Trim$(sString))) Then           ' Only happens once (I hope) - and not in declaration section
                           LinePrint sString

                           If frmMain.chkIndex = vbChecked Then               ' INDEX - Update index-page reference
                              Idx.PIcount = Idx.PIcount + 1
                              ReDim Preserve Idx.ProcIndex(0 To Idx.PIcount)
                              Idx.ProcIndex(Idx.PIcount).File = Mdl(nMdlIndex).File
                              Idx.ProcIndex(Idx.PIcount).Page = nPage
                              Idx.ProcIndex(Idx.PIcount).Procedure.Name = Mdl(nMdlIndex).Proc(n).IndexName
                              Idx.ProcIndex(Idx.PIcount).Procedure.Type = ProcType(nMdlIndex, n)
                           End If

                        Else                                                  ' Just some code or space
                           LinePrint sString
                        End If

                     Next j   ' Code lines

                     If n < nFinalProc Then
                        If frmMain.chkProcPage = vbChecked Then
                           bNextPage = True
                        ElseIf frmMain.chkSeparator = vbChecked Then
                           SeperatorPrint
                        End If
                     End If

                  End If   ' frmMain.chkProcNames = vbChecked
               End If      ' Mdl(nMdlIndex).Proc(n).Selected = vbChecked
            Next n         ' Procedures
         End If

         If UserAbort Then Exit For

         ' ----------------------------------------------------------------------------------------

         If PrintStatus(100) Then Exit For

         If i = nFinalIndex And frmMain.chkIndex = vbUnchecked Then bFinalPage = True
         If nAvailable > -1 Then FooterPrint

      End If   ' Mdl(i).Selected <> vbUnchecked
   Next i

   If Page.Cancelled Then GoTo PrintAbort

   ' Index page(s) ------------------------------------------------------------------------------

   If frmMain.chkIndex = vbChecked Then
      ' Print the index page(s)
      If frmMain.chkResetPage = vbChecked Then nPage = 0
      PrintIndexPage
      If Page.Cancelled Then GoTo PrintAbort
   End If

   ' --------------------------------------------------------------------------------------------

   On Error Resume Next
   Erase Idx.ControlIndex        ' Regain resources
   Erase Idx.DeclareIndex
   Erase Idx.ProcIndex

   TextEndDoc

   If Not Page.Cancelled Then
      If bPreview Then
         Unload frmPreview
      Else
         Unload frmPrint
      End If
   End If
   Exit Sub

PrintError:
   MsgBox "Encounted a printing problem." & vbCrLf & "Error #" & Err.Number & ": " & Err.Description, vbCritical, "Print Error"

PrintAbort:
   On Error Resume Next

   TextKillDoc                   ' Job is aborted - how about killing the print buffer

   Erase Idx.ControlIndex
   Erase Idx.DeclareIndex
   Erase Idx.ProcIndex

   If Not Page.Cancelled Then
      ' Display form is not unloaded yet - do it now
      If bPreview Then
         Unload frmPreview
      Else
         Unload frmPrint
      End If
   End If
End Sub

' --------------------------------------------------------------------------------------------------------
' Test layout (coordinates) of text printing
Sub TextTestPrint()
   Dim sString As String
   Dim i As Integer, j As Integer, _
       nWidth As Integer, nHeight As Integer, _
       n As Integer, nMax As Integer, _
       nLeft As Integer, nRight As Integer, nTop As Integer, nBottom As Integer

   Page.Cancelled = False     ' Either one of them is set to true - Abort has priority.
   bPrintProceed = False   ' (see above)
   bPreview = False
   nPage = 0

   nLeft = Val(frmMain.lblLeft(1))
   nRight = Val(frmMain.lblRight(1))
   nTop = Val(frmMain.lblTop(1))
   nBottom = Val(frmMain.lblBottom(1))

   Load frmPrint
   TextInit PRINT_PRINTER, frmPrint.picPaper
   frmPrint.Show

'   Load frmPreview
'   frmPreview.SetPaperCharSize
'   TextInit PRINT_PREVIEW, frmPreview.picPaper
'   frmPreview.Show

   On Error GoTo TestAbort

   frmPrint.ShowJob "Creating test page layout"
   frmPrint.ShowPageNumber 0
   frmPrint.ShowProgress 0
   frmPrint.Refresh

   If UserAbort Then GoTo TestAbort

   TextSetMargins 0, 0, 0, 0
   nMax = (GetTxtHeight - 2)
   n = 0

   TextStartDoc

   ' Activate the "margin", so the contents area can be shown.
   TextSetMargins nLeft, nRight, nTop, nBottom
   nWidth = GetTxtWidth
   nHeight = GetTxtHeight

   nMax = nMax + (nHeight - 2)

   ' The "margin"
   sString = "+" & String$(nWidth - 2, "-") & "+"

   TextAt 0, 0, sString                        ' Top line
   For i = 1 To (nHeight - 2)                  ' Body
      n = n + 1
      If PrintStatus((n * 100) / nMax) Then Exit For
      TextAt 0, i, "|"
      TextAt (nWidth - 1), i, "|"
   Next
   If Page.Cancelled Then GoTo TestAbort
   TextAt 0, (nHeight - 1), sString         ' Bottom line

   ' Remove margin.
   TextSetMargins 0, 0, 0, 0
   nWidth = GetTxtWidth
   nHeight = GetTxtHeight

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
   TextAt 0, 0, sString                                              ' Top line
   For i = 1 To (nHeight - 2)
      n = n + 1
      If PrintStatus((n * 100) / nMax) Then Exit For

      If j = 10 Then
         TextAt 0, i, "-"
         j = 1
      Else
         TextAt 0, i, Trim$(Str$(j))
         j = j + 1
      End If
   Next
   If Page.Cancelled Then GoTo TestAbort
   TextAt 0, nHeight - 1, sString

   i = nHeight - (nTop + nBottom + 2)
   j = nWidth - (nRight + nLeft + 2)
   If i > 5 And j > 15 Then
      TextAt (nLeft + 2), (nTop + 1), "Page height: " & nHeight
      TextAt (nLeft + 2), (nTop + 2), "      width: " & nWidth
      TextAt (nLeft + 2), (nTop + 3), " Margin top: " & nTop
      TextAt (nLeft + 2), (nTop + 4), "     bottom: " & nBottom
      TextAt (nLeft + 2), (nTop + 5), "       left: " & nLeft
      TextAt (nLeft + 2), (nTop + 6), "      right: " & nRight
   End If

   frmPrint.ShowJob "Printing test page"
   frmPrint.ShowProgress 100
   frmPrint.Refresh

   TextEndDoc

'SetCaption frmPreview.cmdCancel, "Close"
'SetEnabled frmPreview.cmdNextPage, False
'Do While True
'   DoEvents
'   If Page.Cancelled Or bPrintProceed Then Exit Do
'Loop

   On Error Resume Next
   Unload frmPrint
'Unload frmPreview
   Exit Sub

TestAbort:
   On Error Resume Next
   TextKillDoc                             ' Job is aborted - how about killing the print buffer
   If Not Page.Cancelled Then Unload frmPrint 'frmPreview
End Sub

' --------------------------------------------------------------------------------------------------------

' This routine is a SOB...
Private Sub LinePrint(sString As String, Optional bCrLf)
   Dim nLines As Integer, i As Integer
   Dim nTextHeight As Integer
   Dim sText As String
   Dim WrapLine() As String

   ' Remove trailing spaces
   sString = RTrim(sString)

   If IsMissing(bCrLf) Then bCrLf = True

   ' Place line into textbox - frmMain.txtWrap
   SetWrapObject sString
   SaveWrap WrapLine, nLines

   ' No forced pagebreak - check if there's enough room for the this string
   If Not bNextPage Then
      ' If the page height is smaller than then string height just "page wrap" it.
      If GetTxtHeight > nLines Then
         ' It will fit on one page - but is there still enough room for it?
         If nAvailable < nLines Then
            ' No - force page break
            bNextPage = True
         End If
      End If
   End If

   For i = 1 To nLines

      sText = WrapLine(i)

      If Not bNextPage Then
         If nAvailable < 1 Then
            ' There's not enough page height to accomodate this text - go to next page
            bNextPage = True
         End If
      End If

      If bNextPage Then
         If Len(sText) = 0 Then GoTo EndOfLinePrint ' Do not allow empty lines in top of page
         If nAvailable > -1 Then                      ' Footer not printed yet...
            FooterPrint
            If Page.Cancelled Then Exit For
         End If
         HeaderPrint                                  ' Now print the header
         If Page.Cancelled Then Exit For
      End If

      ' Finally print the string
      If i > 1 Then
         TextPrint ">> " & sText, bCrLf
      Else
         TextPrint sText, bCrLf
      End If

      If bCrLf Then nAvailable = nAvailable - 1

EndOfLinePrint:

   Next

End Sub

' -----------------------------------------------------------------------------------
' Use characters as measurement.
Private Sub SetWrapObject(ByVal sText As String, Optional nLength)

   If IsMissing(nLength) Then
      nLength = GetTxtWidth
   Else
      nLength = CInt(nLength)
   End If

   sText = RTrim$(sText)

   nWrapLines = 1
   ReDim WrapText(1 To 1)
   WrapText(1) = ""

   If Len(sText) <= nLength Then
      ' No wrap required... great!
      WrapText(1) = sText
      Exit Sub
   End If

   ' Text must be wrapped (or warped, like my mind)... oh, no..
   Dim nSize As Integer, nMark As Integer, nSymbol As Integer
   Dim sNextLine As String, sMarker As String

'   nLength = nLength + 1
   nSymbol = 0
   sMarker = " "

   Do
      nSize = Len(sNextLine)
      nMark = InStr(sText, sMarker)

      If nMark Then
         If nSize + nMark <= nLength Then
            sNextLine = sNextLine & Left$(sText, nMark)
            sText = Mid$(sText, nMark + 1)
         ElseIf nMark > nLength Then
            'sBuffer = sBuffer & vbCrLf & Left$(sText, nLength)
            nWrapLines = nWrapLines + 1
            ReDim Preserve WrapText(1 To nWrapLines)
            WrapText(nWrapLines) = Left$(sText, nLength)
            sText = Mid$(sText, nLength + 1)
         Else
            'sBuffer = sBuffer & sNextLine & vbCrLf
            WrapText(nWrapLines) = WrapText(nWrapLines) & sNextLine
            nWrapLines = nWrapLines + 1
            ReDim Preserve WrapText(1 To nWrapLines)
            WrapText(nWrapLines) = ""
            sNextLine = ""
         End If
      Else
         If Len(sText) > nLength Then
            If nSymbol < 4 Then nSymbol = nSymbol + 1

            Select Case nSymbol
            Case 1
               sMarker = ","
            Case 2
               sMarker = "="
            Case 3
               sMarker = "\"
            Case 4
               sMarker = "("
            Case Else
               sMarker = " "
               ' We got a problem: No marker-character to wrap and text is too long - Cut if off.

               If nSize Then
                  WrapText(nWrapLines) = WrapText(nWrapLines) & sNextLine
                  sNextLine = ""
                  nWrapLines = nWrapLines + 1
                  ReDim Preserve WrapText(1 To nWrapLines)
                  WrapText(nWrapLines) = ""
               End If

               sNextLine = Left(sText, nLength)
               sText = Mid$(sText, nLength + 1)

               If Len(WrapText(nWrapLines) & sNextLine) > nLength Then
                  nWrapLines = nWrapLines + 1
                  ReDim Preserve WrapText(1 To nWrapLines)
               End If
               WrapText(nWrapLines) = sNextLine
               sNextLine = ""
            End Select

         ElseIf nSize Then
            If nSize + Len(sText) > nLength Then
               'sBuffer = sBuffer & sNextLine & vbCrLf & sText & vbCrLf
               WrapText(nWrapLines) = WrapText(nWrapLines) & sNextLine
               nWrapLines = nWrapLines + 1
               ReDim Preserve WrapText(1 To nWrapLines)
               WrapText(nWrapLines) = sText
            Else
               'sBuffer = sBuffer & sNextLine & sText & vbCrLf
               WrapText(nWrapLines) = WrapText(nWrapLines) & sNextLine & sText
            End If
            Exit Do
         Else
            'sBuffer = sBuffer & sText & vbCrLf
            If Len(WrapText(nWrapLines) & sText) > nLength Then
               nWrapLines = nWrapLines + 1
               ReDim Preserve WrapText(1 To nWrapLines)
               WrapText(nWrapLines) = ""
            End If
            WrapText(nWrapLines) = WrapText(nWrapLines) & sText
            Exit Do
         End If
      End If

   Loop

   If nWrapLines > 1 Then
      ' Subtract any empty lines in the bottom
      For nMark = nWrapLines To 2 Step -1
         If Not EmptyString(WrapText(nMark)) Then
            nWrapLines = nMark
            Exit For
         End If
      Next
   End If

End Sub

Private Function GetWrapText(nGetLine As Integer) As String
   If nWrapLines = 0 Or nGetLine < 1 Or nGetLine > nWrapLines Then
      GetWrapText = ""
      Exit Function
   End If
   GetWrapText = WrapText(nGetLine)
End Function

Private Sub SaveWrap(ByRef TextHolder, ByRef TextLines)
   Dim i As Integer
   TextLines = nWrapLines
   ReDim TextHolder(1 To TextLines)
   For i = 1 To TextLines
      TextHolder(i) = RTrim(GetWrapText(i))
   Next
End Sub

' -----------------------------------------------------------------------------------
' Check if there's enough room to print the sub line with some code (at least 2 lines of code)
'
Private Sub CheckAreaPrint(Optional nExtra)
   Dim H As Integer
   H = 1
   If Not IsMissing(nExtra) Then H = H + nExtra
   If nAvailable < H Then bNextPage = True
End Sub

Private Sub SeperatorPrint(Optional nOffset, Optional nStyle)

   CheckAreaPrint                ' Prevent line at bottom of page without any text below it.

   If bNextPage Then Exit Sub
   If IsMissing(nOffset) Then nOffset = 0
   If IsMissing(nStyle) Then nStyle = vbSolid

   Dim nCurY As Integer
   nCurY = GetTextCurrentY
   
   TextLine CInt(nOffset), nCurY, (GetTxtWidth - (CInt(nOffset) + 1)), , CInt(nStyle)
   TextPSet 0, (nCurY + 1)
   nAvailable = nAvailable - 1
End Sub

Private Sub HeaderPrint()
   Dim nCurX As Integer
   Dim sText As String

   nCurX = GetTextCurrentX

   nPage = nPage + 1
   If Not bPreview And Not Page.Cancelled Then frmPrint.ShowPageNumber nPage

   If frmMain.chkHeader = vbChecked Then

      If nMdlIndex = -1 Then
         sText = ExtractFileName(frmMain.txtProject) & " (Project)"
      ElseIf nMdlIndex = -2 Then
         sText = "Index"
      ElseIf nMdlIndex = -3 Then
         sText = "Icons"
      Else
         sText = Mdl(nMdlIndex).File
         If Mdl(nMdlIndex).Type = MT_MODULE Then
            sText = sText & " (Module - " & Mdl(nMdlIndex).Name & ")"
         ElseIf Mdl(nMdlIndex).Type = MT_CLASS Then
            sText = sText & " (Class - " & Mdl(nMdlIndex).Name & ")"
         Else
            sText = sText & " (Form - " & Mdl(nMdlIndex).Name & ")"
         End If
      End If

      ' Some header text
      If frmMain.optPagePos(0) And frmMain.chkPageNumbers = vbChecked Then
         SetWrapObject sText, (GetTxtWidth - Len("Page " & nPage))
      Else
         SetWrapObject sText
      End If
      sText = RTrim(GetWrapText(1))
      TextPSet 0, 0
      TextPrint sText

      ' Page number?
      If frmMain.optPagePos(0) And frmMain.chkPageNumbers = vbChecked Then
         TextAt (GetTxtWidth - Len("Page " & nPage)), 0, "Page " & nPage
      End If

      ' Print the line
      TextLine 0, 1, GetTxtWidth

      TextPSet 0, 2
      nAvailable = GetTxtHeight - (2 + IIf(frmMain.chkfooter = vbChecked, 3, 0))

   Else
      ' No header...
      TextPSet 0, 0
      nAvailable = GetTxtHeight
   End If

   bNextPage = False

   TextCurrentX nCurX
End Sub

Private Sub FooterPrint()
   If Page.Cancelled Then Exit Sub

   Dim sText As String
   Dim nCurX As Integer, nSize As Integer

   nCurX = GetTextCurrentX

   If frmMain.chkfooter = vbChecked Then
      ' Print the line
      TextLine 0, GetTxtHeight - 3, GetTxtWidth

      sText = ""
      If frmMain.chkDate = vbChecked Then
         sText = Format(Now, "Medium Date")
         If frmMain.chkTime = vbChecked Then
            sText = sText & " - " & Format(Now, "Medium Time")
         End If
      ElseIf frmMain.chkTime = vbChecked Then
         sText = Format(Now, "Medium Time")
      End If

      If Len(Trim(sText)) > 0 Then
         nSize = GetTxtWidth - (Len(sText) + 1)
         SetWrapObject frmMain.txtOwner(0), nSize
         sText = Pad(GetWrapText(1), nSize) & " " & sText
      Else
         SetWrapObject frmMain.txtOwner(0)
         sText = GetWrapText(1)
      End If
      TextAt 0, GetTxtHeight - 2, sText

      If frmMain.optPagePos(1) And frmMain.chkPageNumbers = vbChecked Then
         nSize = GetTxtWidth - (Len("Page " & nPage) + 1)
         SetWrapObject frmMain.txtOwner(1), nSize
         sText = Pad(GetWrapText(1), nSize) & " Page " & nPage
      Else
         SetWrapObject frmMain.txtOwner(1)
         sText = GetWrapText(1)
      End If
      TextAt 0, GetTxtHeight - 1, sText

'TextView ' Debug procedure

   End If

   nPgPrinted = nPgPrinted + 1
   frmMain.lblPrinted = nPgPrinted

   ' Once footer is requested, do not let any printing occur on this page
   nAvailable = -1

   If bPreview And Not Page.Cancelled Then

      If bFinalPage Then
         SetCaption frmPreview.cmdCancel, "Close"
         SetEnabled frmPreview.cmdNextPage, False
      Else
         SetEnabled frmPreview.cmdNextPage, True
      End If
      bPrintProceed = False
      
      Do While True
         DoEvents
         If Page.Cancelled Or bPrintProceed Then Exit Do
      Loop

      If Page.Cancelled Then
         UserAbort
      Else
         SetEnabled frmPreview.cmdNextPage, False
      End If
   End If
   
   If Not Page.Cancelled Then TextNewPage

   TextCurrentX nCurX
End Sub

' --------------------------------------------------------------------------------------------------------

Private Sub PrintProjectPage()
   Dim i As Integer, n As Integer, nMax As Integer, _
       nFileLines As Integer, nNameLines As Integer, _
       nTextOffset As Integer, nNameOffset As Integer, _
       nTextLength As Integer, nNameLength As Integer
   Dim FileWrap() As String, NameWrap() As String
   Dim sString As String
   Dim Pj As ProjectState

   If Page.Cancelled Then Exit Sub

   If bPreview Then
      SetEnabled frmPreview.cmdCancel, False
      Pj = AnalyseVBP(frmMain.txtProject, frmPreview)
      SetEnabled frmPreview.cmdCancel, True
      If Not Pj.Loaded Then Exit Sub

      frmPreview.ShowProgress 0

   Else
      frmPrint.ShowJob "Analysing " & ExtractFileName(frmMain.txtProject)
      frmPrint.ShowPageNumber nPage

      SetEnabled frmPrint.cmdCancel, False
      Pj = AnalyseVBP(frmMain.txtProject, frmPrint)
      SetEnabled frmPrint.cmdCancel, True
      If Not Pj.Loaded Then Exit Sub

      frmPrint.ShowJob "Printing " & ExtractFileName(frmMain.txtProject)
      frmPrint.ShowPageNumber nPage
      frmPrint.ShowProgress 0
   End If

   ' ----------------------------------------------------------------------------------

   If Not InDevelopmentMode Then
      On Error GoTo ProjectPrintError
   End If

   nMdlIndex = -1                      ' It's a project.

   nAvailable = -1                     ' Force header to print on first line to be printed
   bNextPage = True                    ' Force new page

   ' Print the information...

   ' No application icon in text mode.

   LinePrint "General Project Information"
   LinePrint ""
   If PrintStatus(1) Then GoTo ProjectPrintAbort

   ' Obtain the largest title width
   nNameOffset = 0
   nTextOffset = Len("Application Description: ")
   nTextLength = GetTxtWidth - nTextOffset

   TextCurrentX nNameOffset: LinePrint "VBP Filename:", False
   ShortPrint ExtractFileName(frmMain.txtProject), nTextOffset, nTextLength
   If PrintStatus(2) Then GoTo ProjectPrintAbort

   TextCurrentX nNameOffset: LinePrint "Source Path:", False
   ShortPrint ExtractPath(frmMain.txtProject), nTextOffset, nTextLength
   If PrintStatus(3) Then GoTo ProjectPrintAbort

   LinePrint ""

   If Not EmptyString(Pj.Name) Then
      TextCurrentX nNameOffset: LinePrint "Project Name:", False
      ShortPrint Pj.Name, nTextOffset, nTextLength
   End If
   If PrintStatus(4) Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.Description) Then
      TextCurrentX nNameOffset: LinePrint "Application Description:", False
      ShortPrint Pj.Description, nTextOffset, nTextLength
   End If
   If PrintStatus(7) Then GoTo ProjectPrintAbort
   If (Pj.MajorVersion + Pj.MinorVersion + Pj.RevisionVersion) > 0 Then
      TextCurrentX nNameOffset: LinePrint "Version number:", False
      ShortPrint Format(Pj.MajorVersion, "###0") & "." & Format(Pj.MinorVersion, "###0") & "." & Format(Pj.RevisionVersion, "###0") & IIf(Pj.AutoVersion, "  (Auto increment)", ""), nTextOffset, nTextLength
   End If
   If PrintStatus(10) Then GoTo ProjectPrintAbort

   If Not EmptyString(Pj.Name) Or _
      Not EmptyString(Pj.Description) Or _
      (Pj.MajorVersion + Pj.MinorVersion + Pj.RevisionVersion) > 0 Then
      LinePrint ""
   End If

   If Not EmptyString(Pj.Comments) Then
      TextCurrentX nNameOffset: LinePrint "Comments:", False
      ShortPrint Pj.Comments, nTextOffset, nTextLength
   End If
   If PrintStatus(13) Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.CompanyName) Then
      TextCurrentX nNameOffset: LinePrint "Company Name:", False
      ShortPrint Pj.CompanyName, nTextOffset, nTextLength
   End If
   If PrintStatus(16) Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.FileDescription) Then
      TextCurrentX nNameOffset: LinePrint "File Description:", False
      ShortPrint Pj.FileDescription, nTextOffset, nTextLength
   End If
   If PrintStatus(19) Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.Copyright) Then
      TextCurrentX nNameOffset: LinePrint "Legal Copyright:", False
      ShortPrint Pj.Copyright, nTextOffset, nTextLength
   End If
   If PrintStatus(21) Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.TradeMarks) Then
      TextCurrentX nNameOffset: LinePrint "Legal Trademarks:", False
      ShortPrint Pj.TradeMarks, nTextOffset, nTextLength
   End If
   If PrintStatus(24) Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.ProductName) Then
      TextCurrentX nNameOffset: LinePrint "Product Name:", False
      ShortPrint Pj.ProductName, nTextOffset, nTextLength
   End If
   If PrintStatus(27) Then GoTo ProjectPrintAbort
   
   If Not EmptyString(Pj.Comments) Or Not EmptyString(Pj.CompanyName) Or _
      Not EmptyString(Pj.FileDescription) Or Not EmptyString(Pj.Copyright) Or _
      Not EmptyString(Pj.TradeMarks) Or Not EmptyString(Pj.ProductName) Then
      LinePrint ""
   End If

   If Not EmptyString(Pj.Title) Then
      TextCurrentX nNameOffset: LinePrint "Application Title:", False
      ShortPrint Pj.Title, nTextOffset, nTextLength
   End If
   If PrintStatus(30) Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.IconForm) Then
      TextCurrentX nNameOffset: LinePrint "Application Icon in:", False
      ShortPrint Pj.IconForm, nTextOffset, nTextLength
   End If
   If PrintStatus(33) Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.StartupForm) Then
      TextCurrentX nNameOffset: LinePrint "Startup Form:", False
      ShortPrint Pj.StartupForm, nTextOffset, nTextLength
   End If
   If PrintStatus(36) Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.StartMode) Then
      TextCurrentX nNameOffset: LinePrint "Start Mode", False
      ShortPrint Pj.StartMode, nTextOffset, nTextLength
   End If
   If PrintStatus(39) Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.CompileArg) Then
      TextCurrentX nNameOffset: LinePrint "Compilation Arguments:", False
      ShortPrint Pj.CompileArg, nTextOffset, nTextLength
   End If
   If PrintStatus(42) Then GoTo ProjectPrintAbort

   If Not EmptyString(Pj.HelpFile) Then
      If Not EmptyString(Pj.Title) Or Not EmptyString(Pj.IconForm) Or _
         Not EmptyString(Pj.StartupForm) Or Not EmptyString(Pj.StartMode) Or _
         Not EmptyString(Pj.CompileArg) Then
         LinePrint ""
      End If

      TextCurrentX nNameOffset: LinePrint "Help File:", False
      ShortPrint Pj.HelpFile, nTextOffset, nTextLength
   
      TextCurrentX nNameOffset: LinePrint "HelpContextID:", False
      ShortPrint Pj.HelpContextID, nTextOffset, nTextLength
   End If
   If PrintStatus(45) Then GoTo ProjectPrintAbort

   If Not EmptyString(Pj.Title) Or Not EmptyString(Pj.IconForm) Or _
      Not EmptyString(Pj.StartupForm) Or Not EmptyString(Pj.StartMode) Or _
      Not EmptyString(Pj.CompileArg) Or Not EmptyString(Pj.HelpFile) Then
      LinePrint ""
   End If

   If Pj.Bit32 Then
      SeperatorPrint 0, vbDot
      ShortPrint "Specific 32bit (for Windows 95 and NT) Information", 0, GetTxtWidth
      LinePrint ""
      If PrintStatus(48) Then GoTo ProjectPrintAbort

      If Not EmptyString(Pj.ExeName32) Then
         TextCurrentX nNameOffset: LinePrint "Executable Filename:", False
         ShortPrint Pj.ExeName32, nTextOffset, nTextLength
      End If
      If PrintStatus(51) Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.Path32) Then
         TextCurrentX nNameOffset: LinePrint "Path:", False
         ShortPrint Pj.Path32, nTextOffset, nTextLength
      End If
      If PrintStatus(54) Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.Command32) Then
         TextCurrentX nNameOffset: LinePrint "Command Line Arguments:", False
         ShortPrint Pj.Command32, nTextOffset, nTextLength
      End If
      If PrintStatus(57) Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.OLEServer32) Then
         TextCurrentX nNameOffset: LinePrint "Compatible OLE Server:", False
         ShortPrint Pj.OLEServer32, nTextOffset, nTextLength
      End If
      If PrintStatus(58) Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.Resource32) Then
         TextCurrentX nNameOffset: LinePrint "Resource file:", False
         ShortPrint Pj.Resource32, nTextOffset, nTextLength
      End If
      LinePrint ""
   End If
   If PrintStatus(60) Then GoTo ProjectPrintAbort

   If Pj.Bit16 Then
      SeperatorPrint 0, vbDot
      ShortPrint "Specific 16bit (for Windows 3.x) Information", 0, GetTxtWidth
      LinePrint ""
      If PrintStatus(63) Then GoTo ProjectPrintAbort

      If Not EmptyString(Pj.ExeName16) Then
         TextCurrentX nNameOffset: LinePrint "Executable Filename:", False
         ShortPrint Pj.ExeName16, nTextOffset, nTextLength
      End If
      If PrintStatus(66) Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.Path16) Then
         TextCurrentX nNameOffset: LinePrint "Path:", False
         ShortPrint Pj.Path16, nTextOffset, nTextLength
      End If
      If PrintStatus(69) Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.Command16) Then
         TextCurrentX nNameOffset: LinePrint "Command Line Arguments:", False
         ShortPrint Pj.Command16, nTextOffset, nTextLength
      End If
      If PrintStatus(72) Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.OLEServer16) Then
         TextCurrentX nNameOffset: LinePrint "Compatible OLE Server:", False
         ShortPrint Pj.OLEServer16, nTextOffset, nTextLength
      End If
      If PrintStatus(73) Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.Resource16) Then
         TextCurrentX nNameOffset: LinePrint "Resource file:", False
         ShortPrint Pj.Resource16, nTextOffset, nTextLength
      End If
      LinePrint ""
   End If
   If PrintStatus(75) Then GoTo ProjectPrintAbort

   If (Pj.FormCount + Pj.ModuleCount + Pj.ClassCount + Pj.ReferenceCount + Pj.ObjectCount) > 0 Then

      SeperatorPrint

      TextCurrentX 0: LinePrint "Project Files"
      LinePrint ""
      
      If Pj.FormCount > 0 Then
         LinePrint "Total Forms:", False
         TextCurrentX nTextOffset: LinePrint CInt(Pj.FormCount)
      End If
      If Pj.ModuleCount > 0 Then
         LinePrint "Total Modules:", False
         TextCurrentX nTextOffset: LinePrint CInt(Pj.ModuleCount)
      End If
      If Pj.ClassCount > 0 Then
         LinePrint "Total Classes:", False
         TextCurrentX nTextOffset: LinePrint CInt(Pj.ClassCount)
      End If
      If Pj.ReferenceCount > 0 Then
         LinePrint "Total References:", False
         TextCurrentX nTextOffset: LinePrint CInt(Pj.ReferenceCount)
      End If
      If Pj.ObjectCount > 0 Then
         LinePrint "Total Objects:", False
         TextCurrentX nTextOffset: LinePrint CInt(Pj.ObjectCount)
      End If

      LinePrint ""

      nTextOffset = Len("References ")
      nNameOffset = GetTxtWidth * 0.55

      nTextLength = nNameOffset - (nTextOffset + 1)
      nNameLength = GetTxtWidth - nNameOffset

      TextCurrentX nTextOffset: LinePrint "File", False
      TextCurrentX nNameOffset: LinePrint "Name"

      SeperatorPrint 0, vbDot

      If PrintStatus(78) Then GoTo ProjectPrintAbort

      ' Forms (.frm) in project
      If Pj.FormCount > 0 Then
         For i = 1 To Pj.FormCount
            If i = 1 Then
               TextCurrentX 0: LinePrint "Forms", False
            End If

            sString = Pj.Form(i).Name
            If Pj.Form(i).File = Pj.StartupFile Then
               sString = sString & " (App.Start)"
            End If
            If Pj.Form(i).Name = Pj.IconForm Then
               sString = sString & " (App.Icon)"
            End If

            SetWrapObject Pj.Form(i).File, nTextLength
            SaveWrap FileWrap, nFileLines
            SetWrapObject sString, nNameLength
            SaveWrap NameWrap, nNameLines

            nMax = IIf(nFileLines > nNameLines, nFileLines, nNameLines)
            For n = 1 To nMax
               TextCurrentX nTextOffset: If n <= nFileLines Then LinePrint FileWrap(n), False
               TextCurrentX nNameOffset: If n > nNameLines Then LinePrint "" Else LinePrint NameWrap(n)
            Next
         Next
         If (Pj.ModuleCount + Pj.ClassCount + Pj.ReferenceCount + Pj.ObjectCount) > 0 Then SeperatorPrint nTextOffset, vbDot
      End If
      If PrintStatus(81) Then GoTo ProjectPrintAbort

      ' Modules (.bas) in project
      If Pj.ModuleCount > 0 Then
         For i = 1 To Pj.ModuleCount
            If i = 1 Then
               TextCurrentX 0: LinePrint "Modules", False
            End If

            sString = Pj.Module(i).Name
            If Pj.Module(i).File = Pj.StartupFile Then
               sString = sString & " (App.Start)"
            End If

            SetWrapObject Pj.Module(i).File, nTextLength
            SaveWrap FileWrap, nFileLines
            SetWrapObject sString, nNameLength
            SaveWrap NameWrap, nNameLines

            nMax = IIf(nFileLines > nNameLines, nFileLines, nNameLines)
            For n = 1 To nMax
               TextCurrentX nTextOffset: If n <= nFileLines Then LinePrint FileWrap(n), False
               TextCurrentX nNameOffset: If n > nNameLines Then LinePrint "" Else LinePrint NameWrap(n)
            Next
         Next
         If (Pj.ClassCount + Pj.ReferenceCount + Pj.ObjectCount) > 0 Then SeperatorPrint nTextOffset, vbDot
      End If
      If PrintStatus(84) Then GoTo ProjectPrintAbort

      ' Classes (.cls) in project
      If Pj.ClassCount > 0 Then
         For i = 1 To Pj.ClassCount
            If i = 1 Then
               TextCurrentX 0: LinePrint "Classes", False
            End If

            SetWrapObject Pj.Class(i).File, nTextLength
            SaveWrap FileWrap, nFileLines
            SetWrapObject Pj.Class(i).Name, nNameLength
            SaveWrap NameWrap, nNameLines

            nMax = IIf(nFileLines > nNameLines, nFileLines, nNameLines)
            For n = 1 To nMax
               TextCurrentX nTextOffset: If n <= nFileLines Then LinePrint FileWrap(n), False
               TextCurrentX nNameOffset: If n > nNameLines Then LinePrint "" Else LinePrint NameWrap(n)
            Next
         Next
         If (Pj.ReferenceCount + Pj.ObjectCount) > 0 Then SeperatorPrint nTextOffset, vbDot
      End If
      If PrintStatus(87) Then GoTo ProjectPrintAbort

      ' References in project
      If Pj.ReferenceCount > 0 Then
         For i = 1 To Pj.ReferenceCount
            If i = 1 Then
               TextCurrentX 0: LinePrint "References", False
            End If

            SetWrapObject Pj.Reference(i).File, nTextLength
            SaveWrap FileWrap, nFileLines
            SetWrapObject Pj.Reference(i).Name, nNameLength
            SaveWrap NameWrap, nNameLines

            nMax = IIf(nFileLines > nNameLines, nFileLines, nNameLines)
            For n = 1 To nMax
               TextCurrentX nTextOffset: If n <= nFileLines Then LinePrint FileWrap(n), False
               TextCurrentX nNameOffset: If n > nNameLines Then LinePrint "" Else LinePrint NameWrap(n)
            Next
         Next
         If Pj.ObjectCount > 0 Then SeperatorPrint nTextOffset, vbDot
      End If
      If PrintStatus(90) Then GoTo ProjectPrintAbort

      nTextLength = GetTxtWidth - nTextOffset

      ' Objects in project
      If Pj.ObjectCount > 0 Then
         For i = 1 To Pj.ObjectCount
            If i = 1 Then
               TextCurrentX 0: LinePrint "Objects", False
            End If
            ShortPrint Pj.Object(i).File, nTextOffset, nTextLength
         Next
      End If

      SeperatorPrint 0, vbDot
      ShortPrint "(App.Icon) - Location were the application icon is stored", nTextOffset, nTextLength
      ShortPrint "(App.Start) - Specifies which file the application starts", nTextOffset, nTextLength
   End If

   If PrintStatus(100) Then GoTo ProjectPrintAbort

   On Error Resume Next

   If frmMain.chkControlNames = vbUnchecked And _
      frmMain.chkCode = vbUnchecked Then
      bFinalPage = True
   End If
   If nAvailable > -1 Then FooterPrint

   UserAbort

   Exit Sub

ProjectPrintError:
   MsgBox "Encounted a printing problem." & vbCrLf & "Error #" & Err.Number & ": " & Err.Description, vbCritical, "Print Error"

ProjectPrintAbort:
   On Error Resume Next

End Sub

Private Sub ShortPrint(sText As String, nTextOffset As Integer, nTextLength As Integer)
   Dim i As Integer, nLines As Integer
   Dim WrapLine() As String

   SetWrapObject sText, nTextLength
   SaveWrap WrapLine, nLines

   For i = 1 To nLines
      TextCurrentX nTextOffset
      LinePrint WrapLine(i)
   Next
End Sub

' --------------------------------------------------------------------------------------------------------

Private Sub PrintFormControls()
   Dim i As Integer, j As Integer, nIndex As Integer

   ' Load list into listbox (for optional sorting)
   nIndex = IIf(frmMain.chkSortControls = vbChecked, 1, 0)
   frmMain.lstNames(nIndex).Clear
   For i = 1 To Mdl(nMdlIndex).CtrlCount
      frmMain.lstNames(nIndex).AddItem Mdl(nMdlIndex).Control(i).Name
      frmMain.lstNames(nIndex).ItemData(frmMain.lstNames(nIndex).NewIndex) = i
   Next

   If UserAbort Then Exit Sub

   CheckAreaPrint
   LinePrint "(Form Control Objects)"
   LinePrint ""

   If frmMain.chkIndex = vbChecked Then                           ' INDEX - Update index-page reference
      Idx.CICount = Idx.CICount + 1
      ReDim Preserve Idx.ControlIndex(0 To Idx.CICount)
      Idx.ControlIndex(Idx.CICount).File = Mdl(nMdlIndex).File
      Idx.ControlIndex(Idx.CICount).Page = nPage
   End If

   For i = 0 To frmMain.lstNames(nIndex).ListCount - 1

      If UserAbort Then Exit For

      j = frmMain.lstNames(nIndex).ItemData(i)

      TextCurrentX (GetTxtWidth * 0.01): LinePrint Mdl(nMdlIndex).Control(j).Name, False
      TextCurrentX (GetTxtWidth * 0.3):  LinePrint Mdl(nMdlIndex).Control(j).Type, False
      TextCurrentX (GetTxtWidth * 0.5):  LinePrint Mdl(nMdlIndex).Control(j).Library, False

      If Mdl(nMdlIndex).Control(j).Elements > 1 Then
         TextCurrentX (GetTxtWidth * 0.65): LinePrint "Elements: " & Mdl(nMdlIndex).Control(j).Elements
      Else
         LinePrint ""
      End If
   Next i

   If UserAbort Then Exit Sub

   LinePrint ""
   LinePrint "   Total control names: " & Mdl(nMdlIndex).CtrlCount
   LinePrint "Total control elements: " & Mdl(nMdlIndex).CtrlElements

   ' That's it. Either print line or go to next page
   If frmMain.chkControlPage = vbChecked Or frmMain.chkProcPage = vbChecked Then
      bNextPage = True

   ElseIf frmMain.chkCode = vbChecked Or frmMain.chkIcon = vbChecked Then
      ' Only print a separator if code is following
      LinePrint ""
      SeperatorPrint
   End If

End Sub

' --------------------------------------------------------------------------------------------------------

Private Sub PrintIndexPage()
   Dim i As Integer, j As Integer, nMaxGauge As Integer, nCurGauge As Integer, nIndex As Integer
   Dim sString As String
   Dim H As Integer, W As Integer, nPointY As Integer

   If Page.Cancelled Then Exit Sub

   ' Any index info to be printed?
   If Idx.CICount < 0 And Idx.DIcount < 0 And Idx.PIcount < 0 Then Exit Sub
   ' Yep.

   If bPreview Then
      frmPreview.ShowProgress 0
   Else
      frmPrint.ShowJob "Printing Index"
      frmPrint.ShowProgress 0
   End If
   
   If Not InDevelopmentMode Then
      On Error GoTo PIP_ErrorHandler
   End If

   nMaxGauge = (Idx.CICount + 1) + (Idx.DIcount + 1) + (Idx.PIcount + 1)
   nCurGauge = 0

   bNextPage = True          ' Always print index on a new page
   nMdlIndex = -2            ' Let Header() now knows it's a index page

   If Idx.CICount > -1 Then  ' Controls index -------------------------------------------------------------
      If UserAbort Then GoTo PrintIndexAbort

      ' Load list into listbox (for optional sorting)
      nIndex = IIf(frmMain.chkSortIndex = vbChecked, 1, 0)
      frmMain.lstNames(nIndex).Clear
      For i = 0 To Idx.CICount
         frmMain.lstNames(nIndex).AddItem Idx.ControlIndex(i).File
         frmMain.lstNames(nIndex).ItemData(frmMain.lstNames(nIndex).NewIndex) = i
      Next

      TextCurrentX (GetTxtWidth * 0.9): LinePrint "Page", False
      TextCurrentX 0:                    LinePrint "Form Control Object INDEX"
      LinePrint ""

      For i = 0 To frmMain.lstNames(nIndex).ListCount - 1
         nCurGauge = nCurGauge + 1
         If PrintStatus((nCurGauge * 100) / nMaxGauge) Then Exit For

         j = frmMain.lstNames(nIndex).ItemData(i)

         TextCurrentX 1:                    LinePrint Idx.ControlIndex(j).File, False
         TextCurrentX (GetTxtWidth * 0.9): LinePrint CStr(Idx.ControlIndex(j).Page)
      Next
      LinePrint ""
      If Idx.DIcount > -1 Or Idx.PIcount > -1 Then SeperatorPrint
      frmMain.lstNames(nIndex).Clear
   End If

   If Idx.DIcount > -1 Then   ' Declarations index --------------------------------------------------------
      If UserAbort Then GoTo PrintIndexAbort

      ' Load list into listbox (for optional sorting)
      nIndex = IIf(frmMain.chkSortIndex = vbChecked, 1, 0)
      frmMain.lstNames(nIndex).Clear
      For i = 0 To Idx.DIcount
         frmMain.lstNames(nIndex).AddItem Idx.DeclareIndex(i).File
         frmMain.lstNames(nIndex).ItemData(frmMain.lstNames(nIndex).NewIndex) = i
      Next

      TextCurrentX (GetTxtWidth * 0.9): LinePrint "Page", False
      TextCurrentX 0:                    LinePrint "Declarations INDEX"
      LinePrint ""

      For i = 0 To frmMain.lstNames(nIndex).ListCount - 1
         nCurGauge = nCurGauge + 1
         If PrintStatus((nCurGauge * 100) / nMaxGauge) Then Exit For

         j = frmMain.lstNames(nIndex).ItemData(i)

         TextCurrentX 1:                    LinePrint Idx.DeclareIndex(j).File, False
         TextCurrentX (GetTxtWidth * 0.9): LinePrint CStr(Idx.DeclareIndex(j).Page)
      Next
      LinePrint ""
      If Idx.PIcount > -1 Then SeperatorPrint
      frmMain.lstNames(nIndex).Clear
   End If

   If Idx.PIcount > -1 Then   ' Procedures Index -----------------------------------------------------------
      If UserAbort Then GoTo PrintIndexAbort

      ' Load list into listbox (for optional sorting)
      nIndex = IIf(frmMain.chkSortIndex = vbChecked, 1, 0)
      frmMain.lstNames(nIndex).Clear
      For i = 0 To Idx.PIcount
         frmMain.lstNames(nIndex).AddItem Idx.ProcIndex(i).Procedure.Name
         frmMain.lstNames(nIndex).ItemData(frmMain.lstNames(nIndex).NewIndex) = i
      Next

      TextCurrentX (GetTxtWidth * 0.4):  LinePrint "Type", False
      TextCurrentX (GetTxtWidth * 0.65): LinePrint "File", False
      TextCurrentX (GetTxtWidth * 0.9):  LinePrint "Page", False
      TextCurrentX 0:                     LinePrint "Procedures INDEX"
      LinePrint ""

      For i = 0 To frmMain.lstNames(nIndex).ListCount - 1
         nCurGauge = nCurGauge + 1
         If PrintStatus((nCurGauge * 100) / nMaxGauge) Then Exit For

         j = frmMain.lstNames(nIndex).ItemData(i)

         TextCurrentX 1:                     LinePrint Idx.ProcIndex(j).Procedure.Name, False
         TextCurrentX (GetTxtWidth * 0.4):  LinePrint Idx.ProcIndex(j).Procedure.Type, False
         TextCurrentX (GetTxtWidth * 0.65): LinePrint Idx.ProcIndex(j).File, False
         TextCurrentX (GetTxtWidth * 0.9):  LinePrint CStr(Idx.ProcIndex(j).Page)
      Next

      frmMain.lstNames(nIndex).Clear
   End If

   If PrintStatus(100) Then GoTo PrintIndexAbort

   bFinalPage = True
   If nAvailable > -1 Then FooterPrint

   Exit Sub

PIP_ErrorHandler:
   MsgBox "Encounted a printing problem." & vbCrLf & "Error #" & Err.Number & ": " & Err.Description, vbCritical, "Print Error"

PrintIndexAbort:
   On Error Resume Next

End Sub
