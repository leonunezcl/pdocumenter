Attribute VB_Name = "PrintCode"
Option Explicit

' The Print Module !!
'
' Pages (in order):  1) Form Icon (if applicable)  Flags: .chkIcon
' (per module)       2) Controls (if applicable)          .chkControlNames
'                    3) Declaration/Procedures            .chkCode
'
' Additional pages:  a) Project Information               .chkProject
'                    b) Form icons                        .chkFormIcons
'                    c) Index                             .chkIndex
'
' Printer order:     a, b, 1, 2, 3, c
'
' -----------------------------------------------------
' See PnCommon.bas for support procedures and variables
' -----------------------------------------------------

Dim nPage As Integer             ' Page number
Dim nMdlIndex As Integer         ' Let the module know which file is being processed (using array element number)
Dim lp As Integer                ' Layout() pointer (element reference number)

Dim WrapText() As String         ' Used for the text-only wrapping procedures
Dim nWrapLines As Integer        ' Number of lines in text-only wrapped text

Dim bNextPage As Boolean         ' Next line should go to next page
Dim bFinalPage As Boolean        ' True if final page is done (used for preview)

Dim nFontRatio As Double         ' Used to scale fonts in preview picturebox (used in sample view)

Sub PrintJob()
   Dim i As Integer, j As Integer, n As Integer, nIndex As Integer, _
       nFinalIndex As Integer, nFinalProc As Integer
   Dim nSize As Long
   Dim sString As String, sUpper As String
   Dim H As Single, W As Single, nPointY As Single

   bFinalPage = False
   nPage = 0

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

   ' Max height needs some attention
   If frmMain.chkHeader = vbChecked Then
      If Page.Output = OUT_PORT Then
         nYHeight = GetHeight - (2 + IIf(frmMain.chkFooter = vbChecked, 3, 0))
      Else
         FontSet FONT_HEADER, False
         H = GetTextHeight() + 4
         If frmMain.chkFooter = vbChecked Then
            FontSet FONT_FOOTER, False
            H = H + (2 * GetTextHeight(frmMain.txtOwner(0) + frmMain.txtOwner(1))) + 4
         End If
         nYHeight = GetHeight - H
      End If
   Else
      nYHeight = GetHeight
   End If

   ' Height available. This one is easy: by setting it to -1, the header will be forced to be printed.
   nYAvailable = -1

   ' Prep it.
   PrintStartDoc

   If UserAbort Then GoTo PrintAbort

   ' General project information ----------------------------------------------------------------

   If frmMain.chkProject = vbChecked Then
      If frmMain.chkResetPage = vbChecked Then nPage = 0
      PrintProjectPage
      If Page.Cancelled Then GoTo PrintAbort
   End If

   ' Form icons (only to driver) ----------------------------------------------------------------

   If Page.Output = OUT_DRIVER Then
      If frmMain.chkFormIcons = vbChecked Then
         If frmMain.chkResetPage = vbChecked Then nPage = 0
         PrintProjectIcons
         If Page.Cancelled Then GoTo PrintAbort
      End If
   End If

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

         ' Initialise
         nYAvailable = -1                   ' Force header to print on first line to be printed
         bNextPage = True

         ' Does the user wants to abort?
         If UserAbort Then Exit For

         ' Icon (only to driver) -----------------------------------------------------------------
         If Page.Output = OUT_DRIVER And frmMain.chkIcon = vbChecked Then
            ' Icon print - Not available in text mode
            If Not EmptyString(Mdl(nMdlIndex).IconData) Then
               If PrintFormIcon(nMdlIndex, "(Icono Formulario)", 8) Then
                  If frmMain.chkControlPage = vbChecked Then
                     bNextPage = True
                  Else
                     FeedLine
                     SeperatorPrint
                  End If
               End If
            End If
         End If

         ' Controls ------------------------------------------------------------------------------
         If frmMain.chkControlNames = vbChecked And Mdl(nMdlIndex).CtrlSelect Then
            PrintFormControls
         End If

         If UserAbort Then Exit For

         ' Declaration/Procedures (Code) ---------------------------------------------------------
         If frmMain.chkCode = vbChecked And Mdl(nMdlIndex).ProcCount > 0 Then

            ' Get the last procedure
            nFinalProc = Mdl(nMdlIndex).ProcCount
            For n = Mdl(nMdlIndex).ProcCount To 1 Step -1
               If Mdl(nMdlIndex).Proc(n).Selected = vbChecked Then
                  nFinalProc = n
                  Exit For
               End If
            Next n

            CheckAreaPrint

            For n = 1 To Mdl(nMdlIndex).ProcCount

               If UserAbort Then Exit For

               If Mdl(nMdlIndex).Proc(n).Selected = vbChecked Then

                  If frmMain.chkProcNames = vbChecked Then                 ' Procedure names only...

                     If Mdl(nMdlIndex).Proc(n).Type <> PT_DECLARE Then     ' Declarations not allowed

                        FontSet FONT_PROCS                                 ' FontSet() ignores output to port (so let this call it)
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
                        FontSet FONT_PROCS
                        LinePrint "(Declaraciones)"
                        FeedLine

                        If frmMain.chkIndex = vbChecked Then               ' INDEX - Update index-page reference
                           Idx.DIcount = Idx.DIcount + 1
                           ReDim Preserve Idx.DeclareIndex(0 To Idx.DIcount)
                           Idx.DeclareIndex(Idx.DIcount).File = Mdl(nMdlIndex).File
                           Idx.DeclareIndex(Idx.DIcount).Page = nPage
                        End If
                     End If

                     For j = 1 To Mdl(nMdlIndex).Proc(n).Lines

                        sString = Mdl(nMdlIndex).Proc(n).Code(j)
                        sUpper = UCase$(Trim$(sString))

                        If MatchString(sUpper, "'") Then                      ' Comments
                           FontSet FONT_COMMENTS
                           LinePrint sString

                        ElseIf MatchString(sUpper, "#") Then                  ' Compiler directive
                           FontSet FONT_DIRECTIVE
                           LinePrint sString

                        ElseIf IsProcedure(sUpper) Then                       ' Only happens once (I hope) - and not in declaration section
                           FontSet FONT_PROCS
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
                           FontSet FONT_CODE ' Don't worry about repetitive font sets - PrintFont() is intelligent (it won't duplicate).
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
         End If            ' frmMain.chkCode = vbChecked And Mdl(nMdlIndex).ProcCount > 0

         If UserAbort Then Exit For

         ' ----------------------------------------------------------------------------------------

         If i = nFinalIndex And frmMain.chkIndex = vbUnchecked Then bFinalPage = True
         If nYAvailable > -1 Then FooterPrint

      End If   ' Mdl(i).Selected <> vbUnchecked
   Next i      ' i = 1 To MdCount [Cycle thru the files]

   If Page.Cancelled Then GoTo PrintAbort

   ' Index page(s) ------------------------------------------------------------------------------

   If frmMain.chkIndex = vbChecked Then
      ' Print the index page
      If frmMain.chkResetPage = vbChecked Then nPage = 0
      PrintIndexPage
      If Page.Cancelled Then GoTo PrintAbort
   End If

   ' --------------------------------------------------------------------------------------------

   On Error Resume Next
   Erase Idx.ControlIndex           ' Regain resources/memory
   Erase Idx.DeclareIndex
   Erase Idx.ProcIndex

   PrintEndDoc

   Exit Sub

PrintError:
   MsgBox "Problema al imprimir." & vbCrLf & "Error #" & Err.Number & ": " & Err.Description, vbCritical

PrintAbort:
   On Error Resume Next

   PrintKillDoc                     ' Job is aborted - how about killing the print buffer

   Erase Idx.ControlIndex
   Erase Idx.DeclareIndex
   Erase Idx.ProcIndex

End Sub

' (just a brainstorm which never continued, but left in the code just in case I change my mind) ----------
' Prints the module (file).
' Provide Procedure index: -1   for whole file
'                           0   for form controls (if available)
'                           1 > for procedures (1 is mostly declaration)
'
' Index array update only available when whole file is printed
'
'Function PrintModule(nMIndex As Integer, nPIndex As Integer) As Boolean
'   If frmMain.chkResetPage = vbChecked Then nPage = 0
'   nMdlIndex = nMIndex
'End Function
' --------------------------------------------------------------------------------------------------------

Sub FontSet(nIndex As Integer, Optional bSaveFont)
   If Page.Output = OUT_PORT Then Exit Sub
   If IsMissing(bSaveFont) Then bSaveFont = True
   SetFontName frmMain.lblFont(nIndex).FontName
   SetFontSize frmMain.lblFont(nIndex).FONTSIZE
   SetFontColor frmMain.lblFont(nIndex).ForeColor
   SetFontBold frmMain.lblFont(nIndex).FontBold
   SetFontItalic frmMain.lblFont(nIndex).FontItalic
   SetFontStrikethru frmMain.lblFont(nIndex).FontStrikethru
   SetFontUnderline frmMain.lblFont(nIndex).FontUnderline
   If bSaveFont Then PrintFont
End Sub

' ------------------------------------------------------------------------------------------

' This routine is a SOB...
Private Sub LinePrint(sString As String, Optional bCrLf)
   Dim nLines As Integer, i As Integer
   Dim nTextHeight As Single
   Dim sText As String
   Dim WrapLine() As String

   ' Remove trailing spaces
   sString = RTrim(sString)

   If IsMissing(bCrLf) Then bCrLf = True

   ' Place line into textbox - frmMain.txtWrap
   SetWrapObject sString
   SaveWrap WrapLine, nLines

   If Not bNextPage Then
      ' No forced pagebreak - check if there's enough room for the this string
      nTextHeight = nLines * GetTextHeight
      ' If the page height is smaller than then string height just "page wrap" it.
      If GetHeight > nTextHeight Then
         ' It will fit on one page - but is there still enough room for it?
         If nYAvailable < nTextHeight Then
            ' No - force page break
            bNextPage = True
         End If
      End If
   End If

   For i = 1 To nLines

      sText = WrapLine(i)

      If Not bNextPage Then
         If nYAvailable < GetTextHeight(sText) Then
            ' There's not enough page height to accomodate this text - go to next page
            bNextPage = True
         End If
      End If

      If bNextPage Then
         If Len(sText) = 0 Then GoTo EndOfLinePrint   ' Do not allow empty lines in top of page
         If nYAvailable > -1 Then                     ' Footer not printed yet...
            FooterPrint
            If Page.Cancelled Then Exit For
         End If
         HeaderPrint                                  ' Now print the header
         If Page.Cancelled Then Exit For
      End If

      ' Finally print the string
      If i > 1 Then
         If Page.Output = OUT_PORT Then
            PrintPrint ">> " & sText, bCrLf
         Else
            ' Turn attributes off for marker - but keep fontname, size and colour
            PushFont
            SetFontBold False
            SetFontItalic False
            SetFontStrikethru False
            SetFontUnderline False
            PrintFont
            PrintPrint ">> ", False          ' Print marker.
            ' Put it back to how it was
            PopFont
            PrintFont
            PrintPrint sText, bCrLf
         End If
      Else
         PrintPrint sText, bCrLf
      End If

      If bCrLf Then
         If Page.Output = OUT_RTF Then LayoutMode LYO_EOL
         ReduceHeight GetTextHeight(sText)
      End If

EndOfLinePrint:
   
   Next

End Sub

Private Sub FeedLine(Optional nFont, Optional nLines)
   If bNextPage Then Exit Sub
   ' Only process line if a new pages isn't forced

   PushFont

   If IsMissing(nFont) Then
      FontSet FONT_CODE, False
   Else
      FontSet CInt(nFont), False
   End If

   If IsMissing(nLines) Then nLines = 1

   Dim nHeight As Single
   nHeight = GetTextHeight() * nLines
      
   If nYAvailable < nHeight Then
      ' No room - force new page
      bNextPage = True
   Else
      ' Just process it - increment Y and deduct available height
      PrintPSet 0, GetCurrentY + nHeight
      ReduceHeight nHeight
   
      If Page.Output = OUT_RTF Then
         Dim p As Integer
         For p = 1 To nLines
            LayoutMode LYO_EOL
         Next
      End If

   End If

   PopFont
End Sub

' -----------------------------------------------------------------------------------
' 56.7 twips per logical millimeter
'
Private Sub SetWrapObject(sText As String, Optional nLength)
   If Page.Ruler = RULER_CHAR Then
      If IsMissing(nLength) Then
         nLength = GetWidth
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

'      nLength = nLength + 1
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
   
   Else
      If IsMissing(nLength) Then
         nLength = GetWidth
      Else
         nLength = CSng(nLength)
      End If

      ' Set the size (scale 1:1)
      frmMain.txtWrap.Width = nLength * 56.7
      frmMain.txtWrap.Height = GetHeight * 56.7
      ' Set the font
      CloneFont frmMain.txtWrap
      ' Set the text
      frmMain.txtWrap.Text = sText
   End If
End Sub

Private Function GetWrapLines() As Integer
   If Page.Ruler = RULER_CHAR Then
      GetWrapLines = nWrapLines
   Else
      GetWrapLines = SendMessageBynum(frmMain.txtWrap.hWnd, EM_GETLINECOUNT, 0, 0&)
   End If
End Function

Private Function GetWrapText(nGetLine As Integer) As String
   If Page.Ruler = RULER_CHAR Then
      If nWrapLines = 0 Or nGetLine < 1 Or nGetLine > nWrapLines Then
         GetWrapText = ""
         Exit Function
      End If
      GetWrapText = WrapText(nGetLine)

   Else
      Dim nCharOffset As Long, nLineSize As Long
      Dim sLineBuffer As String

      ' Find out the character offset to the first character in the specified line
      nCharOffset = SendMessageBynum(frmMain.txtWrap.hWnd, EM_LINEINDEX, nGetLine - 1, 0&)

      ' The character offset is used to determine the length of the line
      ' containing that character.
      nLineSize = SendMessageBynum(frmMain.txtWrap.hWnd, EM_LINELENGTH, nCharOffset, 0&) + 1

      ' Now allocate a string long enough to hold the result
      sLineBuffer = String$(nLineSize + 2, 0)
      Mid$(sLineBuffer, 1, 1) = Chr$(nLineSize And &HFF)
      Mid$(sLineBuffer, 2, 1) = Chr$(nLineSize \ 256)

      ' Now get the line
      nLineSize = SendMessageByString(frmMain.txtWrap.hWnd, EM_GETLINE, nGetLine - 1, sLineBuffer)

      GetWrapText = Left$(sLineBuffer, nLineSize)
   End If
End Function

Private Sub SaveWrap(ByRef TextHolder, ByRef TextLines)
   Dim i As Integer
   If Page.Ruler = RULER_CHAR Then
      TextLines = nWrapLines
   Else
      TextLines = GetWrapLines
   End If

   ReDim TextHolder(1 To TextLines)
   For i = 1 To TextLines
      TextHolder(i) = RTrim(GetWrapText(i))
   Next
End Sub

' -----------------------------------------------------------------------------------
' Check if there's enough room to print the sub line with some code (at least 2 lines of code)
'
Private Sub CheckAreaPrint(Optional nExtra)
   Dim H As Single

   If Page.Ruler = RULER_CHAR Then
      H = 1
   Else
      PushFont
      FontSet FONT_PROCS, False
      H = GetTextHeight()
      FontSet FONT_CODE, False
      H = H + (GetTextHeight() * 2)
      PopFont
   End If

   If Not IsMissing(nExtra) Then H = H + nExtra

   If nYAvailable < H Then bNextPage = True
End Sub

Private Sub SeperatorPrint(Optional nOffset, Optional nStyle)

   ' Prevent line at bottom of page without any text below it.
   If Page.Ruler = RULER_CHAR Then
      CheckAreaPrint
   Else
      CheckAreaPrint 2
   End If

   If bNextPage Then Exit Sub
   If IsMissing(nOffset) Then nOffset = 0
   If IsMissing(nStyle) Then nStyle = vbSolid

   Dim nCurY As Single
   nCurY = GetCurrentY

   If Page.Ruler = RULER_CHAR Then
      PrintLine CSng(nOffset), nCurY, , , , CInt(nStyle)
      PrintPSet 0, (nCurY + 1)
      ReduceHeight 1
   Else
      PrintLine CSng(nOffset), nCurY + 1, , , , CInt(nStyle)
      PrintPSet 0, (nCurY + 2)
      ReduceHeight 2
   End If
End Sub

Private Sub ReduceHeight(nUnits As Single)
   nYAvailable = nYAvailable - nUnits
   PrintProgress
End Sub

Private Sub HeaderPrint()
   Dim H As Single, nCurX As Single
   Dim sText As String, sPg As String

   nCurX = GetCurrentX
   PushFont

   nPage = nPage + 1
   If Page.Show Then PrintDialog "Formateando pagina " & nPage

   If frmMain.chkHeader = vbChecked Then

      sPg = "Pagina " & IIf(Page.Sample, "999", nPage)

      If Page.Sample Then
         sText = "Archivo (Formulario/Modulo/Clase/Control)"
      ElseIf nMdlIndex = -1 Then
         sText = ExtractFileName(frmMain.txtProject) & " (Proyecto)"
      ElseIf nMdlIndex = -2 Then
         sText = "Indice"
      ElseIf nMdlIndex = -3 Then
         sText = "Iconos"
      Else
         sText = Mdl(nMdlIndex).File
         If Mdl(nMdlIndex).Type = MT_MODULE Then
            sText = sText & " (Modulo - " & Mdl(nMdlIndex).Name & ")"
         ElseIf Mdl(nMdlIndex).Type = MT_CLASS Then
            sText = sText & " (Clase - " & Mdl(nMdlIndex).Name & ")"
         ElseIf Mdl(nMdlIndex).Type = MT_CONTROL Then
            sText = sText & " (Control de Usuario - " & Mdl(nMdlIndex).Name & ")"
         ElseIf Mdl(nMdlIndex).Type = MT_PROPERTY Then
            sText = sText & " (Pagina de Propiedades - " & Mdl(nMdlIndex).Name & ")"
         ElseIf Mdl(nMdlIndex).Type = MT_DOCUMENT Then
            sText = sText & " (Documento de Usuario - " & Mdl(nMdlIndex).Name & ")"
         Else
            sText = sText & " (Formulario - " & Mdl(nMdlIndex).Name & ")"
         End If
      End If

      If Page.Output = OUT_PORT Then
         If frmMain.optPagePos(0) And frmMain.chkPageNumbers = vbChecked Then
            SetWrapObject sText, (GetWidth - (1 + Len(sPg)))
            sText = Pad(RTrim(GetWrapText(1)), GetWidth - (1 + Len(sPg))) & " " & sPg
         Else
            SetWrapObject sText
            sText = RTrim(GetWrapText(1))
         End If
         PrintAt 0, 0, sText
         PrintLine 0, 1       ' Print the line
         PrintPSet 0, 2

         ' Page number?
'         If frmMain.optPagePos(0) And frmMain.chkPageNumbers = vbChecked Then
'            PrintAt (GetWidth - Len(sPg)), 0, sPg
'         End If

      ElseIf Page.Output = OUT_RTF Then
         FontSet FONT_HEADER
         H = GetTextHeight(sText)

         LayoutMode LYO_PAGE

         If frmMain.optPagePos(0) And frmMain.chkPageNumbers = vbChecked Then
            lp = AddLayout(LYO_TABS, 1, 0)
            ReDim Layout(lp).Tabs(0 To 0)
            Layout(lp).Tabs(0) = (GetWidth - (GetTextWidth(sPg) + 2.5)) * 56.7

            SetWrapObject sText, (GetWidth - (GetTextWidth(sPg) + 5))
         Else
            SetWrapObject sText
         End If
         sText = RTrim(GetWrapText(1))
         PrintPSet 1, 1
         PrintPrint sText

         If frmMain.optPagePos(0) And frmMain.chkPageNumbers = vbChecked Then
            LayoutMode LYO_TAB
            PrintAt (GetWidth - (GetTextWidth(sPg) + 2.5)), 1, sPg
         End If

         LayoutMode LYO_EOL

         PrintLine 0, H + 2

         H = H + 4
         PrintPSet 0, H

      Else
         FontSet FONT_HEADER
         H = GetTextHeight(sText)

         ' Print box first, so that the boxfill will not erase the text print
         PrintBox 0, 0, GetWidth - 0.1, H + 2

         If frmMain.optPagePos(0) And frmMain.chkPageNumbers = vbChecked Then
            SetWrapObject sText, (GetWidth - (GetTextWidth(sPg) + 5))
         Else
            SetWrapObject sText
         End If
         sText = RTrim(GetWrapText(1))
         PrintPSet 1, 1
         PrintPrint sText

         If frmMain.optPagePos(0) And frmMain.chkPageNumbers = vbChecked Then
            PrintAt (GetWidth - (GetTextWidth(sPg) + 2.5)), 1, sPg
         End If

         H = H + 4
         PrintPSet 0, H
      End If
   Else
      PrintPSet 0, 0
   End If

   nYAvailable = nYHeight        ' Reset available height
   PrintProgress

   bNextPage = False
   PopFont
   PrintFont

   SetCurrentX nCurX
End Sub

Private Sub FooterPrint()
   If Page.Cancelled Then Exit Sub

   Dim sText As String, sPg As String
   Dim nCurX As Single, H As Single
   Dim nSize As Integer

   nCurX = GetCurrentX
   PushFont

   If frmMain.chkFooter = vbChecked Then

      If Page.Output = OUT_PORT Then
         PrintLine 0, GetHeight - 3
      Else
         FontSet FONT_FOOTER
         H = 2 * GetTextHeight(frmMain.txtOwner(0) + frmMain.txtOwner(1))
         
         If Page.Output = OUT_RTF Then
            ' "Pad" lines to end
            nYAvailable = nYAvailable - (H + 4)
            FontSet FONT_CODE, False
            If nYAvailable >= GetTextHeight Then
               Dim i As Integer
               FontSet FONT_CODE
               nSize = Int(nYAvailable / GetTextHeight)
               For i = 1 To nSize
                  LayoutMode LYO_EOL
               Next
               FontSet FONT_FOOTER
            End If
            PrintLine 0, GetHeight - (H + 2)
         Else
            PrintLine 0, GetHeight - (H + 2)
            PrintLine 0, GetHeight - (H + 1.5)
         End If
      End If

      sText = ""
      If frmMain.chkDate = vbChecked Then
         sText = Format(Now, "Medium Date")
         If frmMain.chkTime = vbChecked Then
            sText = sText & " - " & Format(Now, "Medium Time")
         End If
      ElseIf frmMain.chkTime = vbChecked Then
         sText = Format(Now, "Medium Time")
      End If

      sPg = "Pagina " & IIf(Page.Sample, "999", nPage)

      If Page.Output = OUT_PORT Then
         If Len(Trim(sText)) > 0 Then
            nSize = GetWidth - (Len(sText) + 1)
            SetWrapObject frmMain.txtOwner(0), nSize
            sText = Pad(GetWrapText(1), nSize) & " " & sText
         Else
            SetWrapObject frmMain.txtOwner(0)
            sText = GetWrapText(1)
         End If
         PrintAt 0, GetHeight - 2, sText

         If frmMain.optPagePos(1) And frmMain.chkPageNumbers = vbChecked Then
            nSize = GetWidth - (Len(sPg) + 1)
            SetWrapObject frmMain.txtOwner(1), nSize
            sText = Pad(GetWrapText(1), nSize) & " " & sPg
         Else
            SetWrapObject frmMain.txtOwner(1)
            sText = GetWrapText(1)
         End If
         PrintAt 0, GetHeight - 1, sText

      ElseIf Page.Output = OUT_RTF Then
         If Len(Trim(sText)) > 0 Then
            lp = AddLayout(LYO_TABS, 1, 0)
            ReDim Layout(lp).Tabs(0 To 0)
            Layout(lp).Tabs(0) = (GetWidth - (GetTextWidth(sText) + 0.1)) * 56.7

            SetWrapObject frmMain.txtOwner(0), (GetWidth - (GetTextWidth(sText) + 5))
         Else
            SetWrapObject frmMain.txtOwner(0)
         End If
         PrintAt 0, GetHeight - H, RTrim(GetWrapText(1))
         If Len(Trim(sText)) > 0 Then
            LayoutMode LYO_TAB
            PrintAt GetWidth - (GetTextWidth(sText) + 0.1), GetHeight - H, sText
         End If
         LayoutMode LYO_EOL

         If frmMain.optPagePos(1) And frmMain.chkPageNumbers = vbChecked Then
            lp = AddLayout(LYO_TABS, 1, 0)
            ReDim Layout(lp).Tabs(0 To 0)
            Layout(lp).Tabs(0) = (GetWidth - (GetTextWidth(sPg) + 0.1)) * 56.7

            SetWrapObject frmMain.txtOwner(1), (GetWidth - (GetTextWidth(sPg) + 2.5))
         Else
            SetWrapObject frmMain.txtOwner(1)
         End If
         sText = RTrim(GetWrapText(1))
         PrintAt 0, GetHeight - GetTextHeight(sText), sText
         If frmMain.optPagePos(1) And frmMain.chkPageNumbers = vbChecked Then
            LayoutMode LYO_TAB
            PrintAt (GetWidth - (GetTextWidth(sPg) + 0.1)), GetHeight - GetTextHeight(sPg), sPg
         End If
         LayoutMode LYO_EOL

      Else
         If Len(Trim(sText)) > 0 Then
            PrintAt GetWidth - (GetTextWidth(sText) + 0.1), GetHeight - H, sText
            SetWrapObject frmMain.txtOwner(0), (GetWidth - (GetTextWidth(sText) + 5))
         Else
            SetWrapObject frmMain.txtOwner(0)
         End If
         sText = RTrim(GetWrapText(1))
         PrintAt 0, GetHeight - H, sText

         If frmMain.optPagePos(1) And frmMain.chkPageNumbers = vbChecked Then
            SetWrapObject frmMain.txtOwner(1), (GetWidth - (GetTextWidth(sPg) + 2.5))
            PrintAt (GetWidth - (GetTextWidth(sPg) + 0.1)), GetHeight - GetTextHeight(sPg), sPg
         Else
            SetWrapObject frmMain.txtOwner(1)
         End If
         sText = RTrim(GetWrapText(1))
         PrintAt 0, GetHeight - GetTextHeight(sText), sText
      End If
   End If

   ' Just for updating progress-bar (so that bar will give 100%)
   nYAvailable = 0
   PrintProgress

   ' Once footer is requested, do not let any printing occur on this page
   nYAvailable = -1

   PrintNewPage

   If Page.Sample Then Page.Cancelled = True     ' The end of the sample print...

   PopFont
   PrintFont
   SetCurrentX nCurX
End Sub

' --------------------------------------------------------------------------------------------------------

Private Sub PrintProjectPage()
   Dim i As Integer, n As Integer, nMax As Integer, _
       nFileLines As Integer, nNameLines As Integer
   Dim FileWrap() As String, NameWrap() As String
   Dim sString As String
   Dim nTextOffset As Single, nNameOffset As Single, _
       nTextLength As Single, nNameLength As Single
   Dim Pj As ProjectState
   Dim bBold As Boolean

   If Page.Cancelled Then Exit Sub

   PrintDialog "Analizando " & ExtractFileName(frmMain.txtProject)

   Pj = AnalyseVBP(frmMain.txtProject, Page.Form)
   If Not Pj.Loaded Then Exit Sub

   PrintDialog "Formateando"

   ' ----------------------------------------------------------------------------------

   If Not InDevelopmentMode Then
      On Error GoTo ProjectPrintError
   End If

   nMdlIndex = -1                   ' It's a project.
   
   nYAvailable = -1                 ' Force header to print on first line to be printed
   bNextPage = True                 ' Force new page
'   PrintPSet 0, 0

   ' Print the information...

   If Page.Output = OUT_DRIVER And Pj.IconPoint > -1 Then
      If PrintFormIcon(Pj.IconPoint, "Icono Aplicacion", 8) Then
         FeedLine
         SeperatorPrint
      End If
   End If

   FontSet FONT_TITLES
   LinePrint "Informacion General del Proyecto"
   FontSet FONT_CODE
   FeedLine
   If UserAbort Then GoTo ProjectPrintAbort

   ' Obtain the largest title width
   nNameOffset = 0
   If Page.Ruler = RULER_CHAR Then
      nTextOffset = Len("Descripcion de la Aplicacion: ")
   Else
      nTextOffset = GetTextWidth("Descripcion de la Aplicacion:") + 3
   End If
   nTextLength = GetWidth - nTextOffset

   If Page.Output = OUT_RTF Then
      lp = AddLayout(LYO_TABS, 1, 0)
      ReDim Layout(lp).Tabs(0 To 0)
      Layout(lp).Tabs(0) = nTextOffset * 56.7
   End If

   SetCurrentX nNameOffset: LinePrint "Archivo VBP:", False
   ShortPrint ExtractFileName(frmMain.txtProject), nTextOffset, nTextLength
   If UserAbort Then GoTo ProjectPrintAbort

   SetCurrentX nNameOffset: LinePrint "Path Fuentes:", False
   ShortPrint ExtractPath(frmMain.txtProject), nTextOffset, nTextLength
   If UserAbort Then GoTo ProjectPrintAbort

   FeedLine

   If Not EmptyString(Pj.Name) Then
      SetCurrentX nNameOffset: LinePrint "Nombre Proyecto:", False
      ShortPrint Pj.Name, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.Description) Then
      SetCurrentX nNameOffset: LinePrint "Descripcion Aplicacion:", False
      ShortPrint Pj.Description, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort
   If (Pj.MajorVersion + Pj.MinorVersion + Pj.RevisionVersion) > 0 Then
      SetCurrentX nNameOffset: LinePrint "Version:", False
      ShortPrint Format(Pj.MajorVersion, "###0") & "." & Format(Pj.MinorVersion, "###0") & "." & Format(Pj.RevisionVersion, "###0") & IIf(Pj.AutoVersion, "  (Auto incremento)", ""), nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort

   If Not EmptyString(Pj.Name) Or _
      Not EmptyString(Pj.Description) Or _
      (Pj.MajorVersion + Pj.MinorVersion + Pj.RevisionVersion) > 0 Then
      FeedLine
   End If

   If Not EmptyString(Pj.Comments) Then
      SetCurrentX nNameOffset: LinePrint "Comentarios:", False
      ShortPrint Pj.Comments, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.CompanyName) Then
      SetCurrentX nNameOffset: LinePrint "Nombre Compañia:", False
      ShortPrint Pj.CompanyName, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.FileDescription) Then
      SetCurrentX nNameOffset: LinePrint "Descripcion Archivo:", False
      ShortPrint Pj.FileDescription, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.Copyright) Then
      SetCurrentX nNameOffset: LinePrint "Legal Copyright:", False
      ShortPrint Pj.Copyright, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.TradeMarks) Then
      SetCurrentX nNameOffset: LinePrint "Marcas Legales:", False
      ShortPrint Pj.TradeMarks, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.ProductName) Then
      SetCurrentX nNameOffset: LinePrint "Nombre Producto:", False
      ShortPrint Pj.ProductName, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort
   
   If Not EmptyString(Pj.Comments) Or Not EmptyString(Pj.CompanyName) Or _
      Not EmptyString(Pj.FileDescription) Or Not EmptyString(Pj.Copyright) Or _
      Not EmptyString(Pj.TradeMarks) Or Not EmptyString(Pj.ProductName) Then
      FeedLine
   End If

   If Not EmptyString(Pj.Title) Then
      SetCurrentX nNameOffset: LinePrint "Titulo Aplicacion:", False
      ShortPrint Pj.Title, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.IconForm) Then
      SetCurrentX nNameOffset: LinePrint "Icono Aplicacion:", False
      ShortPrint Pj.IconForm, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.StartupForm) Then
      SetCurrentX nNameOffset: LinePrint "Formulario Inicial:", False
      ShortPrint Pj.StartupForm, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.StartMode) Then
      SetCurrentX nNameOffset: LinePrint "Modo Partida", False
      ShortPrint Pj.StartMode, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort
   If Not EmptyString(Pj.CompileArg) Then
      SetCurrentX nNameOffset: LinePrint "Argumentos Compilacion:", False
      ShortPrint Pj.CompileArg, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort

   If Not EmptyString(Pj.HelpFile) Then
      If Not EmptyString(Pj.Title) Or Not EmptyString(Pj.IconForm) Or _
         Not EmptyString(Pj.StartupForm) Or Not EmptyString(Pj.StartMode) Or _
         Not EmptyString(Pj.CompileArg) Then
         FeedLine
      End If

      SetCurrentX nNameOffset: LinePrint "Archivo Ayuda:", False
      ShortPrint Pj.HelpFile, nTextOffset, nTextLength

      SetCurrentX nNameOffset: LinePrint "HelpContextID:", False
      ShortPrint Pj.HelpContextID, nTextOffset, nTextLength
   End If
   If UserAbort Then GoTo ProjectPrintAbort

   If Not EmptyString(Pj.Title) Or Not EmptyString(Pj.IconForm) Or _
      Not EmptyString(Pj.StartupForm) Or Not EmptyString(Pj.StartMode) Or _
      Not EmptyString(Pj.CompileArg) Or Not EmptyString(Pj.HelpFile) Then
      FeedLine
   End If

   If Pj.Bit32 Then
      'FontSet FONT_CODE
      SeperatorPrint 0, vbDot
      FontSet FONT_TITLES
      ShortPrint "Especificaciones 32bit (para Windows 95, 98 y NT) Informacion", 0, GetWidth, True
      FontSet FONT_CODE
      FeedLine
      If UserAbort Then GoTo ProjectPrintAbort

      If Not EmptyString(Pj.ExeName32) Then
         SetCurrentX nNameOffset: LinePrint "Nombre Ejecutable:", False
         ShortPrint Pj.ExeName32, nTextOffset, nTextLength
      End If
      If UserAbort Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.Path32) Then
         SetCurrentX nNameOffset: LinePrint "Path:", False
         ShortPrint Pj.Path32, nTextOffset, nTextLength
      End If
      If UserAbort Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.Command32) Then
         SetCurrentX nNameOffset: LinePrint "Argumentos Linea de Comando:", False
         ShortPrint Pj.Command32, nTextOffset, nTextLength
      End If
      If UserAbort Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.OLEServer32) Then
         SetCurrentX nNameOffset: LinePrint "Compatible OLE Server:", False
         ShortPrint Pj.OLEServer32, nTextOffset, nTextLength
      End If
      If UserAbort Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.Resource32) Then
         SetCurrentX nNameOffset: LinePrint "Archivo de Recursos:", False
         ShortPrint Pj.Resource32, nTextOffset, nTextLength
      End If
      FeedLine
   End If
   If UserAbort Then GoTo ProjectPrintAbort

   If Pj.Bit16 Then
      'FontSet FONT_CODE
      SeperatorPrint 0, vbDot
      FontSet FONT_TITLES
      ShortPrint "Especificaciones 16bit (para Windows 3.x) Informacion", 0, GetWidth, True
      FontSet FONT_CODE
      FeedLine
      If UserAbort Then GoTo ProjectPrintAbort

      If Not EmptyString(Pj.ExeName16) Then
         SetCurrentX nNameOffset: LinePrint "Nombre Ejecutable:", False
         ShortPrint Pj.ExeName16, nTextOffset, nTextLength
      End If
      If UserAbort Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.Path16) Then
         SetCurrentX nNameOffset: LinePrint "Path:", False
         ShortPrint Pj.Path16, nTextOffset, nTextLength
      End If
      If UserAbort Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.Command16) Then
         SetCurrentX nNameOffset: LinePrint "Argumentos Linea de Comando:", False
         ShortPrint Pj.Command16, nTextOffset, nTextLength
      End If
      If UserAbort Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.OLEServer16) Then
         SetCurrentX nNameOffset: LinePrint "Compatible OLE Server:", False
         ShortPrint Pj.OLEServer16, nTextOffset, nTextLength
      End If
      If UserAbort Then GoTo ProjectPrintAbort
      If Not EmptyString(Pj.Resource16) Then
         SetCurrentX nNameOffset: LinePrint "Archivo de Recursos:", False
         ShortPrint Pj.Resource16, nTextOffset, nTextLength
      End If
      FeedLine
   End If
   If UserAbort Then GoTo ProjectPrintAbort

   If (Pj.FormCount + Pj.ModuleCount + Pj.ClassCount + Pj.ReferenceCount + Pj.ObjectCount) > 0 Then

      SeperatorPrint

      FontSet FONT_TITLES
      SetCurrentX 0: LinePrint "Archivos del Proyecto"
      FontSet FONT_CODE
      FeedLine
      
      If Pj.FormCount > 0 Then
         LinePrint "Total Formularios:", False
         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nTextOffset: LinePrint CInt(Pj.FormCount)
      End If
      If Pj.ModuleCount > 0 Then
         LinePrint "Total Modulos:", False
         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nTextOffset: LinePrint CInt(Pj.ModuleCount)
      End If
      If Pj.ClassCount > 0 Then
         LinePrint "Total Clases:", False
         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nTextOffset: LinePrint CInt(Pj.ClassCount)
      End If

      If Pj.ControlCount > 0 Then
         LinePrint "Total Control de Usuario:", False
         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nTextOffset: LinePrint CInt(Pj.ControlCount)
      End If
      If Pj.PropertyCount > 0 Then
         LinePrint "Total Paginas de Propiedades:", False
         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nTextOffset: LinePrint CInt(Pj.PropertyCount)
      End If
      If Pj.DocumentCount > 0 Then
         LinePrint "Total Documento de Usuario:", False
         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nTextOffset: LinePrint CInt(Pj.DocumentCount)
      End If
      If Pj.RelatedCount > 0 Then
         LinePrint "Total Documentos Asociados:", False
         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nTextOffset: LinePrint CInt(Pj.RelatedCount)
      End If

      If Pj.ReferenceCount > 0 Then
         LinePrint "Total Referencias:", False
         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nTextOffset: LinePrint CInt(Pj.ReferenceCount)
      End If
      If Pj.ObjectCount > 0 Then
         LinePrint "Total Objetos:", False
         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nTextOffset: LinePrint CInt(Pj.ObjectCount)
      End If

      FeedLine

      SetFontBold True
      PrintFont
      bBold = True

      If Page.Ruler = RULER_CHAR Then
         nTextOffset = Len("Documentos Asociados")
         nNameOffset = GetWidth * 0.55

         nTextLength = nNameOffset - (nTextOffset + 1)
         nNameLength = GetWidth - nNameOffset
      Else
         nTextOffset = GetTextWidth("Documentos Asociados") + 3
         nNameOffset = GetWidth * 0.55

         nTextLength = nNameOffset - (nTextOffset + 3)
         nNameLength = GetWidth - nNameOffset
      End If

      If Page.Output = OUT_RTF Then
         lp = AddLayout(LYO_TABS, 2, 1)
         ReDim Layout(lp).Tabs(0 To 1)
         Layout(lp).Tabs(0) = nTextOffset * 56.7
         Layout(lp).Tabs(1) = nNameOffset * 56.7
      End If

      If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
      SetCurrentX nTextOffset: LinePrint "Archivo", False
      If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
      SetCurrentX nNameOffset: LinePrint "Nombre"

      SeperatorPrint 0, vbDot

      If UserAbort Then GoTo ProjectPrintAbort

      ' Forms (.frm) in project
      If Pj.FormCount > 0 Then
         For i = 1 To Pj.FormCount
            If i = 1 Then
               'SetFontBold True
               'PrintFont
               SetCurrentX 0: LinePrint "Formularios", False
               FontSet FONT_CODE
               bBold = False
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
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nTextOffset: If n <= nFileLines Then LinePrint FileWrap(n), False
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nNameOffset: If n > nNameLines Then LinePrint "" Else LinePrint NameWrap(n)
            Next
         Next
         If (Pj.ModuleCount + Pj.ClassCount + Pj.ControlCount + Pj.PropertyCount + Pj.DocumentCount + Pj.RelatedCount + Pj.ReferenceCount + Pj.ObjectCount) > 0 Then SeperatorPrint nTextOffset, vbDot
      End If
      If UserAbort Then GoTo ProjectPrintAbort

      ' Modules (.bas) in project
      If Pj.ModuleCount > 0 Then
         For i = 1 To Pj.ModuleCount
            If i = 1 Then
               If Not bBold Then
                  SetFontBold True
                  PrintFont
               End If
               SetCurrentX 0: LinePrint "Modulos", False
               FontSet FONT_CODE
               bBold = False
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
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nTextOffset: If n <= nFileLines Then LinePrint FileWrap(n), False
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nNameOffset: If n > nNameLines Then LinePrint "" Else LinePrint NameWrap(n)
            Next
         Next
         If (Pj.ClassCount + Pj.ControlCount + Pj.PropertyCount + Pj.DocumentCount + Pj.RelatedCount + Pj.ReferenceCount + Pj.ObjectCount) > 0 Then SeperatorPrint nTextOffset, vbDot
      End If
      If UserAbort Then GoTo ProjectPrintAbort

      ' Classes (.cls) in project
      If Pj.ClassCount > 0 Then
         For i = 1 To Pj.ClassCount
            If i = 1 Then
               If Not bBold Then
                  SetFontBold True
                  PrintFont
               End If
               SetCurrentX 0: LinePrint "Clases", False
               FontSet FONT_CODE
               bBold = False
            End If

            SetWrapObject Pj.Class(i).File, nTextLength
            SaveWrap FileWrap, nFileLines
            SetWrapObject Pj.Class(i).Name, nNameLength
            SaveWrap NameWrap, nNameLines

            nMax = IIf(nFileLines > nNameLines, nFileLines, nNameLines)
            For n = 1 To nMax
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nTextOffset: If n <= nFileLines Then LinePrint FileWrap(n), False
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nNameOffset: If n > nNameLines Then LinePrint "" Else LinePrint NameWrap(n)
            Next
         Next
         If (Pj.ControlCount + Pj.PropertyCount + Pj.DocumentCount + Pj.RelatedCount + Pj.ReferenceCount + Pj.ObjectCount) > 0 Then SeperatorPrint nTextOffset, vbDot
      End If
      If UserAbort Then GoTo ProjectPrintAbort

      ' User Controls (.ctl) in project
      If Pj.ControlCount > 0 Then
         For i = 1 To Pj.ControlCount
            If i = 1 Then
               If Not bBold Then
                  SetFontBold True
                  PrintFont
               End If
               SetCurrentX 0: LinePrint "Controles de Usuario", False
               FontSet FONT_CODE
               bBold = False
            End If

            sString = Pj.UControl(i).Name
'            If Pj.UControl(i).File = Pj.StartupFile Then
'               sString = sString & " (App.Start)"
'            End If

            SetWrapObject Pj.UControl(i).File, nTextLength
            SaveWrap FileWrap, nFileLines
            SetWrapObject sString, nNameLength
            SaveWrap NameWrap, nNameLines

            nMax = IIf(nFileLines > nNameLines, nFileLines, nNameLines)
            For n = 1 To nMax
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nTextOffset: If n <= nFileLines Then LinePrint FileWrap(n), False
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nNameOffset: If n > nNameLines Then LinePrint "" Else LinePrint NameWrap(n)
            Next
         Next
         If (Pj.PropertyCount + Pj.DocumentCount + Pj.RelatedCount + Pj.ReferenceCount + Pj.ObjectCount) > 0 Then SeperatorPrint nTextOffset, vbDot
      End If
      If UserAbort Then GoTo ProjectPrintAbort

      ' Property Page (.pag) in project
      If Pj.PropertyCount > 0 Then
         For i = 1 To Pj.PropertyCount
            If i = 1 Then
               If Not bBold Then
                  SetFontBold True
                  PrintFont
               End If
               SetCurrentX 0: LinePrint "Paginas de Propiedades", False
               FontSet FONT_CODE
               bBold = False
            End If

            sString = Pj.PropertyPg(i).Name
'            If Pj.PropertyPg(i).File = Pj.StartupFile Then
'               sString = sString & " (App.Start)"
'            End If

            SetWrapObject Pj.PropertyPg(i).File, nTextLength
            SaveWrap FileWrap, nFileLines
            SetWrapObject sString, nNameLength
            SaveWrap NameWrap, nNameLines

            nMax = IIf(nFileLines > nNameLines, nFileLines, nNameLines)
            For n = 1 To nMax
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nTextOffset: If n <= nFileLines Then LinePrint FileWrap(n), False
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nNameOffset: If n > nNameLines Then LinePrint "" Else LinePrint NameWrap(n)
            Next
         Next
         If (Pj.DocumentCount + Pj.RelatedCount + Pj.ReferenceCount + Pj.ObjectCount) > 0 Then SeperatorPrint nTextOffset, vbDot
      End If
      If UserAbort Then GoTo ProjectPrintAbort

      ' User Documents (.dob) in project
      If Pj.DocumentCount > 0 Then
         For i = 1 To Pj.DocumentCount
            If i = 1 Then
               If Not bBold Then
                  SetFontBold True
                  PrintFont
               End If
               SetCurrentX 0: LinePrint "Documentos de Usuario", False
               FontSet FONT_CODE
               bBold = False
            End If

            sString = Pj.UDocument(i).Name
'            If Pj.UDocument(i).File = Pj.StartupFile Then
'               sString = sString & " (App.Start)"
'            End If

            SetWrapObject Pj.UDocument(i).File, nTextLength
            SaveWrap FileWrap, nFileLines
            SetWrapObject sString, nNameLength
            SaveWrap NameWrap, nNameLines

            nMax = IIf(nFileLines > nNameLines, nFileLines, nNameLines)
            For n = 1 To nMax
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nTextOffset: If n <= nFileLines Then LinePrint FileWrap(n), False
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nNameOffset: If n > nNameLines Then LinePrint "" Else LinePrint NameWrap(n)
            Next
         Next
         If (Pj.RelatedCount + Pj.ReferenceCount + Pj.ObjectCount) > 0 Then SeperatorPrint nTextOffset, vbDot
      End If
      If UserAbort Then GoTo ProjectPrintAbort

      ' Related Documents (.*) in project
      If Pj.RelatedCount > 0 Then
         For i = 1 To Pj.RelatedCount
            If i = 1 Then
               If Not bBold Then
                  SetFontBold True
                  PrintFont
               End If
               SetCurrentX 0: LinePrint "Documentos Asociados", False
               FontSet FONT_CODE
               bBold = False
            End If

            SetWrapObject Pj.RelatedDoc(i).File, nTextLength
            SaveWrap FileWrap, nFileLines
            SetWrapObject Pj.RelatedDoc(i).Name, nNameLength
            SaveWrap NameWrap, nNameLines

            nMax = IIf(nFileLines > nNameLines, nFileLines, nNameLines)
            For n = 1 To nMax
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nTextOffset: If n <= nFileLines Then LinePrint FileWrap(n), False
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nNameOffset: If n > nNameLines Then LinePrint "" Else LinePrint NameWrap(n)
            Next
         Next
         If (Pj.ReferenceCount + Pj.ObjectCount) > 0 Then SeperatorPrint nTextOffset, vbDot
      End If
      If UserAbort Then GoTo ProjectPrintAbort

      ' References in project
      If Pj.ReferenceCount > 0 Then
         For i = 1 To Pj.ReferenceCount
            If i = 1 Then
               If Not bBold Then
                  SetFontBold True
                  PrintFont
               End If
               SetCurrentX 0: LinePrint "Referencias", False
               FontSet FONT_CODE
               bBold = False
            End If

            SetWrapObject Pj.Reference(i).File, nTextLength
            SaveWrap FileWrap, nFileLines
            SetWrapObject Pj.Reference(i).Name, nNameLength
            SaveWrap NameWrap, nNameLines

            nMax = IIf(nFileLines > nNameLines, nFileLines, nNameLines)
            For n = 1 To nMax
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nTextOffset: If n <= nFileLines Then LinePrint FileWrap(n), False
               If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
               SetCurrentX nNameOffset: If n > nNameLines Then LinePrint "" Else LinePrint NameWrap(n)
            Next
         Next
         If Pj.ObjectCount > 0 Then SeperatorPrint nTextOffset, vbDot
      End If
      If UserAbort Then GoTo ProjectPrintAbort

      nTextLength = GetWidth - nTextOffset

      ' Objects in project
      If Pj.ObjectCount > 0 Then
         For i = 1 To Pj.ObjectCount
            If i = 1 Then
               If Not bBold Then
                  SetFontBold True
                  PrintFont
               End If
               SetCurrentX 0: LinePrint "Objetos", False
               FontSet FONT_CODE
               bBold = False
            End If
            ShortPrint Pj.Object(i).File, nTextOffset, nTextLength
         Next
      End If

      SeperatorPrint 0, vbDot

      FontSet FONT_COMMENTS
      ShortPrint "(App.Icon) - Ubicacion donde el icono de la aplicacion esta", nTextOffset, nTextLength
      ShortPrint "(App.Start) - Especifica cual es el archivo de inicio de la aplicacion", nTextOffset, nTextLength
      FontSet FONT_CODE
   End If

   If UserAbort Then GoTo ProjectPrintAbort

   On Error Resume Next

   If frmMain.chkIcon = vbUnchecked And _
      frmMain.chkControlNames = vbUnchecked And _
      frmMain.chkCode = vbUnchecked And _
      frmMain.chkFormIcons = vbUnchecked Then
      bFinalPage = True
   End If
   If nYAvailable > -1 Then FooterPrint

   Exit Sub

ProjectPrintError:
   MsgBox "Problema al imprimir." & vbCrLf & "Error #" & Err.Number & ": " & Err.Description, vbCritical

ProjectPrintAbort:
   On Error Resume Next

End Sub

Private Sub ShortPrint(sText As String, nTextOffset As Single, nTextLength As Single, Optional bNoTab)
   Dim i As Integer, nLines As Integer
   Dim WrapLine() As String

   If IsMissing(bNoTab) Then bNoTab = False

   SetWrapObject sText, nTextLength
   SaveWrap WrapLine, nLines

   For i = 1 To nLines
      If Page.Output = OUT_RTF And Not bNoTab Then LayoutMode LYO_TAB

      SetCurrentX nTextOffset
      LinePrint WrapLine(i)
   Next

   PrintProgress
End Sub

' --------------------------------------------------------------------------------------------------------

' sTitle = "Application Icon"
'          "(Form Icon)"
'
' Use LoadIcon() function to obtain icon into frmMain.picImage
'
Private Function PrintFormIcon(nIndex As Integer, sTitle As String, Optional nSetX) As Boolean
   On Error GoTo FileImageError

   If Page.Output <> OUT_DRIVER Then GoTo FileImageError
   If Not LoadIcon(nIndex) Then GoTo FileImageError

   frmMain.picImage.ScaleMode = vbMillimeters
   CheckAreaPrint frmMain.picImage.ScaleHeight

   PushFont
   FontSet FONT_TITLES
   LinePrint sTitle
   PopFont

   FeedLine

   ' Shift image?
   If Not IsMissing(nSetX) Then SetCurrentX CSng(nSetX)

   ' Print the image
   PrintPicture nIndex, frmMain.picImage.Width, frmMain.picImage.Height
   ReduceHeight frmMain.picImage.ScaleHeight

   ' Lose the picture, gain resources/memory
   frmMain.picImage.Picture = LoadPicture()

   PrintFormIcon = True
   Exit Function

FileImageError:
   PrintFormIcon = False
End Function

Private Sub PrintProjectIcons()

   If Page.Cancelled Then Exit Sub
   If Page.Output <> OUT_DRIVER Then Exit Sub

   ' ----------------------------------------------------------------------------------

   If Not InDevelopmentMode Then
      On Error GoTo PAI_ErrorHandler
   End If

   Dim i As Integer
   Dim nCurY As Single, nPosY As Single, nPosX As Single

   nMdlIndex = -3                   ' It's the icons
   nYAvailable = -1                  ' Force header to print on first line to be printed
   bNextPage = True

   'FontSet FONT_CODE

   For i = 1 To MdCount

      If UserAbort Then Exit For

'      If Mdl(i).Selected <> vbUnchecked Then

         If Not EmptyString(Mdl(i).IconData) Then
            
            If LoadIcon(i) Then

               frmMain.picImage.ScaleMode = vbMillimeters
               CheckAreaPrint frmMain.picImage.ScaleHeight

               If bNextPage Then
                  HeaderPrint
                  If Page.Cancelled Then Exit For
               End If

               FeedLine

               nPosY = GetCurrentY
               nPosX = frmMain.picImage.ScaleWidth + 8

               SetCurrentX 4

               ' Print the image
               PrintPicture i, frmMain.picImage.Width, frmMain.picImage.Height
               ReduceHeight frmMain.picImage.ScaleHeight

               ' Lose the picture, gain resources/memory
               frmMain.picImage.Picture = LoadPicture()

               ' Store position prior repositioning for the text
               nCurY = GetCurrentY
               PrintPSet nPosX, nPosY

               If EmptyString(Mdl(i).Name) Then
                  FontSet FONT_PROCS
                  LinePrint Mdl(i).File
               Else
                  FontSet FONT_PROCS
                  LinePrint Mdl(i).File & " ", False
                  FontSet FONT_CODE
                  LinePrint "(" & Mdl(i).Name & ")"
               End If

               ' Reposition height
               SetCurrentY nCurY
            End If
         End If
'      End If
   Next i

   If UserAbort Then GoTo PAI_ErrorHandler

   On Error Resume Next

   If frmMain.chkControlNames = vbUnchecked And _
      frmMain.chkCode = vbUnchecked And _
      frmMain.chkIcon = vbUnchecked Then
      bFinalPage = True
   End If
   If nYAvailable > -1 Then FooterPrint

PAI_ErrorHandler:
End Sub

' --------------------------------------------------------------------------------------------------------

Private Sub PrintFormControls()
   Dim i As Integer, j As Integer, nIndex As Integer
   Dim nNameOffset As Single, nTypeOffset As Single, nLibraryOffset As Single, nElementsOffset As Single

   ' Load list into listbox (for optional sorting)
   nIndex = IIf(frmMain.chkSortControls = vbChecked, 1, 0)
   frmMain.lstNames(nIndex).Clear
   For i = 1 To Mdl(nMdlIndex).CtrlCount
      frmMain.lstNames(nIndex).AddItem Mdl(nMdlIndex).Control(i).Name
      frmMain.lstNames(nIndex).ItemData(frmMain.lstNames(nIndex).NewIndex) = i
   Next

   If UserAbort Then Exit Sub

   nNameOffset = GetWidth * 0.01
   nTypeOffset = GetWidth * 0.3
   nLibraryOffset = GetWidth * 0.65
   nElementsOffset = GetWidth * 0.8

   If Page.Output = OUT_RTF Then
      lp = AddLayout(LYO_TABS, 3, 2)
      ReDim Layout(lp).Tabs(0 To 2)
      Layout(lp).Tabs(0) = nTypeOffset * 56.7
      Layout(lp).Tabs(1) = nLibraryOffset * 56.7
      Layout(lp).Tabs(2) = nElementsOffset * 56.7
   End If

   CheckAreaPrint
   FontSet FONT_TITLES
   LinePrint "(Controles Formulario)"
   FontSet FONT_CODE
   FeedLine

   If frmMain.chkIndex = vbChecked Then                           ' INDEX - Update index-page reference
      Idx.CICount = Idx.CICount + 1
      ReDim Preserve Idx.ControlIndex(0 To Idx.CICount)
      Idx.ControlIndex(Idx.CICount).File = Mdl(nMdlIndex).File
      Idx.ControlIndex(Idx.CICount).Page = nPage
   End If

   For i = 0 To frmMain.lstNames(nIndex).ListCount - 1

      If UserAbort Then Exit For

      j = frmMain.lstNames(nIndex).ItemData(i)

      SetCurrentX nNameOffset:    LinePrint Mdl(nMdlIndex).Control(j).Name, False
      If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
      SetCurrentX nTypeOffset:    LinePrint Mdl(nMdlIndex).Control(j).Type, False
      If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
      SetCurrentX nLibraryOffset: LinePrint Mdl(nMdlIndex).Control(j).Library, False

      If Mdl(nMdlIndex).Control(j).Elements > 1 Then
         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nElementsOffset: LinePrint "Elementos: " & Mdl(nMdlIndex).Control(j).Elements
      Else
         LinePrint ""
      End If
   Next i

   If UserAbort Then Exit Sub

   FeedLine
   LinePrint "   Total nombres control: " & Mdl(nMdlIndex).CtrlCount
   LinePrint "Total controles: " & Mdl(nMdlIndex).CtrlElements

   ' That's it. Either print line or go to next page
   If frmMain.chkControlPage = vbChecked Or frmMain.chkProcPage = vbChecked Then
      bNextPage = True

   ElseIf frmMain.chkCode = vbChecked Or frmMain.chkIcon = vbChecked Then
      ' Only print a separator if code is following
      FeedLine
      SeperatorPrint
   End If

End Sub

' --------------------------------------------------------------------------------------------------------

Private Sub PrintIndexPage()
   Dim i As Integer, j As Integer, nIndex As Integer
   Dim sString As String
   Dim H As Single, W As Single, nPointY As Single, _
       nPageOffset As Single, nTypeOffset As Single, nFileOffset As Single

   If Page.Cancelled Then Exit Sub

   ' Any index info to be printed?
   If Idx.CICount < 0 And Idx.DIcount < 0 And Idx.PIcount < 0 Then Exit Sub
   ' Yep.

   If Not InDevelopmentMode Then
      On Error GoTo PIP_ErrorHandler
   End If

   bNextPage = True          ' Always print index on a new page
   nMdlIndex = -2            ' Let Header() now knows it's a index page

   nPageOffset = GetWidth * 0.9

   If Idx.CICount > -1 Then  ' Controls index -------------------------------------------------------------
      If UserAbort Then GoTo PrintIndexAbort

      ' Load list into listbox (for optional sorting)
      nIndex = IIf(frmMain.chkSortIndex = vbChecked, 1, 0)
      frmMain.lstNames(nIndex).Clear
      For i = 0 To Idx.CICount
         frmMain.lstNames(nIndex).AddItem Idx.ControlIndex(i).File
         frmMain.lstNames(nIndex).ItemData(frmMain.lstNames(nIndex).NewIndex) = i
      Next

      If Page.Output = OUT_RTF Then
         lp = AddLayout(LYO_TABS, 1, 0)
         ReDim Layout(lp).Tabs(0 To 0)
         Layout(lp).Tabs(0) = nPageOffset * 56.7
      End If

      FontSet FONT_TITLES
      SetCurrentX 0:           LinePrint "Form Control Object INDEX"
      If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
      FontSet FONT_CODE
      SetCurrentX nPageOffset: LinePrint "Page", False
      FeedLine

      For i = 0 To frmMain.lstNames(nIndex).ListCount - 1

         If UserAbort Then Exit For

         j = frmMain.lstNames(nIndex).ItemData(i)

         If Page.Ruler = RULER_CHAR Then
            SetCurrentX 1: LinePrint Idx.ControlIndex(j).File, False
         Else
            If Page.Output = OUT_DRIVER Then
               If (i Mod 2) <> 0 Then
                  H = GetCurrentY
                  nPointY = H + (GetTextHeight(Idx.ControlIndex(j).File) / 2) + 1
               End If
            End If
            SetCurrentX (GetWidth * 0.01): LinePrint Idx.ControlIndex(j).File, False
            If Page.Output = OUT_DRIVER Then
               If (i Mod 2) <> 0 Then
                  W = (GetWidth * 0.9) - ((GetWidth * 0.01) + GetTextWidth(Idx.ControlIndex(j).File) + 250)
                  If W > 500 Then
                     PrintLine GetCurrentX + 125, nPointY, W, , , vbDot
                     SetCurrentY H
                  End If
               End If
            End If
         End If

         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nPageOffset: LinePrint CStr(Idx.ControlIndex(j).Page)
      Next
      FeedLine
      If Idx.DIcount > -1 Or Idx.PIcount > -1 Then SeperatorPrint
      frmMain.lstNames(nIndex).Clear
   End If

   If Idx.DIcount > -1 Then   ' Declarations index ------------------------------------------------------------------
      If UserAbort Then GoTo PrintIndexAbort

      ' Load list into listbox (for optional sorting)
      nIndex = IIf(frmMain.chkSortIndex = vbChecked, 1, 0)
      frmMain.lstNames(nIndex).Clear
      For i = 0 To Idx.DIcount
         frmMain.lstNames(nIndex).AddItem Idx.DeclareIndex(i).File
         frmMain.lstNames(nIndex).ItemData(frmMain.lstNames(nIndex).NewIndex) = i
      Next

      If Page.Output = OUT_RTF Then
         lp = AddLayout(LYO_TABS, 1, 0)
         ReDim Layout(lp).Tabs(0 To 0)
         Layout(lp).Tabs(0) = nPageOffset * 56.7
      End If

      FontSet FONT_TITLES
      SetCurrentX 0:           LinePrint "Declarations INDEX"
      If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
      FontSet FONT_CODE
      SetCurrentX nPageOffset: LinePrint "Page", False
      FeedLine

      For i = 0 To frmMain.lstNames(nIndex).ListCount - 1

         If UserAbort Then Exit For

         j = frmMain.lstNames(nIndex).ItemData(i)

         If Page.Ruler = RULER_CHAR Then
            SetCurrentX 1: LinePrint Idx.DeclareIndex(j).File, False
         Else
            If Page.Output = OUT_DRIVER Then
               If (i Mod 2) <> 0 Then
                  H = GetCurrentY
                  nPointY = H + (GetTextHeight(Idx.DeclareIndex(j).File) / 2) + 1
               End If
            End If
            SetCurrentX (GetWidth * 0.01): LinePrint Idx.DeclareIndex(j).File, False
            If Page.Output = OUT_DRIVER Then
               If (i Mod 2) <> 0 Then
                  W = (GetWidth * 0.9) - ((GetWidth * 0.01) + GetTextWidth(Idx.DeclareIndex(j).File) + 250)
                  If W > 500 Then
                     PrintLine GetCurrentX + 125, nPointY, W, , , vbDot
                     SetCurrentY H
                  End If
               End If
            End If
         End If

         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nPageOffset: LinePrint CStr(Idx.DeclareIndex(j).Page)
      Next
      FeedLine
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

      nTypeOffset = GetWidth * 0.4
      nFileOffset = GetWidth * 0.65
      nPageOffset = GetWidth * 0.9

      If Page.Output = OUT_RTF Then
         lp = AddLayout(LYO_TABS, 3, 2)
         ReDim Layout(lp).Tabs(0 To 2)
         Layout(lp).Tabs(0) = nTypeOffset * 56.7
         Layout(lp).Tabs(1) = nFileOffset * 56.7
         Layout(lp).Tabs(2) = nPageOffset * 56.7
      End If

      FontSet FONT_TITLES
      SetCurrentX 0:           LinePrint "Procedures INDEX"
      If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
      FontSet FONT_CODE
      SetCurrentX nTypeOffset: LinePrint "Type", False
      If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
      SetCurrentX nFileOffset: LinePrint "File", False
      If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
      SetCurrentX nPageOffset: LinePrint "Page", False
      FeedLine

      For i = 0 To frmMain.lstNames(nIndex).ListCount - 1

         If UserAbort Then Exit For

         j = frmMain.lstNames(nIndex).ItemData(i)

         If Page.Ruler = RULER_CHAR Then
            SetCurrentX 1: LinePrint Idx.ProcIndex(j).Procedure.Name, False
         Else
            If Page.Output = OUT_DRIVER Then
               If (i Mod 2) <> 0 Then
                  H = GetCurrentY
                  nPointY = H + (GetTextHeight(Idx.ProcIndex(j).Procedure.Name) / 2) + 1
               End If
            End If
            SetCurrentX (GetWidth * 0.01): LinePrint Idx.ProcIndex(j).Procedure.Name, False
            If Page.Output = OUT_DRIVER Then
               If (i Mod 2) <> 0 Then
                  W = nTypeOffset - ((GetWidth * 0.01) + GetTextWidth(Idx.ProcIndex(j).Procedure.Name) + 250)
                  If W > 500 Then
                     PrintLine GetCurrentX + 125, nPointY, W, , , vbDot
                     SetCurrentY H
                  End If
               End If
            End If
         End If

         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nTypeOffset: LinePrint Idx.ProcIndex(j).Procedure.Type, False

         If Page.Output = OUT_DRIVER Then
            If (i Mod 2) <> 0 Then
               W = nFileOffset - (nTypeOffset + GetTextWidth(Idx.ProcIndex(j).Procedure.Type) + 250)
               If W > 500 Then
                  PrintLine GetCurrentX + 125, nPointY, W, , , vbDot
                  SetCurrentY H
               End If
            End If
         End If

         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nFileOffset: LinePrint Idx.ProcIndex(j).File, False

         If Page.Output = OUT_DRIVER Then
            If (i Mod 2) <> 0 Then
               W = nPageOffset - (nFileOffset + GetTextWidth(Idx.ProcIndex(j).File) + 250)
               If W > 500 Then
                  PrintLine GetCurrentX + 125, nPointY, W, , , vbDot
                  SetCurrentY H
               End If
            End If
         End If

         If Page.Output = OUT_RTF Then LayoutMode LYO_TAB
         SetCurrentX nPageOffset: LinePrint CStr(Idx.ProcIndex(j).Page)
      Next

      frmMain.lstNames(nIndex).Clear
   End If

   If UserAbort Then GoTo PrintIndexAbort

   bFinalPage = True
   If nYAvailable > -1 Then FooterPrint

   Exit Sub

PIP_ErrorHandler:
   MsgBox "Problema al formatear informe." & vbCrLf & "Error #" & Err.Number & ": " & Err.Description, vbCritical

PrintIndexAbort:
   On Error Resume Next

End Sub

' --------------------------------------------------------------------------------------------------------

Sub InvalidateSamplePage()
    frmMain.picPage.Line (0, 0)-(frmMain.picPage.ScaleWidth, frmMain.picPage.ScaleHeight), QBColor(4)
    frmMain.picPage.Line (frmMain.picPage.ScaleWidth, 0)-(0, frmMain.picPage.ScaleHeight), QBColor(4)
End Sub

Private Sub SamplePrint(sString As String, Optional bCrLf)
   If IsMissing(bCrLf) Then bCrLf = True
   LinePrint sString, bCrLf

   If Page.Cancelled Then
      Err.Raise vbObjectError + 32755, "SamplePrint()", "Cancel was selected"
   End If
End Sub

' Gather print info first (Layout()), then show it in picturebox
'
Sub PrintSamplePage()
   If Page.Show Then Exit Sub

   Dim nColor As Long
   Dim i As Integer, nPointer As Integer
   Dim nHeightRatio As Double, nWidthRatio As Double
   Dim H As Single

   On Error GoTo EndOfSamplePage
   
   nPointer = frmMain.MousePointer
   frmMain.MousePointer = vbHourglass

   Page.Sample = True
   Page.Title = "Pagina de Ejemplo"
   Page.Count = 0                      ' Number of pages
   Page.Cancelled = False

   Page.Show = False                   ' Set by frmPreview or frmPrint
   Set Page.Form = Nothing

   ' Reset the margins
   Page.Margin.Left = 0
   Page.Margin.Right = 0
   Page.Margin.Top = 0
   Page.Margin.Bottom = 0

   Page.Output = OUT_DRIVER         ' Not really, but I have to something here...
   Page.File = ""

   Page.Ruler = RULER_MM            ' Set this first prior giving any dimensions/coordinates !!
   Page.Width = Printer.Width       ' Printer.Width and .Height are twips already...
   Page.Height = Printer.Height

   ' Find out the difference in size (the down-scaling) - used in preview font size setting (it doesn't adhere to .ScaleMode)
   nWidthRatio = frmMain.picPage.Width / Page.Width
   nHeightRatio = frmMain.picPage.Height / Page.Height

   ' Obtain smallest ratio
   If nHeightRatio < nWidthRatio Then
      nFontRatio = nHeightRatio
   Else
      nFontRatio = nWidthRatio
   End If
   
   ' Re-scale picturebox - do NOT resize it!!
   frmMain.picPage.Scale (0, 0)-(Page.Width, Page.Height)
   frmMain.picPage.Cls

   'validar
   If frmMain.lblTop(0) = "n/a" Then frmMain.lblTop(0) = "5"
   If frmMain.lblBottom(0) = "n/a" Then frmMain.lblBottom(0) = "5"
   If frmMain.lblLeft(0) = "n/a" Then frmMain.lblLeft(0) = "5"
   If frmMain.lblRight(0) = "n/a" Then frmMain.lblRight(0) = "5"
   
   ' Show margin areas -------------------------------------------------------------------------
   nColor = QBColor(8) ' RGB(0, 255, 255)  ' colour margin cyan
   frmMain.picPage.Line (0, 0)-(Page.Width, MM2Twips(frmMain.lblTop(0))), nColor, BF                              ' Top margin
   frmMain.picPage.Line (0, (Page.Height - MM2Twips(frmMain.lblBottom(0))))-(Page.Width, Page.Height), nColor, BF ' Bottom margin
   frmMain.picPage.Line (0, 0)-(MM2Twips(frmMain.lblLeft(0)), Page.Height), nColor, BF                            ' Left margin
   frmMain.picPage.Line ((Page.Width - MM2Twips(frmMain.lblRight(0))), 0)-(Page.Width, Page.Height), nColor, BF   ' Right margin

   ' Gather Layout() information.... -----------------------------------------------------------
   
   ' Set margin after margin indicators (areas) are placed
   Page.Margin.Left = CTwips(frmMain.lblLeft(0), True)
   Page.Margin.Right = CTwips(frmMain.lblRight(0), True)
   Page.Margin.Top = CTwips(frmMain.lblTop(0))
   Page.Margin.Bottom = CTwips(frmMain.lblBottom(0))

   ' Max height needs some attention
   If frmMain.chkHeader = vbChecked Then
      If Page.Output = OUT_PORT Then
         nYHeight = GetHeight - (2 + IIf(frmMain.chkFooter = vbChecked, 3, 0))
      Else
         FontSet FONT_HEADER, False
         H = GetTextHeight() + 4
         If frmMain.chkFooter = vbChecked Then
            FontSet FONT_FOOTER, False
            H = H + (2 * GetTextHeight(frmMain.txtOwner(0) + frmMain.txtOwner(1))) + 4
         End If
         nYHeight = GetHeight - H
      End If
   Else
      nYHeight = GetHeight
   End If
   
   nYAvailable = -1                  ' Force header to print on first line to be printed
   bNextPage = True

   bFinalPage = False
   nPage = 0

   PrintStartDoc

   If frmMain.chkIcon = vbChecked Then

      FontSet FONT_TITLES
      SamplePrint "(Form Icon)"
      FontSet FONT_CODE
      SamplePrint ""
   
      SetCurrentX 8!
   
      frmMain.picImage.Picture = frmMain.Icon
      PrintPicture -1, frmMain.picImage.Width, frmMain.picImage.Height

      frmMain.picImage.ScaleMode = vbMillimeters
      ReduceHeight frmMain.picImage.ScaleHeight
   
      ' Lose the picture, gain resources
      frmMain.picImage.Picture = LoadPicture()
   
      SamplePrint ""
      SeperatorPrint
   End If

   If frmMain.chkControlNames = vbChecked Then
      CheckAreaPrint
      FontSet FONT_TITLES
      SamplePrint "(Controles Formulario)"
      FontSet FONT_CODE
      SamplePrint ""

      For i = 1 To 4
         SetCurrentX (GetWidth * 0.01)
         SamplePrint "NombreControl", False

         SetCurrentX (GetWidth * 0.3)
         SamplePrint "Control", False

         SetCurrentX (GetWidth * 0.5)
         SamplePrint "(Libreria)", False

         If i = 2 Then
            SetCurrentX (GetWidth * 0.65)
            SamplePrint "Elementos: 99"
         Else
            SamplePrint ""
         End If
      Next

      SamplePrint ""
      SamplePrint "   Total nombres control: 4"
      SamplePrint "Total elementos control: 103"
      If frmMain.chkCode = vbChecked Then
         SamplePrint ""
         SeperatorPrint
      End If
   End If

   If frmMain.chkCode = vbChecked Then
      ' Show two routines
      If frmMain.chkProcNames <> vbChecked Then
         CheckAreaPrint
         FontSet FONT_PROCS
         SamplePrint "(Declaraciones)"
         FontSet FONT_CODE
         SamplePrint ""
         SamplePrint "Option explicit"
         SamplePrint ""
         SamplePrint "Public Const LINE_HORINZONTAL As Integer = 0"
         SamplePrint "Public Const LINE_VERTICAL As Integer = 1"
         SamplePrint ""
         FontSet FONT_COMMENTS
         SamplePrint "' Line is always vertical or horinzontal, else use PrintDraw()"
         SamplePrint "' Assumes horinzontal as default"
         FontSet FONT_CODE
      End If
      CheckAreaPrint
      FontSet FONT_PROCS
      SamplePrint "Sub PrintLine(nLeft As Single, nTop As Single, nLength As Single, Optional nDirection, Optional nColor)"
      FontSet FONT_CODE
      If frmMain.chkProcNames <> vbChecked Then
         SamplePrint "  If IsMissing(nDirection) Then nDirection = LINE_HORINZONTAL"
         SamplePrint ""
         SamplePrint "  If bSendtoPrinter Then"
         SamplePrint "    If IsMissing(nColor) Then"
         SamplePrint "      If nDirection = LINE_VERTICAL Then"
         SamplePrint "        Printer.Line (nMargin.Left + nLeft, nMargin.Top + nTop)-(nMargin.Left + nLeft, nMargin.Top + nTop + nLength)"
         SamplePrint "      Else"
         SamplePrint "        Printer.Line (nMargin.Left + nLeft, nMargin.Top + nTop)-(nMargin.Left + nLeft + nLength, nMargin.Top + nTop)"
         SamplePrint "      End If"
         SamplePrint "  End If"
         SamplePrint "End Sub"
      End If

      If frmMain.chkProcPage <> vbChecked Then
         If frmMain.chkSeparator = vbChecked Then SeperatorPrint

         If frmMain.chkProcNames <> vbChecked Then
            SamplePrint ""
            FontSet FONT_COMMENTS
            SamplePrint "' Checks wether file exist (handles wildcards too)"
         End If
         FontSet FONT_PROCS
         SamplePrint "Public Function FileExist(ByVal sFile As String) As Boolean"
         FontSet FONT_CODE
         If frmMain.chkProcNames <> vbChecked Then
            SamplePrint ""
            SamplePrint "  If Len(Trim(sFile)) = 0 Then"
            FontSet FONT_COMMENTS
            SamplePrint "    ' Nothing given"
            FontSet FONT_CODE
            SamplePrint "    FileExist = False"
            SamplePrint "    Exit Function"
            SamplePrint "  ElseIf Right(sFile, 1) = " & Chr(34) & "\" & Chr(34) & " Or Right(sFile, 1) = " & Chr(34) & ":" & Chr(34) & " Then"
            FontSet FONT_COMMENTS
            SamplePrint "    ' Just a part of a path or drive... (not complete)"
            FontSet FONT_CODE
            SamplePrint "    FileExist = False"
            SamplePrint "    Exit Function"
            SamplePrint "  ElseIf Dir(sFile) = " & Chr(34) & Chr(34) & " Then"
            FontSet FONT_COMMENTS
            SamplePrint "    ' Not there..."
            FontSet FONT_CODE
            SamplePrint "    FileExist = False"
            SamplePrint "    Exit Function"
            SamplePrint "  End If"
            SamplePrint ""
            FontSet FONT_COMMENTS
            SamplePrint "  ' After all that torture, it must exist..."
            FontSet FONT_CODE
            SamplePrint "  FileExist = True"
            SamplePrint "  Exit Function"
            SamplePrint "ExistErrorHandler:"
            SamplePrint "  FileExist = False"
            SamplePrint "End Function"
         End If
         If frmMain.chkSeparator = vbChecked Then SeperatorPrint
      End If
   End If

   If frmMain.chkControlNames <> vbChecked And _
      frmMain.chkCode <> vbChecked And _
      frmMain.chkIcon <> vbChecked Then
      HeaderPrint
   End If

EndOfSamplePage:

   If nYAvailable > -1 Then FooterPrint

   If InDevelopmentMode Then
      If Err.Number <> (vbObjectError + 32755) Then
         ReportError "Problemas refrescando pagina de ejemplo", Err.Number
      End If
   End If

   frmMain.MousePointer = nPointer
End Sub

' Analyse Layout and print...
Sub ShowSamplePage()
   Dim i As Integer, nStyle As Integer

   For i = 1 To nLineCount
      'Debug.Print i
      Select Case Layout(i).Mode
      Case LYO_TEXT
         frmMain.picPage.CurrentX = Layout(i).X
         frmMain.picPage.CurrentY = Layout(i).Y
         frmMain.picPage.Print Layout(i).Text.Text;

      Case LYO_FONT
         frmMain.picPage.FontName = Layout(i).Font.Name
         frmMain.picPage.FONTSIZE = Layout(i).Font.Size * nFontRatio
         frmMain.picPage.ForeColor = Layout(i).Font.Color
         frmMain.picPage.FontBold = Layout(i).Font.Bold
         frmMain.picPage.FontItalic = Layout(i).Font.Italic
         frmMain.picPage.FontStrikethru = Layout(i).Font.Strikethru
         frmMain.picPage.FontUnderline = Layout(i).Font.Underline

      Case LYO_LINE
         nStyle = frmMain.picPage.DrawStyle
         frmMain.picPage.DrawStyle = Layout(i).Line.Style
         frmMain.picPage.Line (Layout(i).X, Layout(i).Y)-(Layout(i).Line.Width, Layout(i).Line.Height), Layout(i).Line.Color
         frmMain.picPage.DrawStyle = nStyle

      Case LYO_BOX
         nStyle = frmMain.picPage.DrawStyle
         frmMain.picPage.DrawStyle = Layout(i).Line.Style
         frmMain.picPage.Line (Layout(i).X, Layout(i).Y)-(Layout(i).Line.Width, Layout(i).Line.Height), Layout(i).Line.Color, B
         frmMain.picPage.DrawStyle = nStyle

      Case LYO_FILLBOX
         nStyle = frmMain.picPage.FillStyle
         frmMain.picPage.DrawStyle = Layout(i).Line.Style
         frmMain.picPage.Line (Layout(i).X, Layout(i).Y)-(Layout(i).Line.Width, Layout(i).Line.Height), Layout(i).Line.Color, BF
         frmMain.picPage.FillStyle = nStyle

      Case LYO_IMAGE
         ' Save picture data to disk, then load with LoadPicture into a picturebox and then use it.
         If LoadIcon(Layout(i).Image.Index) Then
            frmMain.picPage.PaintPicture frmMain.picImage.Picture, Layout(i).X, Layout(i).Y, Layout(i).Image.Width, Layout(i).Image.Height
         End If

      Case LYO_CIRCLE
         frmMain.picPage.Circle (Layout(i).X, Layout(i).Y), Layout(i).Circles.Radius, Layout(i).Circles.Color

      End Select
   Next

End Sub
