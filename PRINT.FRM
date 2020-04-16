VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPrint 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir"
   ClientHeight    =   1665
   ClientLeft      =   6945
   ClientTop       =   6855
   ClientWidth     =   4950
   Icon            =   "Print.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1665
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPause 
      Caption         =   "Detener"
      Default         =   -1  'True
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   765
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   1665
      TabIndex        =   4
      Top             =   -60
      Width           =   3255
      Begin ComctlLib.ProgressBar ProgressBar 
         Height          =   285
         Left            =   60
         TabIndex        =   5
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         _Version        =   327682
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label lblPrinter 
         Alignment       =   2  'Center
         Height          =   210
         Left            =   75
         TabIndex        =   8
         Top             =   405
         Width           =   3105
      End
      Begin VB.Label lblJob 
         Alignment       =   2  'Center
         Caption         =   "Espere ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   7
         Top             =   165
         Width           =   3105
      End
      Begin VB.Label lblPage 
         Alignment       =   2  'Center
         Caption         =   "Inicializando"
         Height          =   210
         Left            =   75
         TabIndex        =   6
         Top             =   645
         Width           =   3105
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   30
         X2              =   3240
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line Line 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   30
         X2              =   3240
         Y1              =   885
         Y2              =   885
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   1320
      Width           =   2400
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   -75
      Width           =   1620
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1430
         Left            =   285
         ScaleHeight     =   1395
         ScaleWidth      =   990
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1015
         Begin VB.Image imgSleep 
            Height          =   480
            Left            =   270
            Picture         =   "Print.frx":030A
            Top             =   435
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Print the pages as they come... (unlike Preview who's stores it)
' Page number is determined by the caller.

'Dim Copy As Integer
'Dim Copies As Integer
'Dim FromPage As Integer
'Dim ToPage As Integer
'Dim Page As Integer

Private Const PAGE_MARK As String = "-<PB>-"

Dim nPgPrinted As Integer        ' Number of pages printer (to show on main form)

Dim nMaxPages As Integer
Dim nPageCount As Integer
Dim nCurrentPage As Integer

Dim bJobStarted As Boolean       ' True if stuff is opened for printing
Dim bJobDriver As Boolean        ' True if the job was send to printer driver (else a file)
Dim nJobHandle As Integer        ' File handle for non-driver jobs
Dim nFontRatio As Double         ' Used to scale fonts in preview picturebox (set in OpenPrintJob())

Dim bPaused As Boolean

Private Sub cmdPause_Click()
   If bPaused Then
      bPaused = False
      SetCaption cmdPause, "Detener"
      SetVisible imgSleep, False
   Else
      bPaused = True
      SetCaption cmdPause, "Continuar"
      SetVisible imgSleep, True
   End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Private Sub Form_Load()

   If Printer.Orientation = vbPRORLandscape Then
      picPaper.Width = 1430
      picPaper.Height = 1015
      picPaper.Top = 390
      picPaper.Left = 90
   Else
      picPaper.Width = 1015
      picPaper.Height = 1430
      picPaper.Top = 180
      picPaper.Left = 285
   End If

   SetCaption lblJob, Page.Title
   SetCaption lblPrinter, frmMain.lblPrinter  ' Printer.DeviceName & " on " & Printer.Port
   SetCaption lblPage, "Inicializando"

   ShowProgress 0

   FormStayOnTop Me, True
   'CentreForm Me

   bPaused = False

   Page.Show = True
   Set Page.Form = Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Page.Cancelled = True
   Page.Show = False                ' No more calls to this form
   Set Page.Form = Nothing

   If Page.Cancelled Then
      KillPrintJob
   Else
      ClosePrintJob
   End If

   FormStayOnTop Me, False
End Sub

Private Sub cmdCancel_Click()
   ShowDialog "Cancelando"
   Page.Cancelled = True
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Public Sub ShowPageNumber(nPage As Integer, nMaxPages As Integer)
   If nPage = 0 Then
      ShowDialog "Pagina n/a"
   Else
      If nMaxPages > 0 Then
         ShowDialog "Pagina " & nPage & " de " & nMaxPages
      Else
         ShowDialog "Pagina " & nPage
      End If
   End If
End Sub

Public Sub SetMaxPages(nPages As Integer)
   nMaxPages = nPages
End Sub

Private Sub ShowPrinted(bCount As Boolean)
   If bCount Then
      nPgPrinted = nPgPrinted + 1
      SetCaption frmMain.lblPrinted, CStr(nPgPrinted)
   Else
      nPgPrinted = 0
      SetCaption frmMain.lblPrinted, "ninguno"
   End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Public Sub ShowProgress(nPercent As Integer)
   If nPercent > 100 Then
      nPercent = 100
   ElseIf nPercent < 0 Then
      nPercent = 0
   End If
   ProgressBar.Value = nPercent
End Sub

Public Sub ShowDialog(sText As String)
   SetCaption lblPage, sText
End Sub

Public Sub PrintStart()
   ShowPrinted False

   SetCaption lblJob, Page.Title
   SetCaption lblPrinter, frmMain.lblPrinter  ' Printer.DeviceName & " on " & Printer.Port
   ShowDialog "Formateando"
   
   ShowProgress 0

   nMaxPages = 0
   nPageCount = 0
   nCurrentPage = 0
End Sub

Public Sub PrintEnd()

   If nPageCount = 0 Then
      ShowDialog "No hay paginas disponibles"
   ElseIf Page.Cancelled Then
      KillPrintJob
   Else
      ShowProgress 100
      ClosePrintJob
      If Page.Output = OUT_RTF Then
         SaveRTF
         ShowDialog "Guardado RTF a " & ExtractFileName(Page.File) & "(" & nPageCount & " pagina" & IIf(nPageCount = 1, ")", "s)")
         MsgBox "Trabajo de impresion finalizado." + vbCrLf + "RTF guardado a " & ExtractFileName(Page.File) & " (" & nPageCount & " pagina" & IIf(nPageCount = 1, ")", "s)"), vbInformation
      Else
         ShowDialog "Impreso " & nPageCount & " pagina" & IIf(nPageCount = 1, "", "s")
         MsgBox "Trabajo de impresion finalizado." + vbCrLf + "Impreso " & nPageCount & " pagina" & IIf(nPageCount = 1, "", "s"), vbInformation
      End If
   End If
End Sub

Public Sub PrintNewPage()
   Dim i As Integer

   On Error Resume Next
   nPageCount = nPageCount + 1   ' Is always nice to know how many pages are going to be printed

   If Page.PageNo = 0 Then
      nCurrentPage = nCurrentPage + 1
   Else
      nCurrentPage = CInt(Page.PageNo)
   End If
   ShowPageNumber nCurrentPage, Page.Count

   ' Print this page NOW.....
   '
   For i = 1 To Page.Copies
      PrintPage
      If Page.Cancelled Then Exit For
   Next i

   If Page.Cancelled Then Unload Me

End Sub

Public Sub PrintCancel()
   ShowDialog "Cancelado"
   Unload Me
End Sub

' Used in registration form print - I leave this code here, just in case it iss used somewhere else (no details search done yet)
Public Sub PrintDirect()
   nPageCount = nPageCount + 1   ' Count registrations forms too!

   PrintPage
   If Page.Cancelled Then Unload Me
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' This is it - The whole application turns around this routine.
'
Private Sub PrintPage()
   On Error Resume Next

   DoEvents    ' Just in case user clicks on Close button.
   If Page.Cancelled Then Exit Sub

   Dim i As Integer, nStyle As Integer, nValue As Integer

   picPaper.Cls            ' Just a mini preview whiles printing

   If Not OpenPrintJob Then
      ShowDialog "Problemas al accesar dispositivo"
      Page.Cancelled = True
      Exit Sub
   End If

   nValue = ProgressBar.Value
   ShowProgress 0

   Select Case Page.Output
   Case OUT_DRIVER
      ' I have to do this to make sure the top line get's it's proper font setting
      With Printer
         .ScaleMode = vbTwips
         .CurrentX = 0
         .CurrentY = 0
         .FontName = "Arial"
         .FONTSIZE = 10
         .FontBold = True
         .FontItalic = True
      End With

      Printer.Print ""
      
      With Printer
         .FontBold = False
         .FontItalic = False
         .CurrentX = 0
         .CurrentY = 0
      End With
      ' What a hassle, isn't it?

'      FontSet FONT_HEADER, False
'      SetFontItalic True
'      SetFontBold True
'      PrintFont
'      PrintAt 0, 0, ""
'      SetFontItalic False
'      SetFontBold False
'      PrintFont
'      PrintAt 0, 0, ""

      For i = 1 To nLineCount

         If bPaused Then
            Do While bPaused
               DoEvents
               If Page.Cancelled Then Exit Do
            Loop
         Else
            DoEvents    ' Just in case user clicks on Close button.
         End If
         If Page.Cancelled Then Exit For

         ShowProgress ((i * 100) / nLineCount)

         Select Case Layout(i).Mode
         Case LYO_TEXT
            picPaper.CurrentX = Layout(i).X
            picPaper.CurrentY = Layout(i).Y
            picPaper.Print Layout(i).Text.Text;
            '
            Printer.CurrentX = Layout(i).X
            Printer.CurrentY = Layout(i).Y
            Printer.Print Layout(i).Text.Text;

         Case LYO_FONT
            picPaper.FontName = Layout(i).Font.Name
            picPaper.FONTSIZE = Layout(i).Font.Size * nFontRatio
            picPaper.ForeColor = Layout(i).Font.Color
            picPaper.FontBold = Layout(i).Font.Bold
            picPaper.FontItalic = Layout(i).Font.Italic
            picPaper.FontStrikethru = Layout(i).Font.Strikethru
            picPaper.FontUnderline = Layout(i).Font.Underline
            '
            Printer.FontName = Layout(i).Font.Name
            Printer.FONTSIZE = Layout(i).Font.Size
            Printer.ForeColor = Layout(i).Font.Color
            Printer.FontBold = Layout(i).Font.Bold
            Printer.FontItalic = Layout(i).Font.Italic
            Printer.FontStrikethru = Layout(i).Font.Strikethru
            Printer.FontUnderline = Layout(i).Font.Underline

         Case LYO_LINE
            nStyle = picPaper.DrawStyle
            picPaper.DrawStyle = Layout(i).Line.Style
            picPaper.Line (Layout(i).X, Layout(i).Y)-(Layout(i).Line.Width, Layout(i).Line.Height), Layout(i).Line.Color
            picPaper.DrawStyle = nStyle
            '
            nStyle = Printer.DrawStyle
            Printer.DrawStyle = Layout(i).Line.Style
            Printer.Line (Layout(i).X, Layout(i).Y)-(Layout(i).Line.Width, Layout(i).Line.Height), Layout(i).Line.Color
            Printer.DrawStyle = nStyle

         Case LYO_BOX
            nStyle = picPaper.DrawStyle
            picPaper.DrawStyle = Layout(i).Line.Style
            picPaper.Line (Layout(i).X, Layout(i).Y)-(Layout(i).Line.Width, Layout(i).Line.Height), Layout(i).Line.Color, B
            picPaper.DrawStyle = nStyle
            '
            nStyle = Printer.DrawStyle
            Printer.DrawStyle = Layout(i).Line.Style
            Printer.Line (Layout(i).X, Layout(i).Y)-(Layout(i).Line.Width, Layout(i).Line.Height), Layout(i).Line.Color, B
            Printer.DrawStyle = nStyle

         Case LYO_FILLBOX
            nStyle = picPaper.FillStyle
            picPaper.DrawStyle = Layout(i).Line.Style
            picPaper.Line (Layout(i).X, Layout(i).Y)-(Layout(i).Line.Width, Layout(i).Line.Height), Layout(i).Line.Color, BF
            picPaper.FillStyle = nStyle
            '
            nStyle = Printer.FillStyle
            Printer.DrawStyle = Layout(i).Line.Style
            Printer.Line (Layout(i).X, Layout(i).Y)-(Layout(i).Line.Width, Layout(i).Line.Height), Layout(i).Line.Color, BF
            Printer.FillStyle = nStyle

         Case LYO_IMAGE
            ' Save picture data to disk, then load with LoadPicture into a picturebox and then use it.
            If LoadIcon(Layout(i).Image.Index) Then
               picPaper.PaintPicture frmMain.picImage.Picture, Layout(i).X, Layout(i).Y, Layout(i).Image.Width, Layout(i).Image.Height
               '
               Printer.PaintPicture frmMain.picImage.Picture, Layout(i).X, Layout(i).Y, Layout(i).Image.Width, Layout(i).Image.Height
            End If

         Case LYO_CIRCLE
            picPaper.Circle (Layout(i).X, Layout(i).Y), Layout(i).Circles.Radius, Layout(i).Circles.Color
            '
            Printer.Circle (Layout(i).X, Layout(i).Y), Layout(i).Circles.Radius, Layout(i).Circles.Color

         End Select
      Next

      If Not Page.Cancelled Then
         ' Feed the page....
         Printer.NewPage
      End If

   Case OUT_RTF
      Dim sFontName As String
      Dim nFontSize As Integer
      Dim nColor As Long
      Dim bBold As Boolean, bItalic As Boolean, bStrikethru As Boolean, bUnderline As Boolean

      For i = 1 To nLineCount

         If bPaused Then
            Do While bPaused
               DoEvents
               If Page.Cancelled Then Exit Do
            Loop
         Else
            DoEvents    ' Just in case user clicks on Close button.
         End If
         If Page.Cancelled Then Exit For

         ShowProgress ((i * 100) / nLineCount)

         Select Case Layout(i).Mode
         Case LYO_PAGE
            frmMain.RTBox.SelText = PAGE_MARK
         Case LYO_EOL
            frmMain.RTBox.SelText = vbCrLf

         Case LYO_TAB
            frmMain.RTBox.SelText = vbTab       ' "\tab "
         Case LYO_TABS
            frmMain.RTBox.SelTabCount = Layout(i).X
            If Layout(i).X > 0 Then
               For nValue = 0 To Layout(i).Y
                  frmMain.RTBox.SelTabs(nValue) = Layout(i).Tabs(nValue)
               Next
            End If

         Case LYO_TEXT
            picPaper.CurrentX = Layout(i).X
            picPaper.CurrentY = Layout(i).Y
            picPaper.Print Layout(i).Text.Text;
            '
            frmMain.RTBox.SelText = Layout(i).Text.Text

         Case LYO_FONT
            picPaper.FontName = Layout(i).Font.Name
            picPaper.FONTSIZE = Layout(i).Font.Size * nFontRatio
            picPaper.ForeColor = Layout(i).Font.Color
            picPaper.FontBold = Layout(i).Font.Bold
            picPaper.FontItalic = Layout(i).Font.Italic
            picPaper.FontStrikethru = Layout(i).Font.Strikethru
            picPaper.FontUnderline = Layout(i).Font.Underline
            '
            frmMain.RTBox.SelFontName = Layout(i).Font.Name
            frmMain.RTBox.SelFontSize = Layout(i).Font.Size
            frmMain.RTBox.SelColor = Layout(i).Font.Color
            frmMain.RTBox.SelBold = Layout(i).Font.Bold
            frmMain.RTBox.SelItalic = Layout(i).Font.Italic
            frmMain.RTBox.SelStrikeThru = Layout(i).Font.Strikethru
            frmMain.RTBox.SelUnderline = Layout(i).Font.Underline

         Case LYO_LINE
            ' Only horinzontal solid lines accepted
            nStyle = picPaper.DrawStyle
            picPaper.DrawStyle = Layout(i).Line.Style
            picPaper.Line (Layout(i).X, Layout(i).Y)-(Layout(i).Line.Width, Layout(i).Line.Height), Layout(i).Line.Color
            picPaper.DrawStyle = nStyle
            '
            sFontName = frmMain.RTBox.SelFontName
            nFontSize = frmMain.RTBox.SelFontSize
            nColor = frmMain.RTBox.SelColor
            bBold = frmMain.RTBox.SelBold
            bItalic = frmMain.RTBox.SelItalic
            bStrikethru = frmMain.RTBox.SelStrikeThru
            bUnderline = frmMain.RTBox.SelUnderline
            '
            frmMain.RTBox.SelFontName = "Courier New"
            frmMain.RTBox.SelFontSize = 10
            frmMain.RTBox.SelColor = Layout(i).Line.Color
            frmMain.RTBox.SelBold = False
            frmMain.RTBox.SelItalic = False
            frmMain.RTBox.SelStrikeThru = False
            frmMain.RTBox.SelUnderline = False

            If Layout(i).X > 0 Then
               nValue = Int(Layout(i).X / 120)
               frmMain.RTBox.SelText = Space$(nValue)
            End If
            nValue = Int((Layout(i).Line.Width - Layout(i).X) / 120)

            Select Case Layout(i).Line.Style
            Case 0, 1, 3, 6   ' Solid, Dash, Dash-Dot, inside solid.
               frmMain.RTBox.SelText = String$(nValue, "_") & vbCrLf
            Case 2, 4         ' Dot, Dash-Dot-Dot.
               frmMain.RTBox.SelText = Replicate(nValue, "_ ") & vbCrLf
            End Select

            frmMain.RTBox.SelFontName = sFontName
            frmMain.RTBox.SelFontSize = nFontSize
            frmMain.RTBox.SelColor = nColor
            frmMain.RTBox.SelBold = bBold
            frmMain.RTBox.SelItalic = bItalic
            frmMain.RTBox.SelStrikeThru = bStrikethru
            frmMain.RTBox.SelUnderline = bUnderline

         End Select
      Next

   Case OUT_PORT
      Dim nLast As Integer

      If Page.FormFeed Then
         ' Find the last line
         nLast = -1
         For i = nLineCount To 1 Step -1
            If Len(Layout(i).Text.Text) > 0 Then
               nLast = i
               Exit For
            End If
         Next
      Else
         nLast = nLineCount
      End If
      
      For i = 1 To nLast

         If bPaused Then
            Do While bPaused
               DoEvents
               If Page.Cancelled Then Exit Do
            Loop
         Else
            DoEvents    ' Just in case user clicks on Close button.
         End If
         If Page.Cancelled Then Exit For

         ShowProgress ((i * 100) / nLineCount)

         picPaper.CurrentX = Layout(i).X
         picPaper.CurrentY = Layout(i).Y
         picPaper.Print Layout(i).Text.Text;
         '
         Print #nJobHandle, Layout(i).Text.Text

      Next

      If Not Page.Cancelled Then
         If Page.FormFeed Then
            Print #nJobHandle, vbFormFeed
         End If
      End If

   End Select

   If Not Page.Cancelled Then
      ShowPrinted True
      ProgressBar.Value = nValue
   End If

End Sub

Private Function OpenPrintJob()
   If bJobStarted Then
      OpenPrintJob = True
      Exit Function
   End If

   On Error GoTo JobOpenError

   Dim nWidth As Long, nHeight As Long
   Dim nHeightRatio As Double, nWidthRatio As Double

   If Page.Ruler = RULER_CHAR Then
      nWidth = Char2Twips(Page.Width)
      nHeight = Line2Twips(Page.Height)
   Else
      nWidth = Page.Width
      nHeight = Page.Height
   End If

   ' Find out the difference in size (the down-scaling) - used in preview font size setting (it doesn't adhere to .ScaleMode)
   nWidthRatio = picPaper.Width / nWidth
   nHeightRatio = picPaper.Height / nHeight

   ' Obtain smallest ratio
   If nHeightRatio < nWidthRatio Then
      nFontRatio = nHeightRatio
   Else
      nFontRatio = nWidthRatio
   End If
   
   ' Re-scale picturebox - do NOT resize it!!
   picPaper.Scale (0, 0)-(nWidth, nHeight)

   If Page.Ruler = RULER_CHAR Then
      picPaper.FontName = "Courier New"
      picPaper.FONTSIZE = 10 * nFontRatio
      picPaper.ForeColor = QBColor(0)
      picPaper.FontBold = False
      picPaper.FontItalic = False
      picPaper.FontStrikethru = False
      picPaper.FontUnderline = False
   End If

   ' "Open" the device
   Select Case Page.Output
   Case OUT_DRIVER
      SetCaption lblPrinter, Printer.DeviceName & " en " & Printer.Port
      bJobDriver = True
      nJobHandle = -1
      OpenPrintJob = True

   Case OUT_RTF
      SetCaption lblPrinter, "Archivo RTF " & Page.File
      bJobDriver = False
      nJobHandle = -1
      frmMain.RTBox.RightMargin = Page.Width
      frmMain.RTBox.SelLength = 0
      frmMain.RTBox = ""
      OpenPrintJob = True

   Case OUT_PORT
      SetCaption lblPrinter, "Archivo plano en puerto " & Page.File
      bJobDriver = False
      Close
      nJobHandle = FreeFile
      Open Page.File For Output As #nJobHandle
      OpenPrintJob = True

   Case Else
      OpenPrintJob = False
   End Select

   bJobStarted = True

   Exit Function

JobOpenError:
   ReportError "Problemas al imprimir", Err.Number
   OpenPrintJob = False

End Function

Private Sub ClosePrintJob()
   On Error Resume Next
   If bJobStarted Then
      If bJobDriver Then
         Printer.EndDoc
      ElseIf nJobHandle > -1 Then
         Close #nJobHandle
      End If
   End If
   bJobStarted = False
End Sub

Private Sub KillPrintJob()
   On Error Resume Next
   If bJobStarted Then
      If bJobDriver Then
         Printer.KillDoc
      ElseIf nJobHandle > -1 Then
         Close #nJobHandle
      End If
   End If
   bJobStarted = False
End Sub

Private Sub SaveRTF()
   Dim nHandle As Integer, nCount As Integer
   Dim bOpenFile As Boolean
   Dim sText As String
   Dim nOffset As Long

   bOpenFile = False

   On Error GoTo SaveRTFError

   sText = frmMain.RTBox.TextRTF
   nCount = 0

   Do While True
      nOffset = InStr(sText, PAGE_MARK)
      If nOffset = 0 Then Exit Do
      nCount = nCount + 1

      If nCount = 1 Then
         ' No page break in the start...
         sText = Left(sText, nOffset - 1) & Mid(sText, nOffset + 6)
      Else
         Mid(sText, nOffset, 6) = "\page "
      End If
   Loop

   If FileExist(Page.File) Then Kill Page.File
   Close

   nHandle = FreeFile
   Open Page.File For Output As #nHandle
   bOpenFile = True

   Print #nHandle, sText

   Close #nHandle
   Exit Sub

SaveRTFError:
   If bOpenFile Then Close #nHandle
End Sub
