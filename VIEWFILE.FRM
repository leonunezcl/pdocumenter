VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmViewFile 
   Caption         =   "Visualizador"
   ClientHeight    =   1470
   ClientLeft      =   735
   ClientTop       =   2025
   ClientWidth     =   9180
   Icon            =   "ViewFile.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1470
   ScaleWidth      =   9180
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ilButtons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   26
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Save"
            Description     =   ""
            Object.ToolTipText     =   "Guardar a archivo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Print"
            Description     =   ""
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Find"
            Description     =   ""
            Object.ToolTipText     =   "Buscar texto"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Cut"
            Description     =   ""
            Object.ToolTipText     =   "Cortar al portapapeles"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Copy"
            Description     =   ""
            Object.ToolTipText     =   "Copiar al portapapeles"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Paste"
            Description     =   ""
            Object.ToolTipText     =   "Pegar desde el portapapeles"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Undo"
            Description     =   ""
            Object.ToolTipText     =   "Deshacer"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Left"
            Description     =   ""
            Object.ToolTipText     =   "Justificar izquierda"
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   2
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Centre"
            Description     =   ""
            Object.ToolTipText     =   "Centrar texto"
            Object.Tag             =   ""
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Right"
            Description     =   ""
            Object.ToolTipText     =   "Justificar derecha"
            Object.Tag             =   ""
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Bold"
            Description     =   ""
            Object.ToolTipText     =   "Negrita"
            Object.Tag             =   ""
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Italic"
            Description     =   ""
            Object.ToolTipText     =   "Italica"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Strikethru"
            Description     =   ""
            Object.ToolTipText     =   "Subrayado"
            Object.Tag             =   ""
            ImageIndex      =   14
            Style           =   1
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            Object.Tag             =   ""
            ImageIndex      =   15
            Style           =   1
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Color"
            Description     =   ""
            Object.ToolTipText     =   "Color Fuente"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Font"
            Description     =   ""
            Object.ToolTipText     =   "Fuente"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button24 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Sample"
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   2000
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button25 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button26 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Exit"
            Description     =   ""
            Object.ToolTipText     =   "Salir Visualizador"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
      EndProperty
   End
   Begin VB.TextBox lblSample 
      Height          =   285
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Font Sample"
      Top             =   60
      Width           =   2000
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   60
      Top             =   435
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox 
      Height          =   1065
      Left            =   0
      TabIndex        =   0
      Top             =   390
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   1879
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"ViewFile.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ilButtons 
      Left            =   540
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   18
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":03DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":04F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":0602
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":0714
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":0826
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":0938
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":0A4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":0B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":0C6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":0D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":0E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":0FA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":10B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":11C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":12DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":13EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":14FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ViewFile.frx":1818
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmViewFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cbRTF As Integer = &HBF01

Dim bHelpView As Boolean

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Private Sub Form_Load()
   If WinView.State > -1 Then
      ScaleMode = vbTwips
      Me.Left = WinView.Left
      Me.Top = WinView.Top
      Me.Width = WinView.Width
      Me.Height = WinView.Height
      Me.WindowState = WinView.State
   Else
      Width = Screen.Width * 0.75
      If Width < 9300 Then Width = 9300
      Height = Screen.Height * 0.5
      CentreForm Me
   End If

   bHelpView = False

   With lblSample
      .Width = Toolbar.Buttons("Sample").Width
      .Top = Toolbar.Buttons("Sample").Top + ((Toolbar.Buttons("Sample").Height - .Height) / 2)
      .Left = Toolbar.Buttons("Sample").Left
      .ZOrder 0
   End With
   
   Toolbar.ZOrder 1
End Sub

Private Sub Form_Activate()
   lblSample.ZOrder 0
End Sub

Private Sub Form_Resize()
   Static bResizing As Boolean
   If WindowState = vbMinimized Then Exit Sub

   If bResizing Then Exit Sub
   bResizing = True

   With lblSample
      .Width = Toolbar.Buttons("Sample").Width
      .Top = Toolbar.Buttons("Sample").Top + ((Toolbar.Buttons("Sample").Height - .Height) / 2)
      .Left = Toolbar.Buttons("Sample").Left
      .ZOrder 0
   End With

   RichTextBox.Top = Toolbar.Top + Toolbar.Height
   RichTextBox.Width = Width - 130
   RichTextBox.Height = Height - (Toolbar.Height + 385)
   RichTextBox.RightMargin = RichTextBox.Width - 400

   Toolbar.ZOrder 1

   bResizing = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Unload frmFind

   ScaleMode = vbTwips
   If WindowState <> vbMinimized Then
      WinView.Left = Me.Left
      WinView.Top = Me.Top
      WinView.Width = Me.Width
      WinView.Height = Me.Height
      WinView.State = Me.WindowState
   End If
End Sub

Public Sub InitView()
   
   With Toolbar
      .Buttons("Save").Enabled = True
      .Buttons("Cut").Enabled = True
      .Buttons("Paste").Enabled = True
      .Buttons("Undo").Enabled = True
      .Buttons("Left").Enabled = True
      .Buttons("Centre").Enabled = True
      .Buttons("Right").Enabled = True
      .Buttons("Bold").Enabled = True
      .Buttons("Italic").Enabled = True
      .Buttons("Strikethru").Enabled = True
      .Buttons("Underline").Enabled = True
      .Buttons("Font").Enabled = True
      .Buttons("Color").Enabled = True
      .ZOrder 1
   End With

   bHelpView = False
   RichTextBox.Locked = False
   RichTextBox_SelChange

   lblSample.ZOrder 0
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Public Sub SetText(sName As String, sText As String)
   If EmptyString(sText) Then
      RichTextBox.SelColor = RGB(255, 0, 0)
      RichTextBox.TextRTF = "No hay texto que visualizar."
      Caption = "Visualizar"
   Else
      RichTextBox.SelColor = RGB(0, 0, 0)
      RichTextBox.Text = sText
      Caption = sName & " - Visualizar"
   End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Public Sub SetLine(sText As String)
   RichTextBox.SelText = sText
End Sub

Public Sub SetFont(nIndex As Integer)
   RichTextBox.SelFontName = frmMain.lblFont(nIndex).FontName
   RichTextBox.SelFontSize = frmMain.lblFont(nIndex).FontSize
   RichTextBox.SelColor = frmMain.lblFont(nIndex).ForeColor
   RichTextBox.SelBold = frmMain.lblFont(nIndex).FontBold
   RichTextBox.SelItalic = frmMain.lblFont(nIndex).FontItalic
   RichTextBox.SelStrikeThru = frmMain.lblFont(nIndex).FontStrikethru
   RichTextBox.SelUnderline = frmMain.lblFont(nIndex).FontUnderline
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Public Sub SetFileName(sFileName As String)
   On Error Resume Next
   If FileExist(sFileName) Then
      RichTextBox.SelColor = RGB(0, 0, 0)
      If UCase$(ExtractFileExt(sFileName)) = "RTF" Then
         RichTextBox.LoadFile sFileName, rtfRTF
      Else
         RichTextBox.LoadFile sFileName, rtfText
      End If
      Caption = sFileName
   Else
      RichTextBox.SelColor = RGB(255, 0, 0)
      RichTextBox.TextRTF = "Archivo seleccionado a visualizar no fue encontrado."
      Caption = "Visualizar"
   End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Public Sub ShowHelpFile()
   On Error Resume Next
   Caption = "Ayuda Project Printer"
   bHelpView = True

   If FileExist(sHelpFile) Then
      RichTextBox.SelColor = RGB(0, 0, 0)
      RichTextBox.LoadFile sHelpFile, rtfRTF
   Else
      RichTextBox.SelColor = RGB(255, 0, 0)
      RichTextBox.TextRTF = "Archivo ayuda ('" & sHelpFile & "') no fue encontrado."
   End If

   With Toolbar
      .Buttons("Save").Enabled = False
      .Buttons("Cut").Enabled = False
      .Buttons("Paste").Enabled = False
      .Buttons("Undo").Enabled = False
      .Buttons("Left").Enabled = False
      .Buttons("Centre").Enabled = False
      .Buttons("Right").Enabled = False
      .Buttons("Bold").Enabled = False
      .Buttons("Italic").Enabled = False
      .Buttons("Strikethru").Enabled = False
      .Buttons("Underline").Enabled = False
      .Buttons("Font").Enabled = False
      .Buttons("Color").Enabled = False
      .ZOrder 1
   End With
   RichTextBox.Locked = True
   RichTextBox_SelChange
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Private Sub PrintText()
   Dim nSelStart As Long, nSelLength As Long
   nSelStart = RichTextBox.SelStart
   nSelLength = RichTextBox.SelLength

   On Error GoTo PrintRTFCancelled

   With CommonDialog
      .Flags = cdlPDReturnDC Or cdlPDNoPageNums Or cdlPDUseDevModeCopies
      If RichTextBox.SelLength = 0 Then
         .Flags = .Flags Or cdlPDAllPages Or cdlPDNoSelection
      Else
         .Flags = .Flags Or cdlPDSelection
      End If
      .CancelError = True
   End With

   CommonDialog.ShowPrinter

   If (CommonDialog.Flags And cdlPDSelection) <> cdlPDSelection Then
      RichTextBox.SelStart = 0
      RichTextBox.SelLength = 0
   End If

   RichTextBox.SelPrint CommonDialog.hDC

PrintRTFCancelled:
   CommonDialog.CancelError = False
   RichTextBox.SelStart = nSelStart
   RichTextBox.SelLength = nSelLength
   RichTextBox.SetFocus
End Sub

Private Sub SaveText()
   On Error GoTo SaveRTFCancelled

   With CommonDialog
      .DialogTitle = "Guardar texto como ..."
      .Filter = "Archivo RTF (*.rtf)|*.rtf|Text file (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
      .FilterIndex = 1
      .DefaultExt = ".rtf"
      .CancelError = True
      .Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
      .FileName = "VBCode.rtf"
   End With

   CommonDialog.ShowSave

   If UCase$(ExtractFileExt(CommonDialog.FileName)) = "RTF" Then
      RichTextBox.SaveFile CommonDialog.FileName, rtfRTF
   Else
      RichTextBox.SaveFile CommonDialog.FileName, rtfText
   End If

SaveRTFCancelled:
   CommonDialog.CancelError = False
   RichTextBox.SetFocus
End Sub

Private Sub UndoText()
   Dim dl As Long
   If SendMessage(RichTextBox.hwnd, EM_CANUNDO, 0, 0&) = 1 Then
      dl = SendMessage(RichTextBox.hwnd, EM_UNDO, 0, 0&)
   End If
End Sub

Private Sub SelectFont()
   On Error GoTo SelectFontCancel

   With CommonDialog
      .CancelError = True
      .Color = IIf(IsNull(RichTextBox.SelColor), RGB(0, 0, 0), RichTextBox.SelColor)
      .FontBold = IIf(IsNull(RichTextBox.SelBold), False, RichTextBox.SelBold)
      .FontItalic = IIf(IsNull(RichTextBox.SelItalic), False, RichTextBox.SelItalic)
      .FontStrikethru = IIf(IsNull(RichTextBox.SelStrikeThru), False, RichTextBox.SelStrikeThru)
      .FontUnderline = IIf(IsNull(RichTextBox.SelUnderline), False, RichTextBox.SelUnderline)
      .FontName = IIf(IsNull(RichTextBox.SelFontName), "", RichTextBox.SelFontName)
      .FontSize = IIf(IsNull(RichTextBox.SelFontSize), 8, RichTextBox.SelFontSize)
      .Flags = cdlCFEffects Or cdlCFForceFontExist Or cdlCFPrinterFonts Or cdlCFScalableOnly
   End With

   CommonDialog.ShowFont

   With CommonDialog
      RichTextBox.SelColor = .Color
      RichTextBox.SelBold = .FontBold
      RichTextBox.SelItalic = .FontItalic
      RichTextBox.SelStrikeThru = .FontStrikethru
      RichTextBox.SelUnderline = .FontUnderline
      RichTextBox.SelFontName = .FontName
      RichTextBox.SelFontSize = .FontSize
   End With

SelectFontCancel:
   CommonDialog.CancelError = False
   RichTextBox.SetFocus
End Sub

Private Sub SelectColor()
   On Error GoTo SelectColorCancel

   With CommonDialog
      .CancelError = True
      .Color = IIf(IsNull(RichTextBox.SelColor), RGB(0, 0, 0), RichTextBox.SelColor)
      .Flags = cdlCCRGBInit
   End With

   CommonDialog.ShowColor

   RichTextBox.SelColor = CommonDialog.Color

SelectColorCancel:
   CommonDialog.CancelError = False
   RichTextBox.SetFocus
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Private Sub Toolbar_ButtonClick(ByVal Button As Button)
   Select Case Button.Key
   Case "Save"
      SaveText
   Case "Print"
      PrintText
   Case "Find"
      frmFind.Show

   Case "Cut", "Copy"
      Clipboard.Clear
      If RichTextBox.SelLength = 0 Then
         ' Whole document
         Clipboard.SetText RichTextBox.TextRTF, cbRTF
         If Button.Key = "Cut" Then RichTextBox.Text = ""
      Else
         ' Selected text
         Clipboard.SetText RichTextBox.SelText
         If Button.Key = "Cut" Then RichTextBox.SelText = ""
      End If
      RichTextBox.SetFocus
   Case "Paste"
      If Clipboard.GetFormat(vbCFText) Then
         RichTextBox.SelText = Clipboard.GetText
      ElseIf Clipboard.GetFormat(cbRTF) Then
         RichTextBox.SelRTF = Clipboard.GetText(cbRTF)
      End If
      RichTextBox.SetFocus

   Case "Undo"
      UndoText
   'Case "Redo"
   '   RedoText

   Case "Left"
      RichTextBox.SelAlignment = rtfLeft
      RichTextBox.SetFocus
   Case "Centre"
      RichTextBox.SelAlignment = rtfCenter
      RichTextBox.SetFocus
   Case "Right"
      RichTextBox.SelAlignment = rtfRight
      RichTextBox.SetFocus

   Case "Exit"
      Unload Me
   Case Else

      Select Case Button.Key
      Case "Bold"
         If Button.MixedState = True Then Button.MixedState = False
         RichTextBox.SelBold = Abs(RichTextBox.SelBold) - 1
         RichTextBox.SetFocus
      Case "Italic"
         If Button.MixedState = True Then Button.MixedState = False
         RichTextBox.SelItalic = Abs(RichTextBox.SelItalic) - 1
         RichTextBox.SetFocus
      Case "Strikethru"
         If Button.MixedState = True Then Button.MixedState = False
         RichTextBox.SelStrikeThru = Abs(RichTextBox.SelStrikeThru) - 1
         RichTextBox.SetFocus
      Case "Underline"
         If Button.MixedState = True Then Button.MixedState = False
         RichTextBox.SelUnderline = Abs(RichTextBox.SelUnderline) - 1
         RichTextBox.SetFocus

      Case "Font"
         SelectFont

      Case "Color"
         SelectColor
      End Select

      RichTextBox_SelChange
   End Select
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Private Sub RichTextBox_SelChange()
   ' When the insertion point changes, set the Toolbar buttons
   ' to reflect the attributes of the text where the cursor is located.
   ' Use the Select Case statement.
   ' The SelAlignment property returns either 0, 1, 2, or Null.
   On Error Resume Next

   With RichTextBox
      lblSample = IIf(IsNull(.SelFontName), "", .SelFontName)
      lblSample.ForeColor = IIf(IsNull(.SelColor), RGB(0, 0, 0), .SelColor)
      lblSample.FontBold = IIf(IsNull(.SelBold), False, .SelBold)
      lblSample.FontItalic = IIf(IsNull(.SelItalic), False, .SelItalic)
      lblSample.FontStrikethru = IIf(IsNull(.SelStrikeThru), False, .SelStrikeThru)
      lblSample.FontUnderline = IIf(IsNull(.SelUnderline), False, .SelUnderline)
   End With

   Select Case RichTextBox.SelAlignment
   Case Is = rtfLeft ' 0
      Toolbar.Buttons("Left").Value = tbrPressed
   Case Is = rtfRight '1
      Toolbar.Buttons("Right").Value = tbrPressed
   Case Is = rtfCenter '2
      Toolbar.Buttons("Centre").Value = tbrPressed
   Case Else ' Null -- No buttons are shown in the up position.
      Toolbar.Buttons("Left").Value = tbrUnpressed
      Toolbar.Buttons("Right").Value = tbrUnpressed
      Toolbar.Buttons("Centre").Value = tbrUnpressed
   End Select

   ' SelBold returns 0, -1, or Null.  If it's Null then set
   ' the MixedState property to True.
   With Toolbar
   Select Case RichTextBox.SelBold
   Case 0 ' Not bold.
      .Buttons("Bold").MixedState = False
      .Buttons("Bold").Value = tbrUnpressed
   Case -1 ' Bold.
      .Buttons("Bold").MixedState = False
      .Buttons("Bold").Value = tbrPressed
   Case Else ' Mixed state.
      .Buttons("Bold").MixedState = True
   End Select

   ' SelItalic returns 0, -1, or Null.  If it's Null then set
   ' the MixedState property to True.
   Select Case RichTextBox.SelItalic
   Case 0 ' Not italic.
      .Buttons("Italic").MixedState = False
      .Buttons("Italic").Value = tbrUnpressed
   Case -1 ' Italic.
      .Buttons("Italic").MixedState = False
      .Buttons("Italic").Value = tbrPressed
   Case Else ' Mixed State.
      .Buttons("Italic").MixedState = True
   End Select

   ' SelStrikethru returns 0, -1, or Null.  If it's Null then set
   ' the MixedState property to True.
   Select Case RichTextBox.SelStrikeThru
   Case 0 ' Off
      .Buttons("Strikethru").MixedState = False
      .Buttons("Strikethru").Value = tbrUnpressed
   Case -1 ' On
      .Buttons("Strikethru").MixedState = False
      .Buttons("Strikethru").Value = tbrPressed
   Case Else ' Mixed State.
      .Buttons("Strikethru").MixedState = True
   End Select

   ' SelUnderline returns 0, -1, or Null.  If it's Null then set
   ' the MixedState property to True.
   Select Case RichTextBox.SelUnderline
   Case 0 ' Off
      .Buttons("Underline").MixedState = False
      .Buttons("Underline").Value = tbrUnpressed
   Case -1 ' On
      .Buttons("Underline").MixedState = False
      .Buttons("Underline").Value = tbrPressed
   Case Else ' Mixed State.
      .Buttons("Underline").MixedState = True
   End Select
   End With

   ButtonState

   lblSample.ZOrder 0
End Sub

Private Sub ButtonState()
   If bHelpView Then
      If RichTextBox.SelText = "" Then
         Toolbar.Buttons("Copy").ToolTipText = "Copia todo el texto al portapapeles (como RTF)"
      Else
         Toolbar.Buttons("Copy").ToolTipText = "Copy el texto selecionado al portapapeles"
      End If

   Else
      With Toolbar
         If RichTextBox.SelText = "" Then
            .Buttons("Cut").ToolTipText = "Corta todo el texto al portapapeles (como RTF)"
            .Buttons("Copy").ToolTipText = "Copia todo el texto al portapapeles (como RTF)"
         Else
            .Buttons("Cut").ToolTipText = "Corta el texto selecionado al portapapeles"
            .Buttons("Copy").ToolTipText = "Copia el texto seleccionado al portapapeles"
         End If

         If Clipboard.GetFormat(vbCFText) Then
            .Buttons("Paste").Enabled = True
            .Buttons("Paste").ToolTipText = "Pega el texto desde el portapapeles"
         ElseIf Clipboard.GetFormat(cbRTF) Then
            .Buttons("Paste").Enabled = True
            .Buttons("Paste").ToolTipText = "Pega RTF desde portapapeles"
         Else
            .Buttons("Paste").Enabled = False
            .Buttons("Paste").ToolTipText = "Nada que pegar"
         End If

         .Buttons("Undo").Enabled = (SendMessage(RichTextBox.hwnd, EM_CANUNDO, 0, 0&) = 1)
      End With
   End If

   lblSample.ZOrder 0
End Sub
