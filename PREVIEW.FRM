VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   Caption         =   "Print Preview"
   ClientHeight    =   1995
   ClientLeft      =   1905
   ClientTop       =   2790
   ClientWidth     =   4875
   Icon            =   "Preview.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1995
   ScaleWidth      =   4875
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ilButtons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   16
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Print"
            Description     =   "Print"
            Object.ToolTipText     =   "Print document"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy page image to clipboard"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Pages"
            Object.ToolTipText     =   "View multiple pages"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Zoom"
            Object.ToolTipText     =   "Zoom"
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   720
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "First"
            Object.ToolTipText     =   "Go to first page"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Previous"
            Object.ToolTipText     =   "Go to previous page"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Next"
            Object.ToolTipText     =   "Go to next page"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Last"
            Object.ToolTipText     =   "Go to last page"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Close"
            Object.ToolTipText     =   "Exit preview"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Progress"
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   360
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboZoom 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Preview.frx":030A
      Left            =   1440
      List            =   "Preview.frx":0323
      TabIndex        =   1
      Text            =   "100%"
      Top             =   30
      Width           =   735
   End
   Begin VB.PictureBox picMultiView 
      AutoRedraw      =   -1  'True
      Height          =   1110
      Left            =   2010
      ScaleHeight     =   1050
      ScaleWidth      =   1020
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   435
      Visible         =   0   'False
      Width           =   1080
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   30
         TabIndex        =   5
         Top             =   705
         Width           =   960
      End
      Begin VB.Image imgPage 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   5
         Left            =   690
         Picture         =   "Preview.frx":034D
         Stretch         =   -1  'True
         Top             =   360
         Width           =   300
      End
      Begin VB.Image imgPage 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   4
         Left            =   360
         Picture         =   "Preview.frx":044F
         Stretch         =   -1  'True
         Top             =   360
         Width           =   300
      End
      Begin VB.Image imgPage 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   3
         Left            =   30
         Picture         =   "Preview.frx":0551
         Stretch         =   -1  'True
         Top             =   360
         Width           =   300
      End
      Begin VB.Image imgPage 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   2
         Left            =   690
         Picture         =   "Preview.frx":0653
         Stretch         =   -1  'True
         Top             =   30
         Width           =   300
      End
      Begin VB.Image imgPage 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   360
         Picture         =   "Preview.frx":0755
         Stretch         =   -1  'True
         Top             =   30
         Width           =   300
      End
      Begin VB.Image imgPage 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   30
         Picture         =   "Preview.frx":0857
         Stretch         =   -1  'True
         Top             =   30
         Width           =   300
      End
   End
   Begin VB.ComboBox cboPage 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1695
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox picPaper 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   45
      ScaleHeight     =   570
      ScaleWidth      =   540
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   465
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.VScrollBar VScroll 
      Enabled         =   0   'False
      Height          =   1005
      Left            =   1725
      Min             =   -15
      TabIndex        =   2
      Top             =   420
      Width           =   225
   End
   Begin VB.HScrollBar HScroll 
      Enabled         =   0   'False
      Height          =   225
      Left            =   0
      Min             =   -15
      SmallChange     =   100
      TabIndex        =   3
      Top             =   1425
      Width           =   1710
   End
   Begin VB.PictureBox picParent 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   615
      ScaleHeight     =   705
      ScaleWidth      =   1095
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   450
      Width           =   1095
      Begin VB.Image imgView 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   5
         Left            =   735
         Stretch         =   -1  'True
         Tag             =   "0"
         Top             =   345
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image imgView 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   4
         Left            =   390
         Stretch         =   -1  'True
         Tag             =   "0"
         Top             =   345
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image imgView 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   3
         Left            =   45
         Stretch         =   -1  'True
         Tag             =   "0"
         Top             =   345
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image imgView 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   735
         Stretch         =   -1  'True
         Tag             =   "0"
         Top             =   15
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image imgView 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   390
         Stretch         =   -1  'True
         Tag             =   "0"
         Top             =   15
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image imgView 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   45
         Stretch         =   -1  'True
         Tag             =   "0"
         Top             =   15
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   1665
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   1429
            MinWidth        =   176
            Text            =   "Initialising "
            TextSave        =   "Initialising "
            Key             =   "Page"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6712
            MinWidth        =   1764
            TextSave        =   ""
            Key             =   "Bar"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3135
      Top             =   1065
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ilButtons 
      Left            =   3165
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Preview.frx":0959
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Preview.frx":0C73
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Preview.frx":0F8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Preview.frx":12A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Preview.frx":15C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Preview.frx":18DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Preview.frx":1BF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Preview.frx":1F0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Preview.frx":2229
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' picPaper contains the actual page information (which is hidden)
' imgView(0) shows the page (from picPaper) and is scalable.

Private Const VIEWPORT_OFFSET As Integer = 120
Private Const MULTIPAGE_OFFSET As Integer = 60

Private Type OffsetState
   X As Long
   Y As Long
   V As Integer
   H As Integer
   pV As Integer
   pH As Integer
End Type
Dim Offset As OffsetState
Dim bPanning As Boolean

Private Type PagesState
   Layout() As LayoutState
   LineCount As Integer
End Type
Dim Pages() As PagesState
Dim nPageCount As Integer

Private Type MultiState
   Mode As Integer
   Enabled As Boolean
   Across As Integer
   Down As Integer
   Count As Integer           ' Zero based (0 to 5)
End Type
Dim MultiPage As MultiState

Dim bZoomChanged As Boolean
Dim nCurrentPage As Integer
Dim bResizing As Boolean

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Private Sub Form_Load()
   StatusBar.Panels(1).Text = "Inicializando "
   StatusBar.Panels(2).Text = ""

   bZoomChanged = False
   bPanning = False

   bResizing = True

   If WinPreview.State > -1 Then
      ScaleMode = vbTwips
      Me.Move WinPreview.Left, WinPreview.Top, WinPreview.Width, WinPreview.Height
      Me.WindowState = WinPreview.State
   Else
      Width = Screen.Width * 0.75
      Height = Screen.Height * 0.75
      CentreForm Me
   End If

   bResizing = False

   With cboZoom
      .Text = frmMain.cboZoom.Text
      .Move Toolbar.Buttons("Zoom").Left, Toolbar.Buttons("Zoom").Top, Toolbar.Buttons("Zoom").Width
      .ZOrder 0
   End With

   With ProgressBar
      .Visible = False
      .Move Toolbar.Buttons("Progress").Left, Toolbar.Buttons("Progress").Top, 465
      .ZOrder 0
   End With

   HScroll.Min = -VIEWPORT_OFFSET
   VScroll.Min = -VIEWPORT_OFFSET
   HScroll.Value = HScroll.Min
   VScroll.Value = VScroll.Min

   MultiPage.Mode = 0
   MultiPage.Enabled = False
   MultiPage.Across = 1
   MultiPage.Down = 1
   MultiPage.Count = 0

   Page.Show = True
   Set Page.Form = Me
End Sub

Private Sub Form_Resize()
   If WindowState = vbMinimized Then Exit Sub
   If bResizing Then Exit Sub
   bResizing = True
   On Error Resume Next

   picParent.Move 0, Toolbar.Height, ScaleWidth - VScroll.Width, ScaleHeight - Toolbar.Height - (HScroll.Height + StatusBar.Height)

   picParent.Line (0, 0)-(picParent.Width, 3), , BF

   VScroll.Move ScaleWidth - VScroll.Width, picParent.Top, VScroll.Width, picParent.Height
   HScroll.Move 0, picParent.Top + picParent.Height, ScaleWidth - VScroll.Width

   cboZoom.Move Toolbar.Buttons("Zoom").Left, Toolbar.Buttons("Zoom").Top + ((Toolbar.Buttons("Zoom").Height - cboZoom.Height) / 2), Toolbar.Buttons("Zoom").Width

   If ProgressBar.Visible Then ShowProgressBar True

   cboZoom_Click
   cboZoom.ZOrder 0
   bResizing = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Page.Show = False                ' No more calls to this form
   Set Page.Form = Nothing

   Erase Pages                      ' Erase to gain resources/memory
   nPageCount = 0

   If WindowState <> vbMinimized Then  ' Take a snapshot of the window position
      WinPreview.Left = Me.Left
      WinPreview.Top = Me.Top
      WinPreview.Width = Me.Width
      WinPreview.Height = Me.Height
      WinPreview.State = Me.WindowState
   End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Private Sub Toolbar_ButtonClick(ByVal Button As Button)
   If Button.Key <> "Pages" Then SetVisible picMultiView, False

   Select Case Button.Key
   Case Is = "Print"
      Print2Device
   Case Is = "Copy"
      CopyPage
   Case Is = "Pages"
      MultiPages
   Case Is = "Close"
      Unload Me
   Case Is = "First"
      If nCurrentPage > 1 Then LoadPage 1
   Case Is = "Previous"
      If nCurrentPage > 1 Then LoadPage nCurrentPage - 1
   Case Is = "Next"
      If nCurrentPage < nPageCount Then LoadPage nCurrentPage + 1
   Case Is = "Last"
      If nCurrentPage < nPageCount Then LoadPage nPageCount
   End Select
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Private Sub MultiPages()
   picMultiView.Move Toolbar.Buttons("Pages").Left, Toolbar.Buttons("Pages").Top + Toolbar.Height + 6
   SetVisible picMultiView, True
End Sub

' 0  1  2
' 3  4  5
'
Private Sub imgPage_Click(Index As Integer)
   MultiPage.Enabled = False

   Select Case Index
   Case 0
      If MultiPage.Mode = 0 Then Exit Sub
      MultiPage.Mode = 0
      MultiPage.Enabled = False
      MultiPage.Across = 1
      MultiPage.Down = 1
      MultiPage.Count = 0
   Case 1
      If MultiPage.Mode = 1 Then Exit Sub
      MultiPage.Mode = 1
      MultiPage.Enabled = True
      MultiPage.Across = 2
      MultiPage.Down = 1
      MultiPage.Count = 1
   Case 2
      If MultiPage.Mode = 2 Then Exit Sub
      MultiPage.Mode = 2
      MultiPage.Enabled = True
      MultiPage.Across = 3
      MultiPage.Down = 1
      MultiPage.Count = 2
   Case 3
      If MultiPage.Mode = 3 Then Exit Sub
      MultiPage.Mode = 3
      MultiPage.Enabled = True
      MultiPage.Across = 1
      MultiPage.Down = 2
      MultiPage.Count = 1
   Case 4
      If MultiPage.Mode = 4 Then Exit Sub
      MultiPage.Mode = 4
      MultiPage.Enabled = True
      MultiPage.Across = 2
      MultiPage.Down = 2
      MultiPage.Count = 3
   Case 5
      If MultiPage.Mode = 5 Then Exit Sub
      MultiPage.Mode = 5
      MultiPage.Enabled = True
      MultiPage.Across = 3
      MultiPage.Down = 2
      MultiPage.Count = 5
   End Select

   If MultiPage.Enabled Then
      Toolbar.Buttons("Pages").Value = tbrPressed
      cboZoom.ZOrder 0
      cboZoom.Enabled = False
   Else
      Toolbar.Buttons("Pages").Value = tbrUnpressed
      cboZoom.ZOrder 0
      cboZoom.Enabled = True
   End If

   ZoomImage
   LoadPage nCurrentPage

   SetVisible picMultiView, False
End Sub

Private Sub picMultiView_LostFocus()
   SetVisible picMultiView, False
End Sub

Private Sub cmdCancel_Click()
   SetVisible picMultiView, False
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Private Sub cboZoom_Change()
   bZoomChanged = True
End Sub

Private Sub cboZoom_LostFocus()
   If bZoomChanged Then cboZoom_Click
End Sub

Private Sub cboZoom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If bZoomChanged Then cboZoom_Click
   End If
End Sub

Private Sub cboZoom_Click()
   SetVisible picMultiView, False
   ZoomImage
   bZoomChanged = False
End Sub

Private Sub ZoomImage()
   Dim nLeft As Integer, nTop As Integer, nHeight As Integer, nWidth As Integer, i As Integer, j As Integer, n As Integer
   Dim nRatio As Double
   Dim nFactor As Single
   Dim sText As String

   On Error Resume Next

   MousePointer = vbHourglass

   cboZoom.ZOrder 0

   If MultiPage.Enabled Then
      ' Position the pages - ignore the visible flag

      ' Fit all pages in window...
      nWidth = (picParent.Width - ((VIEWPORT_OFFSET * 2) + (MULTIPAGE_OFFSET * (MultiPage.Across - 1)))) / MultiPage.Across
      nHeight = (picParent.Height - ((VIEWPORT_OFFSET * 2) + (MULTIPAGE_OFFSET * (MultiPage.Down - 1)))) / MultiPage.Down

      If picPaper.Height > picPaper.Width Then
         ' Portrait
         nRatio = nHeight / picPaper.Height
         If (picPaper.Width * nRatio) > nWidth Then
            nRatio = nWidth / picPaper.Width
         End If
      Else
         ' Landscape - Width is factor 1
         nRatio = nWidth / picPaper.Width
         If (picPaper.Height * nRatio) > nHeight Then
            nRatio = nHeight / picPaper.Height
         End If
      End If

      nFactor = nRatio
      sText = "Fit"

      ' Size'm
      nHeight = picPaper.Height * nFactor
      nWidth = picPaper.Width * nFactor
      
      ' Total size
      nLeft = (nWidth * MultiPage.Across) + (MULTIPAGE_OFFSET * (MultiPage.Across - 1))
      nTop = (nHeight * MultiPage.Down) + (MULTIPAGE_OFFSET * (MultiPage.Down - 1))

      ' Calculate left-most and top-most coordinates
      nLeft = ((picParent.Width - nLeft) / 2)       ' - VIEWPORT_OFFSET
      nTop = ((picParent.Height - nTop) / 2)    ' - VIEWPORT_OFFSET
      nFactor = nLeft

      ' Positioning ...
      n = 0
      For i = 1 To MultiPage.Down
         For j = 1 To MultiPage.Across
            imgView(n).Move nLeft, nTop, nWidth, nHeight
            n = n + 1
            nLeft = nLeft + nWidth + MULTIPAGE_OFFSET
         Next
         nLeft = nFactor
         nTop = nTop + nHeight + MULTIPAGE_OFFSET
      Next

      ' No scrolling in multipage mode
      HScroll.Enabled = False
      VScroll.Enabled = False
   Else
      nFactor = Int(Val(cboZoom.Text))
      If nFactor <= 0 Then
         ' Fit
         nWidth = picParent.Width - (VIEWPORT_OFFSET * 2)
         nHeight = picParent.Height - (VIEWPORT_OFFSET * 2)

         If picPaper.Height > picPaper.Width Then
            ' Portrait
            nRatio = nHeight / picPaper.Height
            If (picPaper.Width * nRatio) > nWidth Then
               nRatio = nWidth / picPaper.Width
            End If
         Else
            ' Landscape - Width is factor 1
            nRatio = nWidth / picPaper.Width
            If (picPaper.Height * nRatio) > nHeight Then
               nRatio = nHeight / picPaper.Height
            End If
         End If

         nFactor = nRatio
         sText = "Fit"
      Else
         nFactor = nFactor / 100
         sText = Format(nFactor * 100, "##0") & "%"
      End If

      With imgView(0)
         .Height = picPaper.Height * nFactor
         .Width = picPaper.Width * nFactor
      End With

      SetSliders
   End If

   If cboZoom.Text <> sText Then
      cboZoom.Text = sText
      bZoomChanged = False
   End If

   MousePointer = vbDefault
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Private Sub imgView_DblClick(Index As Integer)
   If imgView(Index).Tag < 1 Then Exit Sub

   If MultiPage.Enabled Then
      nCurrentPage = imgView(Index).Tag
      imgPage_Click 0
   ElseIf Int(Val(cboZoom.Text)) <= 0 Then
      ' From Fit to 100%
      cboZoom.Text = "100%"
      ZoomImage
   Else
      ' From Zoom to Fit
      cboZoom.Text = "Fit"
      ZoomImage
   End If
End Sub

' Take a snapshot of current mouse coordinates prior page panning
Private Sub imgView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetVisible picMultiView, False
   If MultiPage.Enabled Then Exit Sub
   If Button = vbLeftButton And Shift = 0 Then
      Offset.X = X
      Offset.Y = Y
      MousePointer = vbSizePointer     ' Let the user know we are panning...
   End If
End Sub

' Update scrollbars after pan
Private Sub imgView_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If MultiPage.Enabled Then Exit Sub
   On Error Resume Next
   If Button = vbLeftButton And Shift = 0 Then
      bPanning = True
      If VScroll.Enabled Then VScroll.Value = -(imgView(0).Top)
      If HScroll.Enabled Then HScroll.Value = -(imgView(0).Left)
      bPanning = False
   End If
   MousePointer = vbDefault            ' Tell the user it's all over now
End Sub

' Pan page
Private Sub imgView_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim nTop As Integer, nLeft As Integer

   On Error Resume Next

   cboZoom.ZOrder 0

   If MultiPage.Enabled Then Exit Sub

   If Button = vbLeftButton And Shift = 0 Then

      ' What would be the new coordinates...?
      With imgView(0)
         nTop = -(.Top + (Y - Offset.Y))
         nLeft = -(.Left + (X - Offset.X))
      End With

      ' Check limitations...
      With VScroll
         If .Enabled Then
            If nTop < .Min Then
               nTop = .Min
            ElseIf nTop > .Max Then
               nTop = .Max
            End If
         Else
            nTop = -imgView(0).Top
         End If
      End With

      With HScroll
         If .Enabled Then
            If nLeft < .Min Then
               nLeft = .Min
            ElseIf nLeft > .Max Then
               nLeft = .Max
            End If
         Else
            nLeft = -imgView(0).Left
         End If
      End With

      ' Go ahead, make my pan.... (Do you feel lucky, punk...)
      ' Generates MouseMove event... (see VB help on MouseMove)
      imgView(0).Move -nLeft, -nTop

   End If
End Sub

Private Sub VScroll_Change()
   If bPanning Then Exit Sub
   SetVisible picMultiView, False
   imgView(0).Top = -VScroll.Value
End Sub

Private Sub HScroll_Change()
   If bPanning Then Exit Sub
   SetVisible picMultiView, False
   imgView(0).Left = -HScroll.Value
End Sub

Private Sub SetSliders()
   Dim nLeft As Long, nTop As Long
   Dim nHeight As Long, nWidth As Long

   With imgView(0)
      nLeft = .Left
      nTop = .Top
      nHeight = .Height
      nWidth = .Width
   End With

   If nHeight > 32767 Then nHeight = 32767      ' 16-bit limitation (only a integer)
   If nWidth > 32767 Then nWidth = 32767

   If picParent.Width < nWidth Then
      ' Enable horinzontal scroll mode
      With HScroll
         .Max = (nWidth - picParent.Width) + (VIEWPORT_OFFSET * 2)
         .LargeChange = picParent.Width
         .Enabled = True
         If .Value > .Max Then .Value = .Max
      End With
   Else
      With HScroll
         .Max = .Min + 1
         .LargeChange = 1
         .Value = .Min
         .Enabled = False
      End With
      ' Centre "page"
      nLeft = (picParent.Width - nWidth) / 2
   End If

   If picParent.Height < nHeight Then
      ' Enable vertical scroll mode
      With VScroll
         .Max = (nHeight - picParent.Height) + (VIEWPORT_OFFSET * 2)
         .LargeChange = picParent.Height
         .Enabled = True
         If .Value > .Max Then .Value = .Max
      End With
   Else
      With VScroll
         .Max = VScroll.Min + 1
         .LargeChange = 1
         .Value = .Min
         .Enabled = False
      End With
      ' Centre "page"
      nTop = (picParent.Height - nHeight) / 2
   End If

   imgView(0).Move nLeft, nTop

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
' Copy current view display to clipboard.
Private Sub CopyPage()
   MousePointer = vbHourglass

   DoEvents                               ' Just in case the system needs refreshing
   Clipboard.SetData imgView(0).Picture, 2

   MousePointer = vbDefault
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Private Sub ShowProgressBar(bFlag As Boolean)
   If bFlag Then
      If (Width - 4860) < 465 Then
         ProgressBar.Visible = False
      Else
         With ProgressBar
            .Move Toolbar.Buttons("Progress").Left, Toolbar.Buttons("Progress").Top + ((Toolbar.Buttons("Zoom").Height - ProgressBar.Height) / 2), Width - 4860
            .Visible = True
            .ZOrder 0
         End With
         cboZoom.ZOrder 0
         Me.Refresh
      End If
   Else
      ProgressBar.Visible = False
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
   StatusBar.Panels(1).Text = sText & " "
   If ProgressBar.Visible Then ShowProgressBar True
End Sub

Public Sub PrintStart()
   Erase Pages
   nPageCount = 0
   Page.Count = 0

   imgView(0).Visible = False
   imgView(0).Picture = LoadPicture()

   imgView(0).Tag = 0
   imgView(1).Tag = 0
   imgView(2).Tag = 0
   imgView(3).Tag = 0
   imgView(4).Tag = 0
   imgView(5).Tag = 0

   ShowProgress 0
   ShowProgressBar True

   ShowDialog "Formateando"

   nCurrentPage = 1

   HScroll.Enabled = False
   VScroll.Enabled = False

   cboZoom.Enabled = False
   With Toolbar
      .Buttons("Print").Enabled = False
      .Buttons("Copy").Enabled = False
      .Buttons("Pages").Enabled = False
      .Buttons("First").Enabled = False
      .Buttons("Previous").Enabled = False
      .Buttons("Next").Enabled = False
      .Buttons("Last").Enabled = False
   End With

   cboZoom.ZOrder 0
End Sub

Public Sub PrintEnd()
   ShowProgressBar False
   StatusBar.Panels(2).Text = Page.Title

   Page.Count = nPageCount

   If nPageCount = 0 Then
      ShowDialog "No hay paginas disponible"
   
      cboZoom.Enabled = False
      With Toolbar
         .Buttons("Print").Enabled = False
         .Buttons("Copy").Enabled = False
         .Buttons("Pages").Enabled = False
         .Buttons("First").Enabled = False
         .Buttons("Previous").Enabled = False
         .Buttons("Next").Enabled = False
         .Buttons("Last").Enabled = False
      End With
   Else
      ShowDialog "Pagina " & Format$(nCurrentPage) & " de " & Format$(nPageCount)
      StatusBar.Refresh

      Toolbar.Buttons("Print").Enabled = True
   End If

   picPaper.Cls
   cboZoom.ZOrder 0
End Sub

Public Sub PrintNewPage()
   On Error Resume Next
   
   nPageCount = nPageCount + 1
   Page.Count = nPageCount

   ReDim Preserve Pages(1 To nPageCount)
   LayoutToPage nPageCount

   If nPageCount = 1 Then
      cboZoom.Enabled = True
      Toolbar.Buttons("Copy").Enabled = True

      imgView(0).Move -HScroll.Value, -VScroll.Value
      'imgView(0).Left = -HScroll.Value
      'imgView(0).Top = -VScroll.Value

      LoadPage 1
      cboZoom_Click
      DoEvents          ' Give the system some time to refresh itself

   ElseIf nPageCount = 2 Then
      ' Enable page VCR
      With Toolbar
         .Buttons("Pages").Enabled = True
         .Buttons("First").Enabled = True
         .Buttons("Previous").Enabled = True
         .Buttons("Next").Enabled = True
         .Buttons("Last").Enabled = True
      End With
      cboZoom.ZOrder 0
   End If
End Sub

' From "Layout()" array to "Pages()" array
' Must be a faster way that this...
Private Sub LayoutToPage(nPage As Integer)
   On Error Resume Next
   Dim i As Integer, nCount As Integer, j As Integer

   nCount = nLineCount

   If Page.Ruler <> RULER_CHAR Then
      ' Filter out the font commands at the bottom
      For i = nLineCount To 1 Step -1
         If Layout(i).Mode <> LYO_FONT Then
            nCount = i
            Exit For
         End If
      Next
   End If

   Pages(nPage).LineCount = nCount

   ReDim Pages(nPage).Layout(1 To nCount)

   For i = 1 To nCount

      Pages(nPage).Layout(i).Mode = Layout(i).Mode
      Pages(nPage).Layout(i).X = Layout(i).X
      Pages(nPage).Layout(i).Y = Layout(i).Y

      Select Case Layout(i).Mode
      Case LYO_TABS
         If Pages(nPage).Layout(i).X > 0 Then
            For j = 0 To Pages(nPage).Layout(i).Y
               Pages(nPage).Layout(i).Tabs(j) = Layout(i).Tabs(j)
            Next
         End If

      Case LYO_TEXT
         Pages(nPage).Layout(i).Text.Text = Layout(i).Text.Text

      Case LYO_FONT
         Pages(nPage).Layout(i).Font.Name = Layout(i).Font.Name
         Pages(nPage).Layout(i).Font.Size = Layout(i).Font.Size
         Pages(nPage).Layout(i).Font.Color = Layout(i).Font.Color
         Pages(nPage).Layout(i).Font.Bold = Layout(i).Font.Bold
         Pages(nPage).Layout(i).Font.Italic = Layout(i).Font.Italic
         Pages(nPage).Layout(i).Font.Strikethru = Layout(i).Font.Strikethru
         Pages(nPage).Layout(i).Font.Underline = Layout(i).Font.Underline

      Case LYO_LINE, LYO_BOX, LYO_FILLBOX
         Pages(nPage).Layout(i).Line.Width = Layout(i).Line.Width
         Pages(nPage).Layout(i).Line.Height = Layout(i).Line.Height
         Pages(nPage).Layout(i).Line.Color = Layout(i).Line.Color
         Pages(nPage).Layout(i).Line.Style = Layout(i).Line.Style

      Case LYO_IMAGE
         Pages(nPage).Layout(i).Image.Index = Layout(i).Image.Index
         Pages(nPage).Layout(i).Image.Width = Layout(i).Image.Width
         Pages(nPage).Layout(i).Image.Height = Layout(i).Image.Height

      Case LYO_CIRCLE
         Pages(nPage).Layout(i).Circles.Radius = Layout(i).Circles.Radius
         Pages(nPage).Layout(i).Circles.Color = Layout(i).Circles.Color

      End Select

   Next

End Sub

' From "Pages()" array back to "Layout()" array
Private Sub PageToLayout(nPage As Integer)
   On Error Resume Next
   Dim i As Integer, j As Integer

   nLineCount = Pages(nPage).LineCount
   ReDim Layout(1 To nLineCount)

   For i = 1 To nLineCount

      Layout(i).Mode = Pages(nPage).Layout(i).Mode
      Layout(i).X = Pages(nPage).Layout(i).X
      Layout(i).Y = Pages(nPage).Layout(i).Y

      Select Case Layout(i).Mode
      Case LYO_TABS
         If Layout(i).X > 0 Then
            For j = 0 To Layout(i).Y
               Layout(i).Tabs(j) = Pages(nPage).Layout(i).Tabs(j)
            Next
         End If

      Case LYO_TEXT
         Layout(i).Text.Text = Pages(nPage).Layout(i).Text.Text

      Case LYO_FONT
         Layout(i).Font.Name = Pages(nPage).Layout(i).Font.Name
         Layout(i).Font.Size = Pages(nPage).Layout(i).Font.Size
         Layout(i).Font.Color = Pages(nPage).Layout(i).Font.Color
         Layout(i).Font.Bold = Pages(nPage).Layout(i).Font.Bold
         Layout(i).Font.Italic = Pages(nPage).Layout(i).Font.Italic
         Layout(i).Font.Strikethru = Pages(nPage).Layout(i).Font.Strikethru
         Layout(i).Font.Underline = Pages(nPage).Layout(i).Font.Underline

      Case LYO_LINE, LYO_BOX, LYO_FILLBOX
         Layout(i).Line.Width = Pages(nPage).Layout(i).Line.Width
         Layout(i).Line.Height = Pages(nPage).Layout(i).Line.Height
         Layout(i).Line.Color = Pages(nPage).Layout(i).Line.Color
         Layout(i).Line.Style = Pages(nPage).Layout(i).Line.Style

      Case LYO_IMAGE
         Layout(i).Image.Index = Pages(nPage).Layout(i).Image.Index
         Layout(i).Image.Width = Pages(nPage).Layout(i).Image.Width
         Layout(i).Image.Height = Pages(nPage).Layout(i).Image.Height

      Case LYO_CIRCLE
         Layout(i).Circles.Radius = Pages(nPage).Layout(i).Circles.Radius
         Layout(i).Circles.Color = Pages(nPage).Layout(i).Circles.Color

      End Select

   Next

End Sub

Public Sub PrintCancel()
   ShowDialog "Cancelado"
   Unload Me
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Private Sub StatusBar_PanelDblClick(ByVal Panel As Panel)
   If Panel.Index <> 2 Then Exit Sub
   ViewLayout nCurrentPage
End Sub

Private Sub StatusBar_PanelClick(ByVal Panel As Panel)
   If Panel.Index <> 1 Then Exit Sub
   If nPageCount < 2 Then Exit Sub

   Dim i As Integer

   cboPage.Top = Height - 660
   cboPage.Width = StatusBar.Panels(1).Width - 15

   cboPage.Clear
   For i = 1 To nPageCount
      cboPage.AddItem "Page " & i
   Next
   cboPage.ListIndex = nCurrentPage - 1

   SetVisible cboPage, True
   cboPage.SetFocus
End Sub

Private Sub cboPage_Click()
   If Not cboPage.Visible Then Exit Sub

   SetVisible cboPage, False
   If cboPage.ListIndex < 0 Then Exit Sub

   Dim i As Integer
   i = cboPage.ListIndex + 1

   If i = nCurrentPage Then Exit Sub
   If i < 1 Or i > nPageCount Then Exit Sub
   LoadPage i
End Sub

Private Sub cboPage_LostFocus()
   SetVisible cboPage, False
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Private Sub Print2Device()
   Dim i As Integer, j As Integer, nFromPage As Integer, nToPage As Integer

   On Error GoTo PrintDeviceCancelled

   Toolbar.Buttons("Print").Enabled = False
   
   Select Case Page.Output
   Case OUT_DRIVER
      PD.Name = Printer.DeviceName
      PD.Port = Printer.Port
      PD.EnableCopies = True

   Case OUT_RTF
      PD.Name = "Archivo formato enriquecido"
      PD.Port = Page.File
      PD.EnableCopies = False

      If Not FileOverwriteDialog(PD.Port, CommonDialog, "Archivo RTF (*.rtf)|*.rtf|Todos los archivos (*.*)|*.*", ".rtf") Then
         GoTo PrintDeviceCancelled
      End If
      Page.File = PD.Port
      frmMain.SetRTFfile PD.Port

   Case OUT_PORT
      PD.Name = "Directo al puerto (text only)"
      PD.Port = Page.File
      PD.EnableCopies = True
   Case Else
      GoTo PrintDeviceCancelled
   End Select

   PD.Min = 1
   PD.Max = IIf(Page.Count < 1, 1, Page.Count)
   PD.FromPage = PD.Min
   PD.ToPage = PD.Max

   frmSelect.Show vbModal

   If PD.Cancelled Then GoTo PrintDeviceCancelled

   Me.MousePointer = vbHourglass
   DoEvents

   If PD.RangeAll Then
      nFromPage = 1
      nToPage = IIf(Page.Count < 1, 1, Page.Count)
   Else
      nFromPage = PD.FromPage
      nToPage = PD.ToPage
   End If

   On Error GoTo PrintDeviceError

   frmPrint.Show

   frmPrint.PrintStart

   If PD.Collate Then
      Page.Copies = 1
      For j = 1 To PD.Copies
         For i = nFromPage To nToPage
            Page.PageNo = i
            PageToLayout i
            frmPrint.PrintNewPage
            If Page.Cancelled Then Exit For
         Next
         If Page.Cancelled Then Exit For
      Next

   Else
      Page.Copies = PD.Copies

      For i = nFromPage To nToPage
         Page.PageNo = i
         PageToLayout i
         frmPrint.PrintNewPage
         If Page.Cancelled Then Exit For
      Next
   End If

   If Page.Show Then frmPrint.PrintEnd

PrintDeviceError:
   On Error Resume Next
   Unload frmPrint
   Page.Show = True
   Set Page.Form = Me
   Page.Cancelled = False

PrintDeviceCancelled:
   On Error Resume Next
   Toolbar.Buttons("Print").Enabled = True
   Me.MousePointer = vbDefault
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
' Use Layout() array to create picPaper - then copy that into imgView(0)
Private Sub LoadPage(nPageNumber As Integer)
   If nPageNumber < 1 Or nPageNumber > nPageCount Then Exit Sub

   Dim i As Integer, j As Integer, n As Integer
   Dim nPg As Variant

   MousePointer = vbHourglass

   ' The page view order - mostly different from the current page view order
   nPg = Array(nPageNumber, 0, 0, 0, 0, 0)

   If MultiPage.Enabled Then
      ShowDialog "Cargando paginas "

      ' Fill page number array
      j = nPageNumber + MultiPage.Count
      If j > nPageCount Then j = nPageCount
      n = 0
      For i = nPageNumber To j
         nPg(n) = i
         n = n + 1
      Next

      ' Any page images can be copied?
      If nPageNumber < nCurrentPage Then
         For i = 5 To 1 Step -1
            For j = (i - 1) To 0 Step -1
               If Val(imgView(j).Tag) = nPg(i) Then
                  ' Copy image
                  imgView(i).Picture = imgView(j).Picture
                  imgView(i).Tag = imgView(j).Tag
                  Exit For
               End If
            Next
         Next
      ElseIf nPageNumber > nCurrentPage Then
         For i = 0 To 4
            For j = (i + 1) To 5
               If Val(imgView(j).Tag) = nPg(i) Then
                  ' Copy image
                  imgView(i).Picture = imgView(j).Picture
                  imgView(i).Tag = imgView(j).Tag
                  Exit For
               End If
            Next
         Next
      End If
   End If

   For i = 0 To 5

      If nPg(i) = 0 Then
         If imgView(i).Tag <> 0 Then
            If imgView(i).Visible Then imgView(i).Visible = False
            imgView(i).Picture = LoadPicture()
            imgView(i).Tag = 0
         End If

      ElseIf nPg(i) <> imgView(i).Tag Then
         ShowDialog "Cargando pagina " & Format$(nPg(i))
         PrintPage nPg(i)
         imgView(i).Picture = picPaper.Image
         imgView(i).Tag = nPg(i)                ' Pagenumber in the tag property
         If Not imgView(i).Visible Then imgView(i).Visible = True
      End If

      If Page.Cancelled Then Exit For
   Next

   If Page.Cancelled Then
      ShowDialog "Cancelado"
      Exit Sub
   End If

   If MultiPage.Enabled Then
      i = nPageNumber + MultiPage.Count
      If i > nPageCount Then i = nPageCount
      ShowDialog "Paginas " & Format$(nPageNumber) & " a " & Format$(i) & " (de " & Format$(nPageCount) & ")"
   Else
      ShowDialog "Pagina " & Format$(nPageNumber) & " de " & Format$(nPageCount)
   End If

   nCurrentPage = nPageNumber

   MousePointer = vbDefault
End Sub

' Fill the picture box.
Private Sub PrintPage(nPage As Variant)
   Dim i As Integer, nStyle As Integer, nValue As Integer
   Dim bVisible As Boolean

   On Error Resume Next
   
   DoEvents    ' Just in case user clicks on Close button.
   If Page.Cancelled Then Exit Sub

   picPaper.Cls
   
   If Page.Ruler = RULER_CHAR Then
      picPaper.Width = Char2Twips(Page.Width)
      picPaper.Height = Line2Twips(Page.Height)

      picPaper.ScaleMode = vbTwips
      picPaper.FontName = "Courier New"
      picPaper.FONTSIZE = 10
      picPaper.ForeColor = QBColor(0)
      picPaper.FontBold = False
      picPaper.FontItalic = False
      picPaper.FontStrikethru = False
      picPaper.FontUnderline = False

   Else
      picPaper.Height = Page.Height
      picPaper.Width = Page.Width
   End If

   bVisible = ProgressBar.Visible
   nValue = ProgressBar.Value

   ShowProgress 0
   If Not bVisible Then ShowProgressBar True

   For i = 1 To Pages(nPage).LineCount

      ShowProgress ((i * 100) / Pages(nPage).LineCount)

      Select Case Pages(nPage).Layout(i).Mode
      Case LYO_TEXT
         picPaper.CurrentX = Pages(nPage).Layout(i).X
         picPaper.CurrentY = Pages(nPage).Layout(i).Y
         picPaper.Print Pages(nPage).Layout(i).Text.Text;

      Case LYO_FONT
         picPaper.FontName = Pages(nPage).Layout(i).Font.Name
         picPaper.FONTSIZE = Pages(nPage).Layout(i).Font.Size
         picPaper.ForeColor = Pages(nPage).Layout(i).Font.Color
         picPaper.FontBold = Pages(nPage).Layout(i).Font.Bold
         picPaper.FontItalic = Pages(nPage).Layout(i).Font.Italic
         picPaper.FontStrikethru = Pages(nPage).Layout(i).Font.Strikethru
         picPaper.FontUnderline = Pages(nPage).Layout(i).Font.Underline

      Case LYO_LINE
         nStyle = picPaper.DrawStyle
         picPaper.DrawStyle = Pages(nPage).Layout(i).Line.Style
         picPaper.Line (Pages(nPage).Layout(i).X, Pages(nPage).Layout(i).Y)-(Pages(nPage).Layout(i).Line.Width, Pages(nPage).Layout(i).Line.Height), Pages(nPage).Layout(i).Line.Color
         picPaper.DrawStyle = nStyle

      Case LYO_BOX
         nStyle = picPaper.DrawStyle
         picPaper.DrawStyle = Pages(nPage).Layout(i).Line.Style
         picPaper.Line (Pages(nPage).Layout(i).X, Pages(nPage).Layout(i).Y)-(Pages(nPage).Layout(i).Line.Width, Pages(nPage).Layout(i).Line.Height), Pages(nPage).Layout(i).Line.Color, B
         picPaper.DrawStyle = nStyle

      Case LYO_FILLBOX
         nStyle = picPaper.FillStyle
         picPaper.DrawStyle = Pages(nPage).Layout(i).Line.Style
         picPaper.Line (Pages(nPage).Layout(i).X, Pages(nPage).Layout(i).Y)-(Pages(nPage).Layout(i).Line.Width, Pages(nPage).Layout(i).Line.Height), Pages(nPage).Layout(i).Line.Color, BF
         picPaper.FillStyle = nStyle

      Case LYO_IMAGE
         ' Save picture data to disk, then load with LoadPicture into a picturebox and then use it.
         If LoadIcon(Pages(nPage).Layout(i).Image.Index) Then
            picPaper.PaintPicture frmMain.picImage.Picture, Pages(nPage).Layout(i).X, Pages(nPage).Layout(i).Y, Pages(nPage).Layout(i).Image.Width, Pages(nPage).Layout(i).Image.Height
         End If

      Case LYO_CIRCLE
         picPaper.Circle (Pages(nPage).Layout(i).X, Pages(nPage).Layout(i).Y), Pages(nPage).Layout(i).Circles.Radius, Pages(nPage).Layout(i).Circles.Color

      End Select

   Next

   ProgressBar.Value = nValue
   If Not bVisible Then ShowProgressBar False

End Sub

Private Sub ViewLayout(nPage As Integer)
   On Error Resume Next
   Dim i As Integer, nValue As Integer
   Dim sText As String, sAttr As String
   Dim bVisible As Boolean

   Const QUOTE As String = """"

   bVisible = ProgressBar.Visible
   nValue = ProgressBar.Value

   ShowProgress 0
   If Not bVisible Then ShowProgressBar True

   If Page.Ruler = RULER_CHAR Then
       sText = " Page number: " & nPage & vbCrLf & _
      vbCrLf & "  Page width: " & CNum(Page.Width) & "  (All values in characters/lines)" & _
      vbCrLf & "      height: " & CNum(Page.Height) & vbCrLf & _
      vbCrLf & " Margin left: " & CNum(Page.Margin.Left) & _
      vbCrLf & "       right: " & CNum(Page.Margin.Right) & _
      vbCrLf & "         top: " & CNum(Page.Margin.Top) & _
      vbCrLf & "      bottom: " & CNum(Page.Margin.Bottom) & vbCrLf

      For i = 1 To Pages(nPage).LineCount

         ShowProgress ((i * 100) / Pages(nPage).LineCount)
      
         sText = sText & vbCrLf & Pages(nPage).Layout(i).Text.Text
      Next
   Else
   
       sText = " Page number: " & nPage & _
      vbCrLf & "Instructions: " & Pages(nPage).LineCount & vbCrLf & _
      vbCrLf & "  Page width: " & CNum(Page.Width) & "  (All values in twips)" & _
      vbCrLf & "      height: " & CNum(Page.Height) & vbCrLf & _
      vbCrLf & " Margin left: " & CNum(Page.Margin.Left) & _
      vbCrLf & "       right: " & CNum(Page.Margin.Right) & _
      vbCrLf & "         top: " & CNum(Page.Margin.Top) & _
      vbCrLf & "      bottom: " & CNum(Page.Margin.Bottom) & vbCrLf

      For i = 1 To Pages(nPage).LineCount

         ShowProgress ((i * 100) / Pages(nPage).LineCount)

         Select Case Pages(nPage).Layout(i).Mode
         Case LYO_TEXT
            sText = sText & vbCrLf & _
                    "   Text: " & _
                    CNum(Pages(nPage).Layout(i).X) & ", " & _
                    CNum(Pages(nPage).Layout(i).Y) & ", " & _
                    QUOTE & Pages(nPage).Layout(i).Text.Text & QUOTE

         Case LYO_FONT
            sText = sText & vbCrLf & _
                    "   Font: " & _
                    Pages(nPage).Layout(i).Font.Name & ", " & _
                    Pages(nPage).Layout(i).Font.Size & ", " & _
                    "Color &H" & Hex(Pages(nPage).Layout(i).Font.Color)

                    sAttr = ""
                    If Pages(nPage).Layout(i).Font.Bold Then sAttr = sAttr & " Bold"
                    If Pages(nPage).Layout(i).Font.Italic Then sAttr = sAttr & " Italic"
                    If Pages(nPage).Layout(i).Font.Strikethru Then sAttr = sAttr & " Strikethru"
                    If Pages(nPage).Layout(i).Font.Underline Then sAttr = sAttr & " Underline"
                    If Not EmptyString(sAttr) Then sText = sText & "," & sAttr

         Case LYO_LINE, LYO_BOX, LYO_FILLBOX
            Select Case Pages(nPage).Layout(i).Mode
            Case LYO_LINE
               sText = sText & vbCrLf & "   Line: "
            Case LYO_BOX
               sText = sText & vbCrLf & "    Box: "
            Case LYO_FILLBOX
               sText = sText & vbCrLf & "FillBox: "
            End Select
            sText = sText & _
                    "(" & Pages(nPage).Layout(i).X & ", " & Pages(nPage).Layout(i).Y & _
                    ")-(" & Pages(nPage).Layout(i).Line.Width & ", " & Pages(nPage).Layout(i).Line.Height & _
                    "), Colour &H" & Hex(Pages(nPage).Layout(i).Line.Color) & _
                    ", Style " & Pages(nPage).Layout(i).Line.Style

         Case LYO_IMAGE
            sText = sText & vbCrLf & _
                    "  Image: " & _
                    CNum(Pages(nPage).Layout(i).X) & ", " & _
                    CNum(Pages(nPage).Layout(i).Y) & ", " & _
                    "[Index " & Pages(nPage).Layout(i).Image.Index & "], " & _
                    "Width " & Pages(nPage).Layout(i).Image.Width & ", " & _
                    "Height " & Pages(nPage).Layout(i).Image.Height

         Case LYO_CIRCLE
            sText = sText & vbCrLf & _
                    " Circle: " & _
                    CNum(Pages(nPage).Layout(i).X) & ", " & _
                    CNum(Pages(nPage).Layout(i).Y) & ", " & _
                    "Radius " & Pages(nPage).Layout(i).Circles.Radius & ", " & _
                    "Color &H" & Hex(Pages(nPage).Layout(i).Circles.Color)

         End Select
      Next
   End If

   ProgressBar.Value = nValue
   If Not bVisible Then ShowProgressBar False

   Load frmViewFile
   frmViewFile.SetText "Page layout" & IIf(Page.Ruler = RULER_CHAR, "", " commands"), sText
   frmViewFile.Show

End Sub

Function CNum(nValue As Long) As String
   CNum = Right$("     " & nValue, 5)
End Function
