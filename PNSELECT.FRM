VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir"
   ClientHeight    =   2655
   ClientLeft      =   885
   ClientTop       =   3150
   ClientWidth     =   6135
   Icon            =   "PnSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2655
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Copias"
      Height          =   1320
      Left            =   3375
      TabIndex        =   10
      Top             =   855
      Width           =   2700
      Begin VB.VScrollBar VScroll 
         Height          =   240
         Left            =   2280
         Max             =   9999
         Min             =   1
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   270
         Value           =   1
         Width           =   180
      End
      Begin VB.TextBox txtCopies 
         Height          =   285
         Left            =   1770
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "1"
         Top             =   240
         Width           =   720
      End
      Begin VB.CheckBox chkCollate 
         Caption         =   "C&ollate"
         Height          =   270
         Left            =   1770
         TabIndex        =   5
         Top             =   735
         Width           =   810
      End
      Begin VB.Image imgCollate 
         Appearance      =   0  'Flat
         Height          =   450
         Index           =   0
         Left            =   135
         Picture         =   "PnSelect.frx":000C
         Top             =   645
         Width           =   1470
      End
      Begin VB.Image imgCollate 
         Appearance      =   0  'Flat
         Height          =   540
         Index           =   1
         Left            =   180
         Picture         =   "PnSelect.frx":06A6
         Top             =   600
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblCopies 
         Caption         =   "Numero de &copias:"
         Height          =   210
         Left            =   165
         TabIndex        =   13
         Top             =   285
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rango Impresion"
      Height          =   1320
      Left            =   45
      TabIndex        =   9
      Top             =   855
      Width           =   3240
      Begin VB.TextBox txtToPage 
         Height          =   285
         Left            =   2475
         MaxLength       =   5
         TabIndex        =   3
         Top             =   750
         Width           =   585
      End
      Begin VB.TextBox txtFromPage 
         Height          =   285
         Left            =   1485
         MaxLength       =   5
         TabIndex        =   2
         Top             =   750
         Width           =   570
      End
      Begin VB.OptionButton optRange 
         Caption         =   "Paginas"
         Height          =   225
         Index           =   1
         Left            =   60
         TabIndex        =   1
         Top             =   780
         Width           =   885
      End
      Begin VB.OptionButton optRange 
         Caption         =   "Todo"
         Height          =   225
         Index           =   0
         Left            =   75
         TabIndex        =   0
         Top             =   375
         Value           =   -1  'True
         Width           =   2835
      End
      Begin VB.Label lblTo 
         Caption         =   "a:"
         Height          =   210
         Left            =   2250
         TabIndex        =   16
         Top             =   780
         Width           =   225
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         Caption         =   "desde:"
         Height          =   195
         Left            =   990
         TabIndex        =   15
         Top             =   795
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impresora"
      Height          =   840
      Left            =   30
      TabIndex        =   8
      Top             =   0
      Width           =   6045
      Begin VB.Label lblPort 
         Caption         =   "LPT1:"
         Height          =   210
         Left            =   915
         TabIndex        =   18
         Top             =   510
         Width           =   4950
      End
      Begin VB.Label lblName 
         Caption         =   "(Sin Titulo)"
         Height          =   210
         Left            =   930
         TabIndex        =   17
         Top             =   255
         Width           =   4935
      End
      Begin VB.Label Label 
         Caption         =   "Donde:"
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   510
         Width           =   600
      End
      Begin VB.Label Label 
         Caption         =   "Nombre:"
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   255
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   4995
      TabIndex        =   7
      Top             =   2265
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3825
      TabIndex        =   6
      Top             =   2265
      Width           =   1080
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   On Error Resume Next
   SetCaption lblName, PD.Name
   SetCaption lblPort, PD.Port

   If PD.Min > PD.Max Then
      Dim i As Integer
      i = PD.Min
      PD.Min = PD.Max
      PD.Max = i
   End If

   If PD.FromPage < PD.Min Then PD.FromPage = PD.Min
   If PD.ToPage > PD.Max Then PD.ToPage = PD.Max

   If PD.Min < 1 Or (PD.Max - PD.Min) < 2 Then
      PD.RangeAll = True
      optRange(0) = True
      SetEnabled optRange(1), False
      SetEnabled lblFrom, False
      SetEnabled txtFromPage, False
      SetEnabled lblTo, False
      SetEnabled txtToPage, False
   Else
      PD.RangeAll = False
      optRange(1) = True
      txtFromPage = PD.FromPage
      txtToPage = PD.ToPage
   End If

   If Not PD.EnableCopies Then
      SetEnabled lblCopies, False
      SetEnabled txtCopies, False
      SetEnabled VScroll, False
      SetEnabled chkCollate, False
   End If

   PD.Copies = 1
   PD.Collate = False
   PD.Cancelled = False

   CentreForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PD.RangeAll = optRange(0)
   PD.FromPage = Val(txtFromPage)
   PD.ToPage = Val(txtToPage)
   PD.Copies = Val(txtCopies)
   PD.Collate = (chkCollate = vbChecked)
End Sub

Private Sub cmdCancel_Click()
   PD.Cancelled = True
   Unload Me
End Sub

Private Sub cmdOK_Click()
   PD.Cancelled = False
   Unload Me
End Sub

Private Sub optRange_Click(Index As Integer)
   On Error Resume Next
   If Index = 1 Then
      txtFromPage.SetFocus
   End If
End Sub

Private Sub txtFromPage_Change()
   On Error Resume Next
   If Val(txtFromPage) > Val(txtToPage) Then
      txtFromPage = txtToPage
   ElseIf Val(txtFromPage) < PD.Min Then
      txtFromPage = PD.Min
   End If
End Sub

Private Sub txtFromPage_KeyPress(KeyAscii As Integer)
   If KeyAscii > 57 Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtToPage_Change()
   On Error Resume Next
   If Val(txtToPage) < Val(txtFromPage) Then
      txtToPage = txtFromPage
   ElseIf Val(txtToPage) > PD.Max Then
      txtToPage = PD.Max
   End If
End Sub

Private Sub txtToPage_KeyPress(KeyAscii As Integer)
   If KeyAscii > 57 Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCopies_Change()
   On Error Resume Next
   If Val(txtCopies) < 1 Then
      txtCopies = 1
   ElseIf Val(txtCopies) > 9999 Then
      txtCopies = 9999
   End If
   VScroll.Value = Val(txtCopies)
End Sub

Private Sub VScroll_Change()
   txtCopies = VScroll.Value
End Sub

Private Sub chkCollate_Click()
   If chkCollate = vbChecked Then
      SetVisible imgCollate(1), True
      SetVisible imgCollate(0), False
   Else
      SetVisible imgCollate(0), True
      SetVisible imgCollate(1), False
   End If
End Sub
