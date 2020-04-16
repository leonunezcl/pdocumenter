VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3120
   ClientLeft      =   3435
   ClientTop       =   3750
   ClientWidth     =   5610
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "About.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3120
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "(1.0.0)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2025
      TabIndex        =   1
      Top             =   990
      Width           =   1305
   End
   Begin VB.Label lblRegister 
      Caption         =   "Freeware"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   480
      TabIndex        =   0
      Top             =   1230
      Width           =   2970
   End
   Begin VB.Image imgSplash 
      Appearance      =   0  'Flat
      Height          =   3120
      Left            =   0
      Picture         =   "About.frx":000C
      Top             =   0
      Width           =   5595
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 0 Then
      KeyAscii = 0
      Unload Me
   End If
End Sub

Private Sub Form_Load()

   'Size form to fit bitmap image
   Width = imgSplash.Width
   Height = imgSplash.Height

   lblVersion = Format(App.Major, "###0") & "." & Format(App.Minor, "###0") & "." & Format(App.Revision, "###0")

   'Centre form on screen
   CentreForm Me

   'Make form a top-most window
   FormStayOnTop frmAbout, True

End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormStayOnTop frmAbout, False
End Sub

Private Sub imgSplash_Click()
   Unload Me
End Sub

Private Sub lblRegister_Click()
   Unload Me
End Sub

Private Sub lblVersion_Click()
   Unload Me
End Sub
