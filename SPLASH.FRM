VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3120
   ClientLeft      =   2205
   ClientTop       =   2115
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
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3120
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin VB.Image imgSplash 
      Appearance      =   0  'Flat
      Height          =   3120
      Left            =   0
      Picture         =   "Splash.frx":000C
      Top             =   0
      Width           =   5595
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    'Size form to fit bitmap image
    Width = imgSplash.Width
    Height = imgSplash.Height

    'Centre form on screen
    CentreForm Me

    'Make form a top-most window
    FormStayOnTop frmSplash, True

    'Force painting of form
    Visible = True
    Refresh

End Sub

' Manual Override....
Private Sub imgSplash_DblClick()
   Unload Me
End Sub
