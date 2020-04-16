VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find..."
   ClientHeight    =   1665
   ClientLeft      =   2265
   ClientTop       =   3060
   ClientWidth     =   5640
   Icon            =   "Find.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1665
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   345
      Left            =   4185
      TabIndex        =   4
      Top             =   1230
      Width           =   1380
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Buscar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   4185
      TabIndex        =   3
      Top             =   825
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   945
      Left            =   75
      TabIndex        =   6
      Top             =   630
      Width           =   4005
      Begin VB.CheckBox chkFind 
         Caption         =   "Busqueda exacta"
         Height          =   225
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   585
         Width           =   3660
      End
      Begin VB.CheckBox chkFind 
         Caption         =   "Solo palabras completas"
         Height          =   225
         Index           =   0
         Left            =   165
         TabIndex        =   1
         Top             =   285
         Width           =   3660
      End
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   75
      TabIndex        =   0
      Top             =   285
      Width           =   4000
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar:"
      Height          =   195
      Left            =   75
      TabIndex        =   5
      Top             =   60
      Width           =   1125
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bFound As Boolean
Dim nLastPos As Long

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()
   On Error GoTo FindErr

   Dim nOptions As Integer, nFoundPos As Long, nFoundLine As Long

   If EmptyString(txtFind) Then Exit Sub

   nOptions = 0

   If chkFind(0) = vbChecked Then nOptions = nOptions + rtfWholeWord
   If chkFind(1) = vbChecked Then nOptions = nOptions + rtfMatchCase

   If bFound Then
      nFoundPos = frmViewFile.RichTextBox.Find(txtFind, nLastPos, , nOptions)
      If nFoundPos = -1 Then
         ' No more text - start from top
         nFoundPos = frmViewFile.RichTextBox.Find(txtFind, 0, , nOptions)
      End If
   Else
      ' Find the text specified in the TextBox control.
      nFoundPos = frmViewFile.RichTextBox.Find(txtFind, 0, , nOptions)
   End If

   ' Show message based on whether the text was found or not.
   If nFoundPos = -1 Then
      Beep

      nLastPos = 0
      bFound = False
      SetCaption cmdFind, "&Buscar"
   Else
      ' Returns number of line containing found text.
      'nFoundLine = frmViewFile.RichTextBox.GetLineFromChar(nFoundPos)

      nLastPos = nFoundPos + Len(txtFind)
      bFound = True
      SetCaption cmdFind, "&Buscar siguiente"
   End If

   Exit Sub

FindErr:
   MsgBox "Error ocurrio en el sistema" & vbCrLf & _
          "Error #" & Err.Number & ": " & Err.Description, _
          vbCritical

End Sub

Private Sub Form_Load()
   bFound = False
   nLastPos = 0

   CentreForm Me
End Sub

Private Sub Form_Resize()

   If WindowState = vbMinimized Then Exit Sub

   FormStayOnTop Me, True

End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormStayOnTop Me, False
End Sub

Private Sub chkFind_Click(Index As Integer)
   If bFound Then
      nLastPos = 0
      bFound = False
      SetCaption cmdFind, "&Find"
   End If
End Sub

Private Sub txtFind_Change()
   If bFound Then
      nLastPos = 0
      bFound = False
      SetCaption cmdFind, "&Find"
   End If
   
   SetEnabled cmdFind, (Not EmptyString(txtFind.Text))
End Sub

Private Sub txtFind_GotFocus()
   SelectText txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not EmptyString(txtFind.Text) Then
         cmdFind_Click
         KeyAscii = 0
      End If
   End If
End Sub
