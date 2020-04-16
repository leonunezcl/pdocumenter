Attribute VB_Name = "Support"
Option Explicit

Const INI_FILE As String = "VPrint32.ini"          ' Try using only 8 letters for filename
Public sIniFile As String
Public sHelpFile As String
Public glbCarga As Boolean
Const INI_RECENT_KEY As String = "Recent Files"    ' INI Key constant for recent files
Const RECENT_COUNT As Integer = 9                  ' Maximum number of recent files (one based)

' Font setting areas
Public Const FONT_PROCS As Integer = 0
Public Const FONT_CODE As Integer = 1
Public Const FONT_COMMENTS As Integer = 2
Public Const FONT_HEADER As Integer = 3
Public Const FONT_FOOTER As Integer = 4
Public Const FONT_DIRECTIVE As Integer = 5
Public Const FONT_TITLES As Integer = 6

' Additional file information for Outline - uses Outline.ItemData() as link.
Public Type RefState
   FilePoint As Integer       ' 0-... = File
   ProcPoint As Integer       ' -1 = File, 0 = Declaration, 1-... = Procedure
End Type
Public ItemRef() As RefState  ' Item reference array

' To remember frmPreview's last window position
Public Type WinPosState
   Left As Integer
   Top As Integer
   Width As Integer
   Height As Integer
   State As Integer
End Type
Public WinPreview As WinPosState
Public WinView As WinPosState

' Used for file base ini - instead of using windows registry database.
Private Declare Function GetPrivateProfileStringByKeyName Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$) As Long
Private Declare Function WritePrivateProfileStringByKeyName Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String) As Long
Private Declare Function WritePrivateProfileStringToDeleteKey Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String) As Long

' Use in MakeTempFile() - to create tempory filenames
Private Const MAX_PATH As Long = 260
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' Used to force window on top.
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long

' Used to find out if running in development mode or standalone executable.
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

' Used in word-wrap module and undo/redo richtextbox (frmViewFile)
Public Const EM_GETLINECOUNT As Long = &HBA
Public Const EM_GETLINE As Long = &HC4
Public Const EM_LINEINDEX As Long = &HBB
Public Const EM_LINELENGTH As Long = &HC1
Public Const EM_CANUNDO As Long = &HC6
Public Const EM_UNDO As Long = &HC7
' Together with these API's
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageBynum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

' Play a wave file - keep this? (yes, for the time being)
Public Const WAVE_ACCESSED As Integer = 0
Public Const WAVE_ANALYSE As Integer = 1
Public Const WAVE_ERROR As Integer = 2
Public Const WAVE_EXIT As Integer = 3
Public Const WAVE_OK As Integer = 4
Public Const WAVE_READY As Integer = 5
Public Const WAVE_SORRY As Integer = 6
Public Const WAVE_STANDBY As Integer = 7
Public Const WAVE_STARTUP As Integer = 8
Public Const WAVE_THANKYOU As Integer = 9

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFaceName As String * 33
End Type

Private Const SND_SYNC As Long = &H0
Private Const SND_ASYNC As Long = &H1
Private Const SND_NODEFAULT As Long = &H2
Private Const SND_LOOP As Long = &H8
Private Const SND_NOSTOP As Long = &H10
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Sub FontStuff(ByVal Titulo As String, picDraw As PictureBox)
    
    On Error GoTo GetOut
    Dim f As LOGFONT, hPrevFont As Long, hFont As Long, FontName As String
    Dim FONTSIZE As Integer
    FONTSIZE = 10 'Val(txtSize.Text)
    
    f.lfEscapement = 10 * 90 'Val(txtDegree.Text) 'rotation angle, in tenths
    FontName = "Tahoma" + Chr$(0) 'null terminated
    f.lfFaceName = FontName
    f.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
    hFont = CreateFontIndirect(f)
    hPrevFont = SelectObject(picDraw.hdc, hFont)
    
    picDraw.CurrentX = 3
    'picDraw.CurrentY = 310
    
    picDraw.CurrentY = picDraw.Height - 10
    picDraw.Print Titulo
    
    '  Clean up, restore original font
    hFont = SelectObject(picDraw.hdc, hPrevFont)
    DeleteObject hFont
    
    Exit Sub
GetOut:
    Exit Sub

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Sub Main()
   ' Load splash screen (will unloaded later)
   glbCarga = True
   frmAcerca.Show vbModal

   sIniFile = AppPathFile(INI_FILE)
   sHelpFile = AppPathFile("VBPrint.rtf")

   MakeSound WAVE_STARTUP

   WinPreview.State = -1
   WinView.State = -1

   MdCount = 0
   MdSelected = 0
   PrCount = 0
   PrSelected = 0

   Page.Show = False       ' Set by frmPreview or frmPrint
   Set Page.Form = Nothing

   Load frmMain

   ' Do a manual unload
   Unload frmAcerca
    glbCarga = False
   MakeSound WAVE_READY
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Sub MakeSound(nEvent As Integer, Optional bWait, Optional bForce)
   If IsMissing(bForce) Then bForce = (GetIniString(sIniFile, "Options", "WaveSounds", "1") = "1")
   If bForce Then

      If IsMissing(bWait) Then bWait = False

      Select Case nEvent
      Case WAVE_ACCESSED
         PlayWave AppPathFile("Accessed.wav"), bWait
      Case WAVE_ANALYSE
         PlayWave AppPathFile("Analyse.wav"), bWait
      Case WAVE_ERROR
         PlayWave AppPathFile("Error.wav"), bWait
      Case WAVE_EXIT
         PlayWave AppPathFile("Exit.wav"), bWait
      Case WAVE_OK
         PlayWave AppPathFile("Ok.wav"), bWait
      Case WAVE_READY
         PlayWave AppPathFile("Ready.wav"), bWait
      Case WAVE_SORRY
         PlayWave AppPathFile("Sorry.wav"), bWait
      Case WAVE_STANDBY
         PlayWave AppPathFile("StandBy.wav"), bWait
      Case WAVE_STARTUP
         PlayWave AppPathFile("Startup.wav"), bWait
      Case WAVE_THANKYOU
         PlayWave AppPathFile("ThankYou.wav"), bWait
      End Select
   End If
End Sub

Function GetSoundFileName(nEvent As Integer)
   Select Case nEvent
   Case WAVE_ACCESSED
      GetSoundFileName = "Accessed.wav"
   Case WAVE_ANALYSE
      GetSoundFileName = "Analyse.wav"
   Case WAVE_ERROR
      GetSoundFileName = "Error.wav"
   Case WAVE_EXIT
      GetSoundFileName = "Exit.wav"
   Case WAVE_OK
      GetSoundFileName = "Ok.wav"
   Case WAVE_READY
      GetSoundFileName = "Ready.wav"
   Case WAVE_SORRY
      GetSoundFileName = "Sorry.wav"
   Case WAVE_STANDBY
      GetSoundFileName = "StandBy.wav"
   Case WAVE_STARTUP
      GetSoundFileName = "Startup.wav"
   Case WAVE_THANKYOU
      GetSoundFileName = "ThankYou.wav"
   Case Else
      GetSoundFileName = "n/a"
   End Select
End Function

Sub PlayWave(sSoundFile As String, Optional bWait)
   If FileExist(sSoundFile) Then
      Dim dl As Long, wFlags As Long
      If IsMissing(bWait) Then bWait = False
      If bWait Then
         wFlags = SND_SYNC Or SND_NODEFAULT
      Else
         wFlags = SND_ASYNC Or SND_NODEFAULT
      End If
      dl = sndPlaySound(sSoundFile, wFlags)
   End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' This function will return true if we are running in the IDE (development) mode else it returns false.
'
' Great for enableling error interception code, eg:
'   If Not InDevelopmentMode Then On Error GoTo ErrorHandler
'
Function InDevelopmentMode() As Boolean
   InDevelopmentMode = Not CBool(GetModuleHandle(App.EXEName))
End Function

Function MatchString(sExpression As String, sContaining As String) As Boolean
   MatchString = (Left(sExpression, Len(sContaining)) = sContaining)
End Function

' To force a windows to stay on top. (usefull for splash screens)
Sub FormStayOnTop(frmHandle As Form, bOnTop As Boolean)
   Dim nFlags As Integer
   nFlags = 2 Or 1    ' &H2 Or &H1 Or &H40 Or &H10

   On Error Resume Next

   Select Case bOnTop
   Case True
      SetWindowPos frmHandle.hWnd, -1, 0, 0, 0, 0, nFlags
   Case False
      SetWindowPos frmHandle.hWnd, -2, 0, 0, 0, 0, nFlags
   End Select

End Sub

'Centre form on screen
Sub CentreForm(frmHandle As Form)
    frmHandle.Move (Screen.Width - frmHandle.Width) / 2, (Screen.Height - frmHandle.Height) / 2
End Sub

Function EmptyString(ByRef sText As String) As Boolean
   If IsNull(sText) Then
      EmptyString = True
   Else
      EmptyString = (Len(Trim(sText)) = 0)
   End If
End Function

' Pads a string with spaces
Function Pad(sString As String, nSize As Integer) As String
   Pad = Left$(sString & Space$(nSize), nSize)
End Function

' Adds the application path to a filename
'
Function AppPathFile(sFileName As String) As String
   Dim sFullName As String
   sFullName = App.Path
   If Right$(sFullName, 1) <> "\" Then sFullName = sFullName & "\"
   AppPathFile = sFullName & sFileName
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' An example of modular programming - This provides a safer interface to GetPrivateProfileString
'
' Obtain string data for specified key
'
Function GetIniString(sFile As String, sSection As String, sKey As String, Optional vDefault) As String
   Dim sKeyValue As String
   Dim nCharacters As Long

   If IsMissing(vDefault) Then
      vDefault = ""
   End If

   sKeyValue = String$(250, 0)

   nCharacters = GetPrivateProfileStringByKeyName(sSection, sKey, vDefault, sKeyValue, 250, sFile)

   If nCharacters > 0 Then sKeyValue = Left$(sKeyValue, nCharacters)

   ' Remove some null characters (just in case)
   nCharacters = InStr(sKeyValue, Chr$(0))
   If nCharacters > 0 Then sKeyValue = Left$(sKeyValue, nCharacters - 1)

   GetIniString = sKeyValue
End Function

' Add/Edit new value and key
'
Function AddIniString(sFile As String, sSection As String, sKey As String, sValue As String) As Boolean
   Dim nSuccess As Long

   ' Write the new key
   nSuccess = WritePrivateProfileStringByKeyName(sSection, sKey, sValue, sFile)

   If nSuccess = 0 Then
      AddIniString = False
      Exit Function
   End If

   AddIniString = True
End Function

' Delete the specified key
'
Function DeleteIniKey(sFile As String, sSection As String, sKey As String) As Boolean
   Dim nSuccess As Long

   nSuccess = WritePrivateProfileStringToDeleteKey(sSection, sKey, 0, sFile)

   If nSuccess = 0 Then
      DeleteIniKey = False
      Exit Function
   End If

   DeleteIniKey = True
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Function MakeTempFile() As String
   Dim sBuffer As String, sPath As String
   Dim nCut As Integer
   Dim dl As Long

   sBuffer = Space$(MAX_PATH)
   dl = GetTempPath(MAX_PATH, sBuffer)
   If dl Then
      sPath = Trim$(Mid$(sBuffer, 1, dl))
   Else
      sPath = App.Path
   End If

   sBuffer = Space$(MAX_PATH)
   dl = GetTempFileName(App.Path, "pn_", 0, sBuffer)
   nCut = InStr(1, sBuffer, Chr(0))
   If nCut Then sBuffer = Trim$(Mid$(sBuffer, 1, nCut - 1))

   If FileExist(sBuffer) Then Kill sBuffer
   MakeTempFile = sBuffer
End Function

' Checks wether file exist (handles wildcards too)
Function FileExist(ByVal sFile As String) As Boolean

   If Len(Trim(sFile)) = 0 Then
      ' Nothing given
      FileExist = False
      Exit Function
   ElseIf Right(sFile, 1) = "\" Or Right(sFile, 1) = ":" Then
      ' Just a part of a path or drive... (not complete)
      FileExist = False
      Exit Function
   ElseIf Dir(sFile) = "" Then
      ' Not there...
      FileExist = False
      Exit Function
   End If

   ' After all that torture, it must exist...
   FileExist = True
   Exit Function
ExistErrorHandler:
   FileExist = False
End Function

Function FileOverwriteDialog(ByRef sFile As String, CDialog As Object, Optional sFilter, Optional sDefaultExt) As Boolean
   If Not FileExist(sFile) Then
      FileOverwriteDialog = True
      Exit Function
   End If

   Select Case MsgBox(sFile + vbCrLf + "El archivo ya existe." + vbCrLf + vbCrLf + "Reemplazarlo?", vbYesNoCancel + vbExclamation + vbDefaultButton3)
   Case vbCancel
      FileOverwriteDialog = False
      Exit Function

   Case vbNo
      ' Pick new file
      On Error GoTo FileOverwriteCancelled

      With CDialog
         .DialogTitle = "Guardar archivo como ..."

         If IsMissing(sFilter) Then
            .Filter = "Todos los archivos (*.*)|*.*"
         Else
            .Filter = sFilter
         End If
         .FilterIndex = 1

         If IsMissing(sDefaultExt) Then
            .DefaultExt = ""
         Else
            .DefaultExt = sDefaultExt
         End If

         .CancelError = True
         .Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist

         .FileName = sFile
      End With

      CDialog.ShowSave
      sFile = CDialog.FileName
   End Select

   FileOverwriteDialog = True
   Exit Function

FileOverwriteCancelled:
   CDialog.CancelError = False
   FileOverwriteDialog = False
End Function
Function ExtractFileExt(sFileName As String) As String
   Dim i As Integer
   For i = Len(sFileName) To 1 Step -1
      If InStr(".", Mid$(sFileName, i, 1)) Then Exit For
   Next
   ExtractFileExt = Right$(sFileName, Len(sFileName) - i)
End Function

Function ExtractFileName(sFileIn As String) As String
   Dim i As Integer
   For i = Len(sFileIn) To 1 Step -1
      If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
   Next
   ExtractFileName = Mid$(sFileIn, i + 1, Len(sFileIn) - i)
End Function

Function ExtractPath(sPathIn As String) As String
   Dim i As Integer
   For i = Len(sPathIn) To 1 Step -1
      If InStr(":\", Mid$(sPathIn, i, 1)) Then Exit For
   Next
   ExtractPath = Left$(sPathIn, i)
End Function

Function FixPath(ByVal sPath As String) As String
   If Len(Trim(sPath)) = 0 Then
      FixPath = ""
   ElseIf Right$(sPath, 1) <> "\" Then
      FixPath = sPath & "\"
   Else
      FixPath = sPath
   End If
End Function

Function AttachPath(sFileName As String, sPath As String) As String
   If Len(Trim(ExtractPath(sFileName))) = 0 Then
      AttachPath = FixPath(sPath) & sFileName
   Else
      AttachPath = sFileName
   End If
End Function

Function LongDirFix(Incomming As String, Max As Integer) As String
   Dim i As Integer, LblLen As Integer, StringLen As Integer
   Dim TempString As String

   TempString = Incomming
   LblLen = Max

   If Len(TempString) <= LblLen Then
      LongDirFix = TempString
      Exit Function
   End If

   LblLen = LblLen - 6

   For i = Len(TempString) - LblLen To Len(TempString)
      If Mid$(TempString, i, 1) = "\" Then Exit For
   Next

   LongDirFix = Left$(TempString, 3) + "..." + Right$(TempString, Len(TempString) - (i - 1))
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' This procedure will set the visible prop of a control
' Use this procedure to reduce control "flicker".
Sub SetVisible(ctrlIn As Control, iTrueFalse As Integer, Optional sCaption)
   Dim iCompare As Integer
   iCompare = Not iTrueFalse

   If ctrlIn.Visible = iCompare Then
      ctrlIn.Visible = iTrueFalse
   End If

   If Not IsMissing(sCaption) Then
      If ctrlIn.Caption <> sCaption Then
         ctrlIn.Caption = sCaption
      End If
   End If

End Sub

' This procedure will set the enabled prop of a control
' Use this procedure to reduce control "flicker".
Sub SetEnabled(ctrlIn As Control, ByVal iTrueFalse As Integer, Optional sCaption)
   Dim iCompare As Integer
   iCompare = Not iTrueFalse

   If ctrlIn.Enabled = iCompare Then
      ctrlIn.Enabled = iTrueFalse
   End If

   If Not IsMissing(sCaption) Then
      If ctrlIn.Caption <> sCaption Then
         ctrlIn.Caption = sCaption
      End If
   End If

End Sub

' Use this procedure to reduce control "flicker".
Sub SetLock(ctrlIn As Control, ByVal iTrueFalse As Integer)
   Dim iCompare As Integer
   iCompare = Not iTrueFalse

   If ctrlIn.Locked = iCompare Then
      ctrlIn.Locked = iTrueFalse
   End If
End Sub

' Use this procedure to reduce control "flicker".
Sub SetCaption(ctrlIn As Control, sCaption As String)
   If ctrlIn.Caption <> sCaption Then
      ctrlIn.Caption = sCaption
   End If
End Sub

' Use this procedure to reduce control "flicker".
Sub SetText(ctrlIn As Control, sText As String)
   If ctrlIn.Text <> sText Then
      ctrlIn.Text = sText
   End If
End Sub

' Use this in method txtctrl_GotFocus() and will select text with the text box
Sub SelectText(ctrlIn As Control)
   ctrlIn.SelStart = 0
   ctrlIn.SelLength = 65000
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Function NumericOnly(KeyAscii As Integer) As Integer
   If KeyAscii > 31 Then
      If KeyAscii < 48 Or KeyAscii > 57 Then
         ' Only numbers allowed
         KeyAscii = 0
      End If
   End If
   NumericOnly = KeyAscii
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Like String(), but can handle sString of longer size than one.
'
Function Replicate(nSize As Integer, sString As String) As String
   Dim nLen As Integer
   nLen = Len(sString)

   If nLen = 0 Or nSize = 0 Then
      Replicate = ""
      Exit Function
   ElseIf nLen = 1 Then
      Replicate = String$(nSize, sString)
      Exit Function
   ElseIf nSize <= nLen Then
      Replicate = Left$(sString, nSize)
      Exit Function
   End If

   Dim sText As String
   Dim i As Integer

   sText = ""
   nLen = Int(nSize / nLen) + 1

   For i = 1 To nLen
      sText = sText & sString
   Next

   Replicate = Left$(sText, nSize)

End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Public Sub ReportError(Optional SMessage, Optional ErrorNumber)
   Dim sText As String, sSource As String, sHelpFile As String
   Dim vContext As Variant
   Static sLastMessage As String

   sSource = Err.Source
   sHelpFile = Err.HelpFile
   vContext = Err.HelpContext

   If IsMissing(ErrorNumber) Then ErrorNumber = Err.Number

   If ErrorNumber <> 0 Then

      If IsMissing(SMessage) Then
         sText = Error(ErrorNumber)
      ElseIf Err.Number < 1 Then
         sText = SMessage
      Else
         sText = SMessage & Chr(10) & "(" & Error(ErrorNumber) & ")"
      End If
      
      If sLastMessage = sText Then Exit Sub
      sLastMessage = sText

      sText = sText & Chr(10) & Chr(10) & "Desplegar mas informacion sobre este error?"

      MakeSound WAVE_ERROR, True

      Select Case MsgBox(sText, vbYesNoCancel + vbCritical + vbDefaultButton2, "Error", sHelpFile, vContext)
      Case vbNo
         sLastMessage = ""

      Case vbYes
         sLastMessage = ""

         If ErrorNumber < 0 Then
            sText = "Error # " & Str(ErrorNumber) & " fue generado por " & sSource & Chr(13) & "Error definido por el usuario"
         Else
            sText = "Error # " & Str(ErrorNumber) & " fue generado por " _
                    & sSource & Chr(13) & Error(ErrorNumber)
         End If
         If FileExist(sHelpFile) Then
            sText = sText & Chr(10) & Chr(10) & "Press F1 to view error explaination."
         End If
         MsgBox sText, vbInformation, "Error", sHelpFile, vContext
      End Select

   End If

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
Sub UpdateRecentFiles(sFileName As String)
   ' Write open filename to INI file. (just pops it on the top of the stack)
   WriteRecentFiles sFileName
   
   ' Update the list of the most recently opened files in the File menu control array.
   GetRecentFiles
End Sub

Sub GetRecentFiles()
   Dim i As Integer, nSlot As Integer
   Dim sValue As String

   On Error Resume Next

   nSlot = 1

   For i = 1 To RECENT_COUNT

      sValue = GetIniString(sIniFile, INI_RECENT_KEY, "RecentFile" & i, "")
      If Not EmptyString(sValue) Then           ' Has a value?
         If nSlot = 1 Then SetVisible frmMain.mnuRecentFile(0), False

         Load frmMain.mnuRecentFile(nSlot)
         frmMain.mnuRecentFile(nSlot).Caption = sValue
         SetVisible frmMain.mnuRecentFile(nSlot), True
         SetEnabled frmMain.mnuRecentFile(nSlot), True

         nSlot = nSlot + 1
      End If
   Next
End Sub

Sub WriteRecentFiles(sFileName As String)
   Dim i As Integer, j As Integer
   Dim aRecent() As String
   Dim sValue As String

   ReDim aRecent(1 To 1)
   aRecent(1) = sFileName                                ' Store given filename
   j = 2                                                 ' Next available slot

   ' Load all what's stored and put it in the "to be stored" array
   For i = 1 To RECENT_COUNT
      sValue = GetIniString(sIniFile, INI_RECENT_KEY, "RecentFile" & i, "")
      If Not EmptyString(sValue) Then                    ' Has a value?
         If UCase$(sValue) <> UCase$(sFileName) Then     ' Not equal to first one?
            ReDim Preserve aRecent(1 To j)
            aRecent(j) = sValue
            j = j + 1
            If j > RECENT_COUNT Then Exit For             ' Array full?
         End If
      End If
   Next i

   ' Storage array ready. Now go store it.
   j = UBound(aRecent)
   For i = 1 To j
      If Not (GetIniString(sIniFile, INI_RECENT_KEY, "RecentFile" & i, "") = "" And aRecent(i) = "") Then
         AddIniString sIniFile, INI_RECENT_KEY, "RecentFile" & i, aRecent(i)
      End If
   Next

End Sub
