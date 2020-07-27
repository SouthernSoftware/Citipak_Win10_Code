Attribute VB_Name = "W2EnableCloseButton"
Option Explicit

Const cstrModuleName As String = "modEnableCloseButton"

Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

'---  Constants to Specify how the position of the menu item should be located...
Public Const MF_BYCOMMAND = &H0
Public Const MF_BYPOSITION = &H400&

'---  Constants that specify whether a menu item is Enabled, Grayed and Disabled, or just Disabled
Public Const MF_ENABLED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_DISABLED = &H2&

'---  Constants that specify the Command to be enabled or disabled...
Const SC_SIZE = &HF000&
Const SC_MOVE = &HF010&
Const SC_MINIMIZE = &HF020&
Const SC_MAXIMIZE = &HF030&
Const SC_NEXTWINDOW = &HF040&
Const SC_PREVWINDOW = &HF050&
Const SC_CLOSE = &HF060&
Const SC_VSCROLL = &HF070&
Const SC_HSCROLL = &HF080&
Const SC_MOUSEMENU = &HF090&
Const SC_KEYMENU = &HF100&
Const SC_ARRANGE = &HF110&
Const SC_RESTORE = &HF120&
Const SC_TASKLIST = &HF130&
Const SC_SCREENSAVE = &HF140&
Const SC_HOTKEY = &HF150&

Public Const SC_DEFAULT = &HF160&

Const MY_SET = SC_HOTKEY Or _
              SC_SCREENSAVE Or _
              SC_TASKLIST Or _
              SC_RESTORE Or _
              SC_ARRANGE Or _
              SC_KEYMENU Or _
              SC_MOUSEMENU Or _
              SC_HSCROLL Or _
              SC_VSCROLL Or _
              SC_CLOSE Or _
              SC_PREVWINDOW Or _
              SC_NEXTWINDOW Or _
              SC_MAXIMIZE Or _
              SC_MINIMIZE Or _
              SC_MOVE Or _
              SC_SIZE

Public Function EnableCloseButton(ByVal hwnd As Long, Optional ByVal blnEnable As Boolean = True) As Long
  Const cstrProcName As String = "EnableCloseButton"
  On Error GoTo Proc_Error
  
  '---  This routine will enable or disable the Close button for the Microsoft Access main window.
  '     (It is helpful to do this if your application behaves unexpectedly if access is closed while it is running.)

  Dim hMenu As Long
  Dim lngPos As Long
  Dim lngFlags As Long, lngSuccess As Long
  
'// For Access, do this... --> hwnd = hWndAccessApp()

  lngPos = SC_CLOSE
  hMenu = GetSystemMenu(hwnd, 0)
  If (blnEnable) Then
    lngFlags = MF_BYCOMMAND Or MF_ENABLED
  Else
    lngFlags = MF_BYCOMMAND Or MF_GRAYED
  End If
  lngSuccess = EnableMenuItem(hMenu, lngPos, lngFlags)
  
  '--- Your "Normal" exit stuff goes here...
  
Proc_Exit:
  '--- Cleanup code goes here...
  EnableCloseButton = lngSuccess
  GoTo Proc_End
  
Proc_Error:
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, cstrModuleName, cstrProcName, Erl)
    Case emrExitProc:
      Resume Proc_Exit
    Case emrResume:
      Resume
    Case emrResumeNext:
      Resume Next
    Case Else
      '--- Technically, this should never happen.
      Resume Proc_Exit
  End Select
  
Proc_End:
End Function

