Attribute VB_Name = "modFAMakeWindowTopMost"
Option Explicit

Const cstrModuleName As String = "modMakeWindowTopMost"

' SetWindowPos Flags
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOZORDER = &H4
Const SWP_NOREDRAW = &H8
Const SWP_NOACTIVATE = &H10
Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOCOPYBITS = &H100
Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

' SetWindowPos() hwndInsertAfter values
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long



Public Sub MakeWindowTopMost(ByVal hwnd As Long, Optional ByVal blnMakeTopMost As Boolean = True)
  Const cstrProcName As String = "MakeWindowTopMost"
  On Error GoTo Proc_Error
  
  '--- Your code goes here...
  
  Dim lngNewWindowPos As Long
  Dim lngFlags As Long
  Dim lngReturn As Long
  
  '// Set flags to not Size or Move the window, we only want to change TopMost property...
  lngFlags = SWP_NOMOVE + SWP_NOSIZE
  
  If (blnMakeTopMost) Then
    lngNewWindowPos = HWND_TOPMOST
  Else
    lngNewWindowPos = HWND_NOTOPMOST
  End If
  lngReturn = SetWindowPos(hwnd, lngNewWindowPos, 0, 0, 0, 0, lngFlags)

  '--- Your "Normal" exit stuff goes here...
  
Proc_Exit:
  '--- Cleanup code goes here...
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
End Sub

