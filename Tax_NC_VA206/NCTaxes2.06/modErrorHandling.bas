Attribute VB_Name = "modErrorHandling"
Option Explicit

Const cstrModuleName As String = "modErrorHandling"


'// Debug Mode...
Global Const gblnDebugMode As Boolean = False

'// Testing program before shipping but not debugging...
Global Const gblnTesting As Boolean = True

'// Stop on errors...
Global Const gblnErrorStop As Boolean = (gblnDebugMode Or gblnTesting)


Public Enum ErrorMessageReturnEnum
  emrResume = vbRetry
  emrResumeNext = vbIgnore
  emrExitProc = vbAbort
End Enum

Public Function ErrorMessage( _
                              ByVal lngErrorNumber As Long, _
                              ByVal strDescription As String, _
                              ByVal strSource As String, _
                              ByVal strModuleName As String, _
                              ByVal strProcName As String, _
                              Optional ByVal lngErl As Long = 0, _
                              Optional ByVal blnLogError As Boolean = True, _
                              Optional ByVal strOptionString As String, _
                              Optional ByVal lngButtonOptions As Long = VBA.vbAbortRetryIgnore + VBA.vbExclamation, _
                              Optional ByVal strAdditionalData As String) As ErrorMessageReturnEnum
  'CHANGED: 2001-01-04: Changed Logging default to True and added code to not log errors if in DebugMode or Testing.
  On Error Resume Next
  
  '===  Whenever possible, you should try to pass "Erl" which is a hidden function that returns the error line.
  '     This will allow for the maximum level of debugging possible.
  
  Dim lngReturn As ErrorMessageReturnEnum
  Dim strMsg As String
  
  strMsg = ""
  strMsg = strMsg & "The following error has occured:" & vbCrLf & vbCrLf
  strMsg = strMsg & "Error Number : " & lngErrorNumber & vbCrLf
  strMsg = strMsg & "Description  : " & strDescription & vbCrLf
  strMsg = strMsg & "Error Source : " & strSource & vbCrLf
  strMsg = strMsg & "Module Name  : " & strModuleName & vbCrLf
  strMsg = strMsg & "Procedure    : " & strProcName & vbCrLf
  If (lngErl > 0) Then strMsg = strMsg & "Error Line   : " & lngErl & vbCrLf
  
  If ((Len(strOptionString) = 0) And (lngButtonOptions = vbAbortRetryIgnore)) Then
    '---  Set the default set of options if no option string specified...
    strOptionString = ""
    strOptionString = strOptionString & "Which do you wish to do: " & vbCrLf
    strOptionString = strOptionString & "   Abort the current procedure, " & vbCrLf
    strOptionString = strOptionString & "   Retry the operation,  or " & vbCrLf
    strOptionString = strOptionString & "   Ignore the error and move on?"
  End If
  
  If (Len(strOptionString) > 0) Then strMsg = strMsg & vbCrLf & strOptionString
  
  If (Len(strAdditionalData) > 0) Then strMsg = strMsg & vbCrLf & vbCrLf & "Additional Data:" & vbCrLf & strAdditionalData
  
  lngReturn = MsgBox(strMsg, lngButtonOptions, "Error")
  If (gblnErrorStop) Then
    ClearInUsePRX
    Stop
  End If
  Select Case lngReturn
    Case vbRetry:
      lngReturn = emrResume
    Case vbIgnore:
      lngReturn = emrResumeNext
    Case vbAbort:
      lngReturn = emrExitProc
    Case Else:
      '---  Return the actual value...
  End Select
  
  '---  Log the error if necessary...
  If ((blnLogError) And (Not (gblnTesting Or gblnDebugMode))) Then
    Call ErrLog(strMsg, lngReturn)
  End If
  
  ErrorMessage = lngReturn
End Function

Private Sub ErrLog(ByVal strMsg As String, ByVal lngResponse As Long)
  On Error Resume Next
  
  Const cstrErrLogFilename As String = "_ErrLog.txt"
  Dim lngFH As Long
  Dim strResponse As String
  
  strResponse = ""
  Select Case lngResponse
    Case vbRetry:
      strResponse = "Resume (Retry)"
    Case vbIgnore:
      strResponse = "Resume Next (Ignore)"
    Case vbAbort:
      strResponse = "Exit_Proc (Abort)"
    Case Else:
      strResponse = "Other (unknown)"
  End Select
  
  lngFH = FreeFile()
  Open cstrErrLogFilename For Append As #lngFH
  Print #lngFH, Format(Now, "mm/dd/yyyy hh:nn:ss AMPM")
  Print #lngFH, strMsg
  Print #lngFH, strResponse
  Print #lngFH, String(40, "=")
  Close #lngFH
End Sub

Public Sub SubErrorTest()
  Const cstrProcName As String = ""
  'on error goto Proc_Error
  
  '--- Your code goes here...

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

Private Function FunctionErrorTest()
  Const cstrProcName As String = ""
  'on error goto Proc_Error
  
  '--- Your code goes here...
  

  '--- Your "Normal" exit stuff goes here...
  
Proc_Exit:
  '--- Cleanup code goes here...
  GoTo Proc_End
  
Proc_Error:
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, cstrModuleName, cstrProcName)
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

Private Property Get PropertyErrorTest()
  Const cstrProcName As String = ""
  'on error goto Proc_Error
  
  '--- Your code goes here...
  

  '--- Your "Normal" exit stuff goes here...
   
Proc_Exit:
  '--- Cleanup code goes here...
  GoTo Proc_End
   
Proc_Error:
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, cstrModuleName, cstrProcName)
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
End Property



